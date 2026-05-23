//! Submodule of `models` — see models/mod.rs.

#![allow(unused_imports)]
use serde::{Deserialize, Serialize};
use std::collections::{HashMap, HashSet, VecDeque};
use super::*;
use crate::domain::parser::{FunctionPurity, formula_purity, FunctionRegistry, Parser};

/// A cell address that includes the sheet name. Used as the key type for
/// the workbook-level cross-sheet dependency graph.
pub type CrossSheetKey = (String, usize, usize);

/// Current on-disk schema version for `.tshts` workbooks. Bump this when a
/// schema change is NOT backwards-compatible (e.g. a new required field or
/// a semantic change to an existing field). Backwards-compatible additions
/// (new optional fields with `#[serde(default)]`) do not require a bump.
pub const WORKBOOK_SCHEMA_VERSION: u32 = 1;

fn default_workbook_version() -> u32 {
    // Pre-versioning files implicitly carry version 1.
    1
}

thread_local! {
    /// Thread-local pointer to the current `Workbook` for the duration of a
    /// workbook-driven recalc. The pointer is non-null only inside
    /// `with_recalc_context`; outside of that scope, sheet-driven recalcs
    /// see `None` and fall back to single-sheet evaluation. Stored as a
    /// raw pointer because the cascade reborrows the workbook mutably (via
    /// `set_cell` on its sheets) while the evaluator only needs immutable
    /// access — and lifetimes can't express "borrowed for the duration of
    /// this dynamic scope" across the workbook ↔ sheet boundary.
    static RECALC_WORKBOOK: std::cell::Cell<*const Workbook> = const { std::cell::Cell::new(std::ptr::null()) };
}

/// Run `f` with `wb` published in the thread-local recalc context. Nested
/// scopes are supported — the previous pointer is restored on exit.
///
/// Inside the closure, callers that ultimately drive `Spreadsheet::set_cell`
/// (which cascades through `recalculate_cell`) will have their evaluator
/// receive the workbook reference, so cross-sheet refs resolve. Outside the
/// closure (e.g. tests that drive `Spreadsheet` directly), the fallback to
/// single-sheet eval still works — cross-sheet refs become `#REF!` as
/// before.
pub fn with_recalc_context<R>(wb: &Workbook, f: impl FnOnce() -> R) -> R {
    // Drop guard restores the prior thread-local pointer even on panic.
    // Without this, a panicking closure would leave a dangling pointer
    // in the thread-local; the next eval on this thread would deref the
    // stale address (UB) before any normal restore could run.
    struct CtxGuard(*const Workbook);
    impl Drop for CtxGuard {
        fn drop(&mut self) {
            RECALC_WORKBOOK.with(|c| c.set(self.0));
        }
    }
    let raw = wb as *const Workbook;
    let prev = RECALC_WORKBOOK.with(|c| c.replace(raw));
    let _guard = CtxGuard(prev);
    f()
}

/// Read the thread-local workbook ref and pass it (as `Option<&Workbook>`)
/// to `f`. Used by `Spreadsheet::recalculate_cell` and `maybe_spill` to
/// enrich their evaluator with workbook context when one is in scope.
///
/// # Safety
///
/// The pointer is set by `with_recalc_context` for the lifetime of its
/// closure. We assume callers respect the discipline that any code reached
/// from inside that closure holds a borrow consistent with the original
/// `&Workbook`. The pointer is only dereferenced through this helper, so
/// any access is bounded by `f`'s scope; we never store the borrow.
pub(crate) fn with_workbook_context<R>(f: impl FnOnce(Option<&Workbook>) -> R) -> R {
    let raw = RECALC_WORKBOOK.with(|c| c.get());
    if raw.is_null() {
        f(None)
    } else {
        // SAFETY: pointer was just written by `with_recalc_context` and is
        // valid for the duration of its closure. We hold no other mutable
        // borrow that would alias.
        let wb: &Workbook = unsafe { &*raw };
        f(Some(wb))
    }
}

/// A workbook containing multiple spreadsheets (tabs).
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Workbook {
    /// On-disk schema version. Files without this field deserialize as
    /// version 1 via the serde default. `load_workbook` rejects versions
    /// it doesn't understand.
    #[serde(default = "default_workbook_version")]
    pub version: u32,
    /// The sheets in this workbook
    pub sheets: Vec<Spreadsheet>,
    /// Names for each sheet
    pub sheet_names: Vec<String>,
    /// Index of the currently active sheet
    pub active_sheet: usize,
    /// Named ranges: name -> cell reference string (e.g., "Revenue" -> "B2:B50")
    pub named_ranges: HashMap<String, String>,
    /// Workbook-level dep graph for CROSS-sheet references only. Same-sheet
    /// deps remain on each Spreadsheet. `dependents[P]` is the set of
    /// cells that reference `P` from a different sheet. Not serialized —
    /// rebuilt on load.
    #[serde(skip)]
    pub cross_sheet_dependents: HashMap<CrossSheetKey, HashSet<CrossSheetKey>>,
    /// Inverse of `cross_sheet_dependents`: `dependencies[X]` is the set of
    /// cells `X` references from other sheets. Used to clear stale deps when
    /// a formula changes.
    #[serde(skip)]
    pub cross_sheet_dependencies: HashMap<CrossSheetKey, HashSet<CrossSheetKey>>,
    /// Cells whose formula contains *any* sheet-qualified reference (cross-
    /// sheet OR self-qualified like `=Sheet1!A1` on Sheet1). Used as a
    /// fast-path check: if no formula anywhere uses a qualified ref, the
    /// per-write workbook clone in `*_on_active` can be skipped because no
    /// evaluator inside the recalc cascade will ever ask for workbook
    /// context. Maintained by `register_cross_sheet_deps`.
    #[serde(skip)]
    pub cells_with_qualified_refs: HashSet<CrossSheetKey>,
    /// Cells whose value may be stale. Populated by every mutation site
    /// (set/clear/structural-edit/sheet-rename/sheet-remove). Drained by
    /// the recalc executor. Not serialized — after a fresh load, on-disk
    /// values are authoritative; dirty starts empty and grows only from
    /// subsequent mutations.
    ///
    /// Foundation for the parallel-calc engine (PR 1+): the level
    /// scheduler walks this set instead of re-deriving "what changed
    /// since last recalc" on every call. Today the existing per-sheet
    /// `recalculate_dependents` cascades remain authoritative; the
    /// graph-driven recalc lands in PR 3.
    #[serde(skip)]
    pub dirty: HashSet<CrossSheetKey>,
    /// Stable per-sheet IDs, parallel to `sheets` and `sheet_names`.
    /// Allocated monotonically; never reused. A sheet's ID survives
    /// rename and shifts in `active_sheet`; removing a sheet drops its
    /// entry from this vec but the next allocation still increments
    /// past `next_sheet_id`. Used as the node-key prefix in
    /// [`WorkbookGraph`](dep_graph::WorkbookGraph).
    #[serde(default)]
    pub sheet_ids: Vec<SheetId>,
    /// Monotonic SheetId allocator. Bumped on every `add_sheet`; never
    /// decremented even when sheets are removed.
    #[serde(default)]
    pub next_sheet_id: u32,
    /// Workbook-level unified dependency graph. Bidirectional; nodes are
    /// `(SheetId, row, col)`. Subsumes both the per-`Spreadsheet`
    /// same-sheet graphs and the workbook's cross-sheet graph. Built
    /// lazily on demand via `build_dep_graph_from_scratch`; PR 3 will
    /// keep it incrementally maintained on writes.
    ///
    /// PR 1 keeps the legacy per-sheet `dependents`/`dependencies` and
    /// the `cross_sheet_*` maps as the runtime authority; this graph is
    /// validated via tests but not yet consulted by the recalc path.
    #[serde(skip)]
    pub graph: WorkbookGraph,
    /// Derived per-cell purity classification, keyed parallel to the
    /// graph. Populated by `build_dep_graph_from_scratch` so PR 4's
    /// executor can partition a level into parallel and serial cells
    /// without re-parsing formulas. Non-formula cells default to `Pure`
    /// and are stored implicitly (absence = `Pure`).
    #[serde(skip)]
    pub cell_purities: HashMap<NodeKey, FunctionPurity>,
    /// Resolved dynamic-dep targets per `VolatileStructural` cell.
    /// Populated by the executor after each VolatileStructural cell's
    /// evaluation: INDIRECT/OFFSET push their resolved cells via
    /// `parser::push_dynamic_target`, the executor drains them and
    /// stores here keyed by the structural cell. `recalc_via_graph`
    /// uses this to skip auto-seeding a structural cell whose targets
    /// are unrelated to the current dirty closure — without it, a
    /// workbook with 1000 OFFSET cells would re-evaluate all 1000 on
    /// every keystroke.
    ///
    /// Absence (no entry) = "never been evaluated" — those cells get
    /// auto-seeded unconditionally so we observe their targets on the
    /// first pass.
    #[serde(skip)]
    pub structural_targets: HashMap<NodeKey, HashSet<NodeKey>>,
}

impl Default for Workbook {
    fn default() -> Self {
        Self {
            version: WORKBOOK_SCHEMA_VERSION,
            sheets: vec![Spreadsheet::default()],
            sheet_names: vec!["Sheet1".to_string()],
            active_sheet: 0,
            named_ranges: HashMap::new(),
            cross_sheet_dependents: HashMap::new(),
            cross_sheet_dependencies: HashMap::new(),
            cells_with_qualified_refs: HashSet::new(),
            dirty: HashSet::new(),
            sheet_ids: vec![SheetId(0)],
            next_sheet_id: 1,
            graph: WorkbookGraph::new(),
            cell_purities: HashMap::new(),
            structural_targets: HashMap::new(),
        }
    }
}

impl Workbook {
    /// Gets a reference to the active sheet.
    pub fn current_sheet(&self) -> &Spreadsheet {
        &self.sheets[self.active_sheet]
    }

    /// Gets a mutable reference to the active sheet.
    pub fn current_sheet_mut(&mut self) -> &mut Spreadsheet {
        &mut self.sheets[self.active_sheet]
    }

    /// Adds a new empty sheet with the given name. Allocates a fresh
    /// [`SheetId`] from the monotonic counter so the new sheet has a
    /// stable identity for the duration of the workbook.
    pub fn add_sheet(&mut self, name: String) {
        self.sheets.push(Spreadsheet::default());
        self.sheet_names.push(name);
        let id = SheetId(self.next_sheet_id);
        self.next_sheet_id = self.next_sheet_id.saturating_add(1);
        self.sheet_ids.push(id);
    }

    /// Returns the [`SheetId`] at the given index, allocating a fresh one
    /// if the parallel `sheet_ids` vec is missing or stale (e.g. after
    /// loading a pre-PR-1 file). Bounded by the live sheet count.
    pub fn sheet_id_at(&mut self, idx: usize) -> Option<SheetId> {
        if idx >= self.sheets.len() {
            return None;
        }
        self.ensure_sheet_ids();
        self.sheet_ids.get(idx).copied()
    }

    /// Look up the sheet index for a given [`SheetId`]. Returns `None` if
    /// the ID is unknown (e.g. the sheet was removed).
    pub fn sheet_idx_of(&self, id: SheetId) -> Option<usize> {
        self.sheet_ids.iter().position(|&s| s == id)
    }

    /// Display name for a [`SheetId`]. Returns `None` if the sheet has
    /// been removed.
    #[allow(dead_code)]
    pub fn sheet_name_of(&self, id: SheetId) -> Option<&str> {
        self.sheet_idx_of(id)
            .and_then(|i| self.sheet_names.get(i).map(|s| s.as_str()))
    }

    /// Ensure `sheet_ids` is parallel to `sheets`. Fills in fresh IDs for
    /// sheets that don't have one yet — happens on load of files written
    /// before SheetId existed, where the serde `#[serde(default)]` falls
    /// back to an empty vec.
    fn ensure_sheet_ids(&mut self) {
        if self.sheet_ids.len() == self.sheets.len() {
            return;
        }
        // Special-case the all-empty legacy load: if sheet_ids is empty
        // AND next_sheet_id is at its serde default of 0, start fresh at
        // SheetId(0) to match the `Default::default()` convention. Without
        // this, the first synthesized ID would be 1 (one off, leaving 0
        // unused).
        if !self.sheet_ids.is_empty() {
            // Some IDs exist already (e.g. some sheets serialized, others
            // added after load). Bump next_sheet_id past the maximum to
            // avoid collisions with surviving IDs.
            let max_existing = self.sheet_ids.iter().map(|s| s.0).max().unwrap_or(0);
            if self.next_sheet_id <= max_existing {
                self.next_sheet_id = max_existing.saturating_add(1);
            }
        }
        while self.sheet_ids.len() < self.sheets.len() {
            let id = SheetId(self.next_sheet_id);
            self.next_sheet_id = self.next_sheet_id.saturating_add(1);
            self.sheet_ids.push(id);
        }
        // If somehow sheet_ids is longer than sheets (corrupt file), trim.
        self.sheet_ids.truncate(self.sheets.len());
    }

    /// Rebuild the workbook-level [`WorkbookGraph`] from scratch by
    /// walking every formula cell on every sheet and registering its
    /// dependencies. Also populates `cell_purities` so the executor can
    /// partition by purity without re-parsing. O(sum of formula
    /// lengths). PR 3+ will maintain both incrementally on writes.
    pub fn build_dep_graph_from_scratch(&mut self) {
        self.ensure_sheet_ids();
        self.graph.clear();
        self.cell_purities.clear();
        // Targets cache is rebuilt by the next eval pass; any held entries
        // would be referenced against new (sheet_id, row, col) keys.
        self.structural_targets.clear();
        let registry = FunctionRegistry::shared_builtin();
        let sheet_count = self.sheets.len();
        // We need an evaluator with workbook context to extract qualified
        // refs (cross-sheet). Snapshot the workbook so the evaluator can
        // borrow it immutably while we walk the sheets.
        let snapshot = self.clone_for_graph_build();
        for sheet_idx in 0..sheet_count {
            let sheet_id = self.sheet_ids[sheet_idx];
            // Collect (node, refs, formula) tuples first so we don't borrow self.
            let edges: Vec<(NodeKey, Vec<NodeKey>, String)> = {
                let sheet = &self.sheets[sheet_idx];
                let names = &self.named_ranges;
                let evaluator =
                    crate::domain::services::FormulaEvaluator::for_workbook(
                        &snapshot, sheet, names,
                    );
                sheet
                    .cells
                    .iter()
                    .filter_map(|(&(r, c), cd)| {
                        cd.formula.as_ref().map(|f| {
                            let qrefs = evaluator.extract_qualified_refs(f);
                            let resolved: Vec<NodeKey> = qrefs
                                .into_iter()
                                .filter_map(|(maybe_sheet, rr, cc)| {
                                    let owner_idx = match maybe_sheet {
                                        Some(name) => snapshot
                                            .sheet_names
                                            .iter()
                                            .position(|n| n.eq_ignore_ascii_case(&name))?,
                                        None => sheet_idx,
                                    };
                                    Some((
                                        snapshot.sheet_ids[owner_idx],
                                        rr,
                                        cc,
                                    ))
                                })
                                .collect();
                            ((sheet_id, r, c), resolved, f.clone())
                        })
                    })
                    .collect()
            };
            for (node, prereqs, formula) in edges {
                self.graph.link(node, prereqs);
                // Compute and cache the per-cell purity. Pure cells are
                // implicit (absent from the map) so the common case
                // doesn't pay a HashMap insert. Volatile/side-effecting
                // cells get an explicit entry.
                let strip_eq = formula.strip_prefix('=').unwrap_or(&formula);
                if let Ok(mut p) = Parser::new(strip_eq) {
                    if let Ok(ast) = p.parse() {
                        let pur = formula_purity(&ast, &registry);
                        if pur != FunctionPurity::Pure {
                            self.cell_purities.insert(node, pur);
                        }
                    }
                }
            }
        }
    }

    /// Look up a cell's cached purity. Returns `Pure` for non-formula
    /// cells and for any cell not present in `cell_purities` (the
    /// implicit-Pure default keeps the hot map small).
    pub fn cell_purity(&self, node: NodeKey) -> FunctionPurity {
        self.cell_purities
            .get(&node)
            .copied()
            .unwrap_or(FunctionPurity::Pure)
    }

    /// Cheap snapshot for graph building. The evaluator only reads cell
    /// values and named ranges, so a full deep-clone is enough; PR 3
    /// will avoid this entirely by passing fields by reference once the
    /// borrow shape can be unified.
    fn clone_for_graph_build(&self) -> Workbook {
        self.clone()
    }

    /// Convert a [`NodeKey`] to the equivalent [`CrossSheetKey`] used by
    /// the legacy cross-sheet graph and the dirty-set. The sheet name is
    /// returned as a fresh String; allocates on every call.
    #[allow(dead_code)]
    pub fn node_to_cross_sheet_key(&self, node: NodeKey) -> Option<CrossSheetKey> {
        self.sheet_idx_of(node.0)
            .and_then(|i| self.sheet_names.get(i).cloned())
            .map(|name| (name, node.1, node.2))
    }

    /// Convert a [`CrossSheetKey`] to the matching [`NodeKey`]. Returns
    /// `None` if the sheet name is unknown.
    #[allow(dead_code)]
    pub fn cross_sheet_key_to_node(&self, key: &CrossSheetKey) -> Option<NodeKey> {
        let idx = self
            .sheet_names
            .iter()
            .position(|n| n.eq_ignore_ascii_case(&key.0))?;
        Some((self.sheet_ids[idx], key.1, key.2))
    }

    /// Iterative recalc for a set of cyclic cells that span sheets.
    /// Loops up to `iter_max` passes; converges when every cell's
    /// numeric delta from the previous pass is below `iter_epsilon`.
    /// Returns `Ok(passes_taken)` on convergence, `Err(iter_max)` on
    /// non-convergence so callers can surface the result.
    ///
    /// Replaces the per-sheet `recalculate_dependents` fallback that
    /// the executors previously used for cyclic remainder — that path
    /// only saw one leg of a cross-sheet cycle and never converged.
    ///
    /// `iter_max` / `iter_epsilon` are read from the active sheet
    /// (matches the per-sheet behavior; if a workbook user wants
    /// different settings per cycle, they can adjust the active sheet
    /// before triggering recalc). PRNG-based volatile cells inside a
    /// cycle remain non-deterministic across iterations — they're a
    /// pathological case the user opts into by enabling iterative_calc.
    pub fn iterative_calc_cyclic(
        &mut self,
        cyclic: &[NodeKey],
    ) -> Result<usize, usize> {
        use crate::domain::services::FormulaEvaluator;

        if cyclic.is_empty() {
            return Ok(0);
        }

        // Pre-resolve (node, sheet_idx, formula) so the hot loop doesn't
        // re-walk sheet_idx_of every iteration.
        let mut work: Vec<(NodeKey, usize, String)> = Vec::with_capacity(cyclic.len());
        for &node in cyclic {
            let Some(sheet_idx) = self.sheet_idx_of(node.0) else {
                continue;
            };
            let Some(formula) = self.sheets[sheet_idx]
                .cells
                .get(&(node.1, node.2))
                .and_then(|cd| cd.formula.clone())
            else {
                continue;
            };
            work.push((node, sheet_idx, formula));
        }
        if work.is_empty() {
            return Ok(0);
        }

        // Settings: take the most-permissive iter_max and tightest
        // iter_epsilon across every sheet that owns a cyclic cell, so
        // a multi-sheet cycle's convergence isn't constrained by the
        // active sheet's (possibly conservative) settings. The legacy
        // per-sheet path used the owning sheet's settings; this is the
        // workbook-level analog.
        let involved_sheets: std::collections::HashSet<usize> =
            work.iter().map(|(_, idx, _)| *idx).collect();
        let iter_max = involved_sheets
            .iter()
            .map(|&i| self.sheets[i].iter_max)
            .max()
            .unwrap_or(100)
            .max(1); // Always allow at least one pass — iter_max=0 would
                     // otherwise leave cyclic cells with stale values
                     // and confuse callers about convergence.
        let iter_epsilon = involved_sheets
            .iter()
            .map(|&i| self.sheets[i].iter_epsilon)
            .fold(f64::INFINITY, f64::min);

        // Track non-numeric stability across two consecutive passes.
        // Pure flip-flop (A → B → A → B ...) would otherwise force
        // every cycle to iter_max with no signal of failure; we detect
        // "value stable for 2 consecutive passes" as convergence even
        // when only string-valued.
        let mut prev_string_values: HashMap<NodeKey, String> = HashMap::new();
        let mut string_stable_for_one_pass: HashSet<NodeKey> = HashSet::new();

        // Iterate. Each pass evaluates against a snapshot that includes
        // results from prior passes (Gauss-Seidel-like: each iteration
        // mutates the live workbook and the next iteration's snapshot
        // reflects it).
        for pass in 1..=iter_max {
            let mut max_delta: f64 = 0.0;
            let mut all_strings_stable = true;
            let snapshot = self.clone();
            let names = snapshot.named_ranges.clone();
            // Collected per pass so we can record outside the borrow of
            // self.sheets that the inner loop holds.
            let mut pass_targets: Vec<(NodeKey, Vec<crate::domain::parser::DynamicTarget>)> =
                Vec::new();
            for (node, sheet_idx, formula) in &work {
                // Drain any leakage from a prior cell's eval so the
                // targets we collect after evaluate_formula are this
                // cell's only. Matches the eval_one drain-before-eval
                // pattern used by the level executors.
                let _ = crate::domain::parser::take_dynamic_targets();
                let new_value = {
                    let snap_sheet = &snapshot.sheets[*sheet_idx];
                    let evaluator =
                        FormulaEvaluator::for_workbook(&snapshot, snap_sheet, &names);
                    evaluator.evaluate_formula(formula)
                };
                // Capture INDIRECT/OFFSET targets so the smart auto-seed
                // for cyclic structural cells has the same cache the
                // acyclic path enjoys.
                let targets = crate::domain::parser::take_dynamic_targets();
                if !targets.is_empty()
                    && self.cell_purity(*node)
                        == crate::domain::parser::FunctionPurity::VolatileStructural
                {
                    pass_targets.push((*node, targets));
                }
                let prev_value = self.sheets[*sheet_idx]
                    .cells
                    .get(&(node.1, node.2))
                    .map(|cd| cd.value.clone())
                    .unwrap_or_default();
                let prev_num: Option<f64> = prev_value.parse().ok();
                let new_num: Option<f64> = new_value.parse().ok();
                match (prev_num, new_num) {
                    (Some(a), Some(b)) => {
                        let delta = (a - b).abs();
                        if delta > max_delta {
                            max_delta = delta;
                        }
                    }
                    _ => {
                        // Non-numeric branch: convergence requires the
                        // value to remain identical for 2 consecutive
                        // passes. First pass after a change → marked
                        // "stable for one pass" but not converged.
                        // Second consecutive pass with the same value
                        // → converged for this cell.
                        if prev_value == new_value
                            && prev_string_values
                                .get(node)
                                .is_some_and(|p| p == &new_value)
                        {
                            // Stable across two passes; nothing to do.
                        } else if prev_value == new_value
                            || prev_string_values
                                .get(node)
                                .is_some_and(|p| p == &new_value)
                        {
                            // Stable across one pass — need another to
                            // confirm. Treat as still-unstable for now.
                            string_stable_for_one_pass.insert(*node);
                            all_strings_stable = false;
                        } else {
                            all_strings_stable = false;
                            string_stable_for_one_pass.remove(node);
                        }
                        prev_string_values.insert(*node, new_value.clone());
                    }
                }
                // Apply the new value with the same post-write
                // maintenance the acyclic path performs: clear CF
                // cache, sweep stale spill ghosts, re-spill array
                // formulas. Wrap in with_recalc_context so a spill
                // re-evaluation resolves cross-sheet refs.
                with_recalc_context(&snapshot, || {
                    let sheet = &mut self.sheets[*sheet_idx];
                    if let Some(cd) = sheet.cells.get_mut(&(node.1, node.2)) {
                        cd.value = new_value;
                    }
                    sheet.cf_cache.lock().unwrap().clear();
                    sheet.sweep_spill_ghosts_for(node.1, node.2);
                    sheet.maybe_spill(node.1, node.2);
                });
            }
            // Push captured structural targets to the workbook-level
            // cache. Done outside the per-cell loop because
            // record_structural_targets borrows &mut self while the
            // cell loop holds &self.sheets[*sheet_idx].
            for (node, targets) in pass_targets {
                self.record_structural_targets(node, &targets);
            }
            // Converged when numeric deltas are below epsilon AND no
            // non-numeric flip-flops remain.
            if max_delta < iter_epsilon && all_strings_stable {
                return Ok(pass);
            }
        }
        Err(iter_max)
    }

    /// Sequential recalc driven by the workbook-level graph and dirty set.
    /// Drains [`Workbook::dirty`], expands the closure via the graph,
    /// computes topological levels, and re-evaluates each cell in level
    /// order. Cyclic cells fall back to the per-sheet iterative-calc loop.
    ///
    /// This is the level-based scheduler from PR 1 of the parallel-calc
    /// plan; PR 3 will extract it into a trait, PR 4 will add a parallel
    /// implementation. Today it's exposed alongside the legacy recalc so
    /// callers can switch over while tests verify behavior parity.
    ///
    /// Behavior notes:
    /// - **OFFSET / INDIRECT** dependencies on value-derived targets
    ///   can't be statically extracted. These cells are tagged
    ///   `VolatileStructural` and auto-seeded into the dirty set on
    ///   every recalc, so changes to their value-derived targets
    ///   propagate through their static dependents within one pass.
    ///   This matches Excel's "always recompute volatile" semantics.
    /// - **Cross-sheet cycles** are handled by `iterative_calc_cyclic`
    ///   (workbook-level iterative loop). Uses the highest iter_max
    ///   and tightest iter_epsilon across sheets participating in the
    ///   cycle.
    /// - **Per-level workbook clone** is O(workbook size) per level.
    ///   A future PR can replace this with `Arc<WorkbookSnapshot>` for
    ///   deeper graphs where the clone dominates.
    /// Public wrapper for the recalc engine. Returns any pass-level
    /// error (e.g. iterative-calc non-convergence) so the App layer can
    /// surface it via status message. Individual cell errors flow
    /// through `Value::Error` and don't reach this signature.
    pub fn recalc_via_graph_result(
        &mut self,
    ) -> Result<(), crate::domain::services::CalcError> {
        self.recalc_via_graph_inner()
    }

    /// Original entry-point — kept for backward compatibility with
    /// internal call sites that don't care about pass-level errors
    /// (e.g. tests that just want behavioral parity). Surfaces errors
    /// to stderr only.
    pub fn recalc_via_graph(&mut self) {
        if let Err(e) = self.recalc_via_graph_inner() {
            eprintln!("tshts: recalc: {}", e);
        }
    }

    fn recalc_via_graph_inner(
        &mut self,
    ) -> Result<(), crate::domain::services::CalcError> {
        // Lazy build: if the graph is empty, rebuild from scratch. PR 3+
        // will keep it incrementally maintained on writes.
        if self.graph.is_empty() {
            self.build_dep_graph_from_scratch();
        }
        if self.dirty.is_empty() && self.cell_purities.is_empty() {
            return Ok(());
        }

        // Map dirty cross-sheet keys to graph node keys.
        let dirty_keys = self.drain_dirty();
        let mut seeds: HashSet<NodeKey> = HashSet::with_capacity(dirty_keys.len());
        for k in dirty_keys {
            if let Some(node) = self.cross_sheet_key_to_node(&k) {
                seeds.insert(node);
            }
        }

        // Smart auto-seed for VolatileStructural cells (OFFSET, INDIRECT, ...).
        // Their dep edges are derived conservatively from literal args, so
        // they don't pick up changes to their value-derived targets via the
        // static graph. We close that gap by re-seeding them when their
        // last-known dynamic targets intersect what the user just dirtied.
        //
        // Policy:
        //   * If we have NO recorded targets for the cell yet (first-time
        //     eval, or just-rebuilt graph), seed it — we must run it once
        //     to learn its targets.
        //   * Otherwise, seed only when at least one recorded target lies
        //     inside the transitive dependents of the user-dirty seeds.
        //     (We use the closure of dependents, not just the seeds, so
        //     a chain `target → ... → user-edited cell` still triggers.)
        //
        // VolatileClock/Random are handled separately by RecalcContext +
        // the thread-local PRNG and don't appear here.
        let user_dirty_closure: HashSet<NodeKey> = if seeds.is_empty() {
            HashSet::new()
        } else {
            self.graph.transitive_dependents(&seeds)
        };
        for (&node, &purity) in &self.cell_purities {
            if purity != crate::domain::parser::FunctionPurity::VolatileStructural {
                continue;
            }
            match self.structural_targets.get(&node) {
                None => {
                    seeds.insert(node);
                }
                Some(targets) => {
                    if targets.iter().any(|t| user_dirty_closure.contains(t)) {
                        seeds.insert(node);
                    }
                }
            }
        }

        if seeds.is_empty() {
            return Ok(());
        }

        let topo = self.graph.topo_levels_from_seeds(&seeds);
        let plan = crate::domain::services::RecalcPlan {
            levels: topo.levels,
            cyclic: topo.cyclic,
        };
        let mut ctx = crate::domain::services::RecalcContext::new();
        use crate::domain::services::RecalcExecutor;
        // Auto-tune: pick Parallel when the workload is large enough to
        // amortize rayon's dispatch overhead. The threshold reflects the
        // bench numbers from `benches/calc_engine.rs` — below ~512 cells
        // Sequential wins on every archetype because the per-level
        // workbook clone dominates the parallel savings. Above that, the
        // crossover depends on archetype shape; users can override via
        // the `TSHTS_PAR_THRESHOLD` environment variable.
        let threshold: usize = std::env::var("TSHTS_PAR_THRESHOLD")
            .ok()
            .and_then(|s| s.parse().ok())
            .unwrap_or(512);
        let max_level_size = plan.levels.iter().map(|l| l.len()).max().unwrap_or(0);
        if max_level_size >= threshold {
            crate::domain::services::ParallelExecutor::new().run(&plan, &mut ctx, self)
        } else {
            crate::domain::services::SequentialExecutor.run(&plan, &mut ctx, self)
        }
    }

    /// Structural row insert on the active sheet, with cross-sheet ref
    /// adjustment on every other sheet. If Sheet2 has `=Sheet1!A5` and we
    /// insert a row above A5 in Sheet1, Sheet2's ref shifts to `=Sheet1!A6`.
    /// Matches Excel's structural-edit semantics.
    pub fn insert_row_on_active(&mut self, at: usize) {
        let active_idx = self.active_sheet;
        self.sheets[active_idx].insert_row(at);
        self.adjust_other_sheets_for_structural(active_idx, |evaluator, formula, target| {
            evaluator.adjust_formula_for_sheet_row_insert(formula, target, at)
        });
        // Dirty AFTER the mutation so keys reflect post-shift coordinates.
        // Marking pre-shift would leave dirty pointing at empty cells while
        // the actually-shifted cells go unrecorded.
        self.mark_all_formula_cells_dirty();
    }

    /// Structural row delete on the active sheet, with cross-sheet ref
    /// adjustment. Refs to the deleted row on other sheets become `#REF!`.
    pub fn delete_row_on_active(&mut self, at: usize) {
        let active_idx = self.active_sheet;
        self.sheets[active_idx].delete_row(at);
        self.adjust_other_sheets_for_structural(active_idx, |evaluator, formula, target| {
            evaluator.adjust_formula_for_sheet_row_delete(formula, target, at)
        });
        self.mark_all_formula_cells_dirty();
    }

    pub fn insert_col_on_active(&mut self, at: usize) {
        let active_idx = self.active_sheet;
        self.sheets[active_idx].insert_col(at);
        self.adjust_other_sheets_for_structural(active_idx, |evaluator, formula, target| {
            evaluator.adjust_formula_for_sheet_col_insert(formula, target, at)
        });
        self.mark_all_formula_cells_dirty();
    }

    pub fn delete_col_on_active(&mut self, at: usize) {
        let active_idx = self.active_sheet;
        self.sheets[active_idx].delete_col(at);
        self.adjust_other_sheets_for_structural(active_idx, |evaluator, formula, target| {
            evaluator.adjust_formula_for_sheet_col_delete(formula, target, at)
        });
        self.mark_all_formula_cells_dirty();
    }

    /// Mark every cell on every sheet that holds a formula as dirty.
    /// Used by structural edits and sheet rename/remove: those operations
    /// can shift refs anywhere, so the conservative-but-correct policy is
    /// to mark the whole formula population. The recalc executor (PR 1+)
    /// will compute the actual closure.
    pub(crate) fn mark_all_formula_cells_dirty(&mut self) {
        for (idx, sheet) in self.sheets.iter().enumerate() {
            let name = self.sheet_names[idx].clone();
            for (&(r, c), cd) in &sheet.cells {
                if cd.formula.is_some() {
                    self.dirty.insert((name.clone(), r, c));
                }
            }
        }
    }

    /// Walk every sheet OTHER than `mutated_idx` and apply `adjust` to each
    /// formula. `adjust` receives a fresh evaluator (bound to the sheet it's
    /// rewriting), the formula source, and the name of the sheet that was
    /// mutated. Used by structural ops to keep cross-sheet refs coherent.
    fn adjust_other_sheets_for_structural<F>(&mut self, mutated_idx: usize, adjust: F)
    where
        F: Fn(&crate::domain::services::FormulaEvaluator, &str, &str) -> String,
    {
        let mutated_name = self.sheet_names[mutated_idx].clone();
        // Walk EVERY sheet, including the mutated one. The mutated sheet's
        // unqualified refs were already shifted by `Spreadsheet::insert_row`
        // / `delete_row` etc.; the qualified-only `adjust` closure here is
        // safe to re-run on it because it only touches refs of the form
        // `<sheet>!<cell>` — leaving the unqualified ones intact. This also
        // catches self-qualified refs like `=Sheet1!A5` on Sheet1, which the
        // same-sheet adjustment misses.
        for (idx, sheet) in self.sheets.iter_mut().enumerate() {
            let updates: Vec<((usize, usize), String)> = {
                let evaluator = crate::domain::services::FormulaEvaluator::new(sheet);
                sheet
                    .cells
                    .iter()
                    .filter_map(|(&(r, c), cd)| {
                        cd.formula
                            .as_ref()
                            .map(|f| ((r, c), adjust(&evaluator, f, &mutated_name)))
                    })
                    .filter(|((r, c), new_f)| {
                        sheet
                            .cells
                            .get(&(*r, *c))
                            .and_then(|cd| cd.formula.as_ref())
                            .map(|orig| orig != new_f)
                            .unwrap_or(false)
                    })
                    .collect()
            };
            for ((r, c), new_formula) in updates {
                if let Some(cd) = sheet.cells.get_mut(&(r, c)) {
                    cd.formula = Some(new_formula);
                }
            }
            sheet.rebuild_dependencies();
            // Spill ranges may reference cells that shifted on the mutated
            // sheet; force a re-evaluation so ghosts stay coherent.
            // Skip the mutated sheet — `Spreadsheet::insert_row` already
            // resweeps it as part of its own bookkeeping.
            if idx != mutated_idx {
                sheet.resweep_all_spills();
            }
        }
        // Adjust workbook-level named-range values. `Revenue → Sheet2!A5:A10`
        // is a string of formula shape sans leading `=`; wrap it temporarily
        // so the same adjust closure works, then strip.
        let evaluator = crate::domain::services::FormulaEvaluator::new(&self.sheets[mutated_idx]);
        let name_updates: Vec<(String, String)> = self
            .named_ranges
            .iter()
            .filter_map(|(k, v)| {
                let wrapped = format!("={}", v);
                let new_wrapped = adjust(&evaluator, &wrapped, &mutated_name);
                if new_wrapped == wrapped {
                    return None;
                }
                let new_v = new_wrapped.strip_prefix('=').unwrap_or(&new_wrapped).to_string();
                Some((k.clone(), new_v))
            })
            .collect();
        for (k, v) in name_updates {
            self.named_ranges.insert(k.clone(), v.clone());
            // Mirror onto each sheet's per-sheet named_ranges cache.
            for sheet in &mut self.sheets {
                sheet.named_ranges.insert(k.clone(), v.clone());
            }
        }
        // Cross-sheet dep keys may reference shifted cells; rebuild.
        self.rebuild_cross_sheet_deps();
    }

    /// Removes a sheet by index. Adjusts active_sheet if needed.
    /// Won't remove the last sheet.
    pub fn remove_sheet(&mut self, index: usize) -> bool {
        if self.sheets.len() <= 1 || index >= self.sheets.len() {
            return false;
        }
        let removed_name = self.sheet_names[index].clone();
        // Purge the removed sheet's nodes from the unified graph so PR 3's
        // incremental maintenance doesn't leak stale entries across
        // remove/re-add cycles. The legacy cross_sheet_* maps are cleaned
        // up further down.
        if index < self.sheet_ids.len() {
            let dead_id = self.sheet_ids[index];
            let dead_nodes: Vec<NodeKey> = self.sheets[index]
                .cells
                .keys()
                .map(|&(r, c)| (dead_id, r, c))
                .collect();
            for node in dead_nodes {
                self.graph.forget_node(node);
                // Drop stale purity entries too — without this the
                // auto-seed loop in `recalc_via_graph` walks dead
                // (sheet_id, r, c) keys forever. Same for cached
                // dynamic targets: their NodeKeys point at a sheet
                // that no longer exists.
                self.cell_purities.remove(&node);
                self.structural_targets.remove(&node);
            }
        }
        // Every surviving formula cell that referenced the removed sheet
        // will be rewritten to #REF!, plus any cross-sheet dependent of
        // those formulas needs its value recomputed. Conservative dirty
        // mark across all formula cells handles both cases.
        self.mark_all_formula_cells_dirty();
        // Cells on the removed sheet itself: their dirty entries become
        // unreachable after removal. We pre-clean them here so the dirty
        // set never references a non-existent sheet.
        self.dirty.retain(|k| !k.0.eq_ignore_ascii_case(&removed_name));
        self.sheets.remove(index);
        self.sheet_names.remove(index);
        // Drop the parallel sheet_ids entry too. ID stays "dead" — never
        // reused even after the slot is removed.
        if index < self.sheet_ids.len() {
            self.sheet_ids.remove(index);
        }
        if self.active_sheet >= self.sheets.len() {
            self.active_sheet = self.sheets.len() - 1;
        } else if self.active_sheet > index {
            self.active_sheet -= 1;
        }
        // Rewrite any surviving formula that referenced the removed sheet
        // (e.g. `=Gone!A1`) to `=#REF!`. Excel does the same — leaving the
        // literal sheet name in the formula would otherwise produce an
        // "Unknown sheet" error instead of the Excel-standard #REF!.
        for sheet in &mut self.sheets {
            let updates: Vec<((usize, usize), String)> = sheet
                .cells
                .iter()
                .filter_map(|(&(r, c), cd)| {
                    cd.formula.as_ref().map(|f| {
                        ((r, c), replace_sheet_refs_with_ref_error(f, &removed_name))
                    })
                })
                .filter(|((r, c), new_f)| {
                    sheet
                        .cells
                        .get(&(*r, *c))
                        .and_then(|cd| cd.formula.as_ref())
                        .map(|orig| orig != new_f)
                        .unwrap_or(false)
                })
                .collect();
            let touched: Vec<(usize, usize)> = updates.iter().map(|(rc, _)| *rc).collect();
            for ((r, c), new_formula) in updates {
                if let Some(cd) = sheet.cells.get_mut(&(r, c)) {
                    cd.formula = Some(new_formula);
                }
            }
            sheet.rebuild_dependencies();
            // The rewrite changed only the formula text; the cached `value`
            // still shows whatever the formula previously evaluated to.
            // Force a recalc so the displayed value matches `#REF!` immediately.
            for (r, c) in touched {
                sheet.refresh_cell_value(r, c);
            }
        }
        // Purge any cross-sheet dep entries that touched the removed sheet.
        self.cross_sheet_dependents
            .retain(|k, _| !k.0.eq_ignore_ascii_case(&removed_name));
        for set in self.cross_sheet_dependents.values_mut() {
            set.retain(|k| !k.0.eq_ignore_ascii_case(&removed_name));
        }
        self.cross_sheet_dependencies
            .retain(|k, _| !k.0.eq_ignore_ascii_case(&removed_name));
        for set in self.cross_sheet_dependencies.values_mut() {
            set.retain(|k| !k.0.eq_ignore_ascii_case(&removed_name));
        }
        self.rebuild_cross_sheet_deps();
        true
    }

    /// Renames the active sheet.
    /// Rename the active sheet. Returns false (no-op) if `new_name` is empty
    /// or duplicates another sheet's name (case-insensitive). On success,
    /// rewrites formulas in every sheet AND named-range values that
    /// referenced the old name.
    pub fn rename_sheet(&mut self, new_name: String) -> bool {
        let old_name = self.sheet_names[self.active_sheet].clone();
        if old_name == new_name {
            return true;
        }
        // Reject empty names (unreferenceable in formulas).
        if new_name.trim().is_empty() {
            return false;
        }
        // Reject duplicates against any OTHER sheet (case-insensitive).
        if self
            .sheet_names
            .iter()
            .enumerate()
            .any(|(i, n)| i != self.active_sheet && n.eq_ignore_ascii_case(&new_name))
        {
            return false;
        }
        // Migrate dirty entries from old name → new name before the rename
        // lands. After the rename, lookups by sheet name would miss entries
        // still keyed under `old_name`.
        let migrated: Vec<CrossSheetKey> = self
            .dirty
            .iter()
            .filter(|k| k.0.eq_ignore_ascii_case(&old_name))
            .cloned()
            .collect();
        for k in migrated {
            self.dirty.remove(&k);
            self.dirty.insert((new_name.clone(), k.1, k.2));
        }
        self.sheet_names[self.active_sheet] = new_name.clone();
        // Mark all formula cells dirty AFTER the rename so the keys land
        // under the new sheet name (mark_all_formula_cells_dirty reads
        // sheet_names at insert time).
        self.mark_all_formula_cells_dirty();
        // Rewrite formulas in every sheet. Track which cells were touched per
        // sheet so we can propagate changes after the global rebuild — the
        // rewrite is value-neutral in the steady state (the formula points
        // at the same physical cell, just under a new name), but a sheet may
        // currently be holding an "Unknown sheet" error string from a
        // previous broken reference, and propagation forces a re-eval.
        let sheet_count = self.sheets.len();
        let mut touched_per_sheet: Vec<Vec<(usize, usize)>> = vec![Vec::new(); sheet_count];
        for (sheet_idx, sheet) in self.sheets.iter_mut().enumerate() {
            let updates: Vec<(usize, usize, String)> = sheet
                .cells
                .iter()
                .filter_map(|(&(r, c), cd)| {
                    cd.formula
                        .as_ref()
                        .map(|f| (r, c, rewrite_sheet_refs(f, &old_name, &new_name)))
                })
                .filter(|(r, c, new_formula)| {
                    sheet
                        .cells
                        .get(&(*r, *c))
                        .and_then(|cd| cd.formula.as_ref())
                        .map(|old| old != new_formula)
                        .unwrap_or(false)
                })
                .collect();
            for (r, c, formula) in updates {
                if let Some(cd) = sheet.cells.get_mut(&(r, c)) {
                    cd.formula = Some(formula);
                    touched_per_sheet[sheet_idx].push((r, c));
                }
            }
            sheet.rebuild_dependencies();
        }
        // Update named-range values too — they often contain sheet-qualified
        // ranges (`Sheet1!A1:B10`) that need the rename.
        let updated: Vec<(String, String)> = self
            .named_ranges
            .iter()
            .map(|(k, v)| (k.clone(), rewrite_sheet_refs_for_name_value(v, &old_name, &new_name)))
            .filter(|(k, v)| self.named_ranges.get(k).map(|orig| orig != v).unwrap_or(false))
            .collect();
        for (k, v) in updated {
            self.named_ranges.insert(k.clone(), v.clone());
            for sheet in &mut self.sheets {
                sheet.named_ranges.insert(k.clone(), v.clone());
            }
        }
        // Cross-sheet dep keys reference sheet names by string; rebuild.
        self.rebuild_cross_sheet_deps();
        // Propagate from every cell whose formula was rewritten so any
        // dependents that previously errored on the old name recompute.
        // Batched per-sheet so we pay one workbook clone per sheet
        // rather than one per touched cell.
        for (sheet_idx, touched) in touched_per_sheet.iter().enumerate() {
            if touched.is_empty() {
                continue;
            }
            let sheet_name = self.sheet_names[sheet_idx].clone();
            self.propagate_cross_sheet_changes_batch(&sheet_name, touched);
        }
        true
    }

    /// Define or replace a named range. Keys are normalized to uppercase so
    /// formulas can reference them case-insensitively.
    pub fn set_name(&mut self, name: &str, value: &str) {
        let key = name.to_uppercase();
        self.named_ranges.insert(key.clone(), value.to_string());
        for sheet in &mut self.sheets {
            sheet.named_ranges.insert(key.clone(), value.to_string());
            sheet.rebuild_dependencies();
        }
        // Any formula referencing the name now resolves to a (potentially
        // different) value or range; conservative dirty mark covers it.
        self.mark_all_formula_cells_dirty();
    }

    /// Remove a named range. Returns true if it existed.
    pub fn remove_name(&mut self, name: &str) -> bool {
        let key = name.to_uppercase();
        let existed = self.named_ranges.remove(&key).is_some();
        for sheet in &mut self.sheets {
            sheet.named_ranges.remove(&key);
            sheet.rebuild_dependencies();
        }
        if existed {
            // Formulas referencing the removed name now produce errors;
            // mark them all so the next recalc picks up the change.
            self.mark_all_formula_cells_dirty();
        }
        existed
    }

    /// Run propagate_cross_sheet_changes for a single cell on the active
    /// sheet. Assumes the caller already registered cross-sheet deps
    /// (via `set_cell_on_active` / `clear_cell_on_active` / direct
    /// `register_cross_sheet_deps`). Use from mutation paths that need to
    /// fan-out value changes to dependents without re-registering.
    ///
    /// Re-registering here would clear and rebuild the cell's
    /// structural_targets cache twice per write, defeating the smart
    /// auto-seed perf win for OFFSET/INDIRECT cells.
    pub fn propagate_active_cell(&mut self, row: usize, col: usize) {
        let sheet_name = self.sheet_names[self.active_sheet].clone();
        self.propagate_cross_sheet_changes(&sheet_name, row, col);
    }

    /// Fast-path predicate: do any cell formulas in the workbook contain a
    /// sheet-qualified reference? When false, the workbook clone in the
    /// `*_on_active` mutators can be skipped because no evaluator inside
    /// the same-sheet recalc cascade will ever consult workbook context.
    fn needs_workbook_context(&self) -> bool {
        !self.cells_with_qualified_refs.is_empty()
    }

    /// Mark a single cell as dirty. The cell's value may be stale and must
    /// be recomputed on the next recalc. Future readers (the parallel calc
    /// executor in PR 1+) walk this set to compute the recalc closure.
    pub fn mark_dirty(&mut self, sheet_name: &str, row: usize, col: usize) {
        self.dirty.insert((sheet_name.to_string(), row, col));
    }

    /// Mark a cell on the currently-active sheet as dirty.
    pub fn mark_dirty_active(&mut self, row: usize, col: usize) {
        let key = (self.sheet_names[self.active_sheet].clone(), row, col);
        self.dirty.insert(key);
    }

    /// Mark every cell on a given sheet as dirty. Reserved for callers
    /// that need maximal coverage (e.g. wholesale sheet replacement like
    /// CSV import) where even non-formula cells should participate in
    /// the next recalc closure. Sheet-name lookup is case-insensitive;
    /// inserted keys use the canonical-cased name from `sheet_names`.
    pub fn mark_sheet_dirty(&mut self, sheet_name: &str) {
        let Some(idx) = self
            .sheet_names
            .iter()
            .position(|n| n.eq_ignore_ascii_case(sheet_name))
        else {
            return;
        };
        let canonical = self.sheet_names[idx].clone();
        for &(r, c) in self.sheets[idx].cells.keys() {
            self.dirty.insert((canonical.clone(), r, c));
        }
    }

    /// Take the current dirty set, leaving the workbook's `dirty` empty.
    /// The recalc executor (PR 1+) calls this at the start of a recalc
    /// pass; mutation sites repopulate it between recalcs.
    pub fn drain_dirty(&mut self) -> HashSet<CrossSheetKey> {
        std::mem::take(&mut self.dirty)
    }

    /// True if no cells are dirty (i.e. a recalc would be a no-op).
    #[allow(dead_code)]
    pub fn dirty_is_empty(&self) -> bool {
        self.dirty.is_empty()
    }

    /// Write a single cell on the active sheet with workbook context
    /// published for the same-sheet recalc cascade. Use this instead of
    /// `current_sheet_mut().set_cell(...)` when downstream cells on the same
    /// sheet may contain cross-sheet refs that need to resolve correctly.
    pub fn set_cell_on_active(&mut self, row: usize, col: usize, data: CellData) {
        // Suppress no-op writes so the dirty set doesn't accumulate entries
        // for unchanged cells (PR 1's executor would otherwise re-eval
        // them needlessly).
        let unchanged = self.sheets[self.active_sheet]
            .cells
            .get(&(row, col))
            .map(|existing| existing == &data)
            .unwrap_or(false);
        if unchanged {
            return;
        }
        self.mark_dirty_active(row, col);
        if self.needs_workbook_context() {
            let snapshot = self.clone();
            with_recalc_context(&snapshot, || {
                self.sheets[self.active_sheet].set_cell(row, col, data);
            });
        } else {
            self.sheets[self.active_sheet].set_cell(row, col, data);
        }
        // Keep the unified graph + purity cache in sync with the new
        // formula. Without this, an interactive edit from `=A1+1` to
        // `=INDIRECT(A1)` would leave the cell classified as Pure and
        // miss the VolatileStructural auto-seed in recalc_via_graph.
        let sheet_name = self.sheet_names[self.active_sheet].clone();
        self.register_cross_sheet_deps(&sheet_name, row, col);
    }

    /// Clear a single cell on the active sheet with workbook context
    /// published for the same-sheet recalc cascade.
    pub fn clear_cell_on_active(&mut self, row: usize, col: usize) {
        // No-op suppression: clearing an already-absent cell shouldn't
        // dirty the set.
        if !self.sheets[self.active_sheet].cells.contains_key(&(row, col)) {
            return;
        }
        self.mark_dirty_active(row, col);
        if self.needs_workbook_context() {
            let snapshot = self.clone();
            with_recalc_context(&snapshot, || {
                self.sheets[self.active_sheet].clear_cell(row, col);
            });
        } else {
            self.sheets[self.active_sheet].clear_cell(row, col);
        }
        // Drop the cell's graph + purity entries; an Indirect-after-
        // clear would otherwise auto-seed on the next recalc and try
        // to evaluate a now-empty cell.
        let sheet_name = self.sheet_names[self.active_sheet].clone();
        self.forget_cell_in_graph(&sheet_name, row, col);
    }

    /// Write a batch of cells to the active sheet, then propagate.
    ///
    /// Replaces the previous "call set_many then loop calling
    /// propagate_cell_change at every site" discipline with a single API.
    /// This is the only mutation path that callers outside of undo/redo
    /// should use; using `current_sheet_mut().set_cell` directly bypasses
    /// cross-sheet propagation and is reserved for the workbook's own
    /// internals (load paths, recalc, undo/redo apply).
    pub fn write_cells_on_active(
        &mut self,
        writes: Vec<(usize, usize, CellData)>,
    ) {
        if writes.is_empty() {
            return;
        }
        let positions: Vec<(usize, usize)> =
            writes.iter().map(|(r, c, _)| (*r, *c)).collect();
        let sheet_name = self.sheet_names[self.active_sheet].clone();
        for &(r, c) in &positions {
            self.dirty.insert((sheet_name.clone(), r, c));
        }
        // Snapshot for cross-sheet ref resolution during this sheet's
        // recalc cascade. Without this, formulas like `=B1 + Sheet2!A1`
        // re-evaluated as part of the dependent recalc lose workbook
        // context and resolve the Sheet2!A1 ref to `#REF!`. Skip when no
        // qualified ref exists anywhere — common in single-sheet workbooks.
        if self.needs_workbook_context() {
            let snapshot = self.clone();
            with_recalc_context(&snapshot, || {
                self.sheets[self.active_sheet].set_many(writes);
            });
        } else {
            self.sheets[self.active_sheet].set_many(writes);
        }
        // Cross-sheet propagation in one shot — single workbook clone +
        // single BFS over the union closure, rather than N per-cell
        // clones (which was O(N × workbook_size)).
        self.register_cross_sheet_deps_batch(&sheet_name, &positions);
        self.propagate_cross_sheet_changes_batch(&sheet_name, &positions);
    }

    /// Clear a batch of cells on the active sheet, then propagate.
    /// Symmetric counterpart to `write_cells_on_active`.
    pub fn clear_cells_on_active(&mut self, positions: Vec<(usize, usize)>) {
        if positions.is_empty() {
            return;
        }
        let sheet_name = self.sheet_names[self.active_sheet].clone();
        for &(r, c) in &positions {
            self.dirty.insert((sheet_name.clone(), r, c));
        }
        // Bulk clear via clear_many — one topological cascade for the
        // entire union of dependents rather than N per-cell cascades.
        if self.needs_workbook_context() {
            let snapshot = self.clone();
            with_recalc_context(&snapshot, || {
                self.sheets[self.active_sheet].clear_many(positions.clone());
            });
        } else {
            self.sheets[self.active_sheet].clear_many(positions.clone());
        }
        // Cross-sheet propagation: one batch call for the whole set
        // rather than N per-cell BFS + workbook clones.
        self.register_cross_sheet_deps_batch(&sheet_name, &positions);
        self.propagate_cross_sheet_changes_batch(&sheet_name, &positions);
    }

    /// Creates a Workbook from a single Spreadsheet (for backward compatibility).
    pub fn from_spreadsheet(sheet: Spreadsheet) -> Self {
        Self {
            version: WORKBOOK_SCHEMA_VERSION,
            sheets: vec![sheet],
            sheet_names: vec!["Sheet1".to_string()],
            active_sheet: 0,
            named_ranges: HashMap::new(),
            cross_sheet_dependents: HashMap::new(),
            cross_sheet_dependencies: HashMap::new(),
            cells_with_qualified_refs: HashSet::new(),
            dirty: HashSet::new(),
            sheet_ids: vec![SheetId(0)],
            next_sheet_id: 1,
            graph: WorkbookGraph::new(),
            cell_purities: HashMap::new(),
            structural_targets: HashMap::new(),
        }
    }

    /// Re-register the cross-sheet dependencies for the cell at
    /// `(sheet_name, row, col)`. Called after every cell write. Removes
    /// stale entries from the old formula and inserts new ones from the
    /// current one.
    pub fn register_cross_sheet_deps(&mut self, sheet_name: &str, row: usize, col: usize) {
        use crate::domain::services::FormulaEvaluator;
        let key: CrossSheetKey = (sheet_name.to_string(), row, col);

        // Step 1: clear old reverse links.
        if let Some(old_precs) = self.cross_sheet_dependencies.remove(&key) {
            for p in old_precs {
                if let Some(set) = self.cross_sheet_dependents.get_mut(&p) {
                    set.remove(&key);
                    if set.is_empty() {
                        self.cross_sheet_dependents.remove(&p);
                    }
                }
            }
        }
        // Clear the qualified-ref membership; re-added in Step 4 if the
        // current formula still has any qualified refs.
        self.cells_with_qualified_refs.remove(&key);

        // Step 2: pull the current formula. Sheet names are case-insensitive
        // (Excel convention); using `==` here silently dropped cross-sheet
        // dep registration when callers passed a different casing.
        let sheet_idx = match self
            .sheet_names
            .iter()
            .position(|n| n.eq_ignore_ascii_case(sheet_name))
        {
            Some(i) => i,
            None => return,
        };
        let formula = match self.sheets[sheet_idx]
            .cells
            .get(&(row, col))
            .and_then(|cd| cd.formula.clone())
        {
            Some(f) => f,
            None => {
                // Formula → literal (or cell cleared). Drop unified-graph
                // outgoing edges, cached purity, and recorded dynamic
                // targets so they don't pin the auto-seed loop on a node
                // whose formula no longer exists. Step 1 already cleared
                // cross_sheet_dependencies; this completes the cleanup
                // on the unified-graph side.
                if let Some(node) = self.cross_sheet_key_to_node(&key) {
                    self.graph.unlink_node(node);
                    self.cell_purities.remove(&node);
                    self.structural_targets.remove(&node);
                }
                return;
            }
        };

        // Step 3: extract qualified refs from the formula. Use a snapshot
        // so the evaluator can borrow the workbook immutably.
        let qualified_refs: Vec<(Option<String>, usize, usize)> = {
            let names = self.named_ranges.clone();
            let evaluator = FormulaEvaluator::for_workbook(
                self,
                &self.sheets[sheet_idx],
                &names,
            );
            evaluator.extract_qualified_refs(&formula)
        };

        // Track "this cell has at least one qualified ref" (cross OR self-
        // qualified). Used by the fast path in `*_on_active` to skip the
        // workbook clone when no formula in the book uses a qualified ref.
        if qualified_refs.iter().any(|(s, _, _)| s.is_some()) {
            self.cells_with_qualified_refs.insert(key.clone());
        }

        // Step 4: register the cross-sheet ones (skip refs back to the same
        // sheet — those already live in the per-sheet dep graph).
        for (ref_sheet, ref_row, ref_col) in qualified_refs {
            // Skip if no explicit sheet or if it points to the same sheet.
            let resolved_sheet = match ref_sheet {
                Some(s) if !s.eq_ignore_ascii_case(sheet_name) => s,
                _ => continue,
            };
            // Normalize to the canonical sheet-name casing in `sheet_names`.
            let canon = self
                .sheet_names
                .iter()
                .find(|n| n.eq_ignore_ascii_case(&resolved_sheet))
                .cloned()
                .unwrap_or(resolved_sheet);
            let prec: CrossSheetKey = (canon, ref_row, ref_col);
            self.cross_sheet_dependencies
                .entry(key.clone())
                .or_default()
                .insert(prec.clone());
            self.cross_sheet_dependents
                .entry(prec)
                .or_default()
                .insert(key.clone());
        }

        // Step 5: keep the unified workbook graph and per-cell purity
        // cache consistent with the formula change. Without this,
        // graph-driven recalc walks a stale graph after any
        // interactive edit. The legacy cross_sheet_* maps above only
        // cover cross-sheet edges; the unified graph covers both
        // same-sheet AND cross-sheet, and `cell_purities` is the
        // engine's only signal that this cell is volatile.
        self.update_unified_graph_and_purity(sheet_idx, key.clone(), &formula);
    }

    /// Re-derive the unified graph edges + purity classification for a
    /// single cell. Called from `register_cross_sheet_deps` so every
    /// interactive cell-mutation path keeps the graph and purity cache
    /// in sync. Replaces the cell's outgoing edges via `set_prereqs`
    /// (which clears the old set first), so a formula edit from
    /// `=INDIRECT(A1)` to `=A1+1` removes the structural-volatile
    /// classification and the dynamic edges.
    fn update_unified_graph_and_purity(
        &mut self,
        sheet_idx: usize,
        key: CrossSheetKey,
        formula: &str,
    ) {
        use crate::domain::services::FormulaEvaluator;
        // Translate the cross-sheet key into a unified graph node key.
        let Some(node) = self.cross_sheet_key_to_node(&key) else {
            return;
        };
        // Pull the workbook-context evaluator from a snapshot so we
        // can borrow `self` immutably while updating the graph.
        let snapshot = self.clone();
        let names = snapshot.named_ranges.clone();
        let evaluator = FormulaEvaluator::for_workbook(
            &snapshot,
            &snapshot.sheets[sheet_idx],
            &names,
        );
        let qrefs = evaluator.extract_qualified_refs(formula);
        let resolved: Vec<NodeKey> = qrefs
            .into_iter()
            .filter_map(|(maybe_sheet, rr, cc)| {
                let owner_idx = match maybe_sheet {
                    Some(name) => snapshot
                        .sheet_names
                        .iter()
                        .position(|n| n.eq_ignore_ascii_case(&name))?,
                    None => sheet_idx,
                };
                Some((snapshot.sheet_ids[owner_idx], rr, cc))
            })
            .collect();
        self.graph.set_prereqs(node, resolved);
        // Recompute purity. Pure cells are stored implicitly (absent
        // from the map) so we explicitly remove the entry when the
        // new formula is Pure — a Pure-after-VolatileStructural edit
        // would otherwise leave a stale Structural entry.
        let registry = crate::domain::parser::FunctionRegistry::shared_builtin();
        let strip_eq = formula.strip_prefix('=').unwrap_or(formula);
        let pur = match crate::domain::parser::Parser::new(strip_eq) {
            Ok(mut p) => p
                .parse()
                .map(|ast| crate::domain::parser::formula_purity(&ast, &registry))
                .unwrap_or(crate::domain::parser::FunctionPurity::Pure),
            Err(_) => crate::domain::parser::FunctionPurity::Pure,
        };
        if pur == crate::domain::parser::FunctionPurity::Pure {
            self.cell_purities.remove(&node);
        } else {
            self.cell_purities.insert(node, pur);
        }
        // Any previously-recorded dynamic targets are stale on formula
        // change. Clear them so the next eval relearns (a now-Structural
        // cell with no targets is force-seeded; a now-Pure cell needs the
        // entry gone so it never re-triggers).
        self.structural_targets.remove(&node);
    }

    /// Drop unified-graph + purity tracking for a cell whose formula
    /// was removed (or the cell itself was deleted). Called from
    /// `clear_cell_on_active` paths so a Pure-after-delete or
    /// Structural-after-delete entry doesn't haunt future recalcs.
    pub fn forget_cell_in_graph(&mut self, sheet_name: &str, row: usize, col: usize) {
        let key: CrossSheetKey = (sheet_name.to_string(), row, col);
        if let Some(node) = self.cross_sheet_key_to_node(&key) {
            self.graph.unlink_node(node);
            self.cell_purities.remove(&node);
            self.structural_targets.remove(&node);
        }
    }

    /// Record the cells INDIRECT/OFFSET resolved to during a
    /// VolatileStructural cell's evaluation. Called by the executor
    /// post-eval so the next recalc's auto-seed can skip this cell
    /// when none of its targets are in the dirty closure.
    pub fn record_structural_targets(
        &mut self,
        node: NodeKey,
        targets: &[crate::domain::parser::DynamicTarget],
    ) {
        let default_sheet_id = node.0;
        let resolved: HashSet<NodeKey> = targets
            .iter()
            .filter_map(|t| {
                let sid = match &t.sheet {
                    Some(name) => {
                        let idx = self
                            .sheet_names
                            .iter()
                            .position(|n| n.eq_ignore_ascii_case(name))?;
                        self.sheet_ids[idx]
                    }
                    None => default_sheet_id,
                };
                Some((sid, t.row, t.col))
            })
            .collect();
        if resolved.is_empty() {
            // Some VolatileStructural functions don't push targets
            // (e.g. an OFFSET that errored at #REF! before reaching the
            // target loop). Keep any prior cached targets — they're
            // the most recent valid observation.
            return;
        }
        self.structural_targets.insert(node, resolved);
    }

    /// Batch variant of `register_cross_sheet_deps`. Calling
    /// `register_cross_sheet_deps` in a loop is O(N × formula_size)
    /// per cell which is fine — the per-call work doesn't BFS. We
    /// expose this as a separate API just for symmetry with the batch
    /// propagation below.
    pub fn register_cross_sheet_deps_batch(
        &mut self,
        sheet_name: &str,
        positions: &[(usize, usize)],
    ) {
        for &(r, c) in positions {
            self.register_cross_sheet_deps(sheet_name, r, c);
        }
    }

    /// Batch variant of `propagate_cross_sheet_changes`. Clones the
    /// workbook ONCE and walks a single BFS over the union of seed
    /// cells. The per-cell variant clones per call — O(N × workbook_size)
    /// for an N-cell batch. The batch version is O(workbook_size +
    /// closure_size).
    pub fn propagate_cross_sheet_changes_batch(
        &mut self,
        sheet_name: &str,
        seeds: &[(usize, usize)],
    ) {
        use crate::domain::services::FormulaEvaluator;
        if seeds.is_empty() {
            return;
        }

        // Quick gate: skip the clone if none of the seed cells have any
        // cross-sheet dependents. The single-cell path had the same
        // optimization; here we widen it to "any seed has deps".
        let has_any_deps = seeds.iter().any(|&(r, c)| {
            self.cross_sheet_dependents
                .get(&(sheet_name.to_string(), r, c))
                .is_some_and(|s| !s.is_empty())
        });
        if !has_any_deps {
            return;
        }

        let mut queue: std::collections::VecDeque<CrossSheetKey> =
            std::collections::VecDeque::new();
        let mut visited: HashSet<CrossSheetKey> = HashSet::new();
        for &(r, c) in seeds {
            queue.push_back((sheet_name.to_string(), r, c));
        }

        let mut snapshot = self.clone();
        let names = snapshot.named_ranges.clone();

        while let Some(key) = queue.pop_front() {
            if !visited.insert(key.clone()) {
                continue;
            }
            let deps: Vec<CrossSheetKey> = self
                .cross_sheet_dependents
                .get(&key)
                .cloned()
                .unwrap_or_default()
                .into_iter()
                .collect();
            if deps.is_empty() {
                continue;
            }
            for dep in deps {
                let (dep_sheet, dep_row, dep_col) = dep.clone();
                let Some(dep_idx) = snapshot
                    .sheet_names
                    .iter()
                    .position(|n| n.eq_ignore_ascii_case(&dep_sheet))
                else {
                    continue;
                };
                let new_value = {
                    let snap_sheet = &snapshot.sheets[dep_idx];
                    let Some(cd) = snap_sheet.cells.get(&(dep_row, dep_col)) else {
                        continue;
                    };
                    let Some(formula) = cd.formula.clone() else {
                        continue;
                    };
                    let evaluator =
                        FormulaEvaluator::for_workbook(&snapshot, snap_sheet, &names);
                    evaluator.evaluate_formula(&formula)
                };
                if let Some(snap_cd) = snapshot.sheets[dep_idx]
                    .cells
                    .get_mut(&(dep_row, dep_col))
                {
                    snap_cd.value = new_value.clone();
                }
                let dep_real_idx = self
                    .sheet_names
                    .iter()
                    .position(|n| n.eq_ignore_ascii_case(&dep_sheet));
                if let Some(idx) = dep_real_idx {
                    let value_actually_changed = self.sheets[idx]
                        .cells
                        .get_mut(&(dep_row, dep_col))
                        .map(|real_cd| {
                            if real_cd.value != new_value {
                                real_cd.value = new_value.clone();
                                true
                            } else {
                                false
                            }
                        })
                        .unwrap_or(false);
                    if value_actually_changed {
                        // Same-sheet cascade on the destination sheet
                        // (same logic as the single-cell propagator).
                        with_recalc_context(&snapshot, || {
                            self.sheets[idx].recalculate_dependents(dep_row, dep_col);
                        });
                        snapshot.sheets[idx] = self.sheets[idx].clone();
                        queue.push_back(dep);
                    }
                }
            }
        }
    }

    /// Recalculate every cell that depends on `(sheet_name, row, col)` via
    /// the cross-sheet graph. Walks transitively (BFS) so chains like
    /// Sheet1!A1 → Sheet2!A1 → Sheet3!A1 all update.
    pub fn propagate_cross_sheet_changes(
        &mut self,
        sheet_name: &str,
        row: usize,
        col: usize,
    ) {
        use crate::domain::services::FormulaEvaluator;
        let mut queue: std::collections::VecDeque<CrossSheetKey> =
            std::collections::VecDeque::new();
        queue.push_back((sheet_name.to_string(), row, col));
        let mut visited: HashSet<CrossSheetKey> = HashSet::new();

        // Check up front whether any dep edges exist — common case has none,
        // and we want to skip the expensive workbook clone in that case.
        let has_any_deps = self
            .cross_sheet_dependents
            .get(&(sheet_name.to_string(), row, col))
            .is_some_and(|s| !s.is_empty());
        if !has_any_deps {
            return;
        }

        // One snapshot for the whole BFS, mutated in-place as we compute new
        // values so chained refs (Sheet1!A1 → Sheet2!A1 → Sheet3!A1) see the
        // freshly-recomputed upstream values on later layers.
        let mut snapshot = self.clone();
        let names = snapshot.named_ranges.clone();

        while let Some(key) = queue.pop_front() {
            if !visited.insert(key.clone()) {
                continue;
            }
            let deps: Vec<CrossSheetKey> = self
                .cross_sheet_dependents
                .get(&key)
                .cloned()
                .unwrap_or_default()
                .into_iter()
                .collect();
            if deps.is_empty() {
                continue;
            }
            for dep in deps {
                let (dep_sheet, dep_row, dep_col) = dep.clone();
                let Some(dep_idx) = snapshot
                    .sheet_names
                    .iter()
                    .position(|n| n.eq_ignore_ascii_case(&dep_sheet))
                else {
                    continue;
                };
                let new_value = {
                    let snap_sheet = &snapshot.sheets[dep_idx];
                    let Some(cd) = snap_sheet.cells.get(&(dep_row, dep_col)) else {
                        continue;
                    };
                    let Some(formula) = cd.formula.clone() else { continue };
                    let evaluator =
                        FormulaEvaluator::for_workbook(&snapshot, snap_sheet, &names);
                    evaluator.evaluate_formula(&formula)
                };
                // Update the snapshot first so downstream layers see it.
                if let Some(snap_cd) = snapshot.sheets[dep_idx]
                    .cells
                    .get_mut(&(dep_row, dep_col))
                {
                    snap_cd.value = new_value.clone();
                }
                // Write to the real workbook.
                let dep_real_idx = self
                    .sheet_names
                    .iter()
                    .position(|n| n.eq_ignore_ascii_case(&dep_sheet));
                if let Some(idx) = dep_real_idx {
                    let value_actually_changed = self.sheets[idx]
                        .cells
                        .get_mut(&(dep_row, dep_col))
                        .map(|real_cd| {
                            if real_cd.value != new_value {
                                real_cd.value = new_value.clone();
                                true
                            } else {
                                false
                            }
                        })
                        .unwrap_or(false);
                    if value_actually_changed {
                        // Same-sheet cascade: dependents on the destination
                        // sheet that referenced the cell we just wrote need
                        // to recompute too. The previous version skipped
                        // this — Sheet1!A1 → Sheet2!A1 → Sheet2!B1 left B1
                        // stale because the cross-sheet engine only walked
                        // cross-sheet edges. Publishing `snapshot` as the
                        // workbook context lets the evaluator inside the
                        // cascade resolve any further cross-sheet refs.
                        with_recalc_context(&snapshot, || {
                            self.sheets[idx].recalculate_dependents(dep_row, dep_col);
                        });
                        // Refresh the snapshot's view of this sheet so a
                        // later layer of the cross-sheet BFS sees the
                        // newly-cascaded same-sheet values.
                        snapshot.sheets[idx] = self.sheets[idx].clone();
                        queue.push_back(dep);
                    }
                }
            }
        }
    }

    /// Check whether adding a new formula at `(sheet_name, row, col)` with
    /// the given precedents would create a cross-sheet cycle. Walks the
    /// existing cross-sheet graph from each precedent; if any path reaches
    /// `(sheet_name, row, col)`, we'd loop. The same-sheet check still runs
    /// separately via `FormulaEvaluator::would_create_circular_reference`.
    pub fn would_create_cross_sheet_cycle(
        &self,
        sheet_name: &str,
        row: usize,
        col: usize,
        candidate_precedents: &[(Option<String>, usize, usize)],
    ) -> bool {
        let target: CrossSheetKey = (sheet_name.to_string(), row, col);
        let mut stack: Vec<CrossSheetKey> = Vec::new();
        for (prec_sheet, prec_row, prec_col) in candidate_precedents {
            // Only consider cross-sheet precedents (same-sheet cycles are
            // caught by the existing AST walker).
            let Some(ps) = prec_sheet else { continue };
            if ps.eq_ignore_ascii_case(sheet_name) {
                continue;
            }
            let canon = self
                .sheet_names
                .iter()
                .find(|n| n.eq_ignore_ascii_case(ps))
                .cloned()
                .unwrap_or_else(|| ps.clone());
            stack.push((canon, *prec_row, *prec_col));
        }
        let mut visited: HashSet<CrossSheetKey> = HashSet::new();
        while let Some(node) = stack.pop() {
            if node == target {
                return true;
            }
            if !visited.insert(node.clone()) {
                continue;
            }
            // Also walk down: the cells that *node* in turn depends on.
            if let Some(deps) = self.cross_sheet_dependencies.get(&node) {
                for d in deps {
                    stack.push(d.clone());
                }
            }
        }
        false
    }

    /// Rebuild the cross-sheet dep graph from scratch by scanning every
    /// formula in every sheet. Called after load (since the graph isn't
    /// serialized) and as a fallback when state drifts.
    pub fn rebuild_cross_sheet_deps(&mut self) {
        self.cross_sheet_dependents.clear();
        self.cross_sheet_dependencies.clear();
        // Entries from since-deleted formula cells (or formula→value
        // transitions before the per-cell cleanup landed) would otherwise
        // make needs_workbook_context() falsely true, forcing unnecessary
        // workbook clones in the *_on_active fast path.
        self.cells_with_qualified_refs.clear();
        // Wipe unified-graph state too. The per-cell `register_cross_sheet_deps`
        // below will repopulate from cells.iter(), but it only refreshes
        // entries for CURRENT (sheet_id, row, col) positions. Stale entries
        // from positions whose cells were shifted by a structural edit
        // (row/col insert/delete) would otherwise leak indefinitely —
        // their NodeKeys are unreachable from cells.iter() so per-cell
        // registration can't touch them, but `recalc_via_graph` would still
        // walk them via cell_purities/structural_targets.
        self.graph.clear();
        self.cell_purities.clear();
        self.structural_targets.clear();
        let cells: Vec<(String, usize, usize)> = self
            .sheet_names
            .iter()
            .enumerate()
            .flat_map(|(idx, name)| {
                self.sheets[idx]
                    .cells
                    .iter()
                    .filter(|(_, cd)| cd.formula.is_some())
                    .map(move |(&(r, c), _)| (name.clone(), r, c))
            })
            .collect();
        for (sheet, r, c) in cells {
            self.register_cross_sheet_deps(&sheet, r, c);
        }
    }
}
