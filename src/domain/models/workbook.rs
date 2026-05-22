//! Submodule of `models` — see models/mod.rs.

#![allow(unused_imports)]
use serde::{Deserialize, Serialize};
use std::collections::{HashMap, HashSet, VecDeque};
use super::*;

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
    let raw = wb as *const Workbook;
    let prev = RECALC_WORKBOOK.with(|c| c.replace(raw));
    let result = f();
    RECALC_WORKBOOK.with(|c| c.set(prev));
    result
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

    /// Adds a new empty sheet with the given name.
    pub fn add_sheet(&mut self, name: String) {
        self.sheets.push(Spreadsheet::default());
        self.sheet_names.push(name);
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
    }

    /// Structural row delete on the active sheet, with cross-sheet ref
    /// adjustment. Refs to the deleted row on other sheets become `#REF!`.
    pub fn delete_row_on_active(&mut self, at: usize) {
        let active_idx = self.active_sheet;
        self.sheets[active_idx].delete_row(at);
        self.adjust_other_sheets_for_structural(active_idx, |evaluator, formula, target| {
            evaluator.adjust_formula_for_sheet_row_delete(formula, target, at)
        });
    }

    pub fn insert_col_on_active(&mut self, at: usize) {
        let active_idx = self.active_sheet;
        self.sheets[active_idx].insert_col(at);
        self.adjust_other_sheets_for_structural(active_idx, |evaluator, formula, target| {
            evaluator.adjust_formula_for_sheet_col_insert(formula, target, at)
        });
    }

    pub fn delete_col_on_active(&mut self, at: usize) {
        let active_idx = self.active_sheet;
        self.sheets[active_idx].delete_col(at);
        self.adjust_other_sheets_for_structural(active_idx, |evaluator, formula, target| {
            evaluator.adjust_formula_for_sheet_col_delete(formula, target, at)
        });
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
        self.sheets.remove(index);
        self.sheet_names.remove(index);
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
        self.sheet_names[self.active_sheet] = new_name.clone();
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
        for (sheet_idx, touched) in touched_per_sheet.iter().enumerate() {
            let sheet_name = self.sheet_names[sheet_idx].clone();
            for &(r, c) in touched {
                self.propagate_cross_sheet_changes(&sheet_name, r, c);
            }
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
    }

    /// Remove a named range. Returns true if it existed.
    pub fn remove_name(&mut self, name: &str) -> bool {
        let key = name.to_uppercase();
        let existed = self.named_ranges.remove(&key).is_some();
        for sheet in &mut self.sheets {
            sheet.named_ranges.remove(&key);
            sheet.rebuild_dependencies();
        }
        existed
    }

    /// Run register_cross_sheet_deps + propagate_cross_sheet_changes for a
    /// single cell on the active sheet. Use from mutation paths that bypass
    /// `write_cells_on_active` / `clear_cells_on_active` (e.g. undo/redo
    /// apply that restores via `set_cell` directly).
    pub fn propagate_active_cell(&mut self, row: usize, col: usize) {
        let sheet_name = self.sheet_names[self.active_sheet].clone();
        self.register_cross_sheet_deps(&sheet_name, row, col);
        self.propagate_cross_sheet_changes(&sheet_name, row, col);
    }

    /// Fast-path predicate: do any cell formulas in the workbook contain a
    /// sheet-qualified reference? When false, the workbook clone in the
    /// `*_on_active` mutators can be skipped because no evaluator inside
    /// the same-sheet recalc cascade will ever consult workbook context.
    fn needs_workbook_context(&self) -> bool {
        !self.cells_with_qualified_refs.is_empty()
    }

    /// Write a single cell on the active sheet with workbook context
    /// published for the same-sheet recalc cascade. Use this instead of
    /// `current_sheet_mut().set_cell(...)` when downstream cells on the same
    /// sheet may contain cross-sheet refs that need to resolve correctly.
    pub fn set_cell_on_active(&mut self, row: usize, col: usize, data: CellData) {
        if self.needs_workbook_context() {
            let snapshot = self.clone();
            with_recalc_context(&snapshot, || {
                self.sheets[self.active_sheet].set_cell(row, col, data);
            });
        } else {
            self.sheets[self.active_sheet].set_cell(row, col, data);
        }
    }

    /// Clear a single cell on the active sheet with workbook context
    /// published for the same-sheet recalc cascade.
    pub fn clear_cell_on_active(&mut self, row: usize, col: usize) {
        if self.needs_workbook_context() {
            let snapshot = self.clone();
            with_recalc_context(&snapshot, || {
                self.sheets[self.active_sheet].clear_cell(row, col);
            });
        } else {
            self.sheets[self.active_sheet].clear_cell(row, col);
        }
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
        for (r, c) in positions {
            self.register_cross_sheet_deps(&sheet_name, r, c);
            self.propagate_cross_sheet_changes(&sheet_name, r, c);
        }
    }

    /// Clear a batch of cells on the active sheet, then propagate.
    /// Symmetric counterpart to `write_cells_on_active`.
    pub fn clear_cells_on_active(&mut self, positions: Vec<(usize, usize)>) {
        if positions.is_empty() {
            return;
        }
        let sheet_name = self.sheet_names[self.active_sheet].clone();
        if self.needs_workbook_context() {
            let snapshot = self.clone();
            with_recalc_context(&snapshot, || {
                for (r, c) in &positions {
                    self.sheets[self.active_sheet].clear_cell(*r, *c);
                }
            });
        } else {
            for (r, c) in &positions {
                self.sheets[self.active_sheet].clear_cell(*r, *c);
            }
        }
        for (r, c) in positions {
            self.register_cross_sheet_deps(&sheet_name, r, c);
            self.propagate_cross_sheet_changes(&sheet_name, r, c);
        }
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
            None => return,
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
