//! Recalc executor — the engine that walks topological levels and
//! re-evaluates dirty cells.
//!
//! Designed for swappability: PR 4 will add a `ParallelExecutor` that
//! implements the same trait via rayon. The `Sequential` impl here is
//! the reference behavior — every parallel variant must produce
//! identical observable results.
//!
//! Architecture (per research/index.md):
//! - The orchestrator (main thread) builds a [`RecalcPlan`] by querying
//!   [`Workbook`] for its dirty set + dep graph + topological levels.
//! - The orchestrator builds a [`RecalcContext`] holding snapshot
//!   values for volatile state (clock; PRNG is thread-local) so all
//!   cells in a single recalc pass see consistent volatile values.
//! - The executor runs `plan.levels` in order. Within each level, it
//!   re-evaluates each cell's formula and writes the new value back to
//!   the live workbook. Cyclic remainder cells fall through to the
//!   per-sheet iterative-calc loop.
//!
//! The trait is the seam at which PR 4 plugs in rayon. Sequential and
//! Parallel impls share `RecalcPlan`/`RecalcContext`/`CalcError` so the
//! orchestrator code at the call site doesn't change between PRs.


use crate::domain::models::{NodeKey, Workbook};

/// Per-recalc-pass snapshot of state that must be consistent across
/// the pass. The orchestrator constructs one of these at pass start and
/// hands a reference to every cell evaluation. PR 4's parallel executor
/// hands the same reference to each worker, so all workers see the same
/// clock value etc.
///
/// PR 2 introduces this struct; PR 4 will extend it with HTTP-cache
/// references and a cancellation flag.
#[derive(Debug)]
pub struct RecalcContext {
    /// Wall-clock value (Excel serial) captured at pass start. NOW/TODAY
    /// read this instead of `SystemTime::now()` so every clock-volatile
    /// cell in a single pass returns the same instant. Published to
    /// `parser::RECALC_CLOCK` by each executor's `run` method.
    pub clock_snapshot: f64,
}

impl RecalcContext {
    pub fn new() -> Self {
        Self {
            clock_snapshot: crate::domain::parser::now_serial(),
        }
    }
}

impl Default for RecalcContext {
    fn default() -> Self {
        Self::new()
    }
}

/// Topological order in which cells must be re-evaluated. Built from the
/// workbook's dep graph and dirty set.
///
/// `levels[i]` is independent — every cell in `levels[i]` can be evaluated
/// concurrently without coordinating with peers because no cell in
/// `levels[i]` depends on another cell in the same level. `levels[i+1]`
/// depends only on cells in `levels[0..=i]`.
///
/// `cyclic` holds cells that participate in (or are downstream of) a
/// cycle and could not be placed in any level — they fall back to the
/// iterative-calc engine.
#[derive(Debug, Default, Clone)]
pub struct RecalcPlan {
    pub levels: Vec<Vec<NodeKey>>,
    pub cyclic: Vec<NodeKey>,
}


/// Errors a recalc pass can surface to the user. Individual cell
/// errors flow through `Value::Error` in the cell's value — this type
/// is only for pass-level failures that warrant a status message.
#[derive(Debug, Clone)]
pub enum CalcError {
    /// A worker thread panicked. Carries the panic message captured by
    /// `catch_unwind`. The sequential path can't reach this (a panic
    /// in the orchestrator propagates normally).
    #[allow(dead_code)]
    WorkerPanic(String),
    /// Iterative calc hit `iter_max` without converging. `cells` is
    /// the number of cyclic cells that participated; users can adjust
    /// per-sheet `iter_max`/`iter_epsilon` to extend the budget or
    /// relax the convergence threshold.
    DidNotConverge {
        iter_max: usize,
        cells: usize,
    },
}

impl std::fmt::Display for CalcError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            CalcError::WorkerPanic(msg) => write!(f, "Recalc worker panicked: {}", msg),
            CalcError::DidNotConverge { iter_max, cells } => write!(
                f,
                "Iterative calc did not converge after {} passes ({} cell{} in cycle)",
                iter_max,
                cells,
                if *cells == 1 { "" } else { "s" }
            ),
        }
    }
}

/// The seam at which the parallel executor (PR 4) plugs in. Implementors
/// receive a plan describing what to recalc, a context with snapshot
/// state, and mutable access to the workbook. They re-evaluate each
/// cell in `plan` and write the new value back.
///
/// Contract: the executor must produce identical observable results
/// across implementations for the same `(plan, ctx, wb)` triple modulo
/// the volatility classes that have inherent within-pass nondeterminism
/// (RAND across worker threads). Tests in PR 4 will compare Sequential
/// and Parallel outputs for the same workload.
pub trait RecalcExecutor {
    fn run(
        &self,
        plan: &RecalcPlan,
        ctx: &mut RecalcContext,
        wb: &mut Workbook,
    ) -> Result<(), CalcError>;
}

/// Reference single-threaded implementation. Walks `plan.levels` in
/// order, re-evaluating each cell. Cyclic cells go through the per-sheet
/// iterative-calc loop. Identical behavior to the legacy
/// `Workbook::recalc_via_graph`, which is now a thin wrapper around
/// this trait call.
pub struct SequentialExecutor;

impl RecalcExecutor for SequentialExecutor {
    fn run(
        &self,
        plan: &RecalcPlan,
        ctx: &mut RecalcContext,
        wb: &mut Workbook,
    ) -> Result<(), CalcError> {
        

        // Publish the clock snapshot so NOW()/TODAY() inside this
        // pass return consistent values. Restored on exit.
        crate::domain::parser::with_recalc_clock(ctx.clock_snapshot, || {
            self.run_inner(plan, wb)
        })
    }
}

impl SequentialExecutor {
    fn run_inner(
        &self,
        plan: &RecalcPlan,
        wb: &mut Workbook,
    ) -> Result<(), CalcError> {
        // ONE snapshot for the whole recalc. We mutate the snapshot's
        // cell values between levels so the next level's evaluator
        // sees prior-level results. Eliminates the N-clones-per-recalc
        // cost that the per-level snapshot pattern paid (10ms+ per
        // clone of a 10k-cell workbook × N levels). The snapshot stays
        // valid because no workers are alive between levels — it's
        // the orchestrator's exclusive working copy.
        let mut snapshot = wb.clone();

        // Acyclic levels — process in topological order.
        for level in &plan.levels {
            // Sort for deterministic test output and to keep test
            // expectations stable across HashMap iteration orders.
            let mut ordered: Vec<NodeKey> = level.clone();
            ordered.sort();

            let mut writes: Vec<(NodeKey, CellEvalOutcome)> =
                Vec::with_capacity(ordered.len());
            for &node in &ordered {
                if let Some(outcome) = eval_one(node, &snapshot) {
                    writes.push((node, outcome));
                }
            }

            // Apply level results to the live workbook with the same
            // post-write maintenance the legacy `set_cell` performs.
            // `maybe_spill` re-evaluates the formula via the workbook
            // context thread-local; publishing `snapshot` ensures any
            // cross-sheet refs inside an array formula resolve.
            //
            // We need an immutable borrow on `snapshot` for the
            // recalc-context publish, but we also need to mutate
            // `snapshot.sheets[].cells[].value` so the next level
            // sees fresh results. Solution: collect (sheet_idx, r, c,
            // value) tuples first while snapshot is borrowed for
            // read, then drop the borrow and apply snapshot writes.
            let mut snapshot_writes: Vec<(usize, usize, usize, String)> =
                Vec::with_capacity(writes.len());
            crate::domain::models::with_recalc_context(&snapshot, || {
                for (node, outcome) in writes {
                    let Some(sheet_idx) = wb.sheet_idx_of(node.0) else {
                        continue;
                    };
                    let (r, c) = (node.1, node.2);
                    let value_clone = outcome.value.clone();
                    let sheet = &mut wb.sheets[sheet_idx];
                    if let Some(cd) = sheet.cells.get_mut(&(r, c)) {
                        cd.value = outcome.value;
                    }
                    sheet.cf_cache.lock().unwrap().clear();
                    sheet.sweep_spill_ghosts_for(r, c);
                    sheet.maybe_spill(r, c);
                    // Record dynamic targets for VolatileStructural
                    // cells so the next recalc's auto-seed can skip
                    // this cell when its targets are unrelated to the
                    // dirty closure.
                    if wb.cell_purity(node)
                        == crate::domain::parser::FunctionPurity::VolatileStructural
                    {
                        wb.record_structural_targets(node, &outcome.dynamic_targets);
                    }
                    // Queue the snapshot write so the next level reads
                    // fresh values. We only update the cell VALUE on
                    // the snapshot — cf_cache and spill ghosts are
                    // render/eval state the workers don't consult.
                    snapshot_writes.push((sheet_idx, r, c, value_clone));
                }
            });
            for (sheet_idx, r, c, value) in snapshot_writes {
                if let Some(cd) = snapshot.sheets[sheet_idx].cells.get_mut(&(r, c)) {
                    cd.value = value;
                }
            }
        }

        // Cyclic remainder: workbook-level iterative-calc that walks
        // every cycle across all sheets. The legacy per-sheet fallback
        // only saw one leg of a cross-sheet cycle and never converged.
        // Non-convergence is reported via `CalcError::DidNotConverge`
        // so the orchestrator can surface a status message; the
        // iterated values are still committed (best-effort).
        if !plan.cyclic.is_empty()
            && let Err(iter_max) = wb.iterative_calc_cyclic(&plan.cyclic)
        {
            return Err(CalcError::DidNotConverge {
                iter_max,
                cells: plan.cyclic.len(),
            });
        }

        Ok(())
    }
}

/// Rayon-based parallel executor. For each topological level, partitions
/// cells into a parallel-safe set (Pure / VolatileClock / VolatileRandom)
/// and a serial-only set (VolatileStructural / SideEffecting). The
/// parallel set is dispatched via `par_iter().with_min_len(min_chunk)`
/// over rayon's global thread pool; the serial set runs on the main
/// thread within the same level barrier.
///
/// Workers see an `Arc<WorkbookSnapshot>` (cheap-`Arc`-cloned for each
/// worker) — the snapshot is constructed once per level so reads against
/// in-progress level results stay consistent. After the level completes,
/// the orchestrator merges the per-cell results into the live workbook.
///
/// Cyclic remainder cells fall back to the per-sheet iterative-calc
/// loop, same as `SequentialExecutor`.
///
/// Configuration:
/// - `min_chunk`: minimum cells per rayon chunk. Smaller values use more
///   workers but pay more dispatch overhead. The default of `64` reflects
///   the research finding that rayon dispatch is ~400ns and a typical
///   formula evaluation is ~500ns–5μs.
/// - `parallel_threshold`: don't bother with parallel dispatch when a
///   level has fewer than this many cells; just run serial.
pub struct ParallelExecutor {
    pub min_chunk: usize,
    pub parallel_threshold: usize,
}

impl ParallelExecutor {
    pub fn new() -> Self {
        Self {
            min_chunk: 64,
            parallel_threshold: 64,
        }
    }
}

impl Default for ParallelExecutor {
    fn default() -> Self {
        Self::new()
    }
}

impl RecalcExecutor for ParallelExecutor {
    fn run(
        &self,
        plan: &RecalcPlan,
        ctx: &mut RecalcContext,
        wb: &mut Workbook,
    ) -> Result<(), CalcError> {
        // Publish clock on the main thread; workers re-publish per call
        // since thread-locals don't cross thread boundaries.
        crate::domain::parser::with_recalc_clock(ctx.clock_snapshot, || {
            self.run_inner(plan, ctx.clock_snapshot, wb)
        })
    }
}

impl ParallelExecutor {
    fn run_inner(
        &self,
        plan: &RecalcPlan,
        clock_snapshot: f64,
        wb: &mut Workbook,
    ) -> Result<(), CalcError> {
        use rayon::prelude::*;
        use std::sync::Arc;

        // ONE snapshot for the whole recalc, wrapped in Arc so workers
        // hold cheap refcounts. Between levels we `Arc::make_mut` to
        // mutate the cell values — refcount drops back to 1 after
        // `par_iter().collect()` joins all workers, so make_mut returns
        // a mut without cloning. Eliminates the per-level deep-clone
        // that the previous design paid (O(workbook_size × N_levels)).
        let mut snapshot = Arc::new(wb.clone());

        for level in &plan.levels {
            // Stable ordering for determinism / debuggability.
            let mut ordered: Vec<NodeKey> = level.clone();
            ordered.sort();

            // Partition: cells whose purity is parallel-safe go into
            // `par_cells`; everything else (VolatileStructural,
            // SideEffecting) runs sequentially.
            let mut par_cells: Vec<NodeKey> = Vec::with_capacity(ordered.len());
            let mut serial_cells: Vec<NodeKey> = Vec::new();
            for &node in &ordered {
                match wb.cell_purity(node) {
                    p if p.is_parallel_safe() => par_cells.push(node),
                    _ => serial_cells.push(node),
                }
            }

            // Sequential cells first — they may write to shared state
            // (RNG seed advances, HTTP cache inserts) so we run them on
            // the main thread within the level barrier.
            let mut writes: Vec<(NodeKey, CellEvalOutcome)> =
                Vec::with_capacity(ordered.len());
            for &node in &serial_cells {
                if let Some(outcome) = eval_one(node, &snapshot) {
                    writes.push((node, outcome));
                }
            }

            // Parallel dispatch for the safe cells. Below the threshold,
            // fall through to sequential — the overhead would exceed the
            // savings.
            if par_cells.len() >= self.parallel_threshold {
                let snap_ref: &Workbook = &snapshot;
                // Each worker publishes the clock on its own thread —
                // thread-locals don't propagate across worker spawns,
                // so we re-publish in the closure body.
                let par_writes: Vec<(NodeKey, CellEvalOutcome)> = par_cells
                    .par_iter()
                    .with_min_len(self.min_chunk)
                    .filter_map(|&node| {
                        crate::domain::parser::with_recalc_clock(
                            clock_snapshot,
                            || eval_one(node, snap_ref).map(|o| (node, o)),
                        )
                    })
                    .collect();
                writes.extend(par_writes);
            } else {
                for &node in &par_cells {
                    if let Some(outcome) = eval_one(node, &snapshot) {
                        writes.push((node, outcome));
                    }
                }
            }

            // Apply level results — same post-write maintenance as
            // SequentialExecutor. After par_iter joined, workers have
            // dropped their Arc clones, so refcount on `snapshot` is
            // back to 1 and `Arc::make_mut` returns a mut without
            // cloning. We use this to thread fresh cell values into
            // the snapshot so the next level reads them.
            let snapshot_for_ctx: &Workbook = &snapshot;
            let mut snapshot_writes: Vec<(usize, usize, usize, String)> =
                Vec::with_capacity(writes.len());
            crate::domain::models::with_recalc_context(snapshot_for_ctx, || {
                for (node, outcome) in writes {
                    let Some(sheet_idx) = wb.sheet_idx_of(node.0) else {
                        continue;
                    };
                    let (r, c) = (node.1, node.2);
                    let value_clone = outcome.value.clone();
                    let sheet = &mut wb.sheets[sheet_idx];
                    if let Some(cd) = sheet.cells.get_mut(&(r, c)) {
                        cd.value = outcome.value;
                    }
                    sheet.cf_cache.lock().unwrap().clear();
                    sheet.sweep_spill_ghosts_for(r, c);
                    sheet.maybe_spill(r, c);
                    if wb.cell_purity(node)
                        == crate::domain::parser::FunctionPurity::VolatileStructural
                    {
                        wb.record_structural_targets(node, &outcome.dynamic_targets);
                    }
                    snapshot_writes.push((sheet_idx, r, c, value_clone));
                }
            });
            // Apply queued writes to the snapshot. `Arc::make_mut`
            // returns a mut because no other Arc handle exists at
            // this point (workers all joined before par_iter returned;
            // the orchestrator holds the only Arc).
            let snap_mut = Arc::make_mut(&mut snapshot);
            for (sheet_idx, r, c, value) in snapshot_writes {
                if let Some(cd) = snap_mut.sheets[sheet_idx].cells.get_mut(&(r, c)) {
                    cd.value = value;
                }
            }
        }

        // Cyclic remainder: workbook-level iterative-calc, same as
        // SequentialExecutor. Surface non-convergence so the App layer
        // can show a status message.
        if !plan.cyclic.is_empty()
            && let Err(iter_max) = wb.iterative_calc_cyclic(&plan.cyclic)
        {
            return Err(CalcError::DidNotConverge {
                iter_max,
                cells: plan.cyclic.len(),
            });
        }

        Ok(())
    }
}

/// Outcome of evaluating one cell. `value` is the new cell value.
/// `dynamic_targets` lists the cells INDIRECT/OFFSET resolved to (if
/// any) — captured via `parser::take_dynamic_targets`. The orchestrator
/// uses `dynamic_targets` to update `Workbook::structural_targets` so
/// the next recalc's auto-seed loop can skip VolatileStructural cells
/// whose targets are unaffected by the dirty closure.
#[derive(Debug)]
pub(super) struct CellEvalOutcome {
    pub value: String,
    pub dynamic_targets: Vec<crate::domain::parser::DynamicTarget>,
}

/// Evaluate a single cell's formula against a workbook snapshot.
/// Returns `None` for non-formula cells (no value to compute) or for
/// cells whose owning sheet has been removed since the plan was built.
///
/// Hoisted so both `SequentialExecutor` and `ParallelExecutor` share
/// one canonical implementation; the closure body that goes into
/// `par_iter` is small and the impls don't drift.
fn eval_one(node: NodeKey, snapshot: &Workbook) -> Option<CellEvalOutcome> {
    use crate::domain::parser::take_dynamic_targets;
    use crate::domain::services::FormulaEvaluator;
    let sheet_idx = snapshot.sheet_idx_of(node.0)?;
    let (r, c) = (node.1, node.2);
    let formula = snapshot.sheets[sheet_idx]
        .cells
        .get(&(r, c))
        .and_then(|cd| cd.formula.clone())?;
    // Drain any prior leakage before evaluating, so the targets we
    // collect after eval are this cell's only. (Per-thread buffer, so
    // workers are independent.)
    let _ = take_dynamic_targets();
    let evaluator = FormulaEvaluator::for_workbook(
        snapshot,
        &snapshot.sheets[sheet_idx],
        &snapshot.named_ranges,
    );
    let value = evaluator.evaluate_formula(&formula);
    let dynamic_targets = take_dynamic_targets();
    Some(CellEvalOutcome { value, dynamic_targets })
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::domain::models::CellData;

    /// Sequential executor produces the same observable results as the
    /// legacy `recalc_via_graph` (which is now a thin wrapper around it).
    #[test]
    fn sequential_executor_propagates_through_levels() {
        let mut wb = Workbook::default();
        wb.add_sheet("S2".to_string());
        wb.sheets[0].cells.insert((0, 0), CellData {
            value: "10".to_string(), formula: None, format: None,
            comment: None, spill_anchor: None,
        });
        wb.sheets[0].cells.insert((1, 0), CellData {
            value: "0".to_string(),
            formula: Some("=A1+1".to_string()),
            format: None, comment: None, spill_anchor: None,
        });
        wb.sheets[1].cells.insert((0, 0), CellData {
            value: "0".to_string(),
            formula: Some("=Sheet1!A1*10".to_string()),
            format: None, comment: None, spill_anchor: None,
        });
        wb.build_dep_graph_from_scratch();
        wb.mark_dirty("Sheet1", 0, 0);

        // Build plan directly so we exercise the trait surface.
        let seeds: std::collections::HashSet<NodeKey> = wb
            .drain_dirty()
            .into_iter()
            .filter_map(|k| wb.cross_sheet_key_to_node(&k))
            .collect();
        let topo = wb.graph.topo_levels_from_seeds(&seeds);
        let plan = RecalcPlan { levels: topo.levels, cyclic: topo.cyclic };

        let mut ctx = RecalcContext::new();
        SequentialExecutor.run(&plan, &mut ctx, &mut wb).expect("recalc");

        assert_eq!(wb.sheets[0].get_cell(1, 0).value, "11");
        assert_eq!(wb.sheets[1].get_cell(0, 0).value, "100");
    }

    /// Empty plan is a no-op.
    #[test]
    fn empty_plan_is_noop() {
        let mut wb = Workbook::default();
        wb.sheets[0].cells.insert((0, 0), CellData {
            value: "5".to_string(), formula: None, format: None,
            comment: None, spill_anchor: None,
        });
        let plan = RecalcPlan::default();
        let mut ctx = RecalcContext::new();
        SequentialExecutor.run(&plan, &mut ctx, &mut wb).expect("noop");
        assert_eq!(wb.sheets[0].get_cell(0, 0).value, "5");
    }

    /// RecalcContext::new captures a clock snapshot.
    #[test]
    fn recalc_context_captures_clock() {
        let ctx = RecalcContext::new();
        // Today's serial should be > 1 (1900-01-02 = serial 2).
        assert!(ctx.clock_snapshot > 1.0);
        // And less than year-3000 (serial 400_000-ish).
        assert!(ctx.clock_snapshot < 400_000.0);
    }

    /// Helper: build a workbook with a small mix of same-sheet and
    /// cross-sheet dependencies. Returns the workbook ready for recalc.
    fn build_parity_workbook() -> Workbook {
        let mut wb = Workbook::default();
        wb.add_sheet("S2".to_string());
        // Sheet1!A1 = 10 (literal)
        wb.sheets[0].cells.insert((0, 0), CellData {
            value: "10".to_string(), formula: None, format: None,
            comment: None, spill_anchor: None,
        });
        // Sheet1!A2 = =A1+1
        wb.sheets[0].cells.insert((1, 0), CellData {
            value: "0".to_string(),
            formula: Some("=A1+1".to_string()),
            format: None, comment: None, spill_anchor: None,
        });
        // Sheet1!B1 = =A2*2
        wb.sheets[0].cells.insert((0, 1), CellData {
            value: "0".to_string(),
            formula: Some("=A2*2".to_string()),
            format: None, comment: None, spill_anchor: None,
        });
        // Sheet2!A1 = =Sheet1!A1*10
        wb.sheets[1].cells.insert((0, 0), CellData {
            value: "0".to_string(),
            formula: Some("=Sheet1!A1*10".to_string()),
            format: None, comment: None, spill_anchor: None,
        });
        // Sheet2!A2 = =A1+5
        wb.sheets[1].cells.insert((1, 0), CellData {
            value: "0".to_string(),
            formula: Some("=A1+5".to_string()),
            format: None, comment: None, spill_anchor: None,
        });
        wb.build_dep_graph_from_scratch();
        wb.mark_dirty("Sheet1", 0, 0);
        wb
    }

    fn plan_from_workbook(wb: &mut Workbook) -> RecalcPlan {
        let seeds: std::collections::HashSet<NodeKey> = wb
            .drain_dirty()
            .into_iter()
            .filter_map(|k| wb.cross_sheet_key_to_node(&k))
            .collect();
        let topo = wb.graph.topo_levels_from_seeds(&seeds);
        RecalcPlan { levels: topo.levels, cyclic: topo.cyclic }
    }

    /// ParallelExecutor produces identical results to SequentialExecutor.
    /// This is the central parity test for PR 4 — every other parallel
    /// guarantee builds on this.
    #[test]
    fn parallel_matches_sequential() {
        let mut seq_wb = build_parity_workbook();
        let seq_plan = plan_from_workbook(&mut seq_wb);
        let mut seq_ctx = RecalcContext::new();
        SequentialExecutor.run(&seq_plan, &mut seq_ctx, &mut seq_wb).expect("seq");

        let mut par_wb = build_parity_workbook();
        let par_plan = plan_from_workbook(&mut par_wb);
        let mut par_ctx = RecalcContext::new();
        // Force parallel dispatch even for the small workload by setting
        // a low threshold.
        let exec = ParallelExecutor {
            min_chunk: 1,
            parallel_threshold: 1,
        };
        exec.run(&par_plan, &mut par_ctx, &mut par_wb).expect("par");

        // Compare each cell.
        for sheet_idx in 0..seq_wb.sheets.len() {
            for &(r, c) in seq_wb.sheets[sheet_idx].cells.keys() {
                let sv = seq_wb.sheets[sheet_idx].get_cell(r, c);
                let pv = par_wb.sheets[sheet_idx].get_cell(r, c);
                assert_eq!(sv.value, pv.value,
                    "mismatch at sheet {} ({}, {}): sequential={} parallel={}",
                    sheet_idx, r, c, sv.value, pv.value);
            }
        }
        // Spot-check the headline values.
        assert_eq!(par_wb.sheets[0].get_cell(1, 0).value, "11");
        assert_eq!(par_wb.sheets[0].get_cell(0, 1).value, "22");
        assert_eq!(par_wb.sheets[1].get_cell(0, 0).value, "100");
        assert_eq!(par_wb.sheets[1].get_cell(1, 0).value, "105");
    }

    /// ParallelExecutor partitions structural-volatile cells to the
    /// serial path. The Sequential path runs them in sequence within a
    /// level — verify the dispatch doesn't break their evaluation.
    #[test]
    fn parallel_handles_structural_volatile() {
        let mut wb = Workbook::default();
        wb.sheets[0].cells.insert((0, 0), CellData {
            value: "42".to_string(), formula: None, format: None,
            comment: None, spill_anchor: None,
        });
        // =OFFSET(A1, 0, 0) — structural volatile
        wb.sheets[0].cells.insert((1, 0), CellData {
            value: "0".to_string(),
            formula: Some("=OFFSET(A1, 0, 0)".to_string()),
            format: None, comment: None, spill_anchor: None,
        });
        // =A2 + 1 — pure, depends on the structural cell
        wb.sheets[0].cells.insert((2, 0), CellData {
            value: "0".to_string(),
            formula: Some("=A2+1".to_string()),
            format: None, comment: None, spill_anchor: None,
        });
        wb.build_dep_graph_from_scratch();
        wb.mark_dirty("Sheet1", 0, 0);

        let plan = plan_from_workbook(&mut wb);
        let mut ctx = RecalcContext::new();
        let exec = ParallelExecutor {
            min_chunk: 1,
            parallel_threshold: 1,
        };
        exec.run(&plan, &mut ctx, &mut wb).expect("par");

        assert_eq!(wb.sheets[0].get_cell(1, 0).value, "42");
        assert_eq!(wb.sheets[0].get_cell(2, 0).value, "43");
    }

    /// ParallelExecutor falls back to sequential below the
    /// parallel_threshold. Behaviorally indistinguishable but exercises
    /// the threshold code path.
    #[test]
    fn parallel_below_threshold_falls_through_to_sequential() {
        let mut wb = build_parity_workbook();
        let plan = plan_from_workbook(&mut wb);
        let mut ctx = RecalcContext::new();
        // Threshold higher than any level's size.
        let exec = ParallelExecutor {
            min_chunk: 64,
            parallel_threshold: 1_000_000,
        };
        exec.run(&plan, &mut ctx, &mut wb).expect("par-below-threshold");
        // Values should still be correct — the threshold just changes
        // how the work is dispatched, not what it computes.
        assert_eq!(wb.sheets[0].get_cell(1, 0).value, "11");
        assert_eq!(wb.sheets[1].get_cell(0, 0).value, "100");
    }

    /// Property test: for many randomly-generated workbooks, both
    /// executors must produce IDENTICAL cell values. This catches any
    /// future divergence (e.g. a snapshot-mutation order bug, a
    /// purity-classification mismatch, a thread-local state leak)
    /// the moment it lands rather than waiting for it to surface as
    /// a real-user bug.
    ///
    /// Strategy:
    ///   1. Generate a workbook of N cells. ~30% are literal numbers,
    ///      ~70% are formulas combining a small set of prior cells
    ///      with random arithmetic operators. Each formula references
    ///      only earlier-defined cells so the graph is acyclic by
    ///      construction (avoids the iterative-calc tangent).
    ///   2. Build twice — once for Sequential, once for Parallel.
    ///   3. Compare every cell. ANY difference fails the test with
    ///      the seed printed so the failure is reproducible.
    ///
    /// Deterministic LCG (no rand dep). Seeds 1..=20 sweep a range
    /// of structures from "small chain" to "wide fanout".
    #[test]
    fn property_parallel_matches_sequential_random_workbooks() {
        for seed in 1u64..=20 {
            let (seq_wb, par_wb) = run_pair_with_seed(seed);
            // Compare every cell on every sheet.
            let n_sheets = seq_wb.sheets.len();
            assert_eq!(n_sheets, par_wb.sheets.len(),
                "seed {}: sheet count mismatch", seed);
            for sheet_idx in 0..n_sheets {
                let seq_cells = &seq_wb.sheets[sheet_idx].cells;
                let par_cells = &par_wb.sheets[sheet_idx].cells;
                assert_eq!(seq_cells.len(), par_cells.len(),
                    "seed {}: sheet {} cell count mismatch (seq={}, par={})",
                    seed, sheet_idx, seq_cells.len(), par_cells.len());
                for (&(r, c), seq_cd) in seq_cells {
                    let par_cd = par_cells.get(&(r, c))
                        .unwrap_or_else(|| panic!(
                            "seed {}: par missing cell sheet={} ({}, {})",
                            seed, sheet_idx, r, c));
                    assert_eq!(seq_cd.value, par_cd.value,
                        "seed {}: divergence at sheet {} ({}, {}): \
                         seq={:?} par={:?} (formula={:?})",
                        seed, sheet_idx, r, c, seq_cd.value, par_cd.value,
                        seq_cd.formula);
                }
            }
        }
    }

    /// Generate one parity workbook with the given seed, run it through
    /// both executors, return both for comparison.
    fn run_pair_with_seed(seed: u64) -> (Workbook, Workbook) {
        let formulas = generate_random_dag(seed);

        let mut seq_wb = build_workbook_from(&formulas);
        seq_wb.recalc_via_graph_result().expect("seq seed runs cleanly");

        let mut par_wb = build_workbook_from(&formulas);
        // Force parallel dispatch even on small workloads — the property
        // tests would otherwise both fall back to Sequential.
        let exec = ParallelExecutor { min_chunk: 1, parallel_threshold: 1 };
        par_wb.build_dep_graph_from_scratch();
        let seeds: std::collections::HashSet<NodeKey> = par_wb
            .drain_dirty()
            .into_iter()
            .filter_map(|k| par_wb.cross_sheet_key_to_node(&k))
            .collect();
        let topo = par_wb.graph.topo_levels_from_seeds(&seeds);
        let plan = RecalcPlan { levels: topo.levels, cyclic: topo.cyclic };
        let mut ctx = RecalcContext::new();
        exec.run(&plan, &mut ctx, &mut par_wb).expect("par seed runs cleanly");

        (seq_wb, par_wb)
    }

    fn build_workbook_from(formulas: &[(usize, usize, String)]) -> Workbook {
        let mut wb = Workbook::default();
        let last = formulas.iter().enumerate();
        for (idx, (r, c, content)) in last {
            // First the cell goes in via the public mutator so the graph
            // and dirty set are maintained. Skip the auto-recalc-per-cell
            // overhead by using set_cell_on_active — it batches into the
            // graph + recalcs each time, but the workload is small.
            let cd = if let Some(rest) = content.strip_prefix('=') {
                // It's a formula — keep the leading `=`.
                let _ = rest;
                CellData {
                    value: "0".to_string(),
                    formula: Some(content.clone()),
                    format: None, comment: None, spill_anchor: None,
                }
            } else {
                CellData {
                    value: content.clone(),
                    formula: None,
                    format: None, comment: None, spill_anchor: None,
                }
            };
            wb.set_cell_on_active(*r, *c, cd);
            let _ = idx;
        }
        wb
    }

    /// Hand-rolled LCG: deterministic, no external dep. Returns a
    /// sequence of (row, col, content) tuples laying down a topologically-
    /// ordered DAG of cells in column A. The first N cells are literals;
    /// later cells are formulas referencing strictly-earlier ones.
    fn generate_random_dag(seed: u64) -> Vec<(usize, usize, String)> {
        let mut state = seed.wrapping_mul(2862933555777941757).wrapping_add(3037000493);
        let mut next = || {
            state = state.wrapping_mul(6364136223846793005).wrapping_add(1442695040888963407);
            (state >> 33) as u32
        };
        let n_cells = 30 + (next() % 30) as usize;       // 30..60 cells
        let n_literals = (n_cells as u32 / 3).max(3);     // ~1/3 literals
        let mut cells: Vec<(usize, usize, String)> = Vec::new();
        // Literals come first so later formulas always have something
        // to reference. Mix integers and small positive numbers so
        // SUMPRODUCT and IF-comparison branches see varied inputs.
        for i in 0..n_literals {
            let v = ((next() % 100) as i64) - 50;
            cells.push((i as usize, 0, v.to_string()));
        }
        // Build a few helper cells in column B so range-aware functions
        // (SUM, AVERAGE, VLOOKUP) have a contiguous block to consult.
        // These are mirrors of column A's first ten literals.
        let n_helpers = n_literals.min(10) as usize;
        for i in 0..n_helpers {
            cells.push((i, 1, format!("=A{}", i + 1)));
        }
        for i in n_literals..(n_cells as u32) {
            let row = i as usize;
            // Pick a random shape per cell — operator chain, IF, SUM,
            // MAX, ABS, or VLOOKUP. Each path exercises a different
            // executor surface; if Sequential and Parallel diverge on
            // any of them, the fuzz catches it.
            let formula = match next() % 6 {
                0 => gen_arith(&mut next, i),
                1 => gen_if(&mut next, i),
                2 => gen_sum_range(&mut next, i, n_helpers),
                3 => gen_max_range(&mut next, i, n_helpers),
                4 => gen_abs(&mut next, i),
                _ => gen_vlookup(&mut next, i, n_helpers),
            };
            cells.push((row, 0, formula));
        }
        cells
    }

    fn gen_arith(next: &mut impl FnMut() -> u32, i: u32) -> String {
        let n_refs = 1 + (next() % 3) as usize;
        let mut f = String::from("=");
        for j in 0..n_refs {
            if j > 0 {
                let op = match next() % 4 {
                    0 => '+',
                    1 => '-',
                    2 => '*',
                    _ => '+', // skip '/' to avoid #DIV/0! noise
                };
                f.push(op);
            }
            let pick = (next() as usize) % (i as usize);
            f.push_str(&format!("A{}", pick + 1));
        }
        f
    }

    fn gen_if(next: &mut impl FnMut() -> u32, i: u32) -> String {
        let a = (next() as usize) % (i as usize) + 1;
        let b = (next() as usize) % (i as usize) + 1;
        let c = (next() as usize) % (i as usize) + 1;
        let d = (next() as usize) % (i as usize) + 1;
        format!("=IF(A{}>A{},A{},A{})", a, b, c, d)
    }

    fn gen_sum_range(next: &mut impl FnMut() -> u32, _i: u32, helpers: usize) -> String {
        if helpers < 2 { return "=0".to_string(); }
        // Sum a sub-range of column B (the helpers).
        let lo = (next() as usize) % helpers;
        let hi_extra = (next() as usize) % (helpers - lo);
        let hi = lo + hi_extra;
        format!("=SUM(B{}:B{})", lo + 1, hi + 1)
    }

    fn gen_max_range(next: &mut impl FnMut() -> u32, _i: u32, helpers: usize) -> String {
        if helpers < 2 { return "=0".to_string(); }
        let lo = (next() as usize) % helpers;
        let hi_extra = (next() as usize) % (helpers - lo);
        let hi = lo + hi_extra;
        format!("=MAX(B{}:B{})", lo + 1, hi + 1)
    }

    fn gen_abs(next: &mut impl FnMut() -> u32, i: u32) -> String {
        let pick = (next() as usize) % (i as usize) + 1;
        format!("=ABS(A{})", pick)
    }

    fn gen_vlookup(next: &mut impl FnMut() -> u32, _i: u32, helpers: usize) -> String {
        if helpers < 2 { return "=0".to_string(); }
        // VLOOKUP into B's helper column. Use approximate-match
        // (TRUE) since helpers may not be sorted — Excel returns the
        // last-matching key in approximate mode regardless. Both
        // executors should behave identically.
        let target = (next() as usize) % helpers + 1;
        format!("=IFERROR(VLOOKUP(A{},B1:B{},1,TRUE),0)", target, helpers)
    }

    /// Single-snapshot recalc must propagate per-level writes into the
    /// snapshot the next level reads from. Without that, a depth-3
    /// chain (A1=lit, A2=A1+1, A3=A2+1) would have A3 reading A2's
    /// pre-recalc value from the stale snapshot.
    ///
    /// This is the load-bearing regression test for the per-level
    /// snapshot mutation in both executors. We exercise both paths
    /// (Sequential, Parallel-with-threshold=1) and a fresh edit
    /// midway to make sure values truly flow level→level→level.
    #[test]
    fn snapshot_updates_between_levels_for_deep_chain() {
        let cases: [(&str, Box<dyn RecalcExecutor>); 2] = [
            ("seq", Box::new(SequentialExecutor)),
            ("par", Box::new(ParallelExecutor { min_chunk: 1, parallel_threshold: 1 })),
        ];
        for (label, exec) in cases {
            let mut wb = Workbook::default();
            // A1=10, A2=A1+1, A3=A2+1, A4=A3+1 — four levels.
            wb.sheets[0].cells.insert((0, 0), CellData {
                value: "10".to_string(), formula: None, format: None,
                comment: None, spill_anchor: None,
            });
            for (r, f) in [(1, "=A1+1"), (2, "=A2+1"), (3, "=A3+1")] {
                wb.sheets[0].cells.insert((r, 0), CellData {
                    value: "0".to_string(),
                    formula: Some(f.to_string()),
                    format: None, comment: None, spill_anchor: None,
                });
            }
            wb.build_dep_graph_from_scratch();
            wb.mark_dirty("Sheet1", 0, 0);
            let plan = plan_from_workbook(&mut wb);
            let mut ctx = RecalcContext::new();
            exec.run(&plan, &mut ctx, &mut wb).expect(label);

            assert_eq!(wb.sheets[0].get_cell(1, 0).value, "11", "{label}: A2");
            assert_eq!(wb.sheets[0].get_cell(2, 0).value, "12", "{label}: A3");
            assert_eq!(wb.sheets[0].get_cell(3, 0).value, "13", "{label}: A4");

            // Re-edit A1 and recalc again — confirms the snapshot is
            // freshly built each recalc (not reused across recalcs).
            wb.sheets[0].cells.insert((0, 0), CellData {
                value: "100".to_string(), formula: None, format: None,
                comment: None, spill_anchor: None,
            });
            wb.mark_dirty("Sheet1", 0, 0);
            let plan2 = plan_from_workbook(&mut wb);
            let mut ctx2 = RecalcContext::new();
            exec.run(&plan2, &mut ctx2, &mut wb).expect(label);
            assert_eq!(wb.sheets[0].get_cell(3, 0).value, "103", "{label}: A4 after re-edit");
        }
    }
}
