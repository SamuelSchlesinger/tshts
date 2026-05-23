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

use std::collections::HashMap;

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
    /// cell in a single pass returns the same instant.
    ///
    /// Today NOW/TODAY in tshts still read `SystemTime::now()` directly
    /// because the evaluator doesn't yet plumb `RecalcContext` through.
    /// PR 4 will swap the implementations. The field is here so the
    /// trait stabilizes before that swap.
    #[allow(dead_code)]
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

impl RecalcPlan {
    pub fn is_empty(&self) -> bool {
        self.levels.is_empty() && self.cyclic.is_empty()
    }

    /// Total cell count across all levels (excluding cyclic remainder).
    #[allow(dead_code)]
    pub fn cell_count(&self) -> usize {
        self.levels.iter().map(|l| l.len()).sum()
    }
}

/// Errors a recalc pass can surface. Today only used for catastrophic
/// failures (worker panic in PR 4); individual cell errors flow through
/// `Value::Error` in the cell's value, not through this type.
#[derive(Debug, Clone)]
pub enum CalcError {
    /// A worker thread panicked. Carries the panic message captured by
    /// `catch_unwind`. PR 4 produces this; PR 3's sequential path can't
    /// reach it (a panic in the orchestrator propagates normally).
    #[allow(dead_code)]
    WorkerPanic(String),
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
        _ctx: &mut RecalcContext,
        wb: &mut Workbook,
    ) -> Result<(), CalcError> {
        use crate::domain::services::FormulaEvaluator;

        // Acyclic levels — process in topological order.
        for level in &plan.levels {
            // Sort for deterministic test output and to keep test
            // expectations stable across HashMap iteration orders.
            let mut ordered: Vec<NodeKey> = level.clone();
            ordered.sort();

            // Snapshot for cross-sheet ref resolution. The evaluator
            // reads from this stable view; we write back to the live
            // workbook after each level so the next level sees the
            // up-to-date values. (PR 4 replaces this clone with
            // `Arc<WorkbookSnapshot>`.)
            let snapshot = wb.clone();
            let names = snapshot.named_ranges.clone();
            let mut writes: Vec<(NodeKey, String)> = Vec::with_capacity(ordered.len());
            for &node in &ordered {
                let Some(sheet_idx) = wb.sheet_idx_of(node.0) else {
                    continue;
                };
                let (r, c) = (node.1, node.2);
                let Some(formula) = snapshot.sheets[sheet_idx]
                    .cells
                    .get(&(r, c))
                    .and_then(|cd| cd.formula.clone())
                else {
                    // Non-formula seed cell — already up-to-date.
                    continue;
                };
                let evaluator = FormulaEvaluator::for_workbook(
                    &snapshot,
                    &snapshot.sheets[sheet_idx],
                    &names,
                );
                let value = evaluator.evaluate_formula(&formula);
                writes.push((node, value));
            }

            // Apply level results to the live workbook with the same
            // post-write maintenance the legacy `set_cell` performs.
            // `maybe_spill` re-evaluates the formula via the workbook
            // context thread-local; publishing `snapshot` ensures any
            // cross-sheet refs inside an array formula resolve.
            crate::domain::models::with_recalc_context(&snapshot, || {
                for (node, value) in writes {
                    let Some(sheet_idx) = wb.sheet_idx_of(node.0) else {
                        continue;
                    };
                    let (r, c) = (node.1, node.2);
                    let sheet = &mut wb.sheets[sheet_idx];
                    if let Some(cd) = sheet.cells.get_mut(&(r, c)) {
                        cd.value = value;
                    }
                    sheet.cf_cache.lock().unwrap().clear();
                    sheet.sweep_spill_ghosts_for(r, c);
                    sheet.maybe_spill(r, c);
                }
            });
        }

        // Cyclic remainder: fall back to per-sheet iterative-calc.
        if !plan.cyclic.is_empty() {
            let mut by_sheet: HashMap<usize, Vec<(usize, usize)>> = HashMap::new();
            for &(sid, r, c) in &plan.cyclic {
                if let Some(idx) = wb.sheet_idx_of(sid) {
                    by_sheet.entry(idx).or_default().push((r, c));
                }
            }
            let snapshot = wb.clone();
            crate::domain::models::with_recalc_context(&snapshot, || {
                for (sheet_idx, cells) in by_sheet {
                    for (r, c) in cells {
                        wb.sheets[sheet_idx].recalculate_dependents(r, c);
                    }
                }
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
        _ctx: &mut RecalcContext,
        wb: &mut Workbook,
    ) -> Result<(), CalcError> {
        use crate::domain::parser::FunctionPurity;
        use crate::domain::services::FormulaEvaluator;
        use rayon::prelude::*;
        use std::sync::Arc;

        for level in &plan.levels {
            // Stable ordering for determinism / debuggability.
            let mut ordered: Vec<NodeKey> = level.clone();
            ordered.sort();

            // Partition: cells whose purity is parallel-safe go into
            // `par_cells`; everything else (VolatileStructural,
            // SideEffecting) runs sequentially. Today the partition is a
            // simple per-cell purity lookup; PR 4+ can refine.
            let mut par_cells: Vec<NodeKey> = Vec::with_capacity(ordered.len());
            let mut serial_cells: Vec<NodeKey> = Vec::new();
            for &node in &ordered {
                match wb.cell_purity(node) {
                    p if p.is_parallel_safe() => par_cells.push(node),
                    _ => serial_cells.push(node),
                }
            }

            // Build a per-level snapshot. Arc-wrap so each rayon worker
            // holds a cheap refcount instead of cloning the workbook
            // N times. Without Arc, every worker would deep-clone via
            // capture-by-value — the exact problem this PR exists to
            // avoid.
            let snapshot = Arc::new(wb.clone());

            // Sequential cells first — they may write to shared state
            // (RNG seed advances, HTTP cache inserts) so we run them on
            // the main thread within the level barrier.
            let mut writes: Vec<(NodeKey, String)> = Vec::with_capacity(ordered.len());
            for &node in &serial_cells {
                if let Some(value) = eval_one(node, &snapshot) {
                    writes.push((node, value));
                }
            }

            // Parallel dispatch for the safe cells. Below the threshold,
            // fall through to sequential — the overhead would exceed the
            // savings.
            if par_cells.len() >= self.parallel_threshold {
                let snap_ref: &Workbook = &snapshot;
                let par_writes: Vec<(NodeKey, String)> = par_cells
                    .par_iter()
                    .with_min_len(self.min_chunk)
                    .filter_map(|&node| {
                        eval_one(node, snap_ref).map(|v| (node, v))
                    })
                    .collect();
                writes.extend(par_writes);
            } else {
                for &node in &par_cells {
                    if let Some(value) = eval_one(node, &snapshot) {
                        writes.push((node, value));
                    }
                }
            }

            // Apply level results — same post-write maintenance as
            // SequentialExecutor, including the with_recalc_context
            // wrap so spill re-evaluation can resolve cross-sheet refs.
            crate::domain::models::with_recalc_context(&snapshot, || {
                for (node, value) in writes {
                    let Some(sheet_idx) = wb.sheet_idx_of(node.0) else {
                        continue;
                    };
                    let (r, c) = (node.1, node.2);
                    let sheet = &mut wb.sheets[sheet_idx];
                    if let Some(cd) = sheet.cells.get_mut(&(r, c)) {
                        cd.value = value;
                    }
                    sheet.cf_cache.lock().unwrap().clear();
                    sheet.sweep_spill_ghosts_for(r, c);
                    sheet.maybe_spill(r, c);
                }
            });
        }

        // Cyclic remainder: same fallback as SequentialExecutor.
        if !plan.cyclic.is_empty() {
            let mut by_sheet: std::collections::HashMap<usize, Vec<(usize, usize)>> =
                std::collections::HashMap::new();
            for &(sid, r, c) in &plan.cyclic {
                if let Some(idx) = wb.sheet_idx_of(sid) {
                    by_sheet.entry(idx).or_default().push((r, c));
                }
            }
            let snapshot = wb.clone();
            crate::domain::models::with_recalc_context(&snapshot, || {
                for (sheet_idx, cells) in by_sheet {
                    for (r, c) in cells {
                        wb.sheets[sheet_idx].recalculate_dependents(r, c);
                    }
                }
            });
        }

        Ok(())
    }
}

/// Evaluate a single cell's formula against a workbook snapshot.
/// Returns `None` for non-formula cells (no value to compute) or for
/// cells whose owning sheet has been removed since the plan was built.
///
/// Hoisted so both `SequentialExecutor` and `ParallelExecutor` share
/// one canonical implementation; the closure body that goes into
/// `par_iter` is small and the impls don't drift.
fn eval_one(node: NodeKey, snapshot: &Workbook) -> Option<String> {
    use crate::domain::services::FormulaEvaluator;
    let sheet_idx = snapshot.sheet_idx_of(node.0)?;
    let (r, c) = (node.1, node.2);
    let formula = snapshot.sheets[sheet_idx]
        .cells
        .get(&(r, c))
        .and_then(|cd| cd.formula.clone())?;
    let evaluator = FormulaEvaluator::for_workbook(
        snapshot,
        &snapshot.sheets[sheet_idx],
        &snapshot.named_ranges,
    );
    Some(evaluator.evaluate_formula(&formula))
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::domain::models::{CellData, SheetId};

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
}
