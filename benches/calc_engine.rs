//! Benchmark: parallel vs sequential recalc engine.
//!
//! Tests both executors against archetype workbooks specified in
//! `research/benchmarking-and-recalc-semantics/index.md`. Run with:
//!
//! ```sh
//! cargo bench --bench calc_engine
//! ```
//!
//! Or for a single archetype:
//!
//! ```sh
//! cargo bench --bench calc_engine -- wide
//! ```
//!
//! Methodology: each iteration creates a fresh workbook (so warm caches
//! from a previous iteration don't pollute), pre-builds the dep graph
//! (excluded from timing), marks every formula cell dirty, and runs the
//! executor on the resulting plan. Cold-cache; recalc only.
//!
//! Caveats:
//! - Single-threaded baselines use `SequentialExecutor`, not the legacy
//!   per-sheet cascade. We benchmark what production uses.
//! - "Parallel" uses rayon's global pool sized to the host's physical
//!   cores. Run with `RAYON_NUM_THREADS=N` to override.

use std::collections::HashSet;
use std::hint::black_box;
use std::time::{Duration, Instant};

use tshts::domain::models::NodeKey;
use tshts::domain::services::{
    ParallelExecutor, RecalcContext, RecalcExecutor, RecalcPlan, SequentialExecutor,
};
use tshts::domain::{CellData, Workbook};

/// "Shallow-wide": N independent formula cells, no dep chains.
/// Ideal parallel workload — every cell can run on its own worker.
fn build_wide(n: usize) -> Workbook {
    let mut wb = Workbook::default();
    wb.sheets[0].rows = n + 10;
    wb.sheets[0].cols = 4;
    for i in 0..n {
        wb.sheets[0].cells.insert(
            (i, 0),
            CellData {
                value: (i as i64).to_string(),
                formula: None,
                format: None,
                comment: None,
                spill_anchor: None,
            },
        );
    }
    for i in 0..n {
        wb.sheets[0].cells.insert(
            (i, 1),
            CellData {
                value: "0".to_string(),
                formula: Some(format!("=A{}*2+1", i + 1)),
                format: None,
                comment: None,
                spill_anchor: None,
            },
        );
    }
    wb.build_dep_graph_from_scratch();
    wb
}

/// "Deep-narrow": a chain of N formula cells. Worst parallel case.
fn build_deep(n: usize) -> Workbook {
    let mut wb = Workbook::default();
    wb.sheets[0].rows = n + 10;
    wb.sheets[0].cols = 1;
    wb.sheets[0].cells.insert(
        (0, 0),
        CellData {
            value: "1".to_string(),
            formula: None,
            format: None,
            comment: None,
            spill_anchor: None,
        },
    );
    for i in 1..n {
        wb.sheets[0].cells.insert(
            (i, 0),
            CellData {
                value: "0".to_string(),
                formula: Some(format!("=A{}+1", i)),
                format: None,
                comment: None,
                spill_anchor: None,
            },
        );
    }
    wb.build_dep_graph_from_scratch();
    wb
}

/// "Fan-out": one seed with N independent dependents at level 1.
fn build_fanout(n: usize) -> Workbook {
    let mut wb = Workbook::default();
    wb.sheets[0].rows = n + 10;
    wb.sheets[0].cols = 2;
    wb.sheets[0].cells.insert(
        (0, 0),
        CellData {
            value: "100".to_string(),
            formula: None,
            format: None,
            comment: None,
            spill_anchor: None,
        },
    );
    for i in 1..=n {
        wb.sheets[0].cells.insert(
            (i, 0),
            CellData {
                value: "0".to_string(),
                formula: Some("=A1*2+1".to_string()),
                format: None,
                comment: None,
                spill_anchor: None,
            },
        );
    }
    wb.build_dep_graph_from_scratch();
    wb
}

/// Mark every formula cell dirty and build a RecalcPlan.
fn full_recalc_plan(wb: &mut Workbook) -> RecalcPlan {
    let sheet_count = wb.sheets.len();
    for sheet_idx in 0..sheet_count {
        let name = wb.sheet_names[sheet_idx].clone();
        let dirty_keys: Vec<_> = wb.sheets[sheet_idx]
            .cells
            .iter()
            .filter_map(|(&(r, c), cd)| {
                if cd.formula.is_some() {
                    Some((name.clone(), r, c))
                } else {
                    None
                }
            })
            .collect();
        for k in dirty_keys {
            wb.dirty.insert(k);
        }
    }
    let seeds: HashSet<NodeKey> = wb
        .drain_dirty()
        .into_iter()
        .filter_map(|k| wb.cross_sheet_key_to_node(&k))
        .collect();
    let topo = wb.graph.topo_levels_from_seeds(&seeds);
    RecalcPlan {
        levels: topo.levels,
        cyclic: topo.cyclic,
    }
}

/// Run `executor` against `builder` `iterations` times,
/// returning median wall-clock per recalc.
fn time_executor<E: RecalcExecutor>(
    executor: &E,
    builder: impl Fn() -> Workbook,
    iterations: usize,
) -> Duration {
    let mut samples: Vec<Duration> = Vec::with_capacity(iterations);
    for _ in 0..iterations {
        let mut wb = builder();
        let plan = full_recalc_plan(&mut wb);
        let mut ctx = RecalcContext::new();
        let start = Instant::now();
        executor.run(&plan, &mut ctx, &mut wb).expect("recalc");
        let elapsed = start.elapsed();
        black_box(wb.sheets[0].get_cell(0, 0).value.clone());
        samples.push(elapsed);
    }
    samples.sort();
    samples[samples.len() / 2]
}

fn run_archetype(name: &str, n: usize, iters: usize, builder: impl Fn() -> Workbook) {
    let seq = time_executor(&SequentialExecutor, &builder, iters);
    let par = time_executor(&ParallelExecutor::new(), &builder, iters);
    let speedup = seq.as_secs_f64() / par.as_secs_f64().max(1e-9);
    println!(
        "{:<24} N={:<6}  seq={:>12.3?}  par={:>12.3?}  speedup={:>5.2}x",
        name, n, seq, par, speedup
    );
}

fn main() {
    let filter = std::env::args().nth(1);
    let want = |name: &str| {
        filter
            .as_ref()
            .map(|f| name.contains(f.as_str()))
            .unwrap_or(true)
    };

    println!();
    println!("Calc engine benchmarks (median over 5 iterations):");
    println!("{}", "─".repeat(86));

    if want("wide") {
        run_archetype("wide_small", 100, 5, || build_wide(100));
        run_archetype("wide_medium", 1_000, 5, || build_wide(1_000));
        run_archetype("wide_large", 10_000, 3, || build_wide(10_000));
    }
    if want("deep") {
        run_archetype("deep_small", 100, 5, || build_deep(100));
        run_archetype("deep_medium", 1_000, 5, || build_deep(1_000));
        run_archetype("deep_large", 5_000, 3, || build_deep(5_000));
    }
    if want("fanout") {
        run_archetype("fanout_small", 100, 5, || build_fanout(100));
        run_archetype("fanout_medium", 1_000, 5, || build_fanout(1_000));
        run_archetype("fanout_large", 10_000, 3, || build_fanout(10_000));
    }

    println!("{}", "─".repeat(86));
    println!(
        "Threads: {}",
        std::env::var("RAYON_NUM_THREADS").unwrap_or_else(|_| {
            std::thread::available_parallelism()
                .map(|n| n.to_string())
                .unwrap_or_else(|_| "?".to_string())
        })
    );
    println!();
}
