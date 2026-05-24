//! Sales-rep leaderboard. Ranks reps by quarterly sales using `RANK.EQ`,
//! pulls the top-3 via `LARGE`, computes percentile rank via simple
//! math. Exercises tie-breaking and ranking semantics.

use super::{lit, CellCheck, Scenario};
use crate::common::scenarios::{enter_cell, enter_cells, Harness};

pub struct Leaderboard;

#[derive(Clone)]
pub struct Rep {
    pub name: &'static str,
    pub sales: f64,
}

#[derive(Clone)]
pub struct Inputs {
    pub reps: Vec<Rep>,
}

pub struct Output {
    pub ranks: Vec<f64>,       // RANK.EQ for each rep (1 = highest)
    pub top1: f64,
    pub top2: f64,
    pub top3: f64,
    pub bottom1: f64,
    pub median: f64,
    pub q3: f64,                // 75th percentile
}

impl Scenario for Leaderboard {
    type Inputs = Inputs;
    type Output = Output;

    fn name(&self) -> &'static str { "leaderboard" }

    fn default_inputs(&self) -> Inputs {
        Inputs {
            reps: vec![
                Rep { name: "Aki",   sales: 412_000.0 },
                Rep { name: "Bea",   sales: 285_000.0 },
                Rep { name: "Cyrus", sales: 612_500.0 },
                Rep { name: "Dora",  sales: 412_000.0 }, // tie with Aki
                Rep { name: "Eve",   sales: 198_000.0 },
                Rep { name: "Finn",  sales: 503_700.0 },
                Rep { name: "Gao",   sales: 367_400.0 },
                Rep { name: "Hugo",  sales:  82_100.0 },
            ],
        }
    }

    fn compute(&self, i: &Inputs) -> Output {
        // RANK.EQ semantics: rank = 1 + count of reps with strictly higher sales.
        // Tied reps get the same rank; the next distinct value skips.
        let ranks: Vec<f64> = i.reps.iter().map(|r| {
            1.0 + i.reps.iter().filter(|x| x.sales > r.sales).count() as f64
        }).collect();
        let mut sorted: Vec<f64> = i.reps.iter().map(|r| r.sales).collect();
        sorted.sort_by(|a, b| b.partial_cmp(a).unwrap());
        let top1 = sorted[0];
        let top2 = sorted[1];
        let top3 = sorted[2];
        let bottom1 = *sorted.last().unwrap();
        // Median: linear interpolation between the two middle values for
        // even-length lists (matches tshts/Excel MEDIAN semantics).
        let n = sorted.len();
        let median = if n.is_multiple_of(2) {
            (sorted[n / 2 - 1] + sorted[n / 2]) / 2.0
        } else {
            sorted[n / 2]
        };
        // 75th percentile via linear interp (Excel PERCENTILE.INC).
        let q3 = percentile_inc(&sorted, 0.75);
        Output { ranks, top1, top2, top3, bottom1, median, q3 }
    }

    fn populate(&self, h: &mut Harness, i: &Inputs) {
        // A: name | B: sales | C: rank
        enter_cells(h, &[
            ("A1", "name"), ("B1", "sales"), ("C1", "rank"),
        ]);
        for (idx, r) in i.reps.iter().enumerate() {
            let row = 2 + idx;
            enter_cell(h, &format!("A{}", row), r.name);
            enter_cell(h, &format!("B{}", row), &lit(r.sales));
            enter_cell(h, &format!("C{}", row),
                &format!("=RANK.EQ(B{},$B$2:$B${})", row, 1 + i.reps.len()));
        }
        let last = 1 + i.reps.len();
        let range = format!("B2:B{}", last);
        enter_cells(h, &[
            ("E1", "top1"), ("E2", "top2"), ("E3", "top3"),
            ("E4", "bottom1"), ("E5", "median"), ("E6", "q3"),
        ]);
        enter_cell(h, "F1", &format!("=LARGE({},1)", range));
        enter_cell(h, "F2", &format!("=LARGE({},2)", range));
        enter_cell(h, "F3", &format!("=LARGE({},3)", range));
        enter_cell(h, "F4", &format!("=SMALL({},1)", range));
        enter_cell(h, "F5", &format!("=MEDIAN({})", range));
        // PERCENTILE.INC with the literal 0.75
        enter_cell(h, "F6", &format!("=PERCENTILE.INC({},0.75)", range));
    }

    fn checks(&self, o: &Output) -> Vec<CellCheck> {
        let mut v = Vec::new();
        for (idx, r) in o.ranks.iter().enumerate() {
            let row = 2 + idx;
            v.push(CellCheck::new(format!("rank_{}", idx + 1),
                format!("C{}", row), *r));
        }
        v.push(CellCheck::new("top1",   "F1", o.top1).with_tolerance(1e-6));
        v.push(CellCheck::new("top2",   "F2", o.top2).with_tolerance(1e-6));
        v.push(CellCheck::new("top3",   "F3", o.top3).with_tolerance(1e-6));
        v.push(CellCheck::new("bottom1","F4", o.bottom1).with_tolerance(1e-6));
        v.push(CellCheck::new("median", "F5", o.median).with_tolerance(1e-6));
        v.push(CellCheck::new("q3",     "F6", o.q3).with_tolerance(1e-3));
        v
    }
}

/// Excel-style PERCENTILE.INC: linear interpolation across the sorted
/// array. `sorted` is in descending order here; reverse internally.
fn percentile_inc(sorted_desc: &[f64], p: f64) -> f64 {
    let mut s: Vec<f64> = sorted_desc.to_vec();
    s.sort_by(|a, b| a.partial_cmp(b).unwrap());
    let n = s.len();
    if n == 0 { return f64::NAN; }
    if n == 1 { return s[0]; }
    let rank = p * (n - 1) as f64;
    let lo = rank.floor() as usize;
    let hi = rank.ceil() as usize;
    if lo == hi {
        s[lo]
    } else {
        let frac = rank - lo as f64;
        s[lo] * (1.0 - frac) + s[hi] * frac
    }
}
