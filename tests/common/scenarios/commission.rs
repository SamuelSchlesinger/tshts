//! Tiered sales commission. A rep earns different commission rates on
//! sales above each threshold. Tests nested `IF` (or `IFS`) chains, which
//! are notoriously easy to write incorrectly.
//!
//! Tier rules (cumulative):
//!   sales ≤ 50k       → 5%
//!   50k < sales ≤ 100k → 8% on the slice above 50k (plus tier-1 amount)
//!   sales > 100k       → 12% on the slice above 100k (plus tiers 1+2)

use super::{lit, CellCheck, Scenario};
use crate::common::scenarios::{enter_cell, enter_cells, Harness};

pub struct Commission;

#[derive(Clone)]
pub struct Inputs {
    /// One sales figure per rep to evaluate.
    pub sales: Vec<f64>,
}

pub struct Output {
    pub commissions: Vec<f64>,
}

fn commission_of(sales: f64) -> f64 {
    if sales <= 50_000.0 {
        sales * 0.05
    } else if sales <= 100_000.0 {
        50_000.0 * 0.05 + (sales - 50_000.0) * 0.08
    } else {
        50_000.0 * 0.05 + 50_000.0 * 0.08 + (sales - 100_000.0) * 0.12
    }
}

impl Scenario for Commission {
    type Inputs = Inputs;
    type Output = Output;

    fn name(&self) -> &'static str { "commission" }

    fn default_inputs(&self) -> Inputs {
        // Cover all three tiers including the exact tier boundaries.
        Inputs {
            sales: vec![
                15_000.0,    // tier 1
                49_999.99,   // just under tier 1 ceiling
                50_000.0,    // exact boundary
                75_000.0,    // mid tier 2
                100_000.0,   // exact boundary
                150_000.0,   // tier 3
                275_000.0,   // high tier 3
            ],
        }
    }

    fn compute(&self, i: &Inputs) -> Output {
        Output { commissions: i.sales.iter().map(|&s| commission_of(s)).collect() }
    }

    fn populate(&self, h: &mut Harness, i: &Inputs) {
        enter_cells(h, &[
            ("A1", "sales"), ("B1", "commission"),
        ]);
        // The formula encodes the cumulative tier semantics. Note that
        // we use `<=` on the boundaries so the boundary itself stays in
        // the lower tier — matching commission_of() above.
        let formula =
            "=IF(A{r}<=50000,A{r}*0.05,\
                 IF(A{r}<=100000,2500+(A{r}-50000)*0.08,\
                 2500+4000+(A{r}-100000)*0.12))";
        for (idx, s) in i.sales.iter().enumerate() {
            let row = 2 + idx;
            enter_cell(h, &format!("A{}", row), &lit(*s));
            enter_cell(h, &format!("B{}", row),
                &formula.replace("{r}", &row.to_string()));
        }
    }

    fn checks(&self, o: &Output) -> Vec<CellCheck> {
        o.commissions
            .iter()
            .enumerate()
            .map(|(idx, c)| {
                let row = 2 + idx;
                CellCheck::new(format!("commission_{}", idx + 1),
                    format!("B{}", row), *c).with_tolerance(1e-4)
            })
            .collect()
    }
}
