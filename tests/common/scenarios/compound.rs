//! Compound interest projection: a recurring contribution at a fixed
//! periodic rate over N periods. Tests FV (future value of an annuity).
//!
//! ```text
//!     FV = P * (1+r)^n + PMT * ((1+r)^n - 1) / r
//! ```
//!
//! Plus a year-by-year balance series for cross-check.

use super::{lit, CellCheck, Scenario};
use crate::common::scenarios::{enter_cell, enter_cells, Harness};

pub struct Compound;

#[derive(Clone)]
pub struct Inputs {
    pub principal: f64,
    pub annual_rate: f64,
    pub years: usize,
    pub annual_contribution: f64,
}

pub struct Output {
    pub future_value: f64,
    pub balances: Vec<f64>, // balance at END of each year
}

impl Scenario for Compound {
    type Inputs = Inputs;
    type Output = Output;

    fn name(&self) -> &'static str { "compound" }

    fn default_inputs(&self) -> Inputs {
        Inputs {
            principal: 5_000.0,
            annual_rate: 0.07,
            years: 10,
            annual_contribution: 6_000.0,
        }
    }

    fn compute(&self, i: &Inputs) -> Output {
        let n = i.years as i32;
        let r = i.annual_rate;
        let fv_principal = i.principal * (1.0 + r).powi(n);
        let fv_annuity = i.annual_contribution * ((1.0 + r).powi(n) - 1.0) / r;
        let future_value = fv_principal + fv_annuity;

        let mut bal = i.principal;
        let mut series = Vec::with_capacity(i.years);
        for _ in 0..i.years {
            bal = bal * (1.0 + r) + i.annual_contribution;
            series.push(bal);
        }
        Output { future_value, balances: series }
    }

    fn populate(&self, h: &mut Harness, i: &Inputs) {
        enter_cells(h, &[
            ("A1", "principal"),    ("B1", &lit(i.principal)),
            ("A2", "rate"),         ("B2", &lit(i.annual_rate)),
            ("A3", "years"),        ("B3", &i.years.to_string()),
            ("A4", "contribution"), ("B4", &lit(i.annual_contribution)),
            // Excel-style FV with the saver's sign convention: PV and PMT
            // are cash OUT (negative), so FV returns a positive future
            // balance. =FV(rate, nper, -pmt, -pv) gives the same number
            // our pure-Rust closed form computes.
            ("A6", "FV"),  ("B6", "=FV(B2,B3,-B4,-B1)"),
            ("A8", "year"), ("B8", "balance"),
        ]);
        for y in 1..=i.years {
            let row = 8 + y;
            enter_cell(h, &format!("A{}", row), &y.to_string());
            let prev = if y == 1 {
                "$B$1".to_string()
            } else {
                format!("B{}", row - 1)
            };
            enter_cell(h, &format!("B{}", row),
                &format!("={}*(1+$B$2)+$B$4", prev));
        }
    }

    fn checks(&self, o: &Output) -> Vec<CellCheck> {
        let mut v = vec![
            CellCheck::new("FV_formula", "B6", o.future_value)
                .with_tolerance(1e-3),
        ];
        for (y, bal) in o.balances.iter().enumerate() {
            let row = 9 + y;
            v.push(CellCheck::new(format!("balance_y{}", y + 1),
                format!("B{}", row), *bal).with_tolerance(1e-4));
        }
        v
    }
}
