//! Stock portfolio: shares × price → position value, plus total and
//! per-position weight. Tests SUMPRODUCT, SUM, and percentage math.

use super::{lit, CellCheck, Scenario};
use crate::common::scenarios::{enter_cell, enter_cells, Harness};

pub struct Portfolio;

#[derive(Clone)]
pub struct Position {
    pub ticker: String,
    pub shares: f64,
    pub price: f64,
}

#[derive(Clone)]
pub struct Inputs {
    pub positions: Vec<Position>,
}

pub struct Output {
    pub position_values: Vec<f64>,
    pub total_value: f64,
    pub weights: Vec<f64>,
}

impl Scenario for Portfolio {
    type Inputs = Inputs;
    type Output = Output;

    fn name(&self) -> &'static str { "portfolio" }

    fn default_inputs(&self) -> Inputs {
        Inputs {
            positions: vec![
                Position { ticker: "AAPL".into(),  shares: 50.0,  price: 187.32 },
                Position { ticker: "MSFT".into(),  shares: 30.0,  price: 421.15 },
                Position { ticker: "GOOG".into(),  shares: 20.0,  price: 168.40 },
                Position { ticker: "NVDA".into(),  shares: 10.0,  price: 905.55 },
                Position { ticker: "BRK.B".into(), shares: 15.0,  price: 432.10 },
            ],
        }
    }

    fn compute(&self, i: &Inputs) -> Output {
        let values: Vec<f64> = i.positions.iter().map(|p| p.shares * p.price).collect();
        let total: f64 = values.iter().sum();
        let weights: Vec<f64> = values.iter().map(|v| v / total).collect();
        Output { position_values: values, total_value: total, weights }
    }

    fn populate(&self, h: &mut Harness, i: &Inputs) {
        // Headers + per-position rows.
        enter_cells(h, &[
            ("A1", "ticker"), ("B1", "shares"), ("C1", "price"),
            ("D1", "value"), ("E1", "weight"),
        ]);
        let n = i.positions.len();
        for (idx, p) in i.positions.iter().enumerate() {
            let row = 2 + idx;
            enter_cell(h, &format!("A{}", row), &p.ticker);
            enter_cell(h, &format!("B{}", row), &lit(p.shares));
            enter_cell(h, &format!("C{}", row), &lit(p.price));
            enter_cell(h, &format!("D{}", row), &format!("=B{}*C{}", row, row));
        }
        let last = 1 + n;
        // Cross-check two ways: SUM over D, and SUMPRODUCT(B, C).
        enter_cells(h, &[
            ("A8", "total_sum"),
            ("A9", "total_sumproduct"),
        ]);
        enter_cell(h, "B8", &format!("=SUM(D2:D{})", last));
        enter_cell(h, "B9", &format!("=SUMPRODUCT(B2:B{},C2:C{})", last, last));
        // Weights per row.
        for idx in 0..n {
            let row = 2 + idx;
            enter_cell(h, &format!("E{}", row), &format!("=D{}/$B$8", row));
        }
    }

    fn checks(&self, o: &Output) -> Vec<CellCheck> {
        let mut v = vec![
            CellCheck::new("total_sum",        "B8", o.total_value).with_tolerance(1e-3),
            CellCheck::new("total_sumproduct", "B9", o.total_value).with_tolerance(1e-3),
        ];
        for (idx, val) in o.position_values.iter().enumerate() {
            let row = 2 + idx;
            v.push(CellCheck::new(format!("value_{}", idx + 1),
                format!("D{}", row), *val).with_tolerance(1e-4));
        }
        for (idx, w) in o.weights.iter().enumerate() {
            let row = 2 + idx;
            v.push(CellCheck::new(format!("weight_{}", idx + 1),
                format!("E{}", row), *w).with_tolerance(1e-6));
        }
        v
    }
}
