//! Statistical regression. Given (x, y) pairs, compute SLOPE, INTERCEPT,
//! R² (RSQ), and the regression's predicted y for each x. Predicted y
//! per row exercises the formula `=SLOPE*X + INTERCEPT` cascading
//! across the dataset.

use super::{lit, CellCheck, Scenario};
use crate::common::scenarios::{enter_cell, enter_cells, Harness};

pub struct Regression;

#[derive(Clone)]
pub struct Inputs {
    pub xs: Vec<f64>,
    pub ys: Vec<f64>,
}

pub struct Output {
    pub slope: f64,
    pub intercept: f64,
    pub r_squared: f64,
    pub predictions: Vec<f64>,
    pub sum_squared_residuals: f64,
}

fn slope_intercept(xs: &[f64], ys: &[f64]) -> (f64, f64) {
    let n = xs.len() as f64;
    let sx: f64 = xs.iter().sum();
    let sy: f64 = ys.iter().sum();
    let sxx: f64 = xs.iter().map(|x| x * x).sum();
    let sxy: f64 = xs.iter().zip(ys).map(|(x, y)| x * y).sum();
    let denom = n * sxx - sx * sx;
    let slope = (n * sxy - sx * sy) / denom;
    let intercept = (sy - slope * sx) / n;
    (slope, intercept)
}

fn r_squared(xs: &[f64], ys: &[f64]) -> f64 {
    let n = xs.len() as f64;
    let mean_x: f64 = xs.iter().sum::<f64>() / n;
    let mean_y: f64 = ys.iter().sum::<f64>() / n;
    let cov: f64 = xs.iter().zip(ys)
        .map(|(x, y)| (x - mean_x) * (y - mean_y))
        .sum();
    let var_x: f64 = xs.iter().map(|x| (x - mean_x).powi(2)).sum();
    let var_y: f64 = ys.iter().map(|y| (y - mean_y).powi(2)).sum();
    let r = cov / (var_x.sqrt() * var_y.sqrt());
    r * r
}

impl Scenario for Regression {
    type Inputs = Inputs;
    type Output = Output;

    fn name(&self) -> &'static str { "regression" }

    fn default_inputs(&self) -> Inputs {
        // Strongly linear: y ~ 2.5x + 1 with small noise.
        Inputs {
            xs: (1..=10).map(|i| i as f64).collect(),
            ys: vec![3.4, 6.1, 8.6, 10.9, 13.4, 16.2, 18.7, 21.4, 23.8, 26.5],
        }
    }

    fn compute(&self, i: &Inputs) -> Output {
        let (slope, intercept) = slope_intercept(&i.xs, &i.ys);
        let r_squared = r_squared(&i.xs, &i.ys);
        let predictions: Vec<f64> = i.xs.iter().map(|x| slope * x + intercept).collect();
        let sum_squared_residuals: f64 = predictions.iter().zip(&i.ys)
            .map(|(p, y)| (p - y).powi(2))
            .sum();
        Output { slope, intercept, r_squared, predictions, sum_squared_residuals }
    }

    fn populate(&self, h: &mut Harness, i: &Inputs) {
        // A: x | B: y | C: predicted y (=slope*A + intercept)
        // D: residual (= predicted - y), D^2 in column E for SSR
        enter_cells(h, &[
            ("A1", "x"), ("B1", "y"), ("C1", "y_hat"), ("D1", "resid"), ("E1", "resid2"),
        ]);
        for (idx, (x, y)) in i.xs.iter().zip(&i.ys).enumerate() {
            let row = 2 + idx;
            enter_cell(h, &format!("A{}", row), &lit(*x));
            enter_cell(h, &format!("B{}", row), &lit(*y));
        }
        let last = 1 + i.xs.len();
        let xr = format!("A2:A{}", last);
        let yr = format!("B2:B{}", last);
        // Summary cells.
        enter_cells(h, &[
            ("G1", "slope"), ("G2", "intercept"), ("G3", "r_squared"),
            ("G4", "ssr"),
        ]);
        // SLOPE and INTERCEPT take ys first, xs second (Excel convention).
        enter_cell(h, "H1", &format!("=SLOPE({},{})", yr, xr));
        enter_cell(h, "H2", &format!("=INTERCEPT({},{})", yr, xr));
        // RSQ also takes ys first, xs second.
        enter_cell(h, "H3", &format!("=RSQ({},{})", yr, xr));
        // Per-row predictions using $H$1 / $H$2.
        for idx in 0..i.xs.len() {
            let row = 2 + idx;
            enter_cell(h, &format!("C{}", row),
                &format!("=$H$1*A{}+$H$2", row));
            enter_cell(h, &format!("D{}", row), &format!("=C{}-B{}", row, row));
            enter_cell(h, &format!("E{}", row), &format!("=D{}*D{}", row, row));
        }
        enter_cell(h, "H4", &format!("=SUM(E2:E{})", last));
    }

    fn checks(&self, o: &Output) -> Vec<CellCheck> {
        let mut v = vec![
            CellCheck::new("slope",     "H1", o.slope).with_tolerance(1e-6),
            CellCheck::new("intercept", "H2", o.intercept).with_tolerance(1e-6),
            CellCheck::new("r_squared", "H3", o.r_squared).with_tolerance(1e-6),
            CellCheck::new("ssr",       "H4", o.sum_squared_residuals).with_tolerance(1e-3),
        ];
        for (idx, p) in o.predictions.iter().enumerate() {
            let row = 2 + idx;
            v.push(CellCheck::new(format!("y_hat_{}", idx + 1),
                format!("C{}", row), *p).with_tolerance(1e-6));
        }
        v
    }
}
