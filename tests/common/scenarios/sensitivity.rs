//! Two-axis sensitivity table — output as f(input_x, input_y). Each cell
//! in the body of the table references its row-header and column-header
//! values from the axes. Stresses absolute/relative ref mixing.
//!
//! Example: NPV of a simple cash-flow stream as a function of (discount
//! rate, growth rate).

use super::{lit, CellCheck, Scenario};
use crate::common::scenarios::{enter_cell, enter_cells, Harness};

pub struct Sensitivity;

#[derive(Clone)]
pub struct Inputs {
    pub initial_cf: f64,
    pub years: usize,
    pub rates: Vec<f64>,    // column axis (discount rates)
    pub growths: Vec<f64>,  // row axis (growth rates)
}

pub struct Output {
    pub grid: Vec<Vec<f64>>, // grid[row=growth][col=rate]
}

fn npv(initial_cf: f64, growth: f64, rate: f64, years: usize) -> f64 {
    let mut sum = 0.0;
    for t in 1..=years {
        let cf = initial_cf * (1.0 + growth).powi(t as i32);
        sum += cf / (1.0 + rate).powi(t as i32);
    }
    sum
}

impl Scenario for Sensitivity {
    type Inputs = Inputs;
    type Output = Output;

    fn name(&self) -> &'static str { "sensitivity" }

    fn default_inputs(&self) -> Inputs {
        Inputs {
            initial_cf: 100.0,
            years: 5,
            rates:   vec![0.05, 0.08, 0.10, 0.12],
            growths: vec![0.02, 0.04, 0.06],
        }
    }

    fn compute(&self, i: &Inputs) -> Output {
        let grid: Vec<Vec<f64>> = i.growths.iter().map(|&g| {
            i.rates.iter().map(|&r| npv(i.initial_cf, g, r, i.years)).collect()
        }).collect();
        Output { grid }
    }

    fn populate(&self, h: &mut Harness, i: &Inputs) {
        // Layout:
        //   B1 = initial_cf, B2 = years
        //   Row 4 col-headers (rates): C4..F4
        //   Col B row-headers (growths): B5..B7
        //   Body: C5..F7 = NPV with row/col axes
        enter_cells(h, &[
            ("A1", "initial_cf"), ("B1", &lit(i.initial_cf)),
            ("A2", "years"),      ("B2", &i.years.to_string()),
            ("A4", "g \\ r"),
        ]);
        // Column headers (rates) in row 4.
        for (j, r) in i.rates.iter().enumerate() {
            let col = column_label(2 + j);
            enter_cell(h, &format!("{}{}", col, 4), &lit(*r));
        }
        // Row headers (growths) in col B.
        for (k, g) in i.growths.iter().enumerate() {
            enter_cell(h, &format!("B{}", 5 + k), &lit(*g));
        }
        // Body cells. Each is the closed-form NPV for the geometric CF
        // stream, expressed in formula:
        //   sum_{t=1..N} cf0 * (1+g)^t / (1+r)^t
        // We expand as a SUMPRODUCT-free explicit sum via a helper column
        // — but to keep the formula self-contained inside the body cell
        // we use the closed-form:
        //   NPV = cf0 * x * (1 - x^N) / (1 - x)   where x = (1+g)/(1+r)
        //
        // We can't write that as a single tshts formula without IF for
        // the x==1 case, but g != r in our default inputs, so the simple
        // form suffices.
        for k in 0..i.growths.len() {
            for j in 0..i.rates.len() {
                let row = 5 + k;
                let col = column_label(2 + j);
                let row_header = format!("$B{}", row);          // growth (col absolute)
                let col_header = format!("{}$4", col);          // rate (row absolute)
                // x = (1 + g) / (1 + r)
                let formula = format!(
                    "=$B$1*(1+{g})/(1+{r})*(1-POWER((1+{g})/(1+{r}),$B$2))/(1-(1+{g})/(1+{r}))",
                    g = row_header, r = col_header,
                );
                enter_cell(h, &format!("{}{}", col, row), &formula);
            }
        }
    }

    fn checks(&self, o: &Output) -> Vec<CellCheck> {
        let mut v = Vec::new();
        for (k, row) in o.grid.iter().enumerate() {
            for (j, val) in row.iter().enumerate() {
                let r = 5 + k;
                let col = column_label(2 + j);
                v.push(CellCheck::new(
                    format!("npv_g{}_r{}", k + 1, j + 1),
                    format!("{}{}", col, r),
                    *val,
                ).with_tolerance(1e-4));
            }
        }
        v
    }
}

fn column_label(col: usize) -> String {
    // 0-indexed. 0 → A, 25 → Z, 26 → AA, ...
    let mut s = String::new();
    let mut n = col + 1;
    while n > 0 {
        n -= 1;
        s.insert(0, (b'A' + (n % 26) as u8) as char);
        n /= 26;
    }
    s
}
