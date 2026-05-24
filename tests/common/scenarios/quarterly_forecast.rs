//! Quarterly forecast roll-up. Monthly bookings get aggregated by
//! quarter via SUMIF over a quarter-tag column. Tests SUMIF with
//! string criteria + a derived bucket column.

use super::{lit, CellCheck, Scenario};
use crate::common::scenarios::{enter_cell, enter_cells, Harness};

pub struct QuarterlyForecast;

#[derive(Clone)]
pub struct MonthlyBooking {
    pub month: u8, // 1..=12
    pub amount: f64,
}

#[derive(Clone)]
pub struct Inputs {
    pub bookings: Vec<MonthlyBooking>,
}

pub struct Output {
    pub quarter_totals: [f64; 4],
    pub annual_total: f64,
    pub max_quarter: f64,
    pub avg_quarter: f64,
}

fn quarter_of(month: u8) -> u8 {
    match month {
        1..=3 => 1,
        4..=6 => 2,
        7..=9 => 3,
        10..=12 => 4,
        _ => 0,
    }
}

impl Scenario for QuarterlyForecast {
    type Inputs = Inputs;
    type Output = Output;

    fn name(&self) -> &'static str { "quarterly_forecast" }

    fn default_inputs(&self) -> Inputs {
        // A realistic-ish SaaS booking pattern with some seasonality.
        Inputs {
            bookings: (1..=12).map(|m| {
                let base = 50_000.0;
                let seasonality = match m {
                    3 | 6 | 9 | 12 => 1.5, // end-of-quarter push
                    1 => 0.7,              // post-holiday slump
                    _ => 1.0,
                };
                MonthlyBooking { month: m, amount: base * seasonality }
            }).collect(),
        }
    }

    fn compute(&self, i: &Inputs) -> Output {
        let mut quarter_totals = [0.0; 4];
        for b in &i.bookings {
            let q = quarter_of(b.month);
            if (1..=4).contains(&q) {
                quarter_totals[(q - 1) as usize] += b.amount;
            }
        }
        let annual_total: f64 = quarter_totals.iter().sum();
        let max_quarter = quarter_totals.iter().copied().fold(f64::NEG_INFINITY, f64::max);
        let avg_quarter = annual_total / 4.0;
        Output { quarter_totals, annual_total, max_quarter, avg_quarter }
    }

    fn populate(&self, h: &mut Harness, i: &Inputs) {
        // Layout:
        //   A: month (1..12) | B: amount | C: quarter (1..4) — derived
        // Rows 2..13 = months.
        enter_cells(h, &[
            ("A1", "month"), ("B1", "amount"), ("C1", "quarter"),
        ]);
        for (idx, b) in i.bookings.iter().enumerate() {
            let row = 2 + idx;
            enter_cell(h, &format!("A{}", row), &b.month.to_string());
            enter_cell(h, &format!("B{}", row), &lit(b.amount));
            // Quarter = INT((month - 1) / 3) + 1 — cleaner via nested IF
            // for the 4-bucket case to exercise the IF chain too.
            enter_cell(h, &format!("C{}", row),
                &format!("=IF(A{r}<=3,1,IF(A{r}<=6,2,IF(A{r}<=9,3,4)))", r = row));
        }
        let last = 1 + i.bookings.len();
        // Quarter totals via SUMIF on the derived quarter column.
        enter_cells(h, &[
            ("E1", "Q1"), ("E2", "Q2"), ("E3", "Q3"), ("E4", "Q4"),
            ("G1", "annual"), ("G2", "max_q"), ("G3", "avg_q"),
        ]);
        for q in 1..=4 {
            enter_cell(h, &format!("F{}", q),
                &format!("=SUMIF(C2:C{l},{q},B2:B{l})", l = last, q = q));
        }
        enter_cell(h, "H1", "=SUM(F1:F4)");
        enter_cell(h, "H2", "=MAX(F1:F4)");
        enter_cell(h, "H3", "=AVERAGE(F1:F4)");
    }

    fn checks(&self, o: &Output) -> Vec<CellCheck> {
        let mut v = Vec::new();
        for (i, total) in o.quarter_totals.iter().enumerate() {
            v.push(CellCheck::new(format!("Q{}", i + 1),
                format!("F{}", i + 1), *total).with_tolerance(1e-4));
        }
        v.push(CellCheck::new("annual",   "H1", o.annual_total).with_tolerance(1e-4));
        v.push(CellCheck::new("max_q",    "H2", o.max_quarter).with_tolerance(1e-4));
        v.push(CellCheck::new("avg_q",    "H3", o.avg_quarter).with_tolerance(1e-4));
        v
    }
}
