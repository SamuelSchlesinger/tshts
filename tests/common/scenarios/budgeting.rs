//! Personal/household budget — monthly cash flow with a running balance.
//!
//! Exercises:
//!   * Per-row arithmetic (net = income - expenses)
//!   * A cumulative chain (balance_t = balance_{t-1} + net_t) — i.e.
//!     each cell depends on the previous row, which is what stresses
//!     the propagation cascade.
//!   * SUM across a row range (annual totals).

use super::{lit, CellCheck, Scenario};
use crate::common::scenarios::{enter_cell, enter_cells, Harness};

pub struct Budgeting;

#[derive(Clone)]
pub struct Inputs {
    pub starting_balance: f64,
    pub monthly_income: Vec<f64>,
    pub monthly_expenses: Vec<f64>,
}

pub struct Output {
    pub net_per_month: Vec<f64>,
    pub running_balance: Vec<f64>,  // balance at END of each month
    pub annual_income: f64,
    pub annual_expense: f64,
    pub annual_net: f64,
    pub year_end_balance: f64,
}

impl Scenario for Budgeting {
    type Inputs = Inputs;
    type Output = Output;

    fn name(&self) -> &'static str { "budgeting" }

    fn default_inputs(&self) -> Inputs {
        // Salaried household: $7800/mo, with a $4000 bonus in December.
        // Expenses bumpy: holidays in Nov/Dec, lower in summer.
        let mut income = vec![7800.0; 12];
        income[11] += 4000.0; // December bonus
        let expenses = vec![
            6500.0, 6200.0, 6300.0, 6100.0, 5900.0, 5700.0,
            5800.0, 6000.0, 6100.0, 6400.0, 7200.0, 8100.0,
        ];
        Inputs {
            starting_balance: 12_000.0,
            monthly_income: income,
            monthly_expenses: expenses,
        }
    }

    fn compute(&self, i: &Inputs) -> Output {
        let n = i.monthly_income.len().min(i.monthly_expenses.len());
        let mut net = Vec::with_capacity(n);
        let mut bal = Vec::with_capacity(n);
        let mut running = i.starting_balance;
        for m in 0..n {
            let nv = i.monthly_income[m] - i.monthly_expenses[m];
            net.push(nv);
            running += nv;
            bal.push(running);
        }
        let annual_income: f64 = i.monthly_income.iter().sum();
        let annual_expense: f64 = i.monthly_expenses.iter().sum();
        Output {
            net_per_month: net,
            running_balance: bal.clone(),
            annual_income,
            annual_expense,
            annual_net: annual_income - annual_expense,
            year_end_balance: *bal.last().unwrap_or(&i.starting_balance),
        }
    }

    fn populate(&self, h: &mut Harness, i: &Inputs) {
        // Layout:
        //   B1 = starting balance
        //   A3 header. Row 4 = column labels (Month, Income, Expense, Net, Balance)
        //   Rows 5..16 = months 1..12
        //   Row 18 = annual totals
        enter_cells(h, &[
            ("A1", "starting_balance"), ("B1", &lit(i.starting_balance)),
            ("A4", "Month"), ("B4", "Income"), ("C4", "Expense"),
            ("D4", "Net"), ("E4", "Balance"),
        ]);
        let n = i.monthly_income.len();
        for m in 0..n {
            let row = 5 + m;
            enter_cell(h, &format!("A{}", row), &(m + 1).to_string());
            enter_cell(h, &format!("B{}", row), &lit(i.monthly_income[m]));
            enter_cell(h, &format!("C{}", row), &lit(i.monthly_expenses[m]));
            enter_cell(h, &format!("D{}", row), &format!("=B{}-C{}", row, row));
            if m == 0 {
                enter_cell(h, &format!("E{}", row), &format!("=B1+D{}", row));
            } else {
                let prev = row - 1;
                enter_cell(h, &format!("E{}", row), &format!("=E{}+D{}", prev, row));
            }
        }
        let last = 4 + n;
        enter_cells(h, &[
            ("A18", "Annual"),
        ]);
        enter_cell(h, "B18", &format!("=SUM(B5:B{})", last));
        enter_cell(h, "C18", &format!("=SUM(C5:C{})", last));
        enter_cell(h, "D18", &format!("=SUM(D5:D{})", last));
        enter_cell(h, "E18", &format!("=E{}", last)); // year-end balance
    }

    fn checks(&self, o: &Output) -> Vec<CellCheck> {
        let mut v = Vec::new();
        // Check every month's net and running balance — this stresses the
        // cumulative-dependency chain end-to-end.
        for (m, (net, bal)) in o.net_per_month.iter().zip(&o.running_balance).enumerate() {
            let row = 5 + m;
            v.push(CellCheck::new(format!("net_m{}", m + 1),
                format!("D{}", row), *net));
            v.push(CellCheck::new(format!("balance_m{}", m + 1),
                format!("E{}", row), *bal));
        }
        v.push(CellCheck::new("annual_income", "B18", o.annual_income));
        v.push(CellCheck::new("annual_expense", "C18", o.annual_expense));
        v.push(CellCheck::new("annual_net", "D18", o.annual_net));
        v.push(CellCheck::new("year_end_balance", "E18", o.year_end_balance));
        v
    }
}
