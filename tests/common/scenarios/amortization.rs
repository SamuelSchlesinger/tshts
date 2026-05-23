//! Loan amortization schedule. PMT splits each payment into interest
//! and principal; balance evolves period-by-period.
//!
//! Standard mortgage / installment-loan math:
//!
//! ```text
//!     monthly_rate = annual_rate / 12
//!     payment      = principal * r / (1 - (1 + r)^-n)
//!     For period t:
//!         interest_t  = balance_{t-1} * r
//!         principal_t = payment - interest_t
//!         balance_t   = balance_{t-1} - principal_t
//! ```
//!
//! Stresses PMT, the recurrence, and the convention that final balance
//! must be ~0 after n payments.

use super::{lit, CellCheck, Scenario};
use crate::common::scenarios::{enter_cell, enter_cells, Harness};

pub struct Amortization;

#[derive(Clone)]
pub struct Inputs {
    pub principal: f64,
    pub annual_rate: f64,
    pub years: usize,
}

pub struct Output {
    pub payment: f64,
    pub interest_per_period: Vec<f64>,
    pub principal_per_period: Vec<f64>,
    pub balance_per_period: Vec<f64>,
    pub total_interest: f64,
}

impl Scenario for Amortization {
    type Inputs = Inputs;
    type Output = Output;

    fn name(&self) -> &'static str { "amortization" }

    fn default_inputs(&self) -> Inputs {
        // Small consumer loan: $25k over 3 years (36 months) at 8.5%.
        // Short enough to fully tabulate, long enough to be interesting.
        Inputs { principal: 25_000.0, annual_rate: 0.085, years: 3 }
    }

    fn compute(&self, i: &Inputs) -> Output {
        let n = (i.years * 12) as i32;
        let r = i.annual_rate / 12.0;
        // PMT (Excel sign convention is negative for an outflow; we use
        // the absolute payment amount throughout, like a consumer would).
        let payment = i.principal * r / (1.0 - (1.0 + r).powi(-n));
        let mut interest = Vec::with_capacity(n as usize);
        let mut principal = Vec::with_capacity(n as usize);
        let mut balance = Vec::with_capacity(n as usize);
        let mut bal = i.principal;
        for _ in 0..n {
            let it = bal * r;
            let pp = payment - it;
            bal -= pp;
            interest.push(it);
            principal.push(pp);
            balance.push(bal);
        }
        let total_interest = interest.iter().sum();
        Output {
            payment,
            interest_per_period: interest,
            principal_per_period: principal,
            balance_per_period: balance,
            total_interest,
        }
    }

    fn populate(&self, h: &mut Harness, i: &Inputs) {
        let n = i.years * 12;
        // Inputs
        enter_cells(h, &[
            ("A1", "principal"),  ("B1", &lit(i.principal)),
            ("A2", "annual_rate"),("B2", &lit(i.annual_rate)),
            ("A3", "years"),      ("B3", &i.years.to_string()),
            ("A4", "monthly_r"),  ("B4", "=B2/12"),
            ("A5", "n_periods"),  ("B5", "=B3*12"),
            // PMT(rate, nper, pv) returns a NEGATIVE payment in Excel
            // convention. Negate so it matches the consumer-facing
            // positive number our Rust compute uses.
            ("A6", "payment"),    ("B6", "=-PMT(B4,B5,B1)"),
            // Schedule header (row 8): period | interest | principal | balance
            ("A8", "period"), ("B8", "interest"), ("C8", "principal"), ("D8", "balance"),
        ]);
        // Schedule rows start at row 9. Use B6 as the constant payment
        // and chain D{prev} → D{curr} for the balance recurrence.
        for t in 1..=n {
            let row = 8 + t;
            enter_cell(h, &format!("A{}", row), &t.to_string());
            // interest_t = balance_{t-1} * B4 (monthly rate)
            let prev_bal = if t == 1 {
                "$B$1".to_string()
            } else {
                format!("D{}", row - 1)
            };
            enter_cell(h, &format!("B{}", row),
                &format!("={}*$B$4", prev_bal));
            enter_cell(h, &format!("C{}", row),
                &format!("=$B$6-B{}", row));
            enter_cell(h, &format!("D{}", row),
                &format!("={}-C{}", prev_bal, row));
        }
        let last_row = 8 + n;
        enter_cells(h, &[("A6", "payment")]);
        enter_cell(h, &format!("F1"), &format!("=SUM(B9:B{})", last_row));
        // F1 = total interest paid.
    }

    fn checks(&self, o: &Output) -> Vec<CellCheck> {
        let mut v = Vec::new();
        // PMT.
        v.push(CellCheck::new("payment", "B6", o.payment).with_tolerance(1e-4));
        // Spot-check a sparse set of periods rather than all 36 — keeps the
        // status-bar reads to ~15s instead of ~60s. Still hits the recurrence
        // at the start, middle, and end so a propagation bug shows up.
        let n = o.interest_per_period.len();
        let probe_periods = [1, 2, 3, n / 4, n / 2, (3 * n) / 4, n - 1, n];
        for &t in &probe_periods {
            if t == 0 || t > n { continue; }
            let row = 8 + t;
            v.push(CellCheck::new(format!("interest_p{}", t),
                format!("B{}", row), o.interest_per_period[t - 1])
                .with_tolerance(1e-4));
            v.push(CellCheck::new(format!("principal_p{}", t),
                format!("C{}", row), o.principal_per_period[t - 1])
                .with_tolerance(1e-4));
            v.push(CellCheck::new(format!("balance_p{}", t),
                format!("D{}", row), o.balance_per_period[t - 1])
                .with_tolerance(1e-4));
        }
        v.push(CellCheck::new("total_interest", "F1", o.total_interest)
            .with_tolerance(1e-2));
        v
    }
}
