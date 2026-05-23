//! Progressive income tax with brackets. The marginal-rate calculation
//! is a classic source of subtle off-by-one bugs (does the threshold
//! itself fall in the lower bracket or the upper?).
//!
//! Uses the simplified 2024 US single-filer brackets (rounded thresholds
//! for round numbers in the test):
//!
//! ```text
//!     0 – 11000       → 10%
//!     11000 – 44725   → 12%
//!     44725 – 95375   → 22%
//!     95375 – 182100  → 24%
//!     182100 – 231250 → 32%
//!     231250 – 578125 → 35%
//!     > 578125        → 37%
//! ```

use super::{lit, CellCheck, Scenario};
use crate::common::scenarios::{enter_cell, enter_cells, Harness};

pub struct Tax;

#[derive(Clone)]
pub struct Inputs {
    pub incomes: Vec<f64>,
}

pub struct Output {
    pub tax_owed: Vec<f64>,
    pub effective_rates: Vec<f64>,
}

const BRACKETS: &[(f64, f64)] = &[
    (11_000.0,  0.10),
    (44_725.0,  0.12),
    (95_375.0,  0.22),
    (182_100.0, 0.24),
    (231_250.0, 0.32),
    (578_125.0, 0.35),
    (f64::INFINITY, 0.37),
];

fn tax_owed(income: f64) -> f64 {
    let mut owed = 0.0;
    let mut prev = 0.0;
    for &(ceiling, rate) in BRACKETS {
        if income <= ceiling {
            owed += (income - prev) * rate;
            return owed;
        }
        owed += (ceiling - prev) * rate;
        prev = ceiling;
    }
    owed
}

impl Scenario for Tax {
    type Inputs = Inputs;
    type Output = Output;

    fn name(&self) -> &'static str { "tax" }

    fn default_inputs(&self) -> Inputs {
        Inputs {
            incomes: vec![
                8_000.0,    // entirely in 10% bracket
                25_000.0,   // spans 10% and 12%
                60_000.0,   // spans 10/12/22
                150_000.0,  // spans 10/12/22/24
                250_000.0,  // spans through 32%
                600_000.0,  // top bracket
            ],
        }
    }

    fn compute(&self, i: &Inputs) -> Output {
        let tax: Vec<f64> = i.incomes.iter().map(|&inc| tax_owed(inc)).collect();
        let rates: Vec<f64> = i.incomes
            .iter()
            .zip(&tax)
            .map(|(inc, t)| if *inc > 0.0 { t / inc } else { 0.0 })
            .collect();
        Output { tax_owed: tax, effective_rates: rates }
    }

    fn populate(&self, h: &mut Harness, i: &Inputs) {
        enter_cells(h, &[
            ("A1", "income"),
            ("B1", "tax_owed"),
            ("C1", "effective_rate"),
        ]);
        // Encode brackets as an IFS chain. We compute the tax as
        // (bracket-base) + (income - bracket_floor) * marginal_rate.
        // The bracket bases are pre-computed cumulative amounts:
        //   base_10 = 0
        //   base_12 = 11000 * 0.10                    = 1100
        //   base_22 = 1100 + (44725 - 11000) * 0.12   = 5147
        //   base_24 = 5147 + (95375 - 44725) * 0.22   = 16290
        //   base_32 = 16290 + (182100 - 95375) * 0.24 = 37104
        //   base_35 = 37104 + (231250 - 182100) * 0.32= 52832
        //   base_37 = 52832 + (578125 - 231250) * 0.35= 174238.25
        let formula = "=IF(A{r}<=11000,A{r}*0.1,\
                        IF(A{r}<=44725,1100+(A{r}-11000)*0.12,\
                        IF(A{r}<=95375,5147+(A{r}-44725)*0.22,\
                        IF(A{r}<=182100,16290+(A{r}-95375)*0.24,\
                        IF(A{r}<=231250,37104+(A{r}-182100)*0.32,\
                        IF(A{r}<=578125,52832+(A{r}-231250)*0.35,\
                        174238.25+(A{r}-578125)*0.37))))))";
        for (idx, inc) in i.incomes.iter().enumerate() {
            let row = 2 + idx;
            enter_cell(h, &format!("A{}", row), &lit(*inc));
            enter_cell(h, &format!("B{}", row),
                &formula.replace("{r}", &row.to_string()));
            enter_cell(h, &format!("C{}", row),
                &format!("=B{}/A{}", row, row));
        }
    }

    fn checks(&self, o: &Output) -> Vec<CellCheck> {
        let mut v = Vec::new();
        for (idx, (tax, rate)) in
            o.tax_owed.iter().zip(&o.effective_rates).enumerate()
        {
            let row = 2 + idx;
            v.push(CellCheck::new(format!("tax_{}", idx + 1),
                format!("B{}", row), *tax).with_tolerance(1e-3));
            v.push(CellCheck::new(format!("rate_{}", idx + 1),
                format!("C{}", row), *rate).with_tolerance(1e-6));
        }
        v
    }
}
