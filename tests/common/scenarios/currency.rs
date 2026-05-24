//! Multi-currency translation. Subsidiary balance sheets in local
//! currency get translated to USD via an FX-rate table. Tests
//! VLOOKUP for the rate, a ratio cascade across cells, and a SUM
//! aggregation.

use super::{lit, CellCheck, Scenario};
use crate::common::scenarios::{enter_cell, enter_cells, Harness};

pub struct Currency;

#[derive(Clone)]
pub struct FxRow {
    pub ccy: &'static str,
    /// Units of THIS currency per 1 USD (e.g. EUR=0.91 → 0.91 EUR per 1 USD).
    pub per_usd: f64,
}

#[derive(Clone)]
pub struct Sub {
    pub name: &'static str,
    pub ccy: &'static str,
    pub local_balance: f64,
}

#[derive(Clone)]
pub struct Inputs {
    pub fx: Vec<FxRow>,
    pub subs: Vec<Sub>,
}

pub struct Output {
    pub rates: Vec<f64>,        // looked-up per-USD for each sub
    pub usd_balances: Vec<f64>, // local / per_usd
    pub group_total_usd: f64,
}

impl Scenario for Currency {
    type Inputs = Inputs;
    type Output = Output;

    fn name(&self) -> &'static str { "currency" }

    fn default_inputs(&self) -> Inputs {
        Inputs {
            fx: vec![
                FxRow { ccy: "USD", per_usd: 1.00 },
                FxRow { ccy: "EUR", per_usd: 0.91 },
                FxRow { ccy: "GBP", per_usd: 0.78 },
                FxRow { ccy: "JPY", per_usd: 149.50 },
                FxRow { ccy: "INR", per_usd: 83.20 },
            ],
            subs: vec![
                Sub { name: "NA",     ccy: "USD", local_balance:  1_250_000.0 },
                Sub { name: "EU",     ccy: "EUR", local_balance:    910_000.0 },
                Sub { name: "UK",     ccy: "GBP", local_balance:    624_000.0 },
                Sub { name: "JP",     ccy: "JPY", local_balance: 74_750_000.0 },
                Sub { name: "IN",     ccy: "INR", local_balance: 41_600_000.0 },
            ],
        }
    }

    fn compute(&self, i: &Inputs) -> Output {
        let lookup = |ccy: &str| -> f64 {
            i.fx.iter().find(|r| r.ccy == ccy).map(|r| r.per_usd).unwrap_or(f64::NAN)
        };
        let rates: Vec<f64> = i.subs.iter().map(|s| lookup(s.ccy)).collect();
        let usd_balances: Vec<f64> = i.subs.iter()
            .map(|s| s.local_balance / lookup(s.ccy))
            .collect();
        let group_total_usd: f64 = usd_balances.iter().sum();
        Output { rates, usd_balances, group_total_usd }
    }

    fn populate(&self, h: &mut Harness, i: &Inputs) {
        // FX table at A1:B(n+1).
        enter_cells(h, &[
            ("A1", "ccy"), ("B1", "per_usd"),
        ]);
        for (idx, r) in i.fx.iter().enumerate() {
            let row = 2 + idx;
            enter_cell(h, &format!("A{}", row), r.ccy);
            enter_cell(h, &format!("B{}", row), &lit(r.per_usd));
        }
        let last_fx = 1 + i.fx.len();
        // Use fully-absolute refs so VLOOKUP's range stays pinned as
        // the formula is copied down.
        let fx_range = format!("$A$2:$B${}", last_fx);

        // Sub table at D1:H(m+1).
        enter_cells(h, &[
            ("D1", "name"), ("E1", "ccy"), ("F1", "local"),
            ("G1", "rate"), ("H1", "usd"),
        ]);
        for (idx, s) in i.subs.iter().enumerate() {
            let row = 2 + idx;
            enter_cell(h, &format!("D{}", row), s.name);
            enter_cell(h, &format!("E{}", row), s.ccy);
            enter_cell(h, &format!("F{}", row), &lit(s.local_balance));
            // VLOOKUP exact-match: `0` means false (exact). tshts treats
            // bare `FALSE` as a CellRef → #NAME?; use the numeric form.
            enter_cell(h, &format!("G{}", row),
                &format!("=VLOOKUP(E{},{},2,0)", row, fx_range));
            enter_cell(h, &format!("H{}", row), &format!("=F{}/G{}", row, row));
        }
        let last_sub = 1 + i.subs.len();
        enter_cell(h, "J1", "group_total_usd");
        enter_cell(h, "K1", &format!("=SUM(H2:H{})", last_sub));
    }

    fn checks(&self, o: &Output) -> Vec<CellCheck> {
        let mut v = Vec::new();
        for (idx, (rate, usd)) in o.rates.iter().zip(&o.usd_balances).enumerate() {
            let row = 2 + idx;
            v.push(CellCheck::new(format!("rate_{}", idx + 1),
                format!("G{}", row), *rate).with_tolerance(1e-6));
            v.push(CellCheck::new(format!("usd_{}", idx + 1),
                format!("H{}", row), *usd).with_tolerance(1e-3));
        }
        v.push(CellCheck::new("group_total_usd", "K1", o.group_total_usd)
            .with_tolerance(1e-2));
        v
    }
}
