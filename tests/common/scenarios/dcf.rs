//! Discounted Cash Flow (DCF) — valuing a business from projected
//! free cash flows plus a terminal value.
//!
//! Model:
//!
//! ```text
//!     For year t in 1..=N:
//!         FCF_t = FCF_0 * (1 + g)^t
//!         PV_t  = FCF_t / (1 + r)^t
//!     TerminalValue = FCF_{N+1} / (r - g_terminal)        (Gordon growth)
//!     PV_terminal   = TerminalValue / (1 + r)^N
//!     EnterpriseValue = sum(PV_t) + PV_terminal
//! ```
//!
//! Exercises arithmetic, POWER, SUM over a range, and a derived terminal
//! value — all things every analyst expects to "just work."

use super::{lit, CellCheck, Scenario};
use crate::common::scenarios::{enter_cell, enter_cells, Harness};

pub struct Dcf;

#[derive(Clone)]
pub struct Inputs {
    pub fcf_year_0: f64,
    pub fcf_growth: f64,
    pub discount_rate: f64,
    pub terminal_growth: f64,
    pub projection_years: usize,
}

pub struct Output {
    pub fcf: Vec<f64>,        // FCF_1 .. FCF_N
    pub pv: Vec<f64>,         // PV_1 .. PV_N
    pub terminal_value: f64,  // TV at year N (undiscounted)
    pub pv_terminal: f64,     // PV of TV at year 0
    pub enterprise_value: f64,
}

impl Scenario for Dcf {
    type Inputs = Inputs;
    type Output = Output;

    fn name(&self) -> &'static str { "dcf" }

    fn default_inputs(&self) -> Inputs {
        Inputs {
            fcf_year_0: 1_000_000.0,
            fcf_growth: 0.08,
            discount_rate: 0.11,
            terminal_growth: 0.025,
            projection_years: 5,
        }
    }

    fn compute(&self, i: &Inputs) -> Output {
        let mut fcf = Vec::with_capacity(i.projection_years);
        let mut pv = Vec::with_capacity(i.projection_years);
        for t in 1..=i.projection_years {
            let c = i.fcf_year_0 * (1.0 + i.fcf_growth).powi(t as i32);
            fcf.push(c);
            pv.push(c / (1.0 + i.discount_rate).powi(t as i32));
        }
        // FCF in year N+1 grows at the TERMINAL rate, not the projection
        // rate (standard textbook DCF convention — projection growth
        // tapers off into perpetual terminal growth).
        let fcf_n = i.fcf_year_0
            * (1.0 + i.fcf_growth).powi(i.projection_years as i32);
        let fcf_next = fcf_n * (1.0 + i.terminal_growth);
        let terminal_value = fcf_next / (i.discount_rate - i.terminal_growth);
        let pv_terminal =
            terminal_value / (1.0 + i.discount_rate).powi(i.projection_years as i32);
        let enterprise_value: f64 = pv.iter().sum::<f64>() + pv_terminal;
        Output { fcf, pv, terminal_value, pv_terminal, enterprise_value }
    }

    fn populate(&self, h: &mut Harness, i: &Inputs) {
        // Layout:
        //   A1 fcf_0      B1 <value>
        //   A2 g          B2 <value>
        //   A3 r          B3 <value>
        //   A4 g_term     B4 <value>
        //   Row 6: header (Year, FCF, DF, PV)
        //   Rows 7..7+N-1: year detail
        //   A14 Terminal value | B14 = B(6+N) * (1 + B4) / (B3 - B4)
        //   A15 PV terminal    | B15 = B14 / (1 + B3) ^ N
        //   A16 EV             | B16 = SUM(D7:D{6+N}) + B15
        enter_cells(h, &[
            ("A1", "fcf_0"),       ("B1", &lit(i.fcf_year_0)),
            ("A2", "g"),           ("B2", &lit(i.fcf_growth)),
            ("A3", "r"),           ("B3", &lit(i.discount_rate)),
            ("A4", "g_terminal"),  ("B4", &lit(i.terminal_growth)),
            ("A6", "Year"), ("B6", "FCF"), ("C6", "DF"), ("D6", "PV"),
        ]);
        let n = i.projection_years;
        for t in 1..=n {
            let row = 6 + t;
            enter_cell(h, &format!("A{}", row), &t.to_string());
            // FCF_t = fcf_0 * (1 + g)^t  →  in cells: =$B$1 * POWER(1+$B$2, A_t)
            enter_cell(h, &format!("B{}", row),
                &format!("=$B$1*POWER(1+$B$2,A{})", row));
            // DF_t = 1 / (1 + r)^t
            enter_cell(h, &format!("C{}", row),
                &format!("=1/POWER(1+$B$3,A{})", row));
            // PV_t = FCF_t * DF_t
            enter_cell(h, &format!("D{}", row),
                &format!("=B{}*C{}", row, row));
        }
        let last_data_row = 6 + n;
        enter_cells(h, &[
            ("A14", "Terminal value"),
            ("A15", "PV terminal"),
            ("A16", "Enterprise value"),
        ]);
        // TV = FCF_{N+1} / (r - g_t) = (FCF_N * (1 + g_t)) / (r - g_t)
        enter_cell(h, "B14",
            &format!("=B{}*(1+$B$4)/($B$3-$B$4)", last_data_row));
        // PV(TV) = TV / (1 + r)^N
        enter_cell(h, "B15",
            &format!("=B14/POWER(1+$B$3,{})", n));
        // EV = sum of PVs + PV(TV)
        enter_cell(h, "B16",
            &format!("=SUM(D7:D{})+B15", last_data_row));
    }

    fn checks(&self, o: &Output) -> Vec<CellCheck> {
        let mut v = Vec::new();
        // Spot-check FCF and PV for each projection year.
        for (t, (fcf, pv)) in o.fcf.iter().zip(o.pv.iter()).enumerate() {
            let row = 7 + t;
            v.push(CellCheck::new(format!("fcf_{}", t + 1),
                format!("B{}", row), *fcf).with_tolerance(1e-4));
            v.push(CellCheck::new(format!("pv_{}", t + 1),
                format!("D{}", row), *pv).with_tolerance(1e-4));
        }
        v.push(CellCheck::new("terminal_value", "B14", o.terminal_value)
            .with_tolerance(1e-3));
        v.push(CellCheck::new("pv_terminal", "B15", o.pv_terminal)
            .with_tolerance(1e-3));
        v.push(CellCheck::new("enterprise_value", "B16", o.enterprise_value)
            .with_tolerance(1e-3));
        v
    }
}
