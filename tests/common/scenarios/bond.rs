//! Bond pricing — present value of coupon payments plus PV of face
//! redemption. The same formula every fixed-income analyst pulls up
//! a hundred times a week.
//!
//! ```text
//!     coupon_payment = face * coupon_rate / freq
//!     n_payments     = years * freq
//!     periodic_yield = ytm / freq
//!
//!     price = sum_{t=1}^{n_payments}( coupon_payment / (1 + periodic_yield)^t )
//!           + face / (1 + periodic_yield)^n_payments
//! ```

use super::{lit, CellCheck, Scenario};
use crate::common::scenarios::{enter_cell, enter_cells, Harness};

pub struct Bond;

#[derive(Clone)]
pub struct Inputs {
    pub face: f64,
    pub coupon_rate: f64,
    pub ytm: f64,
    pub years: usize,
    pub freq: usize, // payments per year (typically 2)
}

pub struct Output {
    pub coupon_payment: f64,
    pub pv_coupons: f64,
    pub pv_face: f64,
    pub price: f64,
}

impl Scenario for Bond {
    type Inputs = Inputs;
    type Output = Output;

    fn name(&self) -> &'static str { "bond" }

    fn default_inputs(&self) -> Inputs {
        // 10-year semi-annual, 5% coupon, trading at 4% YTM → premium bond.
        Inputs { face: 1000.0, coupon_rate: 0.05, ytm: 0.04, years: 10, freq: 2 }
    }

    fn compute(&self, i: &Inputs) -> Output {
        let n = i.years * i.freq;
        let cpn = i.face * i.coupon_rate / i.freq as f64;
        let py = i.ytm / i.freq as f64;
        let mut pv_cpns = 0.0;
        for t in 1..=n {
            pv_cpns += cpn / (1.0 + py).powi(t as i32);
        }
        let pv_face = i.face / (1.0 + py).powi(n as i32);
        let price = pv_cpns + pv_face;
        Output { coupon_payment: cpn, pv_coupons: pv_cpns, pv_face, price }
    }

    fn populate(&self, h: &mut Harness, i: &Inputs) {
        let n = i.years * i.freq;
        enter_cells(h, &[
            ("A1", "face"),         ("B1", &lit(i.face)),
            ("A2", "coupon_rate"),  ("B2", &lit(i.coupon_rate)),
            ("A3", "ytm"),          ("B3", &lit(i.ytm)),
            ("A4", "years"),        ("B4", &i.years.to_string()),
            ("A5", "freq"),         ("B5", &i.freq.to_string()),
            ("A6", "n"),            ("B6", "=B4*B5"),
            ("A7", "cpn"),          ("B7", "=B1*B2/B5"),
            ("A8", "py"),           ("B8", "=B3/B5"),
            ("A10", "t"), ("B10", "pv_cpn"), ("C10", "pv_face"),
        ]);
        for t in 1..=n {
            let row = 10 + t;
            enter_cell(h, &format!("A{}", row), &t.to_string());
            enter_cell(h, &format!("B{}", row),
                &format!("=$B$7/POWER(1+$B$8,A{})", row));
        }
        let last = 10 + n;
        // PV face only at the last period.
        enter_cell(h, &format!("C{}", last),
            &format!("=$B$1/POWER(1+$B$8,A{})", last));
        enter_cells(h, &[
            ("E1", "pv_coupons"),
            ("E2", "pv_face"),
            ("E3", "price"),
        ]);
        enter_cell(h, "F1", &format!("=SUM(B11:B{})", last));
        enter_cell(h, "F2", &format!("=C{}", last));
        enter_cell(h, "F3", "=F1+F2");
    }

    fn checks(&self, o: &Output) -> Vec<CellCheck> {
        vec![
            CellCheck::new("coupon_payment", "B7", o.coupon_payment).with_tolerance(1e-6),
            CellCheck::new("pv_coupons",      "F1", o.pv_coupons).with_tolerance(1e-3),
            CellCheck::new("pv_face",         "F2", o.pv_face).with_tolerance(1e-3),
            CellCheck::new("price",           "F3", o.price).with_tolerance(1e-3),
        ]
    }
}
