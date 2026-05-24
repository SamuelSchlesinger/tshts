//! Lookup-driven workflow: a price list + an order book + per-order
//! cost lookup via VLOOKUP. Tests `VLOOKUP`, `INDEX`/`MATCH`,
//! `IFERROR` for missing SKUs.

use super::{lit, CellCheck, Scenario};
use crate::common::scenarios::{enter_cell, enter_cells, Harness};

pub struct Lookup;

#[derive(Clone)]
pub struct PriceRow {
    pub sku: &'static str,
    pub price: f64,
}

#[derive(Clone)]
pub struct Order {
    pub sku: &'static str,
    pub qty: f64,
}

#[derive(Clone)]
pub struct Inputs {
    pub prices: Vec<PriceRow>,
    pub orders: Vec<Order>,
}

pub struct Output {
    /// Unit price per order (NaN for unknown SKU).
    pub unit_prices: Vec<f64>,
    /// Line totals (0 if unknown SKU — IFERROR fallback).
    pub line_totals: Vec<f64>,
    pub order_count: f64,
    pub known_count: f64,
    pub total_revenue: f64,
}

impl Scenario for Lookup {
    type Inputs = Inputs;
    type Output = Output;

    fn name(&self) -> &'static str { "lookup" }

    fn default_inputs(&self) -> Inputs {
        Inputs {
            prices: vec![
                PriceRow { sku: "APL",  price:  0.50 },
                PriceRow { sku: "BNN",  price:  0.25 },
                PriceRow { sku: "CHR",  price:  3.00 },
                PriceRow { sku: "DRG",  price:  2.50 },
                PriceRow { sku: "EGG",  price:  4.00 },
            ],
            orders: vec![
                Order { sku: "BNN",  qty: 12.0 },
                Order { sku: "APL",  qty:  6.0 },
                Order { sku: "EGG",  qty:  2.0 },
                Order { sku: "ZZZ",  qty:  3.0 }, // unknown SKU
                Order { sku: "CHR",  qty:  4.0 },
                Order { sku: "DRG",  qty:  1.0 },
            ],
        }
    }

    fn compute(&self, i: &Inputs) -> Output {
        let lookup = |sku: &str| -> Option<f64> {
            i.prices.iter().find(|p| p.sku == sku).map(|p| p.price)
        };
        let mut unit_prices = Vec::with_capacity(i.orders.len());
        let mut line_totals = Vec::with_capacity(i.orders.len());
        for o in &i.orders {
            match lookup(o.sku) {
                Some(p) => {
                    unit_prices.push(p);
                    line_totals.push(p * o.qty);
                }
                None => {
                    unit_prices.push(f64::NAN);
                    line_totals.push(0.0); // IFERROR fallback
                }
            }
        }
        let known_count = i.orders.iter().filter(|o| lookup(o.sku).is_some()).count() as f64;
        let total_revenue: f64 = line_totals.iter().sum();
        Output {
            unit_prices,
            line_totals,
            order_count: i.orders.len() as f64,
            known_count,
            total_revenue,
        }
    }

    fn populate(&self, h: &mut Harness, i: &Inputs) {
        // Price list at A1:B(n+1). Header row + price rows.
        enter_cells(h, &[
            ("A1", "sku"), ("B1", "price"),
        ]);
        for (idx, p) in i.prices.iter().enumerate() {
            let row = 2 + idx;
            enter_cell(h, &format!("A{}", row), p.sku);
            enter_cell(h, &format!("B{}", row), &lit(p.price));
        }
        let last_price_row = 1 + i.prices.len();
        // Fully-absolute refs for the VLOOKUP table.
        let price_range = super::abs_range(&format!("A2:B{}", last_price_row));

        // Order book at D1:G(m+1).
        enter_cells(h, &[
            ("D1", "sku"), ("E1", "qty"), ("F1", "unit_price"), ("G1", "line_total"),
        ]);
        for (idx, o) in i.orders.iter().enumerate() {
            let row = 2 + idx;
            enter_cell(h, &format!("D{}", row), o.sku);
            enter_cell(h, &format!("E{}", row), &lit(o.qty));
            // VLOOKUP for the unit price; IFERROR returns 0 if SKU unknown.
            // VLOOKUP exact-match via bare FALSE (the natural form, now
            // that tshts handles bare TRUE/FALSE as literals).
            enter_cell(h, &format!("F{}", row),
                &format!("=IFERROR(VLOOKUP(D{},{},2,FALSE),0)", row, price_range));
            enter_cell(h, &format!("G{}", row), &format!("=E{}*F{}", row, row));
        }
        let last_order_row = 1 + i.orders.len();

        // Summary at I1..
        enter_cells(h, &[
            ("I1", "order_count"),
            ("I2", "known_count"),
            ("I3", "total_revenue"),
            ("I4", "indexmatch_xcheck"),
        ]);
        enter_cell(h, "J1", &format!("=COUNTA(D2:D{})", last_order_row));
        // known_count: SUMPRODUCT trick — count orders where VLOOKUP succeeds.
        // Easier: count orders where F is > 0 (assuming all known prices > 0).
        enter_cell(h, "J2", &format!("=COUNTIF(F2:F{},\">0\")", last_order_row));
        enter_cell(h, "J3", &format!("=SUM(G2:G{})", last_order_row));
        // INDEX/MATCH cross-check: same lookup as VLOOKUP, different path.
        // For the FIRST order, compute via INDEX/MATCH and assert equality
        // (the test asserts both unit_price formulas agree). Pick the first
        // KNOWN order so the result is comparable.
        let first_known_idx = i.orders.iter().position(|o|
            i.prices.iter().any(|p| p.sku == o.sku)
        ).unwrap_or(0);
        let row = 2 + first_known_idx;
        enter_cell(h, "J4", &format!(
            "=INDEX({}, MATCH(D{},$A$2:$A${},0), 2)",
            price_range, row, last_price_row,
        ));
    }

    fn checks(&self, o: &Output) -> Vec<CellCheck> {
        let mut v = Vec::new();
        for (idx, total) in o.line_totals.iter().enumerate() {
            let row = 2 + idx;
            v.push(CellCheck::new(format!("line_total_{}", idx + 1),
                format!("G{}", row), *total).with_tolerance(1e-6));
        }
        v.push(CellCheck::new("order_count", "J1", o.order_count));
        v.push(CellCheck::new("known_count", "J2", o.known_count));
        v.push(CellCheck::new("total_revenue", "J3", o.total_revenue).with_tolerance(1e-3));
        // INDEX/MATCH agreement on the first known order's unit price.
        let first_known_idx = o.unit_prices.iter().position(|p| !p.is_nan()).unwrap_or(0);
        v.push(CellCheck::new("indexmatch_unit_price",
            "J4", o.unit_prices[first_known_idx]).with_tolerance(1e-6));
        v
    }
}
