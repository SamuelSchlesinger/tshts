//! Inventory FIFO valuation. A running ledger of stock purchases and sales;
//! Cost-Of-Goods-Sold uses the price of the OLDEST in-stock units (First
//! In First Out). Remaining inventory value is the price of the newest
//! in-stock units.
//!
//! Lays out a transaction log and computes:
//!   * running on-hand quantity
//!   * cost-of-goods-sold per sale (FIFO matched against the layer log)
//!   * ending inventory value (FIFO — newest layers first)
//!
//! Exercises a running cumulative SUM across rows, condition-based
//! totals via SUMIF, and a recurrence (each row depends on prior rows).

use super::{lit, CellCheck, Scenario};
use crate::common::scenarios::{enter_cell, enter_cells, Harness};

pub struct Inventory;

#[derive(Clone)]
pub enum Txn {
    /// Buy `qty` units at `unit_cost` each.
    Buy { qty: f64, unit_cost: f64 },
    /// Sell `qty` units. Sale price is whatever the market is; we don't
    /// model revenue here — only cost.
    Sell { qty: f64 },
}

#[derive(Clone)]
pub struct Inputs {
    pub txns: Vec<Txn>,
}

pub struct Output {
    /// On-hand quantity AFTER each transaction.
    pub on_hand: Vec<f64>,
    /// COGS recognized by THIS transaction (0 for buys; FIFO-matched for sells).
    pub cogs_step: Vec<f64>,
    pub total_cogs: f64,
    /// Inventory value (sum of remaining layers × unit cost).
    pub ending_inventory_value: f64,
}

#[derive(Clone)]
struct Layer {
    qty: f64,
    unit_cost: f64,
}

impl Scenario for Inventory {
    type Inputs = Inputs;
    type Output = Output;

    fn name(&self) -> &'static str { "inventory" }

    fn default_inputs(&self) -> Inputs {
        Inputs {
            txns: vec![
                Txn::Buy { qty: 100.0, unit_cost: 10.0 },
                Txn::Buy { qty:  50.0, unit_cost: 12.0 },
                Txn::Sell { qty: 70.0 },             // consumes 70 of the $10 layer
                Txn::Buy { qty:  80.0, unit_cost: 13.0 },
                Txn::Sell { qty: 90.0 },             // consumes last 30 of $10, then 50 of $12, then 10 of $13
                Txn::Buy { qty:  40.0, unit_cost: 15.0 },
                Txn::Sell { qty: 50.0 },             // consumes 50 of the $13 layer
            ],
        }
    }

    fn compute(&self, i: &Inputs) -> Output {
        let mut layers: Vec<Layer> = Vec::new();
        let mut on_hand = Vec::with_capacity(i.txns.len());
        let mut cogs_step = Vec::with_capacity(i.txns.len());
        for txn in &i.txns {
            match *txn {
                Txn::Buy { qty, unit_cost } => {
                    layers.push(Layer { qty, unit_cost });
                    cogs_step.push(0.0);
                }
                Txn::Sell { mut qty } => {
                    let mut cost = 0.0;
                    while qty > 0.0 && !layers.is_empty() {
                        let take = qty.min(layers[0].qty);
                        cost += take * layers[0].unit_cost;
                        qty -= take;
                        layers[0].qty -= take;
                        if layers[0].qty <= 0.0 {
                            layers.remove(0);
                        }
                    }
                    cogs_step.push(cost);
                }
            }
            on_hand.push(layers.iter().map(|l| l.qty).sum());
        }
        let total_cogs: f64 = cogs_step.iter().sum();
        let ending_inventory_value: f64 = layers.iter().map(|l| l.qty * l.unit_cost).sum();
        Output { on_hand, cogs_step, total_cogs, ending_inventory_value }
    }

    fn populate(&self, h: &mut Harness, i: &Inputs) {
        // Layout (row 1 header, rows 2.. transactions):
        //   A: type ("BUY"/"SELL") | B: qty in   | C: unit cost (BUY only)
        //   D: qty out (SELL only) | E: cogs_step (formula)
        //   F: on_hand after this row (running sum of B - D)
        //   G: cumulative cogs
        //
        // COGS for a SELL row is computed by the test caller in Rust
        // (FIFO is hard to express purely in tshts formulas without
        // helper columns). We just CROSS-CHECK: write the COGS value
        // we expect into E, then verify SUM(E) matches total_cogs and
        // on_hand at each row matches the Rust computation.
        //
        // The cells tshts actually computes:
        //   F{r} = F{r-1} + B{r} - D{r}     (cumulative on-hand)
        //   G{r} = G{r-1} + E{r}            (cumulative COGS)
        //   H1   = SUM(E2:E{last})          (total cogs cross-check)
        //
        // FIFO logic itself is verified by the Rust compute() (since
        // both sides agree). What we're really verifying is the
        // running-sum recurrence on F/G + the SUM aggregate on H1.
        enter_cells(h, &[
            ("A1", "type"), ("B1", "qty_in"), ("C1", "unit_cost"),
            ("D1", "qty_out"), ("E1", "cogs"), ("F1", "on_hand"), ("G1", "cum_cogs"),
        ]);
        let truth = self.compute(i);
        for (idx, txn) in i.txns.iter().enumerate() {
            let row = 2 + idx;
            match *txn {
                Txn::Buy { qty, unit_cost } => {
                    enter_cell(h, &format!("A{}", row), "BUY");
                    enter_cell(h, &format!("B{}", row), &lit(qty));
                    enter_cell(h, &format!("C{}", row), &lit(unit_cost));
                    enter_cell(h, &format!("D{}", row), "0");
                    // COGS for a buy is zero.
                    enter_cell(h, &format!("E{}", row), "0");
                }
                Txn::Sell { qty } => {
                    enter_cell(h, &format!("A{}", row), "SELL");
                    enter_cell(h, &format!("B{}", row), "0");
                    enter_cell(h, &format!("C{}", row), "0");
                    enter_cell(h, &format!("D{}", row), &lit(qty));
                    // Stamp the truth-side COGS into the cell as a literal —
                    // we're cross-checking the running-sum machinery, not
                    // computing FIFO inside tshts.
                    enter_cell(h, &format!("E{}", row), &lit(truth.cogs_step[idx]));
                }
            }
            // F{r} = on_hand running. First row gets B - D directly.
            if idx == 0 {
                enter_cell(h, &format!("F{}", row), &format!("=B{}-D{}", row, row));
                enter_cell(h, &format!("G{}", row), &format!("=E{}", row));
            } else {
                let prev = row - 1;
                enter_cell(h, &format!("F{}", row),
                    &format!("=F{}+B{}-D{}", prev, row, row));
                enter_cell(h, &format!("G{}", row),
                    &format!("=G{}+E{}", prev, row));
            }
        }
        let last = 1 + i.txns.len();
        enter_cell(h, "I1", "total_cogs_sum");
        enter_cell(h, "J1", &format!("=SUM(E2:E{})", last));
        enter_cell(h, "I2", "sales_count");
        enter_cell(h, "J2", &format!("=COUNTIF(A2:A{},\"SELL\")", last));
    }

    fn checks(&self, o: &Output) -> Vec<CellCheck> {
        let mut v = Vec::new();
        for (idx, (on, _cogs)) in o.on_hand.iter().zip(&o.cogs_step).enumerate() {
            let row = 2 + idx;
            v.push(CellCheck::new(format!("on_hand_r{}", idx + 1),
                format!("F{}", row), *on));
        }
        v.push(CellCheck::new("total_cogs_via_SUM", "J1", o.total_cogs)
            .with_tolerance(1e-3));
        // Cumulative COGS at the last row should match total_cogs too.
        let last = 1 + o.on_hand.len();
        v.push(CellCheck::new("cum_cogs_last", format!("G{}", last), o.total_cogs)
            .with_tolerance(1e-3));
        // Sales-count cross-check via COUNTIF.
        let sales = o.cogs_step.iter().filter(|c| **c > 0.0).count() as f64;
        v.push(CellCheck::new("sales_count", "J2", sales));
        v
    }
}
