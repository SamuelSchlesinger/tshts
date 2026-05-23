//! Break-even analysis: how many units must we sell to cover fixed costs?
//!
//! The classic textbook formula:
//!
//! ```text
//!     Q* = FixedCosts / (UnitPrice - UnitVariableCost)
//! ```
//!
//! Also derives the revenue and contribution margin at break-even, which
//! give us multiple cells to cross-check.

use super::{lit, CellCheck, Scenario};
use crate::common::scenarios::{enter_cells, Harness};

pub struct BreakEven;

#[derive(Clone)]
pub struct Inputs {
    pub fixed_costs: f64,
    pub unit_price: f64,
    pub unit_variable_cost: f64,
}

pub struct Output {
    pub contribution_margin: f64,
    pub break_even_units: f64,
    pub break_even_revenue: f64,
}

impl Scenario for BreakEven {
    type Inputs = Inputs;
    type Output = Output;

    fn name(&self) -> &'static str {
        "break_even"
    }

    fn default_inputs(&self) -> Inputs {
        // Niche SaaS: $120k of overhead, sell a $400/yr seat that costs $90
        // to support. ~387 seats to break even.
        Inputs {
            fixed_costs: 120_000.0,
            unit_price: 400.0,
            unit_variable_cost: 90.0,
        }
    }

    fn compute(&self, i: &Inputs) -> Output {
        let contribution_margin = i.unit_price - i.unit_variable_cost;
        let break_even_units = i.fixed_costs / contribution_margin;
        let break_even_revenue = break_even_units * i.unit_price;
        Output {
            contribution_margin,
            break_even_units,
            break_even_revenue,
        }
    }

    fn populate(&self, h: &mut Harness, i: &Inputs) {
        // Layout:
        //   A1 fixed_costs           B1 <value>
        //   A2 unit_price            B2 <value>
        //   A3 unit_variable_cost    B3 <value>
        //   A5 contribution_margin   B5 =B2-B3
        //   A6 break_even_units      B6 =B1/B5
        //   A7 break_even_revenue    B7 =B6*B2
        enter_cells(h, &[
            ("A1", "fixed_costs"),       ("B1", &lit(i.fixed_costs)),
            ("A2", "unit_price"),        ("B2", &lit(i.unit_price)),
            ("A3", "unit_variable_cost"),("B3", &lit(i.unit_variable_cost)),
            ("A5", "contribution_margin"),("B5", "=B2-B3"),
            ("A6", "break_even_units"),  ("B6", "=B1/B5"),
            ("A7", "break_even_revenue"),("B7", "=B6*B2"),
        ]);
    }

    fn checks(&self, o: &Output) -> Vec<CellCheck> {
        vec![
            CellCheck::new("contribution_margin", "B5", o.contribution_margin),
            CellCheck::new("break_even_units", "B6", o.break_even_units)
                .with_tolerance(1e-6),
            CellCheck::new("break_even_revenue", "B7", o.break_even_revenue)
                .with_tolerance(1e-3),
        ]
    }
}
