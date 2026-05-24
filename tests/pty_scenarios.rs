//! End-to-end PTY tests for characteristic spreadsheet use cases.
//!
//! Each `#[test]` runs one scenario: a real tshts process is spawned in a
//! PTY, the spreadsheet is populated with the model, and the computed
//! cell values are cross-checked against an independent pure-Rust
//! implementation of the same model. See `tests/common/scenarios/mod.rs`
//! for the framework.

mod common;

use common::scenarios::run;

#[test]
fn scenario_break_even() {
    run(&common::scenarios::break_even::BreakEven);
}

#[test]
fn scenario_dcf() {
    run(&common::scenarios::dcf::Dcf);
}

#[test]
fn scenario_budgeting() {
    run(&common::scenarios::budgeting::Budgeting);
}

#[test]
fn scenario_amortization() {
    run(&common::scenarios::amortization::Amortization);
}

#[test]
fn scenario_compound_interest() {
    run(&common::scenarios::compound::Compound);
}

#[test]
fn scenario_portfolio() {
    run(&common::scenarios::portfolio::Portfolio);
}

#[test]
fn scenario_commission() {
    run(&common::scenarios::commission::Commission);
}

#[test]
fn scenario_tax_brackets() {
    run(&common::scenarios::tax::Tax);
}

#[test]
fn scenario_project_schedule() {
    run(&common::scenarios::schedule::Schedule);
}

#[test]
fn scenario_bond_pricing() {
    run(&common::scenarios::bond::Bond);
}

#[test]
fn scenario_sensitivity_table() {
    run(&common::scenarios::sensitivity::Sensitivity);
}

#[test]
fn scenario_sales_pipeline() {
    run(&common::scenarios::pipeline::Pipeline);
}

#[test]
fn scenario_inventory_fifo() {
    run(&common::scenarios::inventory::Inventory);
}

#[test]
fn scenario_lookup_orders() {
    run(&common::scenarios::lookup::Lookup);
}

#[test]
fn scenario_quarterly_forecast() {
    run(&common::scenarios::quarterly_forecast::QuarterlyForecast);
}

#[test]
fn scenario_currency_translation() {
    run(&common::scenarios::currency::Currency);
}

#[test]
fn scenario_sales_leaderboard() {
    run(&common::scenarios::leaderboard::Leaderboard);
}

#[test]
fn scenario_regression() {
    run(&common::scenarios::regression::Regression);
}

#[test]
fn scenario_number_formatting() {
    run(&common::scenarios::formatting::Formatting);
}
