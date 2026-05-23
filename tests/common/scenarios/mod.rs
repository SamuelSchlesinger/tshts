//! Scenario-based PTY testing framework.
//!
//! ## Why this exists
//!
//! Unit tests verify that `=SUM(A1:A3)` returns the right number. They do
//! not catch *semantic drift* — e.g., "this DCF model produces enterprise
//! value 4% off because PMT's sign convention differs from what an analyst
//! would expect." This framework cross-checks tshts's end-to-end behavior
//! against an independent pure-Rust implementation of the same model.
//!
//! ## How it works
//!
//! Each scenario implements three things:
//!
//! 1. **`compute(inputs)`** — a pure-Rust ground-truth calculation.
//! 2. **`populate(harness, inputs)`** — drives the PTY to enter the same
//!    model as a spreadsheet (literals + formulas) on a real tshts instance.
//! 3. **`checks(output)`** — a list of `(cell-address, expected-numeric-value)`
//!    pairs derived from the ground-truth output.
//!
//! The runner:
//!
//! 1. Computes truth in Rust.
//! 2. Populates the spreadsheet.
//! 3. Forces a full recalc.
//! 4. For each check, navigates to the cell, enters Visual mode (which
//!    publishes `SUM=<value>` for the selection — a 1×1 selection's SUM
//!    *is* the cell value), parses the value, and asserts within tolerance.
//!
//! ## Why the SUM= status-bar trick
//!
//! The formula bar shows the FORMULA when a cell has one, not its computed
//! value. Reading values from the rendered grid is fragile because narrow
//! columns clip long numbers. The status bar in Visual mode publishes
//! `SUM=…` for the current selection — full-width, never clipped, and
//! `SUM` of a single cell equals that cell's value.
//!
//! ## Extending
//!
//! Add a new scenario by:
//!
//! 1. Implementing the [`Scenario`] trait in a new module under
//!    `tests/common/scenarios/`.
//! 2. Adding a `#[test]` in `tests/pty_scenarios.rs` that calls
//!    `run(&YourScenario)`.
//!
//! Scenarios should NOT hard-code the expected values from a tshts run —
//! they should DERIVE expected values from the Rust `compute()` so that
//! refactors to the model and to tshts both have to agree. The goal is to
//! catch divergence, not to ossify whatever tshts currently outputs.

use crate::common::Harness;
use std::time::Duration;

/// One assertion in a scenario run.
#[derive(Debug, Clone)]
pub struct CellCheck {
    /// Human-readable label shown in failure messages, e.g.
    /// "enterprise_value" or "month_3_balance". Keep concise.
    pub label: String,
    /// A1-style cell address, e.g. "F12" or "Sheet2!A1".
    pub cell: String,
    /// Ground-truth expected numeric value (from the Rust model).
    pub expected: f64,
    /// Absolute tolerance for the comparison. Use ≥ 1e-6 for plain
    /// arithmetic; loosen to ~0.01 for cells that go through
    /// `=ROUND(...)` or display formatting.
    pub tolerance: f64,
}

impl CellCheck {
    pub fn new(label: impl Into<String>, cell: impl Into<String>, expected: f64) -> Self {
        Self {
            label: label.into(),
            cell: cell.into(),
            expected,
            tolerance: 1e-6,
        }
    }

    pub fn with_tolerance(mut self, tol: f64) -> Self {
        self.tolerance = tol;
        self
    }
}

/// A characteristic spreadsheet use case wired up as a Rust
/// ground-truth model + a tshts population routine + a list of
/// post-recalc cell assertions.
pub trait Scenario {
    /// Inputs that parameterize the model. The framework calls
    /// [`Scenario::default_inputs`] today; future runs may sweep over
    /// inputs to fuzz the model.
    type Inputs: Clone;
    /// Ground-truth output of the Rust model. Opaque to the framework;
    /// only [`Scenario::checks`] reads it.
    type Output;

    /// Short display name used in failure messages and test logging.
    fn name(&self) -> &'static str;

    /// Default inputs — typically the example that the scenario is
    /// motivated by. Implementations may expose more for ad-hoc sweeps.
    fn default_inputs(&self) -> Self::Inputs;

    /// Compute the expected outputs in pure Rust.
    ///
    /// This is the GROUND TRUTH. If tshts disagrees, either tshts has a
    /// bug or this Rust model has a bug — both are valid findings.
    fn compute(&self, inputs: &Self::Inputs) -> Self::Output;

    /// Drive the PTY harness to populate the spreadsheet with literals
    /// and formulas matching the model. Must leave the harness in a
    /// state where the runner can navigate freely (typically Normal
    /// mode with no pending input).
    fn populate(&self, h: &mut Harness, inputs: &Self::Inputs);

    /// Build the list of cell checks from the ground-truth output.
    /// Each check ties a spreadsheet cell back to a Rust-computed value.
    fn checks(&self, output: &Self::Output) -> Vec<CellCheck>;
}

/// Run a scenario end-to-end against a fresh PTY harness.
///
/// Panics with a structured multi-line message if any check fails or any
/// cell value can't be parsed from the status bar.
pub fn run<S: Scenario>(s: &S) {
    let inputs = s.default_inputs();
    let truth = s.compute(&inputs);
    let mut h = Harness::new();

    // Populate the spreadsheet. After this, we own the harness state and
    // can drive whatever we need.
    s.populate(&mut h, &inputs);

    // Force a full graph-driven recalc to flush any deferred propagation
    // (e.g. cross-sheet INDIRECT/OFFSET smart-auto-seed). This is the
    // same path the user gets via the `:recalc` command.
    normal_mode(&mut h);
    h.send_text(":recalc");
    h.send_enter();
    // Let the recalc complete before we start probing cells. The settle
    // sleeps inside send_*() handle small inputs; for a full recalc on a
    // large scenario we want more breathing room.
    std::thread::sleep(Duration::from_millis(200));

    // Auto-fit columns. Doesn't affect the SUM=… status-bar trick we
    // actually rely on for reads, but makes failure-time screen dumps
    // legible.
    normal_mode(&mut h);
    h.send_text("+");

    let all_checks = s.checks(&truth);
    let mut failures = Vec::new();
    for check in &all_checks {
        match read_cell_numeric(&mut h, &check.cell) {
            Ok(v) => {
                let delta = (v - check.expected).abs();
                if delta > check.tolerance {
                    failures.push(format!(
                        "  {label} ({cell}): expected {expected:.6}, tshts shows \
                         {actual:.6} (|Δ|={delta:.6}, tol={tol})",
                        label = check.label,
                        cell = check.cell,
                        expected = check.expected,
                        actual = v,
                        delta = delta,
                        tol = check.tolerance,
                    ));
                }
            }
            Err(msg) => {
                failures.push(format!(
                    "  {label} ({cell}): expected {expected:.6}, read error: {msg}",
                    label = check.label,
                    cell = check.cell,
                    expected = check.expected,
                    msg = msg,
                ));
            }
        }
    }

    if !failures.is_empty() {
        panic!(
            "[scenario:{name}] {failed} of {total} checks failed:\n{detail}\n\n\
             ---- final screen ----\n{screen}",
            name = s.name(),
            failed = failures.len(),
            total = all_checks.len(),
            detail = failures.join("\n"),
            screen = h.screen_contents(),
        );
    }
}

// ---------- Harness driving helpers ----------

/// Bounce through `Esc` to guarantee Normal mode. Cheap and safe; the
/// harness's existing input settle delay handles the redraw.
pub fn normal_mode(h: &mut Harness) {
    h.send_esc();
}

/// Navigate the cursor to the given A1 address via Ctrl-G (goto-cell).
/// Plain `g` in tshts is vim's `gg` prefix, not goto — the goto-cell
/// dialog is bound to Ctrl-G (byte 0x07). Accepts plain addresses
/// ("A1", "F12") and sheet-qualified refs ("Sheet2!A1"). Sends the
/// whole `Esc Ctrl-G <addr> Enter` sequence as one PTY write so we pay
/// the harness settle once instead of four times.
pub fn goto_cell(h: &mut Harness, addr: &str) {
    let mut buf = String::with_capacity(addr.len() + 3);
    buf.push('\x1b');   // Esc → Normal mode
    buf.push('\x07');   // Ctrl-G → start_goto_cell
    buf.push_str(addr);
    buf.push('\r');     // commit goto
    h.send_text(&buf);
}

/// Read the numeric value at the given cell. Uses the Visual-mode
/// status-bar trick: in Visual mode the status bar shows
/// `SUM=… AVG=… COUNT=…` for the current selection, and a 1×1 selection's
/// SUM is the cell's value. Full-width, never column-clipped.
///
/// Returns `Err` if the status bar doesn't expose `SUM=` (e.g. the cell
/// holds a non-numeric value) or if the parse fails.
pub fn read_cell_numeric(h: &mut Harness, addr: &str) -> Result<f64, String> {
    goto_cell(h, addr);
    h.send_text("v");
    // The harness's send_text already sleeps INPUT_SETTLE; we add a
    // little extra so the visual-mode status-bar render lands.
    std::thread::sleep(Duration::from_millis(60));
    let status = h.status_bar();
    h.send_esc();
    parse_sum_from_status(&status)
}

/// Read the literal text the formula bar shows for the cell. For a
/// formula cell that's the FORMULA (e.g. `=A1+B1`); for a literal,
/// the value.
pub fn read_cell_formula_bar(h: &mut Harness, addr: &str) -> String {
    goto_cell(h, addr);
    h.formula_bar()
}

/// Type a literal or formula into the given cell and commit with Enter.
/// The whole `Esc Ctrl-G <addr> Enter i <content> Enter` sequence is one
/// PTY write so we pay the harness's render settle once per cell rather
/// than once per keystroke. For a 12-month budget × 5 cols, that's the
/// difference between ~7s and ~50s of input time per scenario.
pub fn enter_cell(h: &mut Harness, addr: &str, content: &str) {
    let mut buf = String::with_capacity(addr.len() + content.len() + 5);
    buf.push('\x1b');           // Esc → Normal mode
    buf.push('\x07');           // Ctrl-G → start goto-cell dialog
    buf.push_str(addr);
    buf.push('\r');             // commit goto
    buf.push('i');              // enter editing
    buf.push_str(content);
    buf.push('\r');             // commit edit (also moves cursor down)
    h.send_text(&buf);
}

/// Bulk-populate a list of (address, content) pairs.
pub fn enter_cells(h: &mut Harness, cells: &[(&str, &str)]) {
    for (addr, content) in cells {
        enter_cell(h, addr, content);
    }
}

fn parse_sum_from_status(status: &str) -> Result<f64, String> {
    let idx = status
        .find("SUM=")
        .ok_or_else(|| format!("no SUM= in status bar: {:?}", status.trim()))?;
    let rest = &status[idx + 4..];
    // SUM=<num> AVG=… — terminate on whitespace.
    let end = rest
        .find(char::is_whitespace)
        .unwrap_or(rest.len());
    let token = &rest[..end];
    token
        .parse::<f64>()
        .map_err(|e| format!("parse {:?} as f64: {}", token, e))
}

// ---------- Numeric helpers shared by scenarios ----------

/// Format an `f64` for inclusion in a tshts formula. Avoids scientific
/// notation, which the parser doesn't accept, and keeps full precision.
pub fn lit(x: f64) -> String {
    // {:?} on f64 prints the shortest decimal that round-trips, in
    // standard notation for normal-range numbers. Falls back to a fixed
    // representation for very large/small magnitudes to avoid scientific
    // notation in formulas.
    if x.is_finite() && x.abs() < 1e15 && (x == 0.0 || x.abs() >= 1e-4) {
        format!("{:?}", x)
    } else {
        format!("{:.12}", x)
    }
}

pub mod amortization;
pub mod bond;
pub mod break_even;
pub mod budgeting;
pub mod commission;
pub mod compound;
pub mod dcf;
pub mod pipeline;
pub mod portfolio;
pub mod schedule;
pub mod sensitivity;
pub mod tax;

#[cfg(test)]
mod parse_tests {
    use super::*;

    #[test]
    fn parse_sum_handles_normal_output() {
        let s = " -- VISUAL --  1x1 cells | SUM=1234.5 AVG=1234.50 COUNT=1 | y yank...";
        assert!((parse_sum_from_status(s).unwrap() - 1234.5).abs() < 1e-9);
    }

    #[test]
    fn parse_sum_handles_integer() {
        let s = "SUM=42 AVG=42.00 COUNT=1";
        assert_eq!(parse_sum_from_status(s).unwrap(), 42.0);
    }

    #[test]
    fn parse_sum_handles_negative() {
        let s = "SUM=-1.5 AVG=-1.50 COUNT=1";
        assert_eq!(parse_sum_from_status(s).unwrap(), -1.5);
    }

    #[test]
    fn parse_sum_missing_returns_err() {
        assert!(parse_sum_from_status("NORMAL mode | nothing here").is_err());
    }
}
