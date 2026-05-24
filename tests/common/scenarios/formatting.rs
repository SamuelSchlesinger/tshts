//! Number formatting: catches divergence between what tshts computes
//! and what tshts DISPLAYS. The rest of the scenario suite reads raw
//! numeric values via the visual-mode SUM= trick — but that bypasses
//! format_cell_value entirely. A bug in currency / percent / general
//! rendering would slip through. This scenario sets cell values, applies
//! a number format per row, and asserts the grid renders the expected
//! formatted text.
//!
//! Layout: each row holds one (value, format) pair. The runner uses
//! `rendered_text_checks` (introduced for this scenario) to assert the
//! rendered cell text contains the expected formatted substring.

use super::{enter_cell, goto_cell, lit, CellCheck, RenderedTextCheck, Scenario};
use crate::common::Harness;

pub struct Formatting;

#[derive(Clone)]
pub struct Case {
    pub raw: f64,
    /// `:format ...` command suffix, e.g. `"currency"`, `"percent"`, `"number 0"`.
    pub format_command: &'static str,
    /// Substring expected in the rendered cell.
    pub expected_substring: &'static str,
    pub label: &'static str,
}

#[derive(Clone)]
pub struct Inputs {
    pub cases: Vec<Case>,
}

pub struct Output {
    pub cases: Vec<Case>,
}

impl Scenario for Formatting {
    type Inputs = Inputs;
    type Output = Output;

    fn name(&self) -> &'static str { "formatting" }

    fn default_inputs(&self) -> Inputs {
        Inputs {
            cases: vec![
                // Currency: $-prefix + 2 decimals + thousands separator.
                // Kept the magnitude small enough that even with the
                // narrow auto-fit column, the full string survives.
                Case {
                    raw: 9876.54,
                    format_command: "currency",
                    expected_substring: "$9,876.54",
                    label: "currency_2dp",
                },
                // Percent: ×100 + % suffix + 1 decimal by default.
                Case {
                    raw: 0.085,
                    format_command: "percent",
                    expected_substring: "8.5%",
                    label: "percent_1dp",
                },
                // Number format with explicit decimals. Kept short so
                // the rendered text survives auto-fit column width.
                Case {
                    raw: 1234.0,
                    format_command: "number 2",
                    expected_substring: "1234.00",
                    label: "number_2dp",
                },
                // Tiny number with percent — 0.01 → 1.0%.
                Case {
                    raw: 0.01,
                    format_command: "percent",
                    expected_substring: "1.0%",
                    label: "percent_small",
                },
                // Negative currency — sign preserved.
                Case {
                    raw: -42.5,
                    format_command: "currency",
                    expected_substring: "-$42.50",
                    label: "currency_negative",
                },
            ],
        }
    }

    fn compute(&self, i: &Inputs) -> Output {
        Output { cases: i.cases.clone() }
    }

    fn populate(&self, h: &mut Harness, i: &Inputs) {
        // For each case: write the raw value into column A, then enter
        // Visual mode on that cell and apply the :format command. The
        // selection is required because :format operates on the current
        // selection, not the cursor cell.
        for (idx, case) in i.cases.iter().enumerate() {
            let row = 1 + idx;
            let addr = format!("A{}", row);
            enter_cell(h, &addr, &lit(case.raw));
            // Re-navigate (enter_cell leaves cursor on the next row).
            goto_cell(h, &addr);
            // v = enter Visual mode (single-cell selection); the format
            // command then applies to A{row}. Escape after to leave
            // Visual cleanly before the next iteration.
            h.send_text("v");
            // Apply the format via the command palette.
            let mut cmd = String::from(":format ");
            cmd.push_str(case.format_command);
            h.send_text(&cmd);
            h.send_enter();
            // Drop back to Normal mode so the next enter_cell starts clean.
            h.send_esc();
        }
    }

    fn checks(&self, _o: &Output) -> Vec<CellCheck> {
        // No numeric checks — we're testing rendering, not arithmetic.
        // The raw value the user typed IS the value; format_cell_value
        // is what we're auditing.
        Vec::new()
    }

    fn rendered_text_checks(&self, o: &Output) -> Vec<RenderedTextCheck> {
        o.cases
            .iter()
            .enumerate()
            .map(|(idx, case)| {
                // data_row is 1-indexed; row 1 = first data row.
                let row = 1 + idx as u16;
                RenderedTextCheck::new(case.label, row, case.expected_substring)
            })
            .collect()
    }
}
