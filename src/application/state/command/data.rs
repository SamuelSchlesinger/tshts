//! Helper methods for data-shaping commands extracted from the
//! `execute_command` dispatcher: pivot, table create, chart, goalseek.
//!
//! These are the longer arm bodies that benefit from being broken out;
//! the trivial one-line arms (`["bold"] => self.toggle_bold()`) stay
//! inline in the dispatcher.

use super::super::{App, parse_range};
use crate::domain::{CellData, FormulaEvaluator, Spreadsheet};

impl App {
    /// `:table create A1:D10 [name=Sales]` — register a named table over
    /// the range. Each column also gets a `Table[Col]` named range so
    /// structured references resolve via the existing name machinery.
    pub(super) fn cmd_table_create(&mut self, range_str: &str, opts: &[&str], trimmed: &str) {
        let Some((start, end)) = parse_range(range_str) else {
            self.status_message = Some(format!("Invalid range: {}", range_str));
            return;
        };
        let mut name = format!("Table{}", self.workbook.current_sheet().tables.len() + 1);
        // Recover case-preserved opts from the original command so
        // `name=Mine` is stored as "Mine", not "mine".
        let original_opts: Vec<&str> = trimmed.split_whitespace().skip(3).collect();
        for (i, opt) in opts.iter().enumerate() {
            if opt.starts_with("name=") {
                if let Some(orig) = original_opts.get(i)
                    && let Some(n) = orig.strip_prefix("name=")
                {
                    name = n.to_string();
                    continue;
                }
                if let Some(n) = opt.strip_prefix("name=") {
                    name = n.to_string();
                }
            }
        }
        let sheet = self.workbook.current_sheet();
        let cols = end.1 - start.1 + 1;
        let mut headers = Vec::with_capacity(cols);
        for c in start.1..=end.1 {
            let h = sheet.get_cell(start.0, c).value;
            headers.push(if h.is_empty() {
                Spreadsheet::column_label(c)
            } else {
                h
            });
        }
        let table = crate::domain::Table {
            name: name.clone(),
            top_row: start.0,
            left_col: start.1,
            bottom_row: end.0,
            right_col: end.1,
            headers,
        };
        self.workbook.current_sheet_mut().tables.push(table);
        // Also register each column as a named range so `Table1[Col1]`
        // works via existing name resolution.
        let sheet = self.workbook.current_sheet();
        let table = sheet.tables.last().unwrap().clone();
        let body_top = table.top_row + 1; // skip header row
        for (i, header) in table.headers.iter().enumerate() {
            let key = format!("{}[{}]", table.name.to_uppercase(), header.to_uppercase());
            let value = format!(
                "{}:{}",
                Spreadsheet::format_cell_reference(body_top, table.left_col + i, false, false),
                Spreadsheet::format_cell_reference(table.bottom_row, table.left_col + i, false, false),
            );
            self.workbook.named_ranges.insert(key.clone(), value.clone());
            for s in &mut self.workbook.sheets {
                s.named_ranges.insert(key.clone(), value.clone());
            }
        }
        self.dirty = true;
        self.status_message = Some(format!("Created table '{}'", name));
    }

    /// `:goalseek TARGET_CELL EXPECTED INPUT_CELL` — bisect the input cell
    /// until the target cell evaluates to (approximately) the expected
    /// value. Runs on a workbook clone to keep ~80 probe writes from
    /// rippling through real cross-sheet dependents; only the converged
    /// value lands via `set_cell_with_undo`.
    pub(super) fn cmd_goalseek(&mut self, target: &str, expected: &str, input: &str) {
        let target_pos = Spreadsheet::parse_cell_reference(target);
        let input_pos = Spreadsheet::parse_cell_reference(input);
        let expected_v: f64 = expected.parse().unwrap_or(0.0);
        let (Some(t), Some(i)) = (target_pos, input_pos) else {
            self.status_message = Some("goalseek: bad cell reference".to_string());
            return;
        };
        let mut scratch = self.workbook.clone();
        let active_idx = scratch.active_sheet;
        let mut lo = -1e9_f64;
        let mut hi = 1e9_f64;
        let mut result: Option<f64> = None;
        for _ in 0..80 {
            let mid = (lo + hi) / 2.0;
            let cell = CellData {
                value: mid.to_string(),
                formula: None,
                format: None,
                comment: None,
                spill_anchor: None,
            };
            scratch.sheets[active_idx].set_cell(i.0, i.1, cell);
            let cur: f64 = scratch.sheets[active_idx]
                .get_cell(t.0, t.1)
                .value
                .parse()
                .unwrap_or(0.0);
            if (cur - expected_v).abs() < 1e-6 {
                result = Some(mid);
                break;
            }
            if cur < expected_v {
                lo = mid;
            } else {
                hi = mid;
            }
        }
        if let Some(v) = result {
            let final_cell = CellData {
                value: v.to_string(),
                formula: None,
                format: None,
                comment: None,
                spill_anchor: None,
            };
            self.set_cell_with_undo(i.0, i.1, final_cell);
            self.status_message = Some(format!("Goal seek: {} = {:.6}", input, v));
        } else {
            self.status_message = Some("Goal seek did not converge".to_string());
        }
    }
}

impl App {
    /// `:pivot <source> <target> [row=COL] [value=COL] [agg=sum|count|avg|min|max]`
    ///
    /// Writes a one-key pivot table to `target` whose cells are formula
    /// references back into the source range, so subsequent edits to the
    /// source data auto-refresh the pivot.
    pub(super) fn cmd_pivot(&mut self, source: &str, target_str: &str, opts: &[&str]) {
        let (Some((s, e)), Some(t)) = (
            parse_range(source),
            Spreadsheet::parse_cell_reference(target_str),
        ) else {
            self.status_message = Some("pivot: bad source range or target".to_string());
            return;
        };

        let mut row_col: Option<usize> = None;
        let mut val_col: Option<usize> = None;
        let mut agg = "sum".to_string();
        for o in opts {
            if let Some(c) = o.strip_prefix("row=") {
                row_col = Spreadsheet::parse_column_label(c);
            } else if let Some(c) = o.strip_prefix("value=") {
                val_col = Spreadsheet::parse_column_label(c);
            } else if let Some(a) = o.strip_prefix("agg=") {
                agg = a.to_string();
            }
        }
        let Some(row_col) = row_col else {
            self.status_message = Some("pivot: row=COL required".to_string());
            return;
        };
        let val_col = val_col.unwrap_or(row_col);

        // Collect distinct row-keys preserving sort order. The pivot values
        // are written as formulas (`=SUMIF(...)`) so they auto-update when
        // source data changes.
        let mut keys: std::collections::BTreeSet<String> = std::collections::BTreeSet::new();
        {
            let sheet = self.workbook.current_sheet();
            for r in s.0..=e.0 {
                let k = sheet.get_cell(r, row_col).value;
                if !k.is_empty() {
                    keys.insert(k);
                }
            }
        }
        let row_col_label = Spreadsheet::column_label(row_col);
        let val_col_label = Spreadsheet::column_label(val_col);
        let source_range = format!("{}{}:{}{}", row_col_label, s.0 + 1, row_col_label, e.0 + 1);
        let sum_range = format!("{}{}:{}{}", val_col_label, s.0 + 1, val_col_label, e.0 + 1);
        let agg_formula = |key: &str| -> String {
            let key_esc = key.replace('"', "\"\"");
            match agg.as_str() {
                "count" => format!("=COUNTIF({}, \"{}\")", source_range, key_esc),
                "avg" | "average" => {
                    format!("=AVERAGEIF({}, \"{}\", {})", source_range, key_esc, sum_range)
                }
                "min" => format!("=MIN(IF({}=\"{}\", {}))", source_range, key_esc, sum_range),
                "max" => format!("=MAX(IF({}=\"{}\", {}))", source_range, key_esc, sum_range),
                _ => format!("=SUMIF({}, \"{}\", {})", source_range, key_esc, sum_range),
            }
        };

        let mut rows: Vec<(usize, usize, CellData)> = Vec::new();
        rows.push((
            t.0,
            t.1,
            CellData {
                value: "Key".to_string(),
                formula: None,
                format: None,
                comment: None,
                spill_anchor: None,
            },
        ));
        rows.push((
            t.0,
            t.1 + 1,
            CellData {
                value: agg.clone(),
                formula: None,
                format: None,
                comment: None,
                spill_anchor: None,
            },
        ));
        for (i, k) in keys.iter().enumerate() {
            rows.push((
                t.0 + 1 + i,
                t.1,
                CellData {
                    value: k.clone(),
                    formula: None,
                    format: None,
                    comment: None,
                    spill_anchor: None,
                },
            ));
            let formula = agg_formula(k);
            // Evaluate immediately so the cell shows its initial value;
            // set_many will rebuild deps so it tracks.
            let evaluator = FormulaEvaluator::for_workbook(
                &self.workbook,
                self.workbook.current_sheet(),
                &self.workbook.named_ranges,
            );
            // Publish a clock for the same NOW()/TODAY() consistency reason
            // as the editing/autofill/paste paths.
            let initial = crate::domain::parser::with_recalc_clock(
                crate::domain::parser::now_serial(),
                || evaluator.evaluate_formula(&formula),
            );
            rows.push((
                t.0 + 1 + i,
                t.1 + 1,
                CellData {
                    value: initial,
                    formula: Some(formula),
                    format: None,
                    comment: None,
                    spill_anchor: None,
                },
            ));
        }
        self.set_many_with_undo(rows);
        self.status_message = Some(format!(
            "Pivot written to {}{}: {} groups (auto-refreshes via formulas)",
            Spreadsheet::column_label(t.1),
            t.0 + 1,
            keys.len()
        ));
    }
}
