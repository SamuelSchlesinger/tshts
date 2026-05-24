//! Helper methods for analysis commands extracted from `execute_command`:
//! `:trace`, `:trace dependents`, `:cf <col> <predicate> ...`.

use super::super::{App, UndoAction};
use crate::domain::{FormulaEvaluator, Spreadsheet, TerminalColor};

impl App {
    /// `:trace` / `:trace precedents` — list the cells the current cell's
    /// formula reads. Uses the same-sheet evaluator so cross-sheet refs are
    /// listed verbatim (they don't resolve here; that's `:trace dependents`).
    pub(super) fn cmd_trace_precedents(&mut self) {
        let cell = self
            .workbook
            .current_sheet()
            .get_cell(self.selected_row, self.selected_col);
        let Some(formula) = cell.formula else {
            self.status_message = Some("(cell has no formula)".to_string());
            return;
        };
        let evaluator = FormulaEvaluator::new(self.workbook.current_sheet());
        let refs = evaluator.extract_cell_references(&formula);
        if refs.is_empty() {
            self.status_message = Some("(no precedents)".to_string());
            return;
        }
        let refs_str: Vec<String> = refs
            .iter()
            .map(|(r, c)| format!("{}{}", Spreadsheet::column_label(*c), r + 1))
            .collect();
        self.status_message = Some(format!("Precedents: {}", refs_str.join(", ")));
    }

    /// `:trace dependents` — walk the unified workbook graph for cells that
    /// read the current cell. The graph is keyed by `NodeKey = (SheetId,
    /// row, col)` and is the single source of truth post-cascade-removal.
    pub(super) fn cmd_trace_dependents(&mut self) {
        let sheet_idx = self.workbook.active_sheet;
        let sheet_id = self.workbook.sheet_ids[sheet_idx];
        let node = (sheet_id, self.selected_row, self.selected_col);
        let mut seed = std::collections::HashSet::new();
        seed.insert(node);
        let downstream: Vec<_> = self
            .workbook
            .graph
            .transitive_dependents(&seed)
            .into_iter()
            .filter(|d| *d != node) // strip self
            .collect();
        if downstream.is_empty() {
            self.status_message = Some("(no dependents)".to_string());
            return;
        }
        let mut labels: Vec<String> = downstream
            .iter()
            .map(|(sid, r, c)| {
                let sheet_name = self
                    .workbook
                    .sheet_ids
                    .iter()
                    .position(|s| s == sid)
                    .map(|i| self.workbook.sheet_names[i].clone())
                    .unwrap_or_else(|| "?".to_string());
                let same_sheet = *sid == sheet_id;
                let cell_label = format!("{}{}", Spreadsheet::column_label(*c), r + 1);
                if same_sheet {
                    cell_label
                } else {
                    format!("{}!{}", sheet_name, cell_label)
                }
            })
            .collect();
        labels.sort();
        self.status_message = Some(format!("Dependents: {}", labels.join(", ")));
    }

    /// `:cf <col> <predicate> [bg=COLOR] [fg=COLOR] [bold] [underline]`
    ///
    /// The predicate may use `_` for the cell value. We accept multiple
    /// whitespace-separated tokens for the predicate (joining all that
    /// don't look like style keys), because the palette splits on
    /// whitespace.
    pub(super) fn cmd_cf_set(&mut self, col_name: &str, rest: &[&str]) {
        let Some(col) = Spreadsheet::parse_column_label(col_name) else {
            self.status_message = Some(format!("Invalid column: {}", col_name));
            return;
        };
        let mut predicate_parts = Vec::new();
        let mut style = crate::domain::CellStyle::default();
        for tok in rest {
            if let Some(c) = tok.strip_prefix("bg=") {
                style.bg_color = TerminalColor::from_name(c);
            } else if let Some(c) = tok.strip_prefix("fg=") {
                style.fg_color = TerminalColor::from_name(c);
            } else if *tok == "bold" {
                style.bold = true;
            } else if *tok == "underline" {
                style.underline = true;
            } else {
                predicate_parts.push(*tok);
            }
        }
        if predicate_parts.is_empty() {
            self.status_message = Some(
                "Usage: cf <col> <predicate> [bg=color] [fg=color] [bold] [underline]".to_string(),
            );
            return;
        }
        let predicate = predicate_parts.join(" ");
        let sheet_idx = self.workbook.active_sheet;
        let old = self.workbook.current_sheet().conditional_formats.clone();
        let mut new = old.clone();
        new.push(crate::domain::ConditionalFormat {
            column: col,
            predicate: predicate.clone(),
            style,
        });
        {
            let sheet = self.workbook.current_sheet_mut();
            sheet.conditional_formats = new.clone();
            sheet.cf_cache.lock().unwrap().clear();
        }
        self.record_action(UndoAction::ConditionalFormatsReplaced {
            sheet_idx,
            old,
            new,
        });
        self.status_message = Some(format!(
            "Added cf for col {}: {}",
            Spreadsheet::column_label(col),
            predicate
        ));
    }
}
