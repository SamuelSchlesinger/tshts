//! Submodule of `state` — see state/mod.rs.

use super::*;

impl App {
    pub fn start_search(&mut self) {
        self.mode = AppMode::Search;
        self.search_query.clear();
        self.search_results.clear();
        self.search_result_index = 0;
        self.cursor_position = 0;
        self.status_message = None;
    }

    pub fn cancel_search(&mut self) {
        self.mode = AppMode::Normal;
        self.search_query.clear();
        self.search_results.clear();
        self.search_result_index = 0;
        self.cursor_position = 0;
    }

    pub fn perform_search(&mut self) {
        self.search_results.clear();
        self.search_result_index = 0;

        if self.search_query.is_empty() {
            return;
        }

        let matcher = TextMatcher::new(
            &self.search_query,
            self.search_regex,
            self.search_case_sensitive,
        );

        for (&(row, col), cell) in &self.workbook.current_sheet().cells {
            let value_matches = matcher.is_match(&cell.value);
            let formula_matches = cell
                .formula
                .as_ref()
                .map(|f| matcher.is_match(f))
                .unwrap_or(false);
            if value_matches || formula_matches {
                self.search_results.push((row, col));
            }
        }
        self.search_results.sort();
        if !self.search_results.is_empty() {
            self.go_to_current_search_result();
        }
    }

    pub fn next_search_result(&mut self) {
        if !self.search_results.is_empty() {
            self.search_result_index = (self.search_result_index + 1) % self.search_results.len();
            self.go_to_current_search_result();
        }
    }

    pub fn previous_search_result(&mut self) {
        if !self.search_results.is_empty() {
            if self.search_result_index == 0 {
                self.search_result_index = self.search_results.len() - 1;
            } else {
                self.search_result_index -= 1;
            }
            self.go_to_current_search_result();
        }
    }

    fn go_to_current_search_result(&mut self) {
        if let Some(&(row, col)) = self.search_results.get(self.search_result_index) {
            self.selected_row = row;
            self.selected_col = col;
            self.ensure_cursor_visible();
        }
    }

    pub fn finish_search(&mut self) {
        self.mode = AppMode::Normal;

        let num_results = self.search_results.len();
        if num_results > 0 {
            self.status_message = Some(format!(
                "Search completed: {} result{} found for '{}' (n/N to navigate)",
                num_results,
                if num_results == 1 { "" } else { "s" },
                self.search_query
            ));
        } else {
            self.status_message = Some(format!("No results found for '{}'", self.search_query));
        }

        self.search_query.clear();
        // Don't clear search_results — keep them for n/N navigation
        self.cursor_position = 0;
    }

    pub fn start_find_replace(&mut self) {
        self.mode = AppMode::FindReplace;
        self.find_replace_search.clear();
        self.find_replace_replace.clear();
        self.find_replace_on_replace = false;
        self.find_replace_results.clear();
        self.find_replace_index = 0;
        self.cursor_position = 0;
        self.status_message = None;
    }

    pub fn find_replace_search(&mut self) {
        self.find_replace_results.clear();
        self.find_replace_index = 0;
        if self.find_replace_search.is_empty() {
            return;
        }
        let matcher = TextMatcher::new(
            &self.find_replace_search,
            self.search_regex,
            self.search_case_sensitive,
        );
        for row in 0..self.workbook.current_sheet().rows {
            for col in 0..self.workbook.current_sheet().cols {
                let cell = self.workbook.current_sheet().get_cell(row, col);
                if matcher.is_match(&cell.value) {
                    self.find_replace_results.push((row, col));
                }
            }
        }
        if !self.find_replace_results.is_empty() {
            let (row, col) = self.find_replace_results[0];
            self.selected_row = row;
            self.selected_col = col;
            self.ensure_cursor_visible();
        }
    }

    pub fn replace_current(&mut self) {
        if self.find_replace_results.is_empty() {
            return;
        }
        let (row, col) = self.find_replace_results[self.find_replace_index];
        let cell = self.workbook.current_sheet().get_cell(row, col);
        if cell.formula.is_some() {
            return;
        }
        let matcher = TextMatcher::new(
            &self.find_replace_search,
            self.search_regex,
            self.search_case_sensitive,
        );
        let new_value = matcher.replace_all(&cell.value, &self.find_replace_replace);
        let new_cell = CellData {
            value: new_value,
            formula: None,
            format: cell.format.clone(),
            comment: cell.comment.clone(),
        spill_anchor: None,
        };
        self.set_cell_with_undo(row, col, new_cell);
        self.find_replace_search();
    }

    pub fn replace_all(&mut self) {
        if self.find_replace_results.is_empty() {
            return;
        }
        let matcher = TextMatcher::new(
            &self.find_replace_search,
            self.search_regex,
            self.search_case_sensitive,
        );
        let mut batch = Vec::new();
        let results = self.find_replace_results.clone();
        for (row, col) in results {
            let cell = self.workbook.current_sheet().get_cell(row, col);
            if cell.formula.is_some() {
                continue;
            }
            let old_cell = Some(cell.clone());
            let new_value = matcher.replace_all(&cell.value, &self.find_replace_replace);
            let new_cell = CellData {
                value: new_value,
                formula: None,
                format: cell.format.clone(),
                comment: cell.comment.clone(),
            spill_anchor: None,
            };
            batch.push(UndoAction::CellModified {
                row,
                col,
                old_cell,
                new_cell: Some(new_cell.clone()),
            });
            self.workbook.current_sheet_mut().set_cell(row, col, new_cell);
        }
        let count = batch.len();
        if !batch.is_empty() {
            self.record_action(UndoAction::Batch(batch));
        }
        self.status_message = Some(format!("Replaced {} occurrence(s)", count));
        self.find_replace_results.clear();
    }

    pub fn finish_find_replace(&mut self) {
        self.mode = AppMode::Normal;
        self.find_replace_search.clear();
        self.find_replace_replace.clear();
        self.find_replace_results.clear();
        self.cursor_position = 0;
    }

    pub fn start_goto_cell(&mut self) {
        self.mode = AppMode::GoToCell;
        self.goto_cell_input.clear();
        self.cursor_position = 0;
        self.status_message = None;
    }

    pub fn finish_goto_cell(&mut self) {
        if let Some((row, col)) = crate::domain::Spreadsheet::parse_cell_reference(&self.goto_cell_input) {
            if row < self.workbook.current_sheet().rows && col < self.workbook.current_sheet().cols {
                self.selected_row = row;
                self.selected_col = col;
                self.ensure_cursor_visible();
                self.status_message = Some(format!("Jumped to {}{}", crate::domain::Spreadsheet::column_label(col), row + 1));
            } else {
                self.status_message = Some("Cell reference out of range".to_string());
            }
        } else {
            self.status_message = Some("Invalid cell reference".to_string());
        }
        self.mode = AppMode::Normal;
        self.goto_cell_input.clear();
        self.cursor_position = 0;
    }

    pub fn cancel_goto_cell(&mut self) {
        self.mode = AppMode::Normal;
        self.goto_cell_input.clear();
        self.cursor_position = 0;
    }

}
