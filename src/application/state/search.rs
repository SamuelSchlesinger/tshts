//! Search, find-and-replace, and go-to-cell state transitions.

use crate::domain::{CellData, Spreadsheet};
use super::{App, AppMode, UndoAction, case_insensitive_replace};

impl App {
    /// Starts search mode and initializes search state.
    pub fn start_search(&mut self) {
        self.mode = AppMode::Search;
        self.search_query.clear();
        self.search_results.clear();
        self.search_result_index = 0;
        self.cursor_position = 0;
        self.status_message = None;
    }

    /// Cancels search mode and returns to normal mode.
    pub fn cancel_search(&mut self) {
        self.mode = AppMode::Normal;
        self.search_query.clear();
        self.search_results.clear();
        self.search_result_index = 0;
        self.cursor_position = 0;
    }

    /// Performs a search across all cells and updates search results.
    pub fn perform_search(&mut self) {
        self.search_results.clear();
        self.search_result_index = 0;

        if self.search_query.is_empty() {
            return;
        }

        let query_lower = self.search_query.to_lowercase();

        for (&(row, col), cell) in &self.workbook.current_sheet().cells {
            let value_matches = cell.value.to_lowercase().contains(&query_lower);
            let formula_matches = cell.formula
                .as_ref()
                .map(|f| f.to_lowercase().contains(&query_lower))
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

    /// Moves to the next search result.
    pub fn next_search_result(&mut self) {
        if !self.search_results.is_empty() {
            self.search_result_index = (self.search_result_index + 1) % self.search_results.len();
            self.go_to_current_search_result();
        }
    }

    /// Moves to the previous search result.
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

    /// Finishes search and returns to normal mode while keeping the current selection.
    /// Search results are preserved for n/N navigation in normal mode.
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
        self.cursor_position = 0;
    }

    /// Starts find and replace mode.
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

    /// Performs find for find-and-replace.
    pub fn find_replace_search(&mut self) {
        self.find_replace_results.clear();
        self.find_replace_index = 0;
        if self.find_replace_search.is_empty() { return; }
        let query = self.find_replace_search.to_lowercase();
        for (&(row, col), cell) in &self.workbook.current_sheet().cells {
            if cell.value.to_lowercase().contains(&query) {
                self.find_replace_results.push((row, col));
            }
        }
        self.find_replace_results.sort();
        if !self.find_replace_results.is_empty() {
            let (row, col) = self.find_replace_results[0];
            self.selected_row = row;
            self.selected_col = col;
            self.ensure_cursor_visible();
        }
    }

    /// Replaces the current find-replace match.
    pub fn replace_current(&mut self) {
        if self.find_replace_results.is_empty() { return; }
        let (row, col) = self.find_replace_results[self.find_replace_index];
        let cell = self.workbook.current_sheet().get_cell(row, col);
        if cell.formula.is_some() { return; }
        let new_value = case_insensitive_replace(&cell.value, &self.find_replace_search, &self.find_replace_replace);
        let new_cell = CellData { value: new_value, formula: None, format: cell.format.clone(), comment: cell.comment.clone() };
        self.set_cell_with_undo(row, col, new_cell);
        self.find_replace_search();
    }

    /// Replaces all find-replace matches.
    pub fn replace_all(&mut self) {
        if self.find_replace_results.is_empty() { return; }
        let mut batch = Vec::new();
        let results = self.find_replace_results.clone();
        for (row, col) in results {
            let cell = self.workbook.current_sheet().get_cell(row, col);
            if cell.formula.is_some() { continue; }
            let old_cell = Some(cell.clone());
            let new_value = case_insensitive_replace(&cell.value, &self.find_replace_search, &self.find_replace_replace);
            let new_cell = CellData { value: new_value, formula: None, format: cell.format.clone(), comment: cell.comment.clone() };
            batch.push(UndoAction::CellModified { row, col, old_cell, new_cell: Some(new_cell.clone()) });
            self.workbook.current_sheet_mut().set_cell(row, col, new_cell);
        }
        let count = batch.len();
        if !batch.is_empty() {
            self.record_action(UndoAction::Batch(batch));
        }
        self.status_message = Some(format!("Replaced {} occurrence(s)", count));
        self.find_replace_results.clear();
    }

    /// Finishes find-replace mode.
    pub fn finish_find_replace(&mut self) {
        self.mode = AppMode::Normal;
        self.find_replace_search.clear();
        self.find_replace_replace.clear();
        self.find_replace_results.clear();
        self.cursor_position = 0;
    }

    /// Starts go-to cell mode.
    pub fn start_goto_cell(&mut self) {
        self.mode = AppMode::GoToCell;
        self.goto_cell_input.clear();
        self.cursor_position = 0;
        self.status_message = None;
    }

    /// Finishes go-to cell and navigates to the entered cell reference.
    pub fn finish_goto_cell(&mut self) {
        if let Some((row, col)) = Spreadsheet::parse_cell_reference(&self.goto_cell_input) {
            if row < self.workbook.current_sheet().rows && col < self.workbook.current_sheet().cols {
                self.selected_row = row;
                self.selected_col = col;
                self.ensure_cursor_visible();
                self.status_message = Some(format!("Jumped to {}{}", Spreadsheet::column_label(col), row + 1));
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

    /// Cancels go-to cell mode.
    pub fn cancel_goto_cell(&mut self) {
        self.mode = AppMode::Normal;
        self.goto_cell_input.clear();
        self.cursor_position = 0;
    }
}
