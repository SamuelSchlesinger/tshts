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
        // Walk the sparse `cells` map (not rows×cols) and match against value
        // OR formula text. Previously this only matched `value`, so formulas
        // that contained the needle were invisible to :replace even though
        // /-search (`perform_search`) did match them — asymmetric and
        // confusing.
        let sheet = self.workbook.current_sheet();
        for (&(row, col), cell) in &sheet.cells {
            let formula_hit = cell
                .formula
                .as_ref()
                .is_some_and(|f| matcher.is_match(f));
            if matcher.is_match(&cell.value) || formula_hit {
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
        // Refresh the match list (the replacement may itself still match,
        // or may no longer match) and resume from the cell right after the
        // one we just replaced — otherwise repeated Replace would loop on
        // results[0] instead of moving forward through the document.
        self.find_replace_search();
        let next_idx = self
            .find_replace_results
            .iter()
            .position(|&(r, c)| (r, c) > (row, col));
        self.find_replace_index = next_idx.unwrap_or(0);
        if let Some(&(r, c)) = self.find_replace_results.get(self.find_replace_index) {
            self.selected_row = r;
            self.selected_col = c;
            self.ensure_cursor_visible();
        }
    }

    pub fn replace_all(&mut self) {
        if self.find_replace_results.is_empty() {
            return;
        }
        // Any operation user-initiated as "replace_all" should clear redo so
        // stale forward-history can't be re-applied. Even when every match is
        // a formula cell (skipped below) we still treat this as a state
        // transition.
        self.redo_stack.clear();
        let matcher = TextMatcher::new(
            &self.find_replace_search,
            self.search_regex,
            self.search_case_sensitive,
        );
        let mut batch = Vec::new();
        let mut writes: Vec<(usize, usize, CellData)> = Vec::new();
        let mut skipped_formulas = 0usize;
        let results = self.find_replace_results.clone();
        for (row, col) in results {
            let cell = self.workbook.current_sheet().get_cell(row, col);
            if cell.formula.is_some() {
                // Excel/Sheets don't rewrite formula text via Replace. Track
                // the count so we can tell the user why fewer cells changed
                // than the result list suggested.
                skipped_formulas += 1;
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
            writes.push((row, col, new_cell));
        }
        let count = writes.len();
        // Single workbook API call handles same-sheet recalc + cross-sheet
        // propagation for the whole batch.
        self.workbook.write_cells_on_active(writes);
        if !batch.is_empty() {
            self.record_action(UndoAction::Batch(batch));
        }
        let msg = if skipped_formulas > 0 {
            format!("Replaced {} occurrence(s); skipped {} formula cell(s)", count, skipped_formulas)
        } else {
            format!("Replaced {} occurrence(s)", count)
        };
        self.status_message = Some(msg);
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
        // Accept sheet-qualified refs (`Sheet2!A1`) so :goto can jump
        // across sheets, not just inside the active one. parse_qualified_reference
        // returns (sheet_opt, row, col, abs_row, abs_col).
        let parsed = crate::domain::Spreadsheet::parse_qualified_reference(&self.goto_cell_input);
        if let Some((sheet_opt, row, col, _, _)) = parsed {
            // Resolve the target sheet (current sheet if unqualified, by
            // case-insensitive lookup otherwise).
            let target_sheet_idx = if let Some(name) = sheet_opt.as_ref() {
                self.workbook
                    .sheet_names
                    .iter()
                    .position(|n| n.eq_ignore_ascii_case(name))
            } else {
                Some(self.workbook.active_sheet)
            };
            let Some(target_idx) = target_sheet_idx else {
                self.status_message = Some(format!(
                    "Unknown sheet: {}",
                    sheet_opt.as_deref().unwrap_or("")
                ));
                self.mode = AppMode::Normal;
                self.goto_cell_input.clear();
                self.cursor_position = 0;
                return;
            };
            let in_bounds = {
                let sheet = &self.workbook.sheets[target_idx];
                row < sheet.rows && col < sheet.cols
            };
            if in_bounds {
                // Cross-sheet jump: switch the active sheet and invalidate
                // (row, col)-keyed state that doesn't survive a sheet change.
                if target_idx != self.workbook.active_sheet {
                    self.workbook.active_sheet = target_idx;
                    self.scroll_row = 0;
                    self.scroll_col = 0;
                    self.clear_selection();
                    self.invalidate_cross_sheet_state();
                }
                self.selected_row = row;
                self.selected_col = col;
                self.ensure_cursor_visible();
                self.status_message = Some(format!(
                    "Jumped to {}{}{}",
                    sheet_opt
                        .as_ref()
                        .map(|n| format!("{}!", n))
                        .unwrap_or_default(),
                    crate::domain::Spreadsheet::column_label(col),
                    row + 1
                ));
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

#[cfg(test)]
mod tests {
    use super::*;
    use crate::domain::CellData;

    #[test]
    fn test_regex_search_toggle() {
        let mut app = App::default();
        app.workbook.current_sheet_mut().set_cell(0, 0, crate::domain::CellData {
            value: "foo123".to_string(), formula: None, format: None, comment: None,
        spill_anchor: None,
        });
        app.workbook.current_sheet_mut().set_cell(1, 0, crate::domain::CellData {
            value: "FOO456".to_string(), formula: None, format: None, comment: None,
        spill_anchor: None,
        });
        // Default: case-insensitive substring matches both.
        app.search_query = "foo".to_string();
        app.perform_search();
        assert_eq!(app.search_results.len(), 2);

        // Case-sensitive: only first matches.
        app.search_case_sensitive = true;
        app.perform_search();
        assert_eq!(app.search_results.len(), 1);

        // Regex: anchored digit match.
        app.search_case_sensitive = false;
        app.search_regex = true;
        app.search_query = "[0-9]+$".to_string();
        app.perform_search();
        assert_eq!(app.search_results.len(), 2);
    }

    #[test]
    fn test_find_replace_basic() {
        let mut app = App::default();
        app.set_cell_with_undo(0, 0, CellData { value: "hello world".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        app.set_cell_with_undo(1, 0, CellData { value: "hello there".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        app.set_cell_with_undo(2, 0, CellData { value: "goodbye".to_string(), formula: None, format: None, comment: None, spill_anchor: None });

        app.start_find_replace();
        assert!(matches!(app.mode, AppMode::FindReplace));

        app.find_replace_search = "hello".to_string();
        app.find_replace_search();

        assert_eq!(app.find_replace_results.len(), 2);
    }

    #[test]
    fn test_replace_current() {
        let mut app = App::default();
        app.set_cell_with_undo(0, 0, CellData { value: "hello world".to_string(), formula: None, format: None, comment: None, spill_anchor: None });

        app.start_find_replace();
        app.find_replace_search = "hello".to_string();
        app.find_replace_replace = "hi".to_string();
        app.find_replace_search();

        assert_eq!(app.find_replace_results.len(), 1);

        app.replace_current();

        assert_eq!(app.workbook.current_sheet().get_cell(0, 0).value, "hi world");
    }

    #[test]
    fn test_replace_all() {
        let mut app = App::default();
        app.set_cell_with_undo(0, 0, CellData { value: "cat".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        app.set_cell_with_undo(1, 0, CellData { value: "cat food".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        app.set_cell_with_undo(2, 0, CellData { value: "dog".to_string(), formula: None, format: None, comment: None, spill_anchor: None });

        app.start_find_replace();
        app.find_replace_search = "cat".to_string();
        app.find_replace_replace = "kitten".to_string();
        app.find_replace_search();

        app.replace_all();

        assert_eq!(app.workbook.current_sheet().get_cell(0, 0).value, "kitten");
        assert_eq!(app.workbook.current_sheet().get_cell(1, 0).value, "kitten food");
        assert_eq!(app.workbook.current_sheet().get_cell(2, 0).value, "dog"); // Unchanged
    }

    #[test]
    fn test_replace_skips_formula_cells() {
        let mut app = App::default();
        app.set_cell_with_undo(0, 0, CellData {
            value: "hello".to_string(),
            formula: Some("=A2".to_string()),
            format: None,
            comment: None,
        spill_anchor: None,
        });

        app.start_find_replace();
        app.find_replace_search = "hello".to_string();
        app.find_replace_replace = "bye".to_string();
        app.find_replace_search();

        // Should find the cell but not replace it
        app.replace_current();

        // Formula cell value should be unchanged (replace_current skips formula cells)
        assert_eq!(app.workbook.current_sheet().get_cell(0, 0).formula, Some("=A2".to_string()));
    }

    #[test]
    fn test_goto_cell() {
        let mut app = App::default();
        app.start_goto_cell();
        assert!(matches!(app.mode, AppMode::GoToCell));

        app.goto_cell_input = "C5".to_string();
        app.finish_goto_cell();

        assert!(matches!(app.mode, AppMode::Normal));
        assert_eq!(app.selected_row, 4); // 0-indexed
        assert_eq!(app.selected_col, 2); // C = index 2
    }

    #[test]
    fn test_goto_cell_cross_sheet_switches_active() {
        let mut app = App::default();
        app.workbook.add_sheet("Data".to_string());
        // Seed a search result on Sheet1 to verify cross-sheet invalidation.
        app.search_results.push((9, 9));
        app.search_result_index = 0;
        assert_eq!(app.workbook.active_sheet, 0);
        app.start_goto_cell();
        app.goto_cell_input = "Data!B3".to_string();
        app.finish_goto_cell();
        assert!(matches!(app.mode, AppMode::Normal));
        assert_eq!(app.workbook.active_sheet, 1, "must switch to Data sheet");
        assert_eq!(app.selected_row, 2);
        assert_eq!(app.selected_col, 1);
        // Cross-sheet jump should invalidate stale search results.
        assert!(app.search_results.is_empty(), "search results must be cleared on cross-sheet goto");
    }

    #[test]
    fn test_goto_cell_unknown_sheet_errors_gracefully() {
        let mut app = App::default();
        app.start_goto_cell();
        app.goto_cell_input = "NoSuchSheet!A1".to_string();
        app.finish_goto_cell();
        assert!(matches!(app.mode, AppMode::Normal));
        assert_eq!(app.workbook.active_sheet, 0, "must not switch on bogus sheet");
        let msg = app.status_message.as_deref().unwrap_or("");
        assert!(
            msg.contains("Unknown sheet") || msg.contains("NoSuchSheet"),
            "expected an error message about the unknown sheet, got: {}",
            msg
        );
    }

    #[test]
    fn test_goto_cell_invalid() {
        let mut app = App::default();
        app.start_goto_cell();

        app.goto_cell_input = "invalid".to_string();
        app.finish_goto_cell();

        assert!(matches!(app.mode, AppMode::Normal));
        assert!(app.status_message.as_ref().unwrap().contains("Invalid cell reference"));
        assert_eq!(app.selected_row, 0); // Unchanged
        assert_eq!(app.selected_col, 0);
    }

    #[test]
    fn test_goto_cell_cancel() {
        let mut app = App::default();
        app.selected_row = 5;
        app.selected_col = 3;
        app.start_goto_cell();

        app.goto_cell_input = "A1".to_string();
        app.cancel_goto_cell();

        assert!(matches!(app.mode, AppMode::Normal));
        assert_eq!(app.selected_row, 5); // Unchanged
        assert_eq!(app.selected_col, 3);
    }

}
