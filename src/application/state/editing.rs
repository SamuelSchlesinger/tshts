//! Submodule of `state` — see state/mod.rs.

use super::*;

impl App {
    pub fn start_editing(&mut self) {
        let cell = self.workbook.current_sheet().get_cell(self.selected_row, self.selected_col);
        if let Some(anchor) = cell.spill_anchor {
            self.status_message = Some(format!(
                "Read-only spill cell (anchor at {}{}). Edit the anchor to change it.",
                crate::domain::Spreadsheet::column_label(anchor.1),
                anchor.0 + 1
            ));
            return;
        }
        self.mode = AppMode::Editing;
        self.input = cell.formula.unwrap_or(cell.value);
        self.cursor_position = self.input.chars().count();
    }

    pub fn vim_enter_insert(&mut self) {
        self.start_editing();
        self.cursor_position = 0;
    }

    pub fn vim_enter_insert_at_end(&mut self) {
        self.start_editing();
        self.cursor_position = self.input.chars().count();
    }

    pub fn vim_open_row_below(&mut self) {
        let at = self.selected_row + 1;
        let sheet_idx = self.workbook.active_sheet;
        // Route through the Workbook so cross-sheet refs to rows >= `at`
        // shift along with the same-sheet ones.
        self.workbook.insert_row_on_active(at);
        self.record_action(UndoAction::RowInserted { sheet_idx, at });
        self.selected_row = at;
        self.ensure_cursor_visible();
        self.start_editing();
    }

    pub fn vim_open_row_above(&mut self) {
        let at = self.selected_row;
        let sheet_idx = self.workbook.active_sheet;
        self.workbook.insert_row_on_active(at);
        self.record_action(UndoAction::RowInserted { sheet_idx, at });
        self.ensure_cursor_visible();
        self.start_editing();
    }

    pub fn vim_substitute_cell(&mut self) {
        self.clear_cell_with_undo(self.selected_row, self.selected_col);
        self.start_editing();
    }

    pub fn vim_substitute_row(&mut self) {
        self.vim_apply_line_op(VimOperator::Delete, 1);
        self.selected_col = 0;
        self.start_editing();
    }

    fn move_after_edit(&mut self, dir: EditExitDir) {
        match dir {
            EditExitDir::Down => {
                if self.selected_row < self.workbook.current_sheet().rows - 1 {
                    self.selected_row += 1;
                }
            }
            EditExitDir::Up => {
                if self.selected_row > 0 {
                    self.selected_row -= 1;
                }
            }
            EditExitDir::Right => {
                if self.selected_col < self.workbook.current_sheet().cols - 1 {
                    self.selected_col += 1;
                }
            }
            EditExitDir::Left => {
                if self.selected_col > 0 {
                    self.selected_col -= 1;
                }
            }
        }
    }

    fn finish_editing_in_direction(&mut self, dir: EditExitDir) {
        let existing = self
            .workbook
            .current_sheet()
            .get_cell(self.selected_row, self.selected_col);
        let mut cell_data = CellData {
            format: existing.format.clone(),
            comment: existing.comment.clone(),
            ..CellData::default()
        };

        if self.input.starts_with('=') {
            let names = self.workbook.named_ranges.clone();
            let evaluator = FormulaEvaluator::for_workbook(
                &self.workbook,
                self.workbook.current_sheet(),
                &names,
            );
            // Same-sheet cycle check.
            if !self.iterative_calc
                && evaluator.would_create_circular_reference(
                    &self.input,
                    (self.selected_row, self.selected_col),
                )
            {
                // Reject — but still exit editing mode so the user isn't stuck.
                self.status_message = Some("Circular reference rejected".to_string());
                self.mode = AppMode::Normal;
                self.input.clear();
                self.cursor_position = 0;
                return;
            }
            // Cross-sheet cycle check: walk the workbook graph from the new
            // formula's precedents.
            if !self.iterative_calc {
                let precedents = evaluator.extract_qualified_refs(&self.input);
                let sheet_name = self
                    .workbook
                    .sheet_names[self.workbook.active_sheet]
                    .clone();
                if self.workbook.would_create_cross_sheet_cycle(
                    &sheet_name,
                    self.selected_row,
                    self.selected_col,
                    &precedents,
                ) {
                    self.status_message =
                        Some("Cross-sheet circular reference rejected".to_string());
                    self.mode = AppMode::Normal;
                    self.input.clear();
                    self.cursor_position = 0;
                    return;
                }
            }
            cell_data.formula = Some(self.input.clone());
            cell_data.value = evaluator.evaluate_formula(&self.input);
        } else {
            cell_data.value = self.input.clone();
        }

        self.set_cell_with_undo(self.selected_row, self.selected_col, cell_data);
        self.move_after_edit(dir);
        self.mode = AppMode::Normal;
        self.input.clear();
        self.cursor_position = 0;
    }

    pub fn finish_editing(&mut self) {
        self.finish_editing_in_direction(EditExitDir::Down);
    }

    pub fn finish_editing_move_right(&mut self) {
        self.finish_editing_in_direction(EditExitDir::Right);
    }

    pub fn finish_editing_move_left(&mut self) {
        self.finish_editing_in_direction(EditExitDir::Left);
    }

    /// Commit the edit and move the cursor up — used by the Up arrow in
    /// Editing mode, matching Excel/Sheets behavior where pressing Up
    /// confirms the value and goes to the cell above.
    pub fn finish_editing_move_up(&mut self) {
        self.finish_editing_in_direction(EditExitDir::Up);
    }

    /// Commit the edit and move the cursor down — alias for `finish_editing`
    /// for symmetry with the Up/Right variants. Provided so the Down arrow
    /// path is self-documenting at the call site.
    pub fn finish_editing_move_down(&mut self) {
        self.finish_editing_in_direction(EditExitDir::Down);
    }

    pub fn cancel_editing(&mut self) {
        self.mode = AppMode::Normal;
        self.input.clear();
        self.cursor_position = 0;
    }

}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::domain::CellData;

    #[test]
    fn test_start_editing_empty_cell() {
        let mut app = App::default();
        app.start_editing();
        
        assert!(matches!(app.mode, AppMode::Editing));
        assert!(app.input.is_empty()); // Empty cell should give empty input
        assert_eq!(app.cursor_position, 0);
    }

    #[test]
    fn test_start_editing_cell_with_value() {
        let mut app = App::default();
        
        // Set a cell with value
        let cell_data = CellData {
            value: "Hello".to_string(),
            formula: None,
            format: None,
            comment: None,
        spill_anchor: None,
        };
        app.workbook.current_sheet_mut().set_cell(0, 0, cell_data);
        
        app.start_editing();
        
        assert!(matches!(app.mode, AppMode::Editing));
        assert_eq!(app.input, "Hello");
        assert_eq!(app.cursor_position, 5); // End of "Hello"
    }

    #[test]
    fn test_start_editing_cell_with_formula() {
        let mut app = App::default();
        
        // Set a cell with formula
        let cell_data = CellData {
            value: "42".to_string(),
            formula: Some("=6*7".to_string()),
            format: None,
            comment: None,
        spill_anchor: None,
        };
        app.workbook.current_sheet_mut().set_cell(0, 0, cell_data);
        
        app.start_editing();
        
        assert!(matches!(app.mode, AppMode::Editing));
        assert_eq!(app.input, "=6*7"); // Should load formula, not value
        assert_eq!(app.cursor_position, 4); // End of "=6*7"
    }

    #[test]
    fn test_finish_editing_simple_value() {
        let mut app = App::default();
        app.start_editing();
        app.input = "Test Value".to_string();
        
        app.finish_editing();
        
        assert!(matches!(app.mode, AppMode::Normal));
        assert!(app.input.is_empty());
        assert_eq!(app.cursor_position, 0);
        
        let cell = app.workbook.current_sheet().get_cell(0, 0);
        assert_eq!(cell.value, "Test Value");
        assert!(cell.formula.is_none());
    }

    #[test]
    fn test_finish_editing_formula() {
        let mut app = App::default();
        app.start_editing();
        app.input = "=2+3".to_string();
        
        app.finish_editing();
        
        assert!(matches!(app.mode, AppMode::Normal));
        assert!(app.input.is_empty());
        assert_eq!(app.cursor_position, 0);
        
        let cell = app.workbook.current_sheet().get_cell(0, 0);
        assert_eq!(cell.value, "5"); // Evaluated result
        assert_eq!(cell.formula.unwrap(), "=2+3"); // Original formula
    }

    #[test]
    fn test_finish_editing_circular_reference() {
        let mut app = App::default();
        app.start_editing();
        app.input = "=A1+1".to_string(); // Self-reference
        
        let original_cell = app.workbook.current_sheet().get_cell(0, 0).clone();
        app.finish_editing();
        
        // Should remain in editing mode and not change the cell
        let cell_after = app.workbook.current_sheet().get_cell(0, 0);
        assert_eq!(original_cell.value, cell_after.value);
        assert_eq!(original_cell.formula, cell_after.formula);
    }

    #[test]
    fn test_cancel_editing() {
        let mut app = App::default();
        app.start_editing();
        app.input = "Some input".to_string();
        app.cursor_position = 5;
        
        app.cancel_editing();
        
        assert!(matches!(app.mode, AppMode::Normal));
        assert!(app.input.is_empty());
        assert_eq!(app.cursor_position, 0);
        
        // Cell should remain unchanged
        let cell = app.workbook.current_sheet().get_cell(0, 0);
        assert!(cell.value.is_empty());
        assert!(cell.formula.is_none());
    }

}
