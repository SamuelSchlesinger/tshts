//! Cell editing lifecycle: start, finish, cancel.

use crate::domain::{CellData, FormulaEvaluator};
use super::{App, AppMode};

impl App {
    /// Switches to editing mode for the currently selected cell.
    pub fn start_editing(&mut self) {
        self.mode = AppMode::Editing;
        let cell = self.workbook.current_sheet().get_cell(self.selected_row, self.selected_col);
        self.input = cell.formula.unwrap_or(cell.value);
        self.cursor_position = self.input.len();
    }

    /// Completes editing and moves down one cell (Enter key behavior).
    pub fn finish_editing(&mut self) {
        let existing = self.workbook.current_sheet().get_cell(self.selected_row, self.selected_col);
        let mut cell_data = CellData {
            format: existing.format.clone(),
            comment: existing.comment.clone(),
            ..CellData::default()
        };

        if self.input.starts_with('=') {
            let evaluator = FormulaEvaluator::new(self.workbook.current_sheet());
            if evaluator.would_create_circular_reference(&self.input, (self.selected_row, self.selected_col)) {
                return;
            }
            cell_data.formula = Some(self.input.clone());
            cell_data.value = evaluator.evaluate_formula(&self.input);
        } else {
            cell_data.value = self.input.clone();
        }

        self.set_cell_with_undo(self.selected_row, self.selected_col, cell_data);

        if self.selected_row < self.workbook.current_sheet().rows - 1 {
            self.selected_row += 1;
        }

        self.mode = AppMode::Normal;
        self.input.clear();
        self.cursor_position = 0;
    }

    /// Completes editing and moves right (Tab key behavior).
    pub fn finish_editing_move_right(&mut self) {
        let existing = self.workbook.current_sheet().get_cell(self.selected_row, self.selected_col);
        let mut cell_data = CellData {
            format: existing.format.clone(),
            comment: existing.comment.clone(),
            ..CellData::default()
        };

        if self.input.starts_with('=') {
            let evaluator = FormulaEvaluator::new(self.workbook.current_sheet());
            if evaluator.would_create_circular_reference(&self.input, (self.selected_row, self.selected_col)) {
                return;
            }
            cell_data.formula = Some(self.input.clone());
            cell_data.value = evaluator.evaluate_formula(&self.input);
        } else {
            cell_data.value = self.input.clone();
        }

        self.set_cell_with_undo(self.selected_row, self.selected_col, cell_data);

        if self.selected_col < self.workbook.current_sheet().cols - 1 {
            self.selected_col += 1;
        }

        self.mode = AppMode::Normal;
        self.input.clear();
        self.cursor_position = 0;
    }

    /// Cancels editing and returns to normal mode without saving changes.
    pub fn cancel_editing(&mut self) {
        self.mode = AppMode::Normal;
        self.input.clear();
        self.cursor_position = 0;
    }
}
