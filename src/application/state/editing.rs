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
        self.workbook.current_sheet_mut().insert_row(at);
        self.selected_row = at;
        self.dirty = true;
        self.ensure_cursor_visible();
        self.start_editing();
    }

    pub fn vim_open_row_above(&mut self) {
        let at = self.selected_row;
        self.workbook.current_sheet_mut().insert_row(at);
        self.dirty = true;
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
            EditExitDir::Right => {
                if self.selected_col < self.workbook.current_sheet().cols - 1 {
                    self.selected_col += 1;
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
            let evaluator = FormulaEvaluator::with_workbook(
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

    pub fn cancel_editing(&mut self) {
        self.mode = AppMode::Normal;
        self.input.clear();
        self.cursor_position = 0;
    }

}
