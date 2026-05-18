//! Submodule of `state` — see state/mod.rs.

use super::*;

impl App {
    pub fn switch_next_sheet(&mut self) {
        if self.workbook.active_sheet < self.workbook.sheets.len() - 1 {
            self.workbook.active_sheet += 1;
            self.selected_row = 0;
            self.selected_col = 0;
            self.scroll_row = 0;
            self.scroll_col = 0;
            self.clear_selection();
            self.status_message = Some(format!("Sheet: {}", self.workbook.sheet_names[self.workbook.active_sheet]));
        }
    }

    pub fn switch_prev_sheet(&mut self) {
        if self.workbook.active_sheet > 0 {
            self.workbook.active_sheet -= 1;
            self.selected_row = 0;
            self.selected_col = 0;
            self.scroll_row = 0;
            self.scroll_col = 0;
            self.clear_selection();
            self.status_message = Some(format!("Sheet: {}", self.workbook.sheet_names[self.workbook.active_sheet]));
        }
    }

    pub fn jump_to_home(&mut self) {
        self.selected_row = 0;
        self.selected_col = 0;
        self.scroll_row = 0;
        self.scroll_col = 0;
    }

    pub fn jump_to_end(&mut self) {
        let mut max_row = 0;
        let mut max_col = 0;
        for &(row, col) in self.workbook.current_sheet().cells.keys() {
            if !self.workbook.current_sheet().get_cell(row, col).value.is_empty() {
                max_row = max_row.max(row);
                max_col = max_col.max(col);
            }
        }
        self.selected_row = max_row;
        self.selected_col = max_col;
        self.ensure_cursor_visible();
    }

}
