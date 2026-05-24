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
            self.invalidate_cross_sheet_state();
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
            self.invalidate_cross_sheet_state();
            self.status_message = Some(format!("Sheet: {}", self.workbook.sheet_names[self.workbook.active_sheet]));
        }
    }

    /// Drop per-sheet state that would be misleading on the new sheet:
    /// search results are stored as (row, col) only, so following them with
    /// n/N after a sheet switch would jump to phantom cells.
    pub(crate) fn invalidate_cross_sheet_state(&mut self) {
        self.search_results.clear();
        self.search_results_set.clear();
        self.search_result_index = 0;
        self.find_replace_results.clear();
        self.find_replace_index = 0;
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

#[cfg(test)]
#[allow(clippy::field_reassign_with_default)]
mod tests {
    use super::*;
    use crate::domain::CellData;

    #[test]
    fn test_jump_to_home() {
        let mut app = App::default();
        app.selected_row = 10;
        app.selected_col = 5;
        app.scroll_row = 8;
        app.scroll_col = 3;

        app.jump_to_home();

        assert_eq!(app.selected_row, 0);
        assert_eq!(app.selected_col, 0);
        assert_eq!(app.scroll_row, 0);
        assert_eq!(app.scroll_col, 0);
    }

    #[test]
    fn test_jump_to_end() {
        let mut app = App::default();
        app.set_cell_with_undo(5, 3, CellData { value: "data".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        app.set_cell_with_undo(10, 7, CellData { value: "last".to_string(), formula: None, format: None, comment: None, spill_anchor: None });

        app.jump_to_end();

        assert_eq!(app.selected_row, 10);
        assert_eq!(app.selected_col, 7);
    }

    #[test]
    fn test_default_workbook_has_one_sheet() {
        let app = App::default();
        assert_eq!(app.workbook.sheets.len(), 1);
        assert_eq!(app.workbook.sheet_names[0], "Sheet1");
        assert_eq!(app.workbook.active_sheet, 0);
    }

    #[test]
    fn test_cannot_delete_last_sheet() {
        let mut app = App::default();
        app.start_command_palette();
        app.command_input = "sheet delete".to_string();
        app.execute_command();

        assert_eq!(app.workbook.sheets.len(), 1); // Still 1 sheet
        assert!(app.status_message.as_ref().unwrap().contains("Cannot delete"));
    }

    #[test]
    fn test_switch_sheets() {
        let mut app = App::default();
        app.workbook.add_sheet("Sheet2".to_string());

        // Set data in sheet 1
        app.set_cell_with_undo(0, 0, CellData { value: "Sheet1Data".to_string(), formula: None, format: None, comment: None, spill_anchor: None });

        // Switch to sheet 2
        app.switch_next_sheet();
        assert_eq!(app.workbook.active_sheet, 1);
        assert!(app.workbook.current_sheet().get_cell(0, 0).value.is_empty());

        // Set data in sheet 2
        app.set_cell_with_undo(0, 0, CellData { value: "Sheet2Data".to_string(), formula: None, format: None, comment: None, spill_anchor: None });

        // Switch back to sheet 1
        app.switch_prev_sheet();
        assert_eq!(app.workbook.active_sheet, 0);
        assert_eq!(app.workbook.current_sheet().get_cell(0, 0).value, "Sheet1Data");

        // Verify sheet 2 still has its data
        app.switch_next_sheet();
        assert_eq!(app.workbook.current_sheet().get_cell(0, 0).value, "Sheet2Data");
    }

    #[test]
    fn test_switch_prev_at_first_sheet() {
        let mut app = App::default();
        app.switch_prev_sheet();
        assert_eq!(app.workbook.active_sheet, 0); // Stays at 0
    }

    #[test]
    fn test_switch_next_at_last_sheet() {
        let mut app = App::default();
        app.switch_next_sheet();
        assert_eq!(app.workbook.active_sheet, 0); // Stays at 0 (only 1 sheet)
    }

}
