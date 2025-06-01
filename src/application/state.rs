//! Application state management for the terminal spreadsheet.
//!
//! This module contains the main application state and mode management
//! for the terminal user interface.

use crate::domain::{Spreadsheet, CellData, FormulaEvaluator};
use std::collections::VecDeque;

/// Represents the current mode of the application.
///
/// The application can be in different modes that determine how user input
/// is interpreted and what UI elements are displayed.
#[derive(Debug)]
pub enum AppMode {
    /// Normal navigation mode - arrow keys move selection, shortcuts available
    Normal,
    /// Cell editing mode - user is typing into a cell
    Editing,
    /// Help screen is displayed
    Help,
    /// Save dialog is open
    SaveAs,
    /// Load dialog is open
    LoadFile,
    /// CSV export dialog is open
    ExportCsv,
    /// CSV import dialog is open
    ImportCsv,
    /// Search mode - user is typing a search query
    Search,
}

/// Represents an action that can be undone/redone.
#[derive(Debug, Clone)]
pub enum UndoAction {
    /// Cell was modified (row, col, old_value, new_value)
    CellModified {
        row: usize,
        col: usize,
        old_cell: Option<CellData>,
        new_cell: Option<CellData>,
    },
}

/// Main application state containing the spreadsheet and UI state.
///
/// This structure holds all the data needed to render the terminal UI
/// and manage user interactions with the spreadsheet.
///
/// # Examples
///
/// ```
/// use tshts::application::App;
///
/// let app = App::default();
/// assert_eq!(app.selected_row, 0);
/// assert_eq!(app.selected_col, 0);
/// ```
#[derive(Debug)]
pub struct App {
    /// The spreadsheet data structure
    pub spreadsheet: Spreadsheet,
    /// Currently selected row (zero-based)
    pub selected_row: usize,
    /// Currently selected column (zero-based)
    pub selected_col: usize,
    /// Top-left row visible in the viewport
    pub scroll_row: usize,
    /// Left-most column visible in the viewport
    pub scroll_col: usize,
    /// Current application mode
    pub mode: AppMode,
    /// Current input buffer (for editing mode)
    pub input: String,
    /// Cursor position within the input buffer
    pub cursor_position: usize,
    /// Current filename (if file has been saved/loaded)
    pub filename: Option<String>,
    /// Scroll position in help text
    pub help_scroll: usize,
    /// Temporary status message to display
    pub status_message: Option<String>,
    /// Input buffer for filename entry
    pub filename_input: String,
    /// Undo stack for tracking changes
    pub undo_stack: VecDeque<UndoAction>,
    /// Redo stack for tracking undone changes
    pub redo_stack: VecDeque<UndoAction>,
    /// Search query input buffer
    pub search_query: String,
    /// Search results as (row, col) coordinates
    pub search_results: Vec<(usize, usize)>,
    /// Current search result index
    pub search_result_index: usize,
    /// Selection start position (row, col)
    pub selection_start: Option<(usize, usize)>,
    /// Selection end position (row, col) 
    pub selection_end: Option<(usize, usize)>,
    /// Whether we're in drag selection mode
    pub selecting: bool,
    /// Viewport height in rows (for scrolling calculations)
    pub viewport_rows: usize,
    /// Viewport width in columns (for scrolling calculations) 
    pub viewport_cols: usize,
}

impl Default for App {
    fn default() -> Self {
        Self {
            spreadsheet: Spreadsheet::default(),
            selected_row: 0,
            selected_col: 0,
            scroll_row: 0,
            scroll_col: 0,
            mode: AppMode::Normal,
            input: String::new(),
            cursor_position: 0,
            filename: None,
            help_scroll: 0,
            status_message: None,
            filename_input: String::new(),
            undo_stack: VecDeque::new(),
            redo_stack: VecDeque::new(),
            search_query: String::new(),
            search_results: Vec::new(),
            search_result_index: 0,
            selection_start: None,
            selection_end: None,
            selecting: false,
            viewport_rows: 20,  // Default reasonable size
            viewport_cols: 8,   // Default reasonable size
        }
    }
}

impl App {
    /// Switches to editing mode for the currently selected cell.
    ///
    /// Loads the cell's formula (if present) or value into the input buffer
    /// and positions the cursor at the end.
    pub fn start_editing(&mut self) {
        self.mode = AppMode::Editing;
        let cell = self.spreadsheet.get_cell(self.selected_row, self.selected_col);
        self.input = cell.formula.unwrap_or(cell.value);
        self.cursor_position = self.input.len();
    }

    /// Completes editing and updates the cell with the input content.
    ///
    /// If the input starts with '=', it's treated as a formula and evaluated.
    /// Checks for circular references before applying the formula.
    /// Returns to normal mode after completion.
    pub fn finish_editing(&mut self) {
        let mut cell_data = CellData::default();
        
        if self.input.starts_with('=') {
            let evaluator = FormulaEvaluator::new(&self.spreadsheet);
            if evaluator.would_create_circular_reference(&self.input, (self.selected_row, self.selected_col)) {
                return;
            }
            cell_data.formula = Some(self.input.clone());
            cell_data.value = evaluator.evaluate_formula(&self.input);
        } else {
            cell_data.value = self.input.clone();
        }

        self.set_cell_with_undo(self.selected_row, self.selected_col, cell_data);
        
        // Move down one cell after editing
        if self.selected_row < self.spreadsheet.rows - 1 {
            self.selected_row += 1;
        }
        
        self.mode = AppMode::Normal;
        self.input.clear();
        self.cursor_position = 0;
    }

    /// Cancels editing and returns to normal mode without saving changes.
    ///
    /// Clears the input buffer and resets cursor position.
    pub fn cancel_editing(&mut self) {
        self.mode = AppMode::Normal;
        self.input.clear();
        self.cursor_position = 0;
    }

    /// Switches to save-as mode to prompt for a filename.
    ///
    /// Initializes the filename input with the current filename or default.
    pub fn start_save_as(&mut self) {
        self.mode = AppMode::SaveAs;
        self.filename_input = self.filename.clone().unwrap_or_else(|| "spreadsheet.tshts".to_string());
        self.cursor_position = self.filename_input.len();
        self.status_message = None;
    }

    /// Switches to load-file mode to prompt for a filename.
    ///
    /// Initializes the filename input with the current filename or default.
    pub fn start_load_file(&mut self) {
        self.mode = AppMode::LoadFile;
        self.filename_input = self.filename.clone().unwrap_or_else(|| "spreadsheet.tshts".to_string());
        self.cursor_position = self.filename_input.len();
        self.status_message = None;
    }

    /// Cancels filename input and returns to normal mode.
    ///
    /// Clears the filename input buffer and resets cursor position.
    pub fn cancel_filename_input(&mut self) {
        self.mode = AppMode::Normal;
        self.filename_input.clear();
        self.cursor_position = 0;
    }

    /// Processes the result of a save operation.
    ///
    /// Updates the current filename and status message based on whether
    /// the save was successful. Returns to normal mode.
    ///
    /// # Arguments
    ///
    /// * `result` - Result of the save operation (filename or error message)
    pub fn set_save_result(&mut self, result: Result<String, String>) {
        match result {
            Ok(filename) => {
                self.filename = Some(filename.clone());
                self.status_message = Some(format!("Saved to {}", filename));
            }
            Err(error) => {
                self.status_message = Some(format!("Save failed: {}", error));
            }
        }
        
        self.mode = AppMode::Normal;
        self.filename_input.clear();
        self.cursor_position = 0;
    }

    /// Processes the result of a load operation.
    ///
    /// Updates the spreadsheet data and resets the view if successful.
    /// Sets appropriate status message and returns to normal mode.
    ///
    /// # Arguments
    ///
    /// * `result` - Result of the load operation (spreadsheet and filename, or error)
    pub fn set_load_result(&mut self, result: Result<(Spreadsheet, String), String>) {
        match result {
            Ok((spreadsheet, filename)) => {
                self.spreadsheet = spreadsheet;
                self.filename = Some(filename.clone());
                self.selected_row = 0;
                self.selected_col = 0;
                self.scroll_row = 0;
                self.scroll_col = 0;
                self.status_message = Some(format!("Loaded from {}", filename));
            }
            Err(error) => {
                self.status_message = Some(format!("Load failed: {}", error));
            }
        }
        
        self.mode = AppMode::Normal;
        self.filename_input.clear();
        self.cursor_position = 0;
    }

    /// Gets the filename to use for saving.
    ///
    /// Returns the filename input if not empty, otherwise returns a default filename.
    ///
    /// # Returns
    ///
    /// The filename to use for saving
    pub fn get_save_filename(&self) -> String {
        if self.filename_input.is_empty() {
            "spreadsheet.tshts".to_string()
        } else {
            self.filename_input.clone()
        }
    }

    /// Gets the filename to use for loading.
    ///
    /// Returns the filename input if not empty, otherwise returns a default filename.
    ///
    /// # Returns
    ///
    /// The filename to use for loading
    pub fn get_load_filename(&self) -> String {
        if self.filename_input.is_empty() {
            "spreadsheet.tshts".to_string()
        } else {
            self.filename_input.clone()
        }
    }

    /// Switches to CSV export mode to prompt for a filename.
    ///
    /// Initializes the filename input with a default CSV filename.
    pub fn start_csv_export(&mut self) {
        self.mode = AppMode::ExportCsv;
        self.filename_input = self.filename
            .as_ref()
            .map(|f| f.replace(".tshts", ".csv"))
            .unwrap_or_else(|| "spreadsheet.csv".to_string());
        self.cursor_position = self.filename_input.len();
        self.status_message = None;
    }

    /// Gets the filename to use for CSV export.
    ///
    /// Returns the filename input if not empty, otherwise returns a default CSV filename.
    ///
    /// # Returns
    ///
    /// The filename to use for CSV export
    pub fn get_csv_export_filename(&self) -> String {
        if self.filename_input.is_empty() {
            "spreadsheet.csv".to_string()
        } else {
            self.filename_input.clone()
        }
    }

    /// Processes the result of a CSV export operation.
    ///
    /// Sets appropriate status message based on whether the export was successful.
    /// Returns to normal mode.
    ///
    /// # Arguments
    ///
    /// * `result` - Result of the CSV export operation (filename or error message)
    pub fn set_csv_export_result(&mut self, result: Result<String, String>) {
        match result {
            Ok(filename) => {
                self.status_message = Some(format!("Exported to {}", filename));
            }
            Err(error) => {
                self.status_message = Some(format!("Export failed: {}", error));
            }
        }
        
        self.mode = AppMode::Normal;
        self.filename_input.clear();
        self.cursor_position = 0;
    }

    /// Switches to CSV import mode to prompt for a filename.
    ///
    /// Initializes the filename input with a default CSV filename.
    pub fn start_csv_import(&mut self) {
        self.mode = AppMode::ImportCsv;
        self.filename_input = "data.csv".to_string();
        self.cursor_position = self.filename_input.len();
        self.status_message = None;
    }

    /// Gets the filename to use for CSV import.
    ///
    /// Returns the filename input if not empty, otherwise returns a default CSV filename.
    ///
    /// # Returns
    ///
    /// The filename to use for CSV import
    pub fn get_csv_import_filename(&self) -> String {
        if self.filename_input.is_empty() {
            "data.csv".to_string()
        } else {
            self.filename_input.clone()
        }
    }

    /// Processes the result of a CSV import operation.
    ///
    /// Updates the spreadsheet data and resets the view if successful.
    /// Sets appropriate status message and returns to normal mode.
    ///
    /// # Arguments
    ///
    /// * `result` - Result of the CSV import operation (spreadsheet or error message)
    pub fn set_csv_import_result(&mut self, result: Result<Spreadsheet, String>) {
        match result {
            Ok(spreadsheet) => {
                self.spreadsheet = spreadsheet;
                self.selected_row = 0;
                self.selected_col = 0;
                self.scroll_row = 0;
                self.scroll_col = 0;
                self.status_message = Some("CSV data imported successfully".to_string());
                // Don't set filename since this is imported CSV data, not a saved spreadsheet
            }
            Err(error) => {
                self.status_message = Some(format!("Import failed: {}", error));
            }
        }
        
        self.mode = AppMode::Normal;
        self.filename_input.clear();
        self.cursor_position = 0;
    }

    /// Records an action for undo/redo functionality.
    ///
    /// Adds the action to the undo stack and clears the redo stack.
    /// Limits the undo stack to 100 actions.
    fn record_action(&mut self, action: UndoAction) {
        const MAX_UNDO_STACK_SIZE: usize = 100;
        
        // Add to undo stack
        self.undo_stack.push_back(action);
        
        // Limit stack size
        if self.undo_stack.len() > MAX_UNDO_STACK_SIZE {
            self.undo_stack.pop_front();
        }
        
        // Clear redo stack since we made a new change
        self.redo_stack.clear();
    }

    /// Performs an undo operation.
    ///
    /// Reverts the last action and moves it to the redo stack.
    pub fn undo(&mut self) {
        if let Some(action) = self.undo_stack.pop_back() {
            match action.clone() {
                UndoAction::CellModified { row, col, old_cell, new_cell: _ } => {
                    // Apply the old cell value
                    if let Some(old_data) = old_cell {
                        self.spreadsheet.set_cell(row, col, old_data);
                    } else {
                        self.spreadsheet.clear_cell(row, col);
                    }
                }
            }
            
            // Add to redo stack
            self.redo_stack.push_back(action);
        }
    }

    /// Performs a redo operation.
    ///
    /// Reapplies the last undone action and moves it back to the undo stack.
    pub fn redo(&mut self) {
        if let Some(action) = self.redo_stack.pop_back() {
            match action.clone() {
                UndoAction::CellModified { row, col, old_cell: _, new_cell } => {
                    // Apply the new cell value
                    if let Some(new_data) = new_cell {
                        self.spreadsheet.set_cell(row, col, new_data);
                    } else {
                        self.spreadsheet.clear_cell(row, col);
                    }
                }
            }
            
            // Add back to undo stack
            self.undo_stack.push_back(action);
        }
    }

    /// Sets a cell value and records the action for undo/redo.
    ///
    /// This is a wrapper around the spreadsheet's set_cell method that also
    /// tracks the change for undo functionality.
    pub fn set_cell_with_undo(&mut self, row: usize, col: usize, new_data: CellData) {
        // Get the old cell data
        let old_cell = if self.spreadsheet.cells.contains_key(&(row, col)) {
            Some(self.spreadsheet.get_cell(row, col))
        } else {
            None
        };
        
        // Record the action
        let action = UndoAction::CellModified {
            row,
            col,
            old_cell,
            new_cell: Some(new_data.clone()),
        };
        self.record_action(action);
        
        // Apply the change
        self.spreadsheet.set_cell(row, col, new_data);
    }

    /// Clears a cell and records the action for undo/redo.
    ///
    /// This is a wrapper around the spreadsheet's clear_cell method that also
    /// tracks the change for undo functionality.
    pub fn clear_cell_with_undo(&mut self, row: usize, col: usize) {
        // Get the old cell data
        let old_cell = if self.spreadsheet.cells.contains_key(&(row, col)) {
            Some(self.spreadsheet.get_cell(row, col))
        } else {
            None
        };
        
        // Only record if there was actually a cell to clear
        if old_cell.is_some() {
            let action = UndoAction::CellModified {
                row,
                col,
                old_cell,
                new_cell: None,
            };
            self.record_action(action);
        }
        
        // Apply the change
        self.spreadsheet.clear_cell(row, col);
    }

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

        // Search through all cells
        for row in 0..self.spreadsheet.rows {
            for col in 0..self.spreadsheet.cols {
                let cell = self.spreadsheet.get_cell(row, col);
                
                // Search in both value and formula (if present)
                let value_matches = cell.value.to_lowercase().contains(&query_lower);
                let formula_matches = cell.formula
                    .as_ref()
                    .map(|f| f.to_lowercase().contains(&query_lower))
                    .unwrap_or(false);

                if value_matches || formula_matches {
                    self.search_results.push((row, col));
                }
            }
        }

        // Move to first result if any found
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

    /// Moves the cursor to the current search result.
    fn go_to_current_search_result(&mut self) {
        if let Some(&(row, col)) = self.search_results.get(self.search_result_index) {
            self.selected_row = row;
            self.selected_col = col;
            self.ensure_cursor_visible();
        }
    }

    /// Finishes search and returns to normal mode while keeping the current selection.
    pub fn finish_search(&mut self) {
        self.mode = AppMode::Normal;
        
        let num_results = self.search_results.len();
        if num_results > 0 {
            self.status_message = Some(format!(
                "Search completed: {} result{} found for '{}'", 
                num_results,
                if num_results == 1 { "" } else { "s" },
                self.search_query
            ));
        } else {
            self.status_message = Some(format!("No results found for '{}'", self.search_query));
        }
        
        self.search_query.clear();
        self.search_results.clear();
        self.search_result_index = 0;
        self.cursor_position = 0;
    }

    /// Starts selection at the current position
    pub fn start_selection(&mut self) {
        self.selection_start = Some((self.selected_row, self.selected_col));
        self.selection_end = Some((self.selected_row, self.selected_col));
        self.selecting = true;
    }

    /// Updates the selection end position
    pub fn update_selection(&mut self, row: usize, col: usize) {
        if self.selecting {
            self.selection_end = Some((row, col));
        }
    }

    /// Ends selection mode
    pub fn end_selection(&mut self) {
        self.selecting = false;
    }

    /// Clears the current selection
    pub fn clear_selection(&mut self) {
        self.selection_start = None;
        self.selection_end = None;
        self.selecting = false;
    }

    /// Gets the normalized selection range (top-left to bottom-right)
    pub fn get_selection_range(&self) -> Option<((usize, usize), (usize, usize))> {
        if let (Some(start), Some(end)) = (self.selection_start, self.selection_end) {
            let min_row = start.0.min(end.0);
            let max_row = start.0.max(end.0);
            let min_col = start.1.min(end.1);
            let max_col = start.1.max(end.1);
            Some(((min_row, min_col), (max_row, max_col)))
        } else {
            None
        }
    }

    /// Checks if a cell is within the current selection
    pub fn is_cell_selected(&self, row: usize, col: usize) -> bool {
        if let Some(((min_row, min_col), (max_row, max_col))) = self.get_selection_range() {
            row >= min_row && row <= max_row && col >= min_col && col <= max_col
        } else {
            false
        }
    }

    /// Updates the viewport size for proper scrolling calculations.
    pub fn update_viewport_size(&mut self, rows: usize, cols: usize) {
        self.viewport_rows = rows;
        self.viewport_cols = cols;
    }

    /// Ensures the selected cell is visible by adjusting scroll position.
    pub fn ensure_cursor_visible(&mut self) {
        // Vertical scrolling
        if self.selected_row < self.scroll_row {
            self.scroll_row = self.selected_row;
        } else if self.selected_row >= self.scroll_row + self.viewport_rows {
            self.scroll_row = self.selected_row.saturating_sub(self.viewport_rows - 1);
        }
        
        // Horizontal scrolling
        if self.selected_col < self.scroll_col {
            self.scroll_col = self.selected_col;
        } else if self.selected_col >= self.scroll_col + self.viewport_cols {
            self.scroll_col = self.selected_col.saturating_sub(self.viewport_cols - 1);
        }
    }

    /// Performs autofill operation on the current selection.
    ///
    /// Copies the formula from the top-left cell of the selection to all other
    /// cells in the selection, adjusting cell references relatively.
    pub fn autofill_selection(&mut self) {
        if let Some(((start_row, start_col), (end_row, end_col))) = self.get_selection_range() {
            // Get the source cell (top-left of selection)
            let source_cell = self.spreadsheet.get_cell(start_row, start_col);
            
            // Only proceed if the source cell has content
            if source_cell.value.is_empty() && source_cell.formula.is_none() {
                return;
            }

            // Collect all the changes first to avoid borrowing conflicts
            let mut changes = Vec::new();
            
            // Fill each cell in the selection
            for row in start_row..=end_row {
                for col in start_col..=end_col {
                    // Skip the source cell
                    if row == start_row && col == start_col {
                        continue;
                    }
                    
                    let row_offset = row as i32 - start_row as i32;
                    let col_offset = col as i32 - start_col as i32;
                    
                    let new_cell_data = if let Some(ref formula) = source_cell.formula {
                        use crate::domain::services::FormulaEvaluator;
                        let evaluator = FormulaEvaluator::new(&self.spreadsheet);
                        
                        // Adjust the formula with relative references
                        let adjusted_formula = evaluator.adjust_formula_references(formula, row_offset, col_offset);
                        
                        // Check for circular references
                        if evaluator.would_create_circular_reference(&adjusted_formula, (row, col)) {
                            continue; // Skip this cell to avoid circular reference
                        }
                        
                        let new_value = evaluator.evaluate_formula(&adjusted_formula);
                        CellData {
                            value: new_value,
                            formula: Some(adjusted_formula),
                        }
                    } else {
                        // Simple value copy (no formula)
                        CellData {
                            value: source_cell.value.clone(),
                            formula: None,
                        }
                    };
                    
                    changes.push((row, col, new_cell_data));
                }
            }
            
            // Apply all changes
            for (row, col, cell_data) in changes {
                self.set_cell_with_undo(row, col, cell_data);
            }
            
            self.status_message = Some(format!(
                "Autofilled {} cells from {}{}",
                (end_row - start_row + 1) * (end_col - start_col + 1) - 1,
                Spreadsheet::column_label(start_col),
                start_row + 1
            ));
        }
    }

}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::domain::CellData;

    #[test]
    fn test_app_default() {
        let app = App::default();
        assert_eq!(app.selected_row, 0);
        assert_eq!(app.selected_col, 0);
        assert_eq!(app.scroll_row, 0);
        assert_eq!(app.scroll_col, 0);
        assert!(matches!(app.mode, AppMode::Normal));
        assert!(app.input.is_empty());
        assert_eq!(app.cursor_position, 0);
        assert!(app.filename.is_none());
        assert_eq!(app.help_scroll, 0);
        assert!(app.status_message.is_none());
        assert!(app.filename_input.is_empty());
    }

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
        };
        app.spreadsheet.set_cell(0, 0, cell_data);
        
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
        };
        app.spreadsheet.set_cell(0, 0, cell_data);
        
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
        
        let cell = app.spreadsheet.get_cell(0, 0);
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
        
        let cell = app.spreadsheet.get_cell(0, 0);
        assert_eq!(cell.value, "5"); // Evaluated result
        assert_eq!(cell.formula.unwrap(), "=2+3"); // Original formula
    }

    #[test]
    fn test_finish_editing_circular_reference() {
        let mut app = App::default();
        app.start_editing();
        app.input = "=A1+1".to_string(); // Self-reference
        
        let original_cell = app.spreadsheet.get_cell(0, 0).clone();
        app.finish_editing();
        
        // Should remain in editing mode and not change the cell
        let cell_after = app.spreadsheet.get_cell(0, 0);
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
        let cell = app.spreadsheet.get_cell(0, 0);
        assert!(cell.value.is_empty());
        assert!(cell.formula.is_none());
    }

    #[test]
    fn test_start_save_as() {
        let mut app = App::default();
        app.start_save_as();
        
        assert!(matches!(app.mode, AppMode::SaveAs));
        assert_eq!(app.filename_input, "spreadsheet.tshts"); // Default filename
        assert_eq!(app.cursor_position, "spreadsheet.tshts".len());
        assert!(app.status_message.is_none());
    }

    #[test]
    fn test_start_save_as_with_existing_filename() {
        let mut app = App::default();
        app.filename = Some("existing.tshts".to_string());
        
        app.start_save_as();
        
        assert!(matches!(app.mode, AppMode::SaveAs));
        assert_eq!(app.filename_input, "existing.tshts");
        assert_eq!(app.cursor_position, "existing.tshts".len());
    }

    #[test]
    fn test_start_load_file() {
        let mut app = App::default();
        app.start_load_file();
        
        assert!(matches!(app.mode, AppMode::LoadFile));
        assert_eq!(app.filename_input, "spreadsheet.tshts");
        assert_eq!(app.cursor_position, "spreadsheet.tshts".len());
        assert!(app.status_message.is_none());
    }

    #[test]
    fn test_cancel_filename_input() {
        let mut app = App::default();
        app.start_save_as();
        app.filename_input = "test.tshts".to_string();
        app.cursor_position = 5;
        
        app.cancel_filename_input();
        
        assert!(matches!(app.mode, AppMode::Normal));
        assert!(app.filename_input.is_empty());
        assert_eq!(app.cursor_position, 0);
    }

    #[test]
    fn test_set_save_result_success() {
        let mut app = App::default();
        app.start_save_as();
        app.filename_input = "test.tshts".to_string();
        
        app.set_save_result(Ok("test.tshts".to_string()));
        
        assert!(matches!(app.mode, AppMode::Normal));
        assert_eq!(app.filename.unwrap(), "test.tshts");
        assert!(app.status_message.unwrap().contains("Saved to test.tshts"));
        assert!(app.filename_input.is_empty());
        assert_eq!(app.cursor_position, 0);
    }

    #[test]
    fn test_set_save_result_failure() {
        let mut app = App::default();
        app.start_save_as();
        
        app.set_save_result(Err("Permission denied".to_string()));
        
        assert!(matches!(app.mode, AppMode::Normal));
        assert!(app.filename.is_none()); // Filename unchanged on failure
        assert!(app.status_message.unwrap().contains("Save failed: Permission denied"));
        assert!(app.filename_input.is_empty());
        assert_eq!(app.cursor_position, 0);
    }

    #[test]
    fn test_set_load_result_success() {
        let mut app = App::default();
        app.selected_row = 5;
        app.selected_col = 3;
        app.scroll_row = 2;
        app.scroll_col = 1;
        
        let mut new_sheet = Spreadsheet::default();
        new_sheet.set_cell(0, 0, CellData {
            value: "Loaded".to_string(),
            formula: None,
        });
        
        app.set_load_result(Ok((new_sheet, "loaded.tshts".to_string())));
        
        assert!(matches!(app.mode, AppMode::Normal));
        assert_eq!(app.filename.unwrap(), "loaded.tshts");
        assert!(app.status_message.unwrap().contains("Loaded from loaded.tshts"));
        
        // Position should be reset
        assert_eq!(app.selected_row, 0);
        assert_eq!(app.selected_col, 0);
        assert_eq!(app.scroll_row, 0);
        assert_eq!(app.scroll_col, 0);
        
        // Spreadsheet should be updated
        let cell = app.spreadsheet.get_cell(0, 0);
        assert_eq!(cell.value, "Loaded");
    }

    #[test]
    fn test_set_load_result_failure() {
        let mut app = App::default();
        let original_sheet = app.spreadsheet.clone();
        
        app.set_load_result(Err("File not found".to_string()));
        
        assert!(matches!(app.mode, AppMode::Normal));
        assert!(app.filename.is_none());
        assert!(app.status_message.unwrap().contains("Load failed: File not found"));
        
        // Spreadsheet should remain unchanged
        assert_eq!(app.spreadsheet.rows, original_sheet.rows);
        assert_eq!(app.spreadsheet.cols, original_sheet.cols);
    }

    #[test]
    fn test_get_save_filename() {
        let mut app = App::default();
        
        // Empty filename input should return default
        assert_eq!(app.get_save_filename(), "spreadsheet.tshts");
        
        // Non-empty filename input should return that
        app.filename_input = "custom.tshts".to_string();
        assert_eq!(app.get_save_filename(), "custom.tshts");
    }

    #[test]
    fn test_get_load_filename() {
        let mut app = App::default();
        
        // Empty filename input should return default
        assert_eq!(app.get_load_filename(), "spreadsheet.tshts");
        
        // Non-empty filename input should return that
        app.filename_input = "custom.tshts".to_string();
        assert_eq!(app.get_load_filename(), "custom.tshts");
    }

    #[test]
    fn test_app_mode_transitions() {
        let mut app = App::default();
        
        // Normal -> Editing -> Normal
        assert!(matches!(app.mode, AppMode::Normal));
        app.start_editing();
        assert!(matches!(app.mode, AppMode::Editing));
        app.finish_editing();
        assert!(matches!(app.mode, AppMode::Normal));
        
        // Normal -> SaveAs -> Normal
        app.start_save_as();
        assert!(matches!(app.mode, AppMode::SaveAs));
        app.cancel_filename_input();
        assert!(matches!(app.mode, AppMode::Normal));
        
        // Normal -> LoadFile -> Normal
        app.start_load_file();
        assert!(matches!(app.mode, AppMode::LoadFile));
        app.cancel_filename_input();
        assert!(matches!(app.mode, AppMode::Normal));
    }

    #[test]
    fn test_status_message_handling() {
        let mut app = App::default();
        
        // Initially no status message
        assert!(app.status_message.is_none());
        
        // Save success sets status message
        app.set_save_result(Ok("test.tshts".to_string()));
        assert!(app.status_message.is_some());
        
        // Starting save dialog clears status message
        app.start_save_as();
        assert!(app.status_message.is_none());
        
        // Load failure sets status message
        app.set_load_result(Err("Error".to_string()));
        assert!(app.status_message.is_some());
        
        // Starting load dialog clears status message
        app.start_load_file();
        assert!(app.status_message.is_none());
    }

    #[test]
    fn test_csv_import_mode() {
        let mut app = App::default();
        
        // Initially in normal mode
        assert!(matches!(app.mode, AppMode::Normal));
        assert!(app.filename_input.is_empty());
        
        // Start CSV import mode
        app.start_csv_import();
        
        // Should be in ImportCsv mode with default filename
        assert!(matches!(app.mode, AppMode::ImportCsv));
        assert_eq!(app.filename_input, "data.csv");
        assert_eq!(app.cursor_position, "data.csv".len());
        assert!(app.status_message.is_none());
        
        // Test getting import filename
        assert_eq!(app.get_csv_import_filename(), "data.csv");
        
        // Test with custom filename
        app.filename_input = "custom.csv".to_string();
        assert_eq!(app.get_csv_import_filename(), "custom.csv");
        
        // Test with empty filename
        app.filename_input.clear();
        assert_eq!(app.get_csv_import_filename(), "data.csv");
        
        // Test cancel
        app.cancel_filename_input();
        assert!(matches!(app.mode, AppMode::Normal));
        assert!(app.filename_input.is_empty());
        assert_eq!(app.cursor_position, 0);
    }

    #[test]
    fn test_csv_import_result_handling() {
        let mut app = App::default();
        app.start_csv_import();
        
        // Set initial position away from origin
        app.selected_row = 5;
        app.selected_col = 3;
        app.scroll_row = 2;
        app.scroll_col = 1;
        
        // Test successful import
        let mut new_sheet = Spreadsheet::default();
        new_sheet.set_cell(0, 0, CellData {
            value: "Imported".to_string(),
            formula: None,
        });
        
        app.set_csv_import_result(Ok(new_sheet));
        
        // Should return to normal mode with success message
        assert!(matches!(app.mode, AppMode::Normal));
        assert!(app.status_message.as_ref().unwrap().contains("imported successfully"));
        assert!(app.filename_input.is_empty());
        assert_eq!(app.cursor_position, 0);
        
        // Position should be reset to origin
        assert_eq!(app.selected_row, 0);
        assert_eq!(app.selected_col, 0);
        assert_eq!(app.scroll_row, 0);
        assert_eq!(app.scroll_col, 0);
        
        // Spreadsheet should be updated
        let cell = app.spreadsheet.get_cell(0, 0);
        assert_eq!(cell.value, "Imported");
        
        // Test failed import
        app.start_csv_import();
        app.set_csv_import_result(Err("File not found".to_string()));
        
        assert!(matches!(app.mode, AppMode::Normal));
        assert!(app.status_message.as_ref().unwrap().contains("Import failed: File not found"));
    }

    #[test]
    fn test_selection_functionality() {
        let mut app = App::default();
        
        // Initially no selection
        assert!(app.get_selection_range().is_none());
        assert!(!app.is_cell_selected(0, 0));
        
        // Start selection
        app.start_selection();
        assert_eq!(app.get_selection_range(), Some(((0, 0), (0, 0))));
        assert!(app.is_cell_selected(0, 0));
        
        // Update selection
        app.update_selection(1, 2);
        assert_eq!(app.get_selection_range(), Some(((0, 0), (1, 2))));
        assert!(app.is_cell_selected(0, 1));
        assert!(app.is_cell_selected(1, 2));
        assert!(!app.is_cell_selected(2, 0));
        
        // Clear selection
        app.clear_selection();
        assert!(app.get_selection_range().is_none());
        assert!(!app.is_cell_selected(0, 0));
    }

    #[test]
    fn test_autofill_simple_values() {
        let mut app = App::default();
        
        // Set up a simple value in A1
        app.set_cell_with_undo(0, 0, CellData {
            value: "Hello".to_string(),
            formula: None,
        });
        
        // Select A1:B2
        app.selection_start = Some((0, 0));
        app.selection_end = Some((1, 1));
        
        // Autofill
        app.autofill_selection();
        
        // Check that the value was copied
        assert_eq!(app.spreadsheet.get_cell(0, 1).value, "Hello");
        assert_eq!(app.spreadsheet.get_cell(1, 0).value, "Hello");
        assert_eq!(app.spreadsheet.get_cell(1, 1).value, "Hello");
    }

    #[test]
    fn test_autofill_formula_horizontal() {
        let mut app = App::default();
        
        // Set up cells with values
        app.set_cell_with_undo(0, 1, CellData { value: "10".to_string(), formula: None }); // B1 = 10
        app.set_cell_with_undo(1, 1, CellData { value: "20".to_string(), formula: None }); // B2 = 20
        
        // Set up a formula in A1 that references B1:B2
        app.set_cell_with_undo(0, 0, CellData {
            value: "30".to_string(),
            formula: Some("=SUM(B1:B2)".to_string()),
        });
        
        // Select A1:C1 (horizontal autofill)
        app.selection_start = Some((0, 0));
        app.selection_end = Some((0, 2));
        
        // Autofill
        app.autofill_selection();
        
        // Check that formulas were adjusted horizontally
        let b1_cell = app.spreadsheet.get_cell(0, 1);
        assert_eq!(b1_cell.formula, Some("=SUM(C1:C2)".to_string()));
        
        let c1_cell = app.spreadsheet.get_cell(0, 2);
        assert_eq!(c1_cell.formula, Some("=SUM(D1:D2)".to_string()));
    }

    #[test]
    fn test_viewport_and_scrolling() {
        let mut app = App::default();
        
        // Test initial viewport size
        assert_eq!(app.viewport_rows, 20);
        assert_eq!(app.viewport_cols, 8);
        
        // Test updating viewport size
        app.update_viewport_size(15, 10);
        assert_eq!(app.viewport_rows, 15);
        assert_eq!(app.viewport_cols, 10);
        
        // Test ensure_cursor_visible - cursor within viewport
        app.selected_row = 5;
        app.selected_col = 3;
        app.scroll_row = 0;
        app.scroll_col = 0;
        app.ensure_cursor_visible();
        assert_eq!(app.scroll_row, 0);  // No need to scroll
        assert_eq!(app.scroll_col, 0);
        
        // Test ensure_cursor_visible - cursor beyond bottom/right
        app.selected_row = 20;  // Beyond viewport (15 rows)
        app.selected_col = 12;  // Beyond viewport (10 cols)
        app.ensure_cursor_visible();
        assert_eq!(app.scroll_row, 6);  // 20 - 15 + 1 = 6
        assert_eq!(app.scroll_col, 3);  // 12 - 10 + 1 = 3
        
        // Test ensure_cursor_visible - cursor before top/left
        app.selected_row = 2;
        app.selected_col = 1;
        app.ensure_cursor_visible();
        assert_eq!(app.scroll_row, 2);  // Scroll to show cursor
        assert_eq!(app.scroll_col, 1);
    }

    #[test]
    fn test_autofill_formula_vertical() {
        let mut app = App::default();
        
        // Set up cells with values
        app.set_cell_with_undo(1, 0, CellData { value: "10".to_string(), formula: None }); // A2 = 10
        app.set_cell_with_undo(1, 1, CellData { value: "20".to_string(), formula: None }); // B2 = 20
        
        // Set up a formula in A1 that references A2+B2
        app.set_cell_with_undo(0, 0, CellData {
            value: "30".to_string(),
            formula: Some("=A2+B2".to_string()),
        });
        
        // Select A1:A3 (vertical autofill)
        app.selection_start = Some((0, 0));
        app.selection_end = Some((2, 0));
        
        // Autofill
        app.autofill_selection();
        
        // Check that formulas were adjusted vertically
        let a2_cell = app.spreadsheet.get_cell(1, 0);
        assert_eq!(a2_cell.formula, Some("=A3+B3".to_string()));
        
        let a3_cell = app.spreadsheet.get_cell(2, 0);
        assert_eq!(a3_cell.formula, Some("=A4+B4".to_string()));
    }
}