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
        // Don't clear search_results — keep them for n/N navigation
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

    /// Performs autofill operation on the current selection with pattern recognition.
    ///
    /// Analyzes non-empty cells in the selection to detect patterns (arithmetic sequences,
    /// text with numbers, days/months, etc.) and fills empty cells with the continuation
    /// of the detected pattern. For cells containing formulas, adjusts cell references
    /// relatively (original behavior).
    ///
    /// Fill direction is determined by selection shape:
    /// - Tall selection (rows > cols): Fill down (column-wise pattern)
    /// - Wide selection (cols > rows): Fill right (row-wise pattern)
    /// - Square: Default to fill down
    pub fn autofill_selection(&mut self) {
        if let Some(((start_row, start_col), (end_row, end_col))) = self.get_selection_range() {
            let num_rows = end_row - start_row + 1;
            let num_cols = end_col - start_col + 1;

            // Determine fill direction: true = fill down (by rows), false = fill right (by cols)
            let fill_down = num_rows >= num_cols;

            // Collect cells along the fill direction
            // For fill_down: iterate through rows for each column
            // For fill_right: iterate through columns for each row
            let mut changes = Vec::new();
            let mut pattern_desc = String::new();

            if fill_down {
                // Process each column independently
                for col in start_col..=end_col {
                    let (filled, desc) = self.autofill_column(start_row, end_row, col);
                    changes.extend(filled);
                    if pattern_desc.is_empty() && !desc.is_empty() {
                        pattern_desc = desc;
                    }
                }
            } else {
                // Process each row independently
                for row in start_row..=end_row {
                    let (filled, desc) = self.autofill_row(row, start_col, end_col);
                    changes.extend(filled);
                    if pattern_desc.is_empty() && !desc.is_empty() {
                        pattern_desc = desc;
                    }
                }
            }

            // Apply all changes
            let num_changes = changes.len();
            for (row, col, cell_data) in changes {
                self.set_cell_with_undo(row, col, cell_data);
            }

            if num_changes > 0 {
                self.status_message = Some(format!(
                    "Autofilled {} cells using {}",
                    num_changes,
                    pattern_desc
                ));
            } else {
                self.status_message = Some("No cells to fill".to_string());
            }
        }
    }

    /// Autofill a single column from start_row to end_row.
    /// Returns the changes to apply and the pattern description.
    fn autofill_column(&self, start_row: usize, end_row: usize, col: usize) -> (Vec<(usize, usize, CellData)>, String) {
        use crate::domain::services::{FormulaEvaluator, AutofillPattern};

        let mut changes = Vec::new();

        // Collect non-empty cells (pattern cells) and empty cells (target cells)
        let mut pattern_cells: Vec<(usize, CellData)> = Vec::new();
        let mut target_rows: Vec<usize> = Vec::new();

        for row in start_row..=end_row {
            let cell = self.spreadsheet.get_cell(row, col);
            if !cell.value.is_empty() || cell.formula.is_some() {
                pattern_cells.push((row, cell.clone()));
            } else {
                target_rows.push(row);
            }
        }

        // If no pattern cells or no targets, nothing to do
        if pattern_cells.is_empty() || target_rows.is_empty() {
            return (changes, String::new());
        }

        // Check if any pattern cell has a formula - if so, use formula-based fill
        let has_formula = pattern_cells.iter().any(|(_, cell)| cell.formula.is_some());

        if has_formula {
            // Use the first cell with a formula as source, adjust references for targets
            let (source_row, source_cell) = pattern_cells.iter()
                .find(|(_, cell)| cell.formula.is_some())
                .unwrap();

            let evaluator = FormulaEvaluator::new(&self.spreadsheet);

            for target_row in &target_rows {
                let row_offset = *target_row as i32 - *source_row as i32;

                if let Some(ref formula) = source_cell.formula {
                    let adjusted_formula = evaluator.adjust_formula_references(formula, row_offset, 0);

                    if evaluator.would_create_circular_reference(&adjusted_formula, (*target_row, col)) {
                        continue;
                    }

                    let new_value = evaluator.evaluate_formula(&adjusted_formula);
                    changes.push((*target_row, col, CellData {
                        value: new_value,
                        formula: Some(adjusted_formula),
                    }));
                }
            }

            return (changes, "formula".to_string());
        }

        // Extract values from pattern cells for pattern detection
        let values: Vec<String> = pattern_cells.iter()
            .map(|(_, cell)| cell.value.clone())
            .collect();

        let pattern = AutofillPattern::detect(&values);
        let pattern_desc = pattern.description();

        // Generate values for target cells
        // The pattern index for targets continues from where pattern cells left off
        let pattern_len = pattern_cells.len();

        for (i, target_row) in target_rows.iter().enumerate() {
            let pattern_index = pattern_len + i;
            let generated_value = pattern.generate(pattern_index);

            changes.push((*target_row, col, CellData {
                value: generated_value,
                formula: None,
            }));
        }

        (changes, pattern_desc)
    }

    /// Autofill a single row from start_col to end_col.
    /// Returns the changes to apply and the pattern description.
    fn autofill_row(&self, row: usize, start_col: usize, end_col: usize) -> (Vec<(usize, usize, CellData)>, String) {
        use crate::domain::services::{FormulaEvaluator, AutofillPattern};

        let mut changes = Vec::new();

        // Collect non-empty cells (pattern cells) and empty cells (target cells)
        let mut pattern_cells: Vec<(usize, CellData)> = Vec::new();
        let mut target_cols: Vec<usize> = Vec::new();

        for col in start_col..=end_col {
            let cell = self.spreadsheet.get_cell(row, col);
            if !cell.value.is_empty() || cell.formula.is_some() {
                pattern_cells.push((col, cell.clone()));
            } else {
                target_cols.push(col);
            }
        }

        // If no pattern cells or no targets, nothing to do
        if pattern_cells.is_empty() || target_cols.is_empty() {
            return (changes, String::new());
        }

        // Check if any pattern cell has a formula - if so, use formula-based fill
        let has_formula = pattern_cells.iter().any(|(_, cell)| cell.formula.is_some());

        if has_formula {
            // Use the first cell with a formula as source, adjust references for targets
            let (source_col, source_cell) = pattern_cells.iter()
                .find(|(_, cell)| cell.formula.is_some())
                .unwrap();

            let evaluator = FormulaEvaluator::new(&self.spreadsheet);

            for target_col in &target_cols {
                let col_offset = *target_col as i32 - *source_col as i32;

                if let Some(ref formula) = source_cell.formula {
                    let adjusted_formula = evaluator.adjust_formula_references(formula, 0, col_offset);

                    if evaluator.would_create_circular_reference(&adjusted_formula, (row, *target_col)) {
                        continue;
                    }

                    let new_value = evaluator.evaluate_formula(&adjusted_formula);
                    changes.push((row, *target_col, CellData {
                        value: new_value,
                        formula: Some(adjusted_formula),
                    }));
                }
            }

            return (changes, "formula".to_string());
        }

        // Extract values from pattern cells for pattern detection
        let values: Vec<String> = pattern_cells.iter()
            .map(|(_, cell)| cell.value.clone())
            .collect();

        let pattern = AutofillPattern::detect(&values);
        let pattern_desc = pattern.description();

        // Generate values for target cells
        // The pattern index for targets continues from where pattern cells left off
        let pattern_len = pattern_cells.len();

        for (i, target_col) in target_cols.iter().enumerate() {
            let pattern_index = pattern_len + i;
            let generated_value = pattern.generate(pattern_index);

            changes.push((row, *target_col, CellData {
                value: generated_value,
                formula: None,
            }));
        }

        (changes, pattern_desc)
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

        // Set up a simple value in A1 (pattern cell)
        app.set_cell_with_undo(0, 0, CellData {
            value: "Hello".to_string(),
            formula: None,
        });

        // Select A1:A3 (vertical selection, A1 has value, A2-A3 are empty)
        app.selection_start = Some((0, 0));
        app.selection_end = Some((2, 0));

        // Autofill
        app.autofill_selection();

        // Check that the value was copied to empty cells only
        assert_eq!(app.spreadsheet.get_cell(0, 0).value, "Hello"); // Original
        assert_eq!(app.spreadsheet.get_cell(1, 0).value, "Hello"); // Filled
        assert_eq!(app.spreadsheet.get_cell(2, 0).value, "Hello"); // Filled
    }

    #[test]
    fn test_autofill_formula_horizontal() {
        let mut app = App::default();

        // Set up cells with values for reference
        app.set_cell_with_undo(0, 1, CellData { value: "10".to_string(), formula: None }); // B1 = 10
        app.set_cell_with_undo(1, 1, CellData { value: "20".to_string(), formula: None }); // B2 = 20
        app.set_cell_with_undo(0, 2, CellData { value: "30".to_string(), formula: None }); // C1 = 30
        app.set_cell_with_undo(1, 2, CellData { value: "40".to_string(), formula: None }); // C2 = 40

        // Set up a formula in A1 that references B1:B2
        app.set_cell_with_undo(0, 0, CellData {
            value: "30".to_string(),
            formula: Some("=SUM(B1:B2)".to_string()),
        });

        // Select A1:D1 (horizontal autofill, A1 has formula, B1 has value, C1 has value, D1 is empty)
        app.selection_start = Some((0, 0));
        app.selection_end = Some((0, 3));

        // Autofill - should only fill D1 since it's the only empty cell
        app.autofill_selection();

        // Check that only the empty cell D1 got the adjusted formula
        let d1_cell = app.spreadsheet.get_cell(0, 3);
        // The formula from A1 is adjusted by 3 columns: B->E, so =SUM(E1:E2)
        assert_eq!(d1_cell.formula, Some("=SUM(E1:E2)".to_string()));

        // Verify B1 and C1 still have their original values (not overwritten)
        assert_eq!(app.spreadsheet.get_cell(0, 1).value, "10");
        assert_eq!(app.spreadsheet.get_cell(0, 2).value, "30");
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

        // Set up cells with values for reference
        app.set_cell_with_undo(1, 0, CellData { value: "10".to_string(), formula: None }); // A2 = 10
        app.set_cell_with_undo(1, 1, CellData { value: "20".to_string(), formula: None }); // B2 = 20
        app.set_cell_with_undo(2, 0, CellData { value: "30".to_string(), formula: None }); // A3 = 30
        app.set_cell_with_undo(2, 1, CellData { value: "40".to_string(), formula: None }); // B3 = 40

        // Set up a formula in A1 that references A2+B2
        app.set_cell_with_undo(0, 0, CellData {
            value: "30".to_string(),
            formula: Some("=A2+B2".to_string()),
        });

        // Select A1:A4 (vertical autofill, A1 has formula, A2-A3 have values, A4 is empty)
        app.selection_start = Some((0, 0));
        app.selection_end = Some((3, 0));

        // Autofill - should only fill A4 since it's the only empty cell
        app.autofill_selection();

        // Check that only the empty cell A4 got the adjusted formula
        let a4_cell = app.spreadsheet.get_cell(3, 0);
        // The formula from A1 is adjusted by 3 rows: A2->A5, B2->B5, so =A5+B5
        assert_eq!(a4_cell.formula, Some("=A5+B5".to_string()));

        // Verify A2 and A3 still have their original values (not overwritten)
        assert_eq!(app.spreadsheet.get_cell(1, 0).value, "10");
        assert_eq!(app.spreadsheet.get_cell(2, 0).value, "30");
    }

    #[test]
    fn test_autofill_pattern_arithmetic() {
        let mut app = App::default();

        // Set up arithmetic pattern: 1, 2, 3
        app.set_cell_with_undo(0, 0, CellData { value: "1".to_string(), formula: None });
        app.set_cell_with_undo(1, 0, CellData { value: "2".to_string(), formula: None });
        app.set_cell_with_undo(2, 0, CellData { value: "3".to_string(), formula: None });

        // Select A1:A6 (A1-A3 have values, A4-A6 are empty)
        app.selection_start = Some((0, 0));
        app.selection_end = Some((5, 0));

        // Autofill
        app.autofill_selection();

        // Check pattern continuation
        assert_eq!(app.spreadsheet.get_cell(3, 0).value, "4");
        assert_eq!(app.spreadsheet.get_cell(4, 0).value, "5");
        assert_eq!(app.spreadsheet.get_cell(5, 0).value, "6");
    }

    #[test]
    fn test_autofill_pattern_days() {
        let mut app = App::default();

        // Set up days pattern: Mon, Tue
        app.set_cell_with_undo(0, 0, CellData { value: "Mon".to_string(), formula: None });
        app.set_cell_with_undo(1, 0, CellData { value: "Tue".to_string(), formula: None });

        // Select A1:A5 (A1-A2 have values, A3-A5 are empty)
        app.selection_start = Some((0, 0));
        app.selection_end = Some((4, 0));

        // Autofill
        app.autofill_selection();

        // Check pattern continuation
        assert_eq!(app.spreadsheet.get_cell(2, 0).value, "Wed");
        assert_eq!(app.spreadsheet.get_cell(3, 0).value, "Thu");
        assert_eq!(app.spreadsheet.get_cell(4, 0).value, "Fri");
    }

    #[test]
    fn test_autofill_pattern_prefixed() {
        let mut app = App::default();

        // Set up prefixed pattern: Item1, Item2
        app.set_cell_with_undo(0, 0, CellData { value: "Item1".to_string(), formula: None });
        app.set_cell_with_undo(1, 0, CellData { value: "Item2".to_string(), formula: None });

        // Select A1:A5 (A1-A2 have values, A3-A5 are empty)
        app.selection_start = Some((0, 0));
        app.selection_end = Some((4, 0));

        // Autofill
        app.autofill_selection();

        // Check pattern continuation
        assert_eq!(app.spreadsheet.get_cell(2, 0).value, "Item3");
        assert_eq!(app.spreadsheet.get_cell(3, 0).value, "Item4");
        assert_eq!(app.spreadsheet.get_cell(4, 0).value, "Item5");
    }

    #[test]
    fn test_autofill_pattern_months_short() {
        let mut app = App::default();

        // Set up months pattern: Jan, Feb, Mar
        app.set_cell_with_undo(0, 0, CellData { value: "Jan".to_string(), formula: None });
        app.set_cell_with_undo(1, 0, CellData { value: "Feb".to_string(), formula: None });
        app.set_cell_with_undo(2, 0, CellData { value: "Mar".to_string(), formula: None });

        // Select A1:A7 (A1-A3 have values, A4-A7 are empty)
        app.selection_start = Some((0, 0));
        app.selection_end = Some((6, 0));

        // Autofill
        app.autofill_selection();

        // Check pattern continuation
        assert_eq!(app.spreadsheet.get_cell(3, 0).value, "Apr");
        assert_eq!(app.spreadsheet.get_cell(4, 0).value, "May");
        assert_eq!(app.spreadsheet.get_cell(5, 0).value, "Jun");
        assert_eq!(app.spreadsheet.get_cell(6, 0).value, "Jul");
    }

    #[test]
    fn test_autofill_pattern_months_full() {
        let mut app = App::default();

        // Set up full months pattern: January, February
        app.set_cell_with_undo(0, 0, CellData { value: "January".to_string(), formula: None });
        app.set_cell_with_undo(1, 0, CellData { value: "February".to_string(), formula: None });

        // Select A1:A5 (A1-A2 have values, A3-A5 are empty)
        app.selection_start = Some((0, 0));
        app.selection_end = Some((4, 0));

        // Autofill
        app.autofill_selection();

        // Check pattern continuation
        assert_eq!(app.spreadsheet.get_cell(2, 0).value, "March");
        assert_eq!(app.spreadsheet.get_cell(3, 0).value, "April");
        assert_eq!(app.spreadsheet.get_cell(4, 0).value, "May");
    }

    #[test]
    fn test_autofill_pattern_quarters() {
        let mut app = App::default();

        // Set up quarters pattern: Q1, Q2
        app.set_cell_with_undo(0, 0, CellData { value: "Q1".to_string(), formula: None });
        app.set_cell_with_undo(1, 0, CellData { value: "Q2".to_string(), formula: None });

        // Select A1:A6 (A1-A2 have values, A3-A6 are empty)
        app.selection_start = Some((0, 0));
        app.selection_end = Some((5, 0));

        // Autofill
        app.autofill_selection();

        // Check pattern continuation with wrap-around
        assert_eq!(app.spreadsheet.get_cell(2, 0).value, "Q3");
        assert_eq!(app.spreadsheet.get_cell(3, 0).value, "Q4");
        assert_eq!(app.spreadsheet.get_cell(4, 0).value, "Q1"); // Wraps
        assert_eq!(app.spreadsheet.get_cell(5, 0).value, "Q2"); // Wraps
    }

    #[test]
    fn test_autofill_pattern_months_wrap() {
        let mut app = App::default();

        // Set up months pattern starting near end: Oct, Nov, Dec
        app.set_cell_with_undo(0, 0, CellData { value: "Oct".to_string(), formula: None });
        app.set_cell_with_undo(1, 0, CellData { value: "Nov".to_string(), formula: None });
        app.set_cell_with_undo(2, 0, CellData { value: "Dec".to_string(), formula: None });

        // Select A1:A6 (A1-A3 have values, A4-A6 are empty)
        app.selection_start = Some((0, 0));
        app.selection_end = Some((5, 0));

        // Autofill
        app.autofill_selection();

        // Check pattern continuation with wrap-around
        assert_eq!(app.spreadsheet.get_cell(3, 0).value, "Jan"); // Wraps
        assert_eq!(app.spreadsheet.get_cell(4, 0).value, "Feb");
        assert_eq!(app.spreadsheet.get_cell(5, 0).value, "Mar");
    }

    #[test]
    fn test_autofill_horizontal_pattern() {
        let mut app = App::default();

        // Set up arithmetic pattern horizontally: 10, 20 in A1, B1
        app.set_cell_with_undo(0, 0, CellData { value: "10".to_string(), formula: None });
        app.set_cell_with_undo(0, 1, CellData { value: "20".to_string(), formula: None });

        // Select A1:E1 (A1-B1 have values, C1-E1 are empty) - wide selection = fill right
        app.selection_start = Some((0, 0));
        app.selection_end = Some((0, 4));

        // Autofill
        app.autofill_selection();

        // Check pattern continuation
        assert_eq!(app.spreadsheet.get_cell(0, 2).value, "30");
        assert_eq!(app.spreadsheet.get_cell(0, 3).value, "40");
        assert_eq!(app.spreadsheet.get_cell(0, 4).value, "50");
    }
}