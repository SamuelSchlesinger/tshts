//! Application state management for the terminal spreadsheet.
//!
//! This module contains the main application state and mode management
//! for the terminal user interface.

use crate::domain::{Spreadsheet, Workbook, CellData, CellFormat, NumberFormat, TerminalColor, FormulaEvaluator};
use std::collections::{HashSet, VecDeque};

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
    /// Go-to cell mode - user is typing a cell reference
    GoToCell,
    /// Find and replace mode
    FindReplace,
    /// Command palette mode
    CommandPalette,
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
    /// Multiple actions that should be undone/redone atomically
    Batch(Vec<UndoAction>),
}

/// Data stored in the internal clipboard for copy/paste operations.
#[derive(Debug, Clone)]
pub struct ClipboardData {
    /// Cell data relative to top-left of copied region: (row_offset, col_offset, cell_data)
    pub cells: Vec<(usize, usize, CellData)>,
    /// Original top-left position (for cut to know where to clear from)
    pub source_row: usize,
    pub source_col: usize,
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
    /// The workbook containing spreadsheet data
    pub workbook: Workbook,
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
    /// Input buffer for go-to cell reference
    pub goto_cell_input: String,
    /// Internal clipboard for copy/paste
    pub clipboard: Option<ClipboardData>,
    /// Find and replace: search field
    pub find_replace_search: String,
    /// Find and replace: replace field
    pub find_replace_replace: String,
    /// Find and replace: which field is active (false=search, true=replace)
    pub find_replace_on_replace: bool,
    /// Find and replace: search results
    pub find_replace_results: Vec<(usize, usize)>,
    /// Find and replace: current result index
    pub find_replace_index: usize,
    /// Command palette input
    pub command_input: String,
    /// Frozen rows (number of rows frozen from top)
    pub frozen_rows: usize,
    /// Frozen columns (number of columns frozen from left)
    pub frozen_cols: usize,
    /// Hidden rows (for column filtering)
    pub hidden_rows: HashSet<usize>,
    /// Active filter column and criteria
    pub filter_column: Option<usize>,
    pub filter_value: Option<String>,
}

impl Default for App {
    fn default() -> Self {
        Self {
            workbook: Workbook::default(),
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
            goto_cell_input: String::new(),
            clipboard: None,
            find_replace_search: String::new(),
            find_replace_replace: String::new(),
            find_replace_on_replace: false,
            find_replace_results: Vec::new(),
            find_replace_index: 0,
            command_input: String::new(),
            frozen_rows: 0,
            frozen_cols: 0,
            hidden_rows: HashSet::new(),
            filter_column: None,
            filter_value: None,
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
        let cell = self.workbook.current_sheet().get_cell(self.selected_row, self.selected_col);
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
        
        // Move down one cell after editing
        if self.selected_row < self.workbook.current_sheet().rows - 1 {
            self.selected_row += 1;
        }
        
        self.mode = AppMode::Normal;
        self.input.clear();
        self.cursor_position = 0;
    }

    /// Completes editing and moves right (for Tab key).
    pub fn finish_editing_move_right(&mut self) {
        let mut cell_data = CellData::default();

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

        // Move right one cell after editing
        if self.selected_col < self.workbook.current_sheet().cols - 1 {
            self.selected_col += 1;
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
    /// Processes the result of a workbook load operation.
    pub fn set_load_workbook_result(&mut self, result: Result<(Workbook, String), String>) {
        match result {
            Ok((workbook, filename)) => {
                self.workbook = workbook;
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
                *self.workbook.current_sheet_mut() = spreadsheet;
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
            self.apply_undo(&action);
            self.redo_stack.push_back(action);
        }
    }

    fn apply_undo(&mut self, action: &UndoAction) {
        match action {
            UndoAction::CellModified { row, col, old_cell, new_cell: _ } => {
                if let Some(old_data) = old_cell {
                    self.workbook.current_sheet_mut().set_cell(*row, *col, old_data.clone());
                } else {
                    self.workbook.current_sheet_mut().clear_cell(*row, *col);
                }
            }
            UndoAction::Batch(actions) => {
                // Undo in reverse order
                for a in actions.iter().rev() {
                    self.apply_undo(a);
                }
            }
        }
    }

    /// Performs a redo operation.
    ///
    /// Reapplies the last undone action and moves it back to the undo stack.
    pub fn redo(&mut self) {
        if let Some(action) = self.redo_stack.pop_back() {
            self.apply_redo(&action);
            self.undo_stack.push_back(action);
        }
    }

    fn apply_redo(&mut self, action: &UndoAction) {
        match action {
            UndoAction::CellModified { row, col, old_cell: _, new_cell } => {
                if let Some(new_data) = new_cell {
                    self.workbook.current_sheet_mut().set_cell(*row, *col, new_data.clone());
                } else {
                    self.workbook.current_sheet_mut().clear_cell(*row, *col);
                }
            }
            UndoAction::Batch(actions) => {
                for a in actions {
                    self.apply_redo(a);
                }
            }
        }
    }

    /// Sets a cell value and records the action for undo/redo.
    ///
    /// This is a wrapper around the spreadsheet's set_cell method that also
    /// tracks the change for undo functionality.
    pub fn set_cell_with_undo(&mut self, row: usize, col: usize, new_data: CellData) {
        // Get the old cell data
        let old_cell = if self.workbook.current_sheet().cells.contains_key(&(row, col)) {
            Some(self.workbook.current_sheet().get_cell(row, col))
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
        self.workbook.current_sheet_mut().set_cell(row, col, new_data);
    }

    /// Clears a cell and records the action for undo/redo.
    ///
    /// This is a wrapper around the spreadsheet's clear_cell method that also
    /// tracks the change for undo functionality.
    pub fn clear_cell_with_undo(&mut self, row: usize, col: usize) {
        // Get the old cell data
        let old_cell = if self.workbook.current_sheet().cells.contains_key(&(row, col)) {
            Some(self.workbook.current_sheet().get_cell(row, col))
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
        self.workbook.current_sheet_mut().clear_cell(row, col);
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
        for row in 0..self.workbook.current_sheet().rows {
            for col in 0..self.workbook.current_sheet().cols {
                let cell = self.workbook.current_sheet().get_cell(row, col);
                
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

    /// Copies the current selection to both the internal and system clipboards.
    pub fn copy_selection(&mut self) {
        let range = if let Some(range) = self.get_selection_range() {
            range
        } else {
            // No selection: copy current cell
            ((self.selected_row, self.selected_col), (self.selected_row, self.selected_col))
        };
        let ((start_row, start_col), (end_row, end_col)) = range;
        let mut cells = Vec::new();
        for row in start_row..=end_row {
            for col in start_col..=end_col {
                let cell = self.workbook.current_sheet().get_cell(row, col);
                if !cell.value.is_empty() || cell.formula.is_some() {
                    cells.push((row - start_row, col - start_col, cell));
                }
            }
        }
        let count = (end_row - start_row + 1) * (end_col - start_col + 1);

        // Build TSV string for system clipboard
        let mut tsv = String::new();
        for row in start_row..=end_row {
            for col in start_col..=end_col {
                if col > start_col {
                    tsv.push('\t');
                }
                let cell = self.workbook.current_sheet().get_cell(row, col);
                tsv.push_str(&cell.value);
            }
            tsv.push('\n');
        }
        // Write to system clipboard (best-effort)
        if let Ok(mut board) = arboard::Clipboard::new() {
            let _ = board.set_text(tsv);
        }

        self.clipboard = Some(ClipboardData {
            cells,
            source_row: start_row,
            source_col: start_col,
        });
        self.status_message = Some(format!("Copied {} cell(s)", count));
    }

    /// Cuts the current selection to the internal clipboard.
    pub fn cut_selection(&mut self) {
        let range = if let Some(range) = self.get_selection_range() {
            range
        } else {
            ((self.selected_row, self.selected_col), (self.selected_row, self.selected_col))
        };
        let ((start_row, start_col), (end_row, end_col)) = range;
        let mut cells = Vec::new();
        for row in start_row..=end_row {
            for col in start_col..=end_col {
                let cell = self.workbook.current_sheet().get_cell(row, col);
                if !cell.value.is_empty() || cell.formula.is_some() {
                    cells.push((row - start_row, col - start_col, cell));
                }
            }
        }
        let count = (end_row - start_row + 1) * (end_col - start_col + 1);
        self.clipboard = Some(ClipboardData {
            cells,
            source_row: start_row,
            source_col: start_col,
        });
        // Clear the cut cells
        let mut batch = Vec::new();
        for row in start_row..=end_row {
            for col in start_col..=end_col {
                let old = if self.workbook.current_sheet().cells.contains_key(&(row, col)) {
                    Some(self.workbook.current_sheet().get_cell(row, col))
                } else {
                    None
                };
                if old.is_some() {
                    batch.push(UndoAction::CellModified { row, col, old_cell: old, new_cell: None });
                    self.workbook.current_sheet_mut().clear_cell(row, col);
                }
            }
        }
        if !batch.is_empty() {
            self.record_action(UndoAction::Batch(batch));
        }
        self.status_message = Some(format!("Cut {} cell(s)", count));
    }

    /// Pastes clipboard contents at the current cursor position.
    /// Falls back to system clipboard if internal clipboard is empty.
    pub fn paste(&mut self) {
        let clipboard = if let Some(ref cb) = self.clipboard {
            cb.clone()
        } else {
            // Try system clipboard
            if let Ok(mut board) = arboard::Clipboard::new() {
                if let Ok(text) = board.get_text() {
                    if !text.is_empty() {
                        self.paste_tsv(&text);
                        return;
                    }
                }
            }
            self.status_message = Some("Nothing to paste".to_string());
            return;
        };

        let dest_row = self.selected_row;
        let dest_col = self.selected_col;

        // Compute all new cells first (evaluator borrows spreadsheet immutably)
        let new_cells: Vec<_> = {
            let evaluator = crate::domain::FormulaEvaluator::new(self.workbook.current_sheet());
            clipboard.cells.iter().filter_map(|(row_off, col_off, cell)| {
                let target_row = dest_row + row_off;
                let target_col = dest_col + col_off;
                if target_row >= self.workbook.current_sheet().rows || target_col >= self.workbook.current_sheet().cols {
                    return None;
                }
                let new_cell = if let Some(ref formula) = cell.formula {
                    let row_offset = target_row as i32 - (clipboard.source_row + row_off) as i32;
                    let col_offset = target_col as i32 - (clipboard.source_col + col_off) as i32;
                    let adjusted = evaluator.adjust_formula_references(formula, row_offset, col_offset);
                    let value = evaluator.evaluate_formula(&adjusted);
                    CellData { value, formula: Some(adjusted), format: None, comment: None }
                } else {
                    cell.clone()
                };
                Some((target_row, target_col, new_cell))
            }).collect()
        };

        // Now apply changes (mutably borrows spreadsheet)
        let mut batch = Vec::new();
        for (target_row, target_col, new_cell) in &new_cells {
            let old = if self.workbook.current_sheet().cells.contains_key(&(*target_row, *target_col)) {
                Some(self.workbook.current_sheet().get_cell(*target_row, *target_col))
            } else {
                None
            };
            batch.push(UndoAction::CellModified {
                row: *target_row,
                col: *target_col,
                old_cell: old,
                new_cell: Some(new_cell.clone()),
            });
            self.workbook.current_sheet_mut().set_cell(*target_row, *target_col, new_cell.clone());
        }

        if !batch.is_empty() {
            self.record_action(UndoAction::Batch(batch));
        }
        self.status_message = Some(format!("Pasted {} cell(s)", clipboard.cells.len()));
    }

    /// Pastes TSV text from the system clipboard at the current cursor position.
    fn paste_tsv(&mut self, text: &str) {
        let dest_row = self.selected_row;
        let dest_col = self.selected_col;
        let mut batch = Vec::new();
        let mut count = 0;

        for (row_offset, line) in text.lines().enumerate() {
            if line.is_empty() { continue; }
            for (col_offset, value) in line.split('\t').enumerate() {
                let target_row = dest_row + row_offset;
                let target_col = dest_col + col_offset;
                if target_row >= self.workbook.current_sheet().rows || target_col >= self.workbook.current_sheet().cols {
                    continue;
                }
                let old = if self.workbook.current_sheet().cells.contains_key(&(target_row, target_col)) {
                    Some(self.workbook.current_sheet().get_cell(target_row, target_col))
                } else {
                    None
                };
                let new_cell = CellData { value: value.to_string(), formula: None, format: None, comment: None };
                batch.push(UndoAction::CellModified {
                    row: target_row,
                    col: target_col,
                    old_cell: old,
                    new_cell: Some(new_cell.clone()),
                });
                self.workbook.current_sheet_mut().set_cell(target_row, target_col, new_cell);
                count += 1;
            }
        }
        if !batch.is_empty() {
            self.record_action(UndoAction::Batch(batch));
        }
        self.status_message = Some(format!("Pasted {} cell(s) from system clipboard", count));
    }

    /// Inserts a row above the current cursor position.
    pub fn insert_row(&mut self) {
        let insert_at = self.selected_row;
        self.workbook.current_sheet_mut().insert_row(insert_at);
        self.status_message = Some(format!("Inserted row at {}", insert_at + 1));
    }

    /// Deletes the current row.
    pub fn delete_row(&mut self) {
        let delete_at = self.selected_row;
        self.workbook.current_sheet_mut().delete_row(delete_at);
        if self.selected_row >= self.workbook.current_sheet().rows {
            self.selected_row = self.workbook.current_sheet().rows.saturating_sub(1);
        }
        self.status_message = Some(format!("Deleted row {}", delete_at + 1));
    }

    /// Inserts a column to the left of the current cursor position.
    pub fn insert_col(&mut self) {
        let insert_at = self.selected_col;
        self.workbook.current_sheet_mut().insert_col(insert_at);
        self.status_message = Some(format!("Inserted column at {}", crate::domain::Spreadsheet::column_label(insert_at)));
    }

    /// Deletes the current column.
    pub fn delete_col(&mut self) {
        let delete_at = self.selected_col;
        self.workbook.current_sheet_mut().delete_col(delete_at);
        if self.selected_col >= self.workbook.current_sheet().cols {
            self.selected_col = self.workbook.current_sheet().cols.saturating_sub(1);
        }
        self.status_message = Some(format!("Deleted column {}", crate::domain::Spreadsheet::column_label(delete_at)));
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
        for row in 0..self.workbook.current_sheet().rows {
            for col in 0..self.workbook.current_sheet().cols {
                let cell = self.workbook.current_sheet().get_cell(row, col);
                if cell.value.to_lowercase().contains(&query) {
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

    /// Replaces the current find-replace match.
    pub fn replace_current(&mut self) {
        if self.find_replace_results.is_empty() { return; }
        let (row, col) = self.find_replace_results[self.find_replace_index];
        let cell = self.workbook.current_sheet().get_cell(row, col);
        if cell.formula.is_some() { return; } // Don't replace in formula cells
        let new_value = cell.value.replace(&self.find_replace_search, &self.find_replace_replace);
        let new_cell = CellData { value: new_value, formula: None, format: None, comment: None };
        self.set_cell_with_undo(row, col, new_cell);
        // Re-search and move to next
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
            let new_value = cell.value.replace(&self.find_replace_search, &self.find_replace_replace);
            let new_cell = CellData { value: new_value, formula: None, format: None, comment: None };
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

    /// Starts command palette mode.
    pub fn start_command_palette(&mut self) {
        self.mode = AppMode::CommandPalette;
        self.command_input.clear();
        self.cursor_position = 0;
        self.status_message = None;
    }

    /// Executes a command from the command palette.
    pub fn execute_command(&mut self) {
        // Handle rename specially to preserve case
        let trimmed = self.command_input.trim().to_string();
        if trimmed.starts_with("rename ") || trimmed.starts_with("RENAME ") {
            let name = trimmed[7..].trim().to_string();
            if !name.is_empty() {
                self.workbook.rename_sheet(name.clone());
                self.status_message = Some(format!("Renamed sheet to '{}'", name));
            } else {
                self.status_message = Some("Usage: rename <name>".to_string());
            }
            self.mode = AppMode::Normal;
            self.command_input.clear();
            self.cursor_position = 0;
            return;
        }

        let cmd = trimmed.to_lowercase();
        let parts: Vec<&str> = cmd.split_whitespace().collect();
        match parts.as_slice() {
            ["ir"] | ["insert", "row"] => self.insert_row(),
            ["dr"] | ["delete", "row"] => self.delete_row(),
            ["ic"] | ["insert", "col"] | ["insert", "column"] => self.insert_col(),
            ["dc"] | ["delete", "col"] | ["delete", "column"] => self.delete_col(),
            ["sort", "asc"] => self.sort_column_asc(),
            ["sort", "desc"] => self.sort_column_desc(),
            ["freeze"] => {
                self.frozen_rows = self.selected_row;
                self.frozen_cols = self.selected_col;
                self.status_message = Some(format!("Frozen {} rows, {} cols", self.frozen_rows, self.frozen_cols));
            }
            ["unfreeze"] => {
                self.frozen_rows = 0;
                self.frozen_cols = 0;
                self.status_message = Some("Unfrozen all panes".to_string());
            }
            ["format", "general"] => self.set_selection_format(NumberFormat::General),
            ["format", "number"] => self.set_selection_format(NumberFormat::Number { decimals: 2, thousands_sep: false }),
            ["format", "number", d] => {
                if let Ok(decimals) = d.parse::<u32>() {
                    self.set_selection_format(NumberFormat::Number { decimals, thousands_sep: false });
                } else {
                    self.status_message = Some("Invalid decimal count".to_string());
                }
            }
            ["format", "currency"] => self.set_selection_format(NumberFormat::Currency { symbol: "$".to_string(), decimals: 2 }),
            ["format", "currency", sym] => self.set_selection_format(NumberFormat::Currency { symbol: sym.to_string(), decimals: 2 }),
            ["format", "percent"] | ["format", "percentage"] => self.set_selection_format(NumberFormat::Percentage { decimals: 1 }),
            ["format", "percent", d] | ["format", "percentage", d] => {
                if let Ok(decimals) = d.parse::<u32>() {
                    self.set_selection_format(NumberFormat::Percentage { decimals });
                } else {
                    self.status_message = Some("Invalid decimal count".to_string());
                }
            }
            ["bold"] => self.toggle_bold(),
            ["underline"] => self.toggle_underline(),
            ["color", color_name] => {
                if *color_name == "none" || *color_name == "default" {
                    self.set_selection_fg_color(None);
                } else if let Some(c) = TerminalColor::from_name(color_name) {
                    self.set_selection_fg_color(Some(c));
                } else {
                    self.status_message = Some(format!("Unknown color: {}", color_name));
                }
            }
            ["bg", color_name] => {
                if *color_name == "none" || *color_name == "default" {
                    self.set_selection_bg_color(None);
                } else if let Some(c) = TerminalColor::from_name(color_name) {
                    self.set_selection_bg_color(Some(c));
                } else {
                    self.status_message = Some(format!("Unknown color: {}", color_name));
                }
            }
            ["sheet", "new"] | ["new", "sheet"] | ["addsheet"] => {
                let name = format!("Sheet{}", self.workbook.sheets.len() + 1);
                self.workbook.add_sheet(name.clone());
                self.workbook.active_sheet = self.workbook.sheets.len() - 1;
                self.selected_row = 0;
                self.selected_col = 0;
                self.scroll_row = 0;
                self.scroll_col = 0;
                self.status_message = Some(format!("Added sheet '{}'", name));
            }
            ["sheet", "delete"] | ["delsheet"] => {
                let name = self.workbook.sheet_names[self.workbook.active_sheet].clone();
                if self.workbook.remove_sheet(self.workbook.active_sheet) {
                    self.selected_row = 0;
                    self.selected_col = 0;
                    self.scroll_row = 0;
                    self.scroll_col = 0;
                    self.status_message = Some(format!("Deleted sheet '{}'", name));
                } else {
                    self.status_message = Some("Cannot delete the last sheet".to_string());
                }
            }
            ["sheet", "next"] | ["sn"] => {
                self.switch_next_sheet();
            }
            ["sheet", "prev"] | ["sp"] => {
                self.switch_prev_sheet();
            }
            ["comment", ..] => {
                let text = parts[1..].join(" ");
                if text == "clear" || text == "none" {
                    self.set_cell_comment(None);
                } else {
                    self.set_cell_comment(Some(text));
                }
            }
            ["filter", column_name] => {
                if let Some(col) = Spreadsheet::parse_column_label(column_name) {
                    self.apply_filter(col, None);
                } else {
                    self.status_message = Some(format!("Invalid column: {}", column_name));
                }
            }
            ["filter", column_name, ..] => {
                if let Some(col) = Spreadsheet::parse_column_label(column_name) {
                    let criteria = parts[2..].join(" ");
                    self.apply_filter(col, Some(criteria));
                } else {
                    self.status_message = Some(format!("Invalid column: {}", column_name));
                }
            }
            ["unfilter"] | ["clearfilter"] | ["clear", "filter"] => {
                self.clear_filter();
            }
            _ => {
                self.status_message = Some(format!("Unknown command: {}", self.command_input));
            }
        }
        self.mode = AppMode::Normal;
        self.command_input.clear();
        self.cursor_position = 0;
    }

    /// Sorts all data rows by the current column, ascending.
    pub fn sort_column_asc(&mut self) {
        self.sort_column(true);
    }

    /// Sorts all data rows by the current column, descending.
    pub fn sort_column_desc(&mut self) {
        self.sort_column(false);
    }

    fn sort_column(&mut self, ascending: bool) {
        let col = self.selected_col;
        // Find data bounds
        let mut max_row = 0;
        let mut max_col = 0;
        for &(r, c) in self.workbook.current_sheet().cells.keys() {
            max_row = max_row.max(r);
            max_col = max_col.max(c);
        }
        if max_row == 0 { return; }

        // Collect all rows as Vec of cell data
        let mut rows: Vec<Vec<Option<CellData>>> = Vec::new();
        for row in 0..=max_row {
            let mut row_data = Vec::new();
            for c in 0..=max_col {
                if self.workbook.current_sheet().cells.contains_key(&(row, c)) {
                    row_data.push(Some(self.workbook.current_sheet().get_cell(row, c)));
                } else {
                    row_data.push(None);
                }
            }
            rows.push(row_data);
        }

        // Sort rows by the selected column
        rows.sort_by(|a, b| {
            let a_val = a.get(col).and_then(|c| c.as_ref()).map(|c| &c.value);
            let b_val = b.get(col).and_then(|c| c.as_ref()).map(|c| &c.value);

            let cmp = match (a_val, b_val) {
                (None, None) => std::cmp::Ordering::Equal,
                (None, Some(_)) => std::cmp::Ordering::Greater,
                (Some(_), None) => std::cmp::Ordering::Less,
                (Some(a), Some(b)) => {
                    match (a.parse::<f64>(), b.parse::<f64>()) {
                        (Ok(an), Ok(bn)) => an.partial_cmp(&bn).unwrap_or(std::cmp::Ordering::Equal),
                        _ => a.cmp(b),
                    }
                }
            };
            if ascending { cmp } else { cmp.reverse() }
        });

        // Apply sorted rows back (batch undo via snapshot)
        let mut batch = Vec::new();
        for (row_idx, row_data) in rows.iter().enumerate() {
            for (col_idx, cell_opt) in row_data.iter().enumerate() {
                let old = if self.workbook.current_sheet().cells.contains_key(&(row_idx, col_idx)) {
                    Some(self.workbook.current_sheet().get_cell(row_idx, col_idx))
                } else {
                    None
                };
                let new = cell_opt.clone();
                if old != new {
                    batch.push(UndoAction::CellModified {
                        row: row_idx,
                        col: col_idx,
                        old_cell: old,
                        new_cell: new.clone(),
                    });
                    if let Some(cell) = new {
                        self.workbook.current_sheet_mut().set_cell(row_idx, col_idx, cell);
                    } else {
                        self.workbook.current_sheet_mut().clear_cell(row_idx, col_idx);
                    }
                }
            }
        }
        if !batch.is_empty() {
            self.record_action(UndoAction::Batch(batch));
        }
        let dir = if ascending { "ascending" } else { "descending" };
        self.status_message = Some(format!("Sorted by column {} {}", crate::domain::Spreadsheet::column_label(col), dir));
    }

    /// Sets the number format on the current selection or current cell.
    pub fn set_selection_format(&mut self, number_format: NumberFormat) {
        let range = if let Some(range) = self.get_selection_range() {
            range
        } else {
            ((self.selected_row, self.selected_col), (self.selected_row, self.selected_col))
        };
        let ((start_row, start_col), (end_row, end_col)) = range;
        let fmt_name = match &number_format {
            NumberFormat::General => "General",
            NumberFormat::Number { .. } => "Number",
            NumberFormat::Currency { .. } => "Currency",
            NumberFormat::Percentage { .. } => "Percentage",
        };
        let mut count = 0;
        for row in start_row..=end_row {
            for col in start_col..=end_col {
                let mut cell = self.workbook.current_sheet().get_cell(row, col);
                let format = match &number_format {
                    NumberFormat::General => None,
                    _ => Some(CellFormat { number_format: number_format.clone(), ..CellFormat::default() }),
                };
                cell.format = format;
                self.workbook.current_sheet_mut().set_cell(row, col, cell);
                count += 1;
            }
        }
        self.status_message = Some(format!("Applied {} format to {} cell(s)", fmt_name, count));
    }

    /// Toggles bold on the current selection or current cell.
    pub fn toggle_bold(&mut self) {
        let range = if let Some(range) = self.get_selection_range() {
            range
        } else {
            ((self.selected_row, self.selected_col), (self.selected_row, self.selected_col))
        };
        let ((start_row, start_col), (end_row, end_col)) = range;
        // Check current state from first cell to determine toggle direction
        let first_cell = self.workbook.current_sheet().get_cell(start_row, start_col);
        let currently_bold = first_cell.format.as_ref().map(|f| f.style.bold).unwrap_or(false);
        let new_bold = !currently_bold;

        for row in start_row..=end_row {
            for col in start_col..=end_col {
                let mut cell = self.workbook.current_sheet().get_cell(row, col);
                let mut fmt = cell.format.unwrap_or_default();
                fmt.style.bold = new_bold;
                cell.format = Some(fmt);
                self.workbook.current_sheet_mut().set_cell(row, col, cell);
            }
        }
        self.status_message = Some(format!("Bold {}", if new_bold { "on" } else { "off" }));
    }

    /// Toggles underline on the current selection or current cell.
    pub fn toggle_underline(&mut self) {
        let range = if let Some(range) = self.get_selection_range() {
            range
        } else {
            ((self.selected_row, self.selected_col), (self.selected_row, self.selected_col))
        };
        let ((start_row, start_col), (end_row, end_col)) = range;
        let first_cell = self.workbook.current_sheet().get_cell(start_row, start_col);
        let currently_underline = first_cell.format.as_ref().map(|f| f.style.underline).unwrap_or(false);
        let new_underline = !currently_underline;

        for row in start_row..=end_row {
            for col in start_col..=end_col {
                let mut cell = self.workbook.current_sheet().get_cell(row, col);
                let mut fmt = cell.format.unwrap_or_default();
                fmt.style.underline = new_underline;
                cell.format = Some(fmt);
                self.workbook.current_sheet_mut().set_cell(row, col, cell);
            }
        }
        self.status_message = Some(format!("Underline {}", if new_underline { "on" } else { "off" }));
    }

    /// Sets the foreground color on the current selection or current cell.
    pub fn set_selection_fg_color(&mut self, color: Option<TerminalColor>) {
        let range = if let Some(range) = self.get_selection_range() {
            range
        } else {
            ((self.selected_row, self.selected_col), (self.selected_row, self.selected_col))
        };
        let ((start_row, start_col), (end_row, end_col)) = range;
        for row in start_row..=end_row {
            for col in start_col..=end_col {
                let mut cell = self.workbook.current_sheet().get_cell(row, col);
                let mut fmt = cell.format.unwrap_or_default();
                fmt.style.fg_color = color.clone();
                cell.format = Some(fmt);
                self.workbook.current_sheet_mut().set_cell(row, col, cell);
            }
        }
        let color_name = color.as_ref().map(|c| format!("{:?}", c)).unwrap_or("default".to_string());
        self.status_message = Some(format!("Set foreground color to {}", color_name));
    }

    /// Sets the background color on the current selection or current cell.
    pub fn set_selection_bg_color(&mut self, color: Option<TerminalColor>) {
        let range = if let Some(range) = self.get_selection_range() {
            range
        } else {
            ((self.selected_row, self.selected_col), (self.selected_row, self.selected_col))
        };
        let ((start_row, start_col), (end_row, end_col)) = range;
        for row in start_row..=end_row {
            for col in start_col..=end_col {
                let mut cell = self.workbook.current_sheet().get_cell(row, col);
                let mut fmt = cell.format.unwrap_or_default();
                fmt.style.bg_color = color.clone();
                cell.format = Some(fmt);
                self.workbook.current_sheet_mut().set_cell(row, col, cell);
            }
        }
        let color_name = color.as_ref().map(|c| format!("{:?}", c)).unwrap_or("default".to_string());
        self.status_message = Some(format!("Set background color to {}", color_name));
    }

    /// Switches to the next sheet.
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

    /// Switches to the previous sheet.
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

    /// Sets a comment on the currently selected cell.
    pub fn set_cell_comment(&mut self, comment: Option<String>) {
        let row = self.selected_row;
        let col = self.selected_col;
        let mut cell = self.workbook.current_sheet().get_cell(row, col);
        let old_cell = if self.workbook.current_sheet().cells.contains_key(&(row, col)) {
            Some(cell.clone())
        } else {
            None
        };
        cell.comment = comment.clone();
        self.workbook.current_sheet_mut().set_cell(row, col, cell.clone());
        self.record_action(UndoAction::CellModified {
            row, col,
            old_cell,
            new_cell: Some(cell),
        });
        if let Some(ref text) = comment {
            self.status_message = Some(format!("Comment set: {}", text));
        } else {
            self.status_message = Some("Comment cleared".to_string());
        }
    }

    /// Applies a filter on a column. If criteria is None, shows all unique values.
    /// If criteria is Some, hides rows where the column value doesn't match.
    pub fn apply_filter(&mut self, col: usize, criteria: Option<String>) {
        self.hidden_rows.clear();
        self.filter_column = Some(col);
        if let Some(ref criteria) = criteria {
            self.filter_value = Some(criteria.clone());
            let criteria_lower = criteria.to_lowercase();
            // Find data extent
            let max_row = self.workbook.current_sheet().cells.keys()
                .map(|&(r, _)| r)
                .max()
                .unwrap_or(0);
            for row in 0..=max_row {
                let cell = self.workbook.current_sheet().get_cell(row, col);
                if !cell.value.to_lowercase().contains(&criteria_lower) {
                    self.hidden_rows.insert(row);
                }
            }
            let hidden_count = self.hidden_rows.len();
            self.status_message = Some(format!("Filter applied: {} rows hidden", hidden_count));
        } else {
            self.filter_value = None;
            self.status_message = Some(format!("Filter set on column {}", Spreadsheet::column_label(col)));
        }
    }

    /// Clears any active filter, showing all rows.
    pub fn clear_filter(&mut self) {
        self.hidden_rows.clear();
        self.filter_column = None;
        self.filter_value = None;
        self.status_message = Some("Filter cleared".to_string());
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

    /// Cancels go-to cell mode.
    pub fn cancel_goto_cell(&mut self) {
        self.mode = AppMode::Normal;
        self.goto_cell_input.clear();
        self.cursor_position = 0;
    }

    /// Jumps to cell A1 (Ctrl+Home).
    pub fn jump_to_home(&mut self) {
        self.selected_row = 0;
        self.selected_col = 0;
        self.scroll_row = 0;
        self.scroll_col = 0;
    }

    /// Jumps to the last cell with data (Ctrl+End).
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

    /// Computes aggregate stats (SUM, AVERAGE, COUNT) for the current selection.
    pub fn get_selection_stats(&self) -> Option<(f64, f64, usize)> {
        let ((start_row, start_col), (end_row, end_col)) = self.get_selection_range()?;
        if start_row == end_row && start_col == end_col {
            return None; // Single cell, no stats
        }
        let mut sum = 0.0;
        let mut count = 0usize;
        for row in start_row..=end_row {
            for col in start_col..=end_col {
                let cell = self.workbook.current_sheet().get_cell(row, col);
                if let Ok(n) = cell.value.parse::<f64>() {
                    sum += n;
                    count += 1;
                }
            }
        }
        if count > 0 {
            Some((sum, sum / count as f64, count))
        } else {
            None
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
            let cell = self.workbook.current_sheet().get_cell(row, col);
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

            let evaluator = FormulaEvaluator::new(self.workbook.current_sheet());

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
                        format: None,
                        comment: None,
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
                format: None,
                comment: None,
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
            let cell = self.workbook.current_sheet().get_cell(row, col);
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

            let evaluator = FormulaEvaluator::new(self.workbook.current_sheet());

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
                        format: None,
                        comment: None,
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
                format: None,
                comment: None,
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
            format: None,
            comment: None,
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
    fn test_set_load_workbook_result_success() {
        let mut app = App::default();
        app.selected_row = 5;
        app.selected_col = 3;
        app.scroll_row = 2;
        app.scroll_col = 1;

        let mut new_sheet = Spreadsheet::default();
        new_sheet.set_cell(0, 0, CellData {
            value: "Loaded".to_string(),
            formula: None,
            format: None,
            comment: None,
        });
        let workbook = Workbook::from_spreadsheet(new_sheet);

        app.set_load_workbook_result(Ok((workbook, "loaded.tshts".to_string())));

        assert!(matches!(app.mode, AppMode::Normal));
        assert_eq!(app.filename.unwrap(), "loaded.tshts");
        assert!(app.status_message.unwrap().contains("Loaded from loaded.tshts"));

        // Position should be reset
        assert_eq!(app.selected_row, 0);
        assert_eq!(app.selected_col, 0);
        assert_eq!(app.scroll_row, 0);
        assert_eq!(app.scroll_col, 0);

        // Spreadsheet should be updated
        let cell = app.workbook.current_sheet().get_cell(0, 0);
        assert_eq!(cell.value, "Loaded");
    }

    #[test]
    fn test_set_load_workbook_result_failure() {
        let mut app = App::default();
        let original_sheet = app.workbook.current_sheet().clone();

        app.set_load_workbook_result(Err("File not found".to_string()));

        assert!(matches!(app.mode, AppMode::Normal));
        assert!(app.filename.is_none());
        assert!(app.status_message.unwrap().contains("Load failed: File not found"));

        // Spreadsheet should remain unchanged
        assert_eq!(app.workbook.current_sheet().rows, original_sheet.rows);
        assert_eq!(app.workbook.current_sheet().cols, original_sheet.cols);
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
        app.set_load_workbook_result(Err("Error".to_string()));
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
            format: None,
            comment: None,
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
        let cell = app.workbook.current_sheet().get_cell(0, 0);
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
            format: None,
            comment: None,
        });

        // Select A1:A3 (vertical selection, A1 has value, A2-A3 are empty)
        app.selection_start = Some((0, 0));
        app.selection_end = Some((2, 0));

        // Autofill
        app.autofill_selection();

        // Check that the value was copied to empty cells only
        assert_eq!(app.workbook.current_sheet().get_cell(0, 0).value, "Hello"); // Original
        assert_eq!(app.workbook.current_sheet().get_cell(1, 0).value, "Hello"); // Filled
        assert_eq!(app.workbook.current_sheet().get_cell(2, 0).value, "Hello"); // Filled
    }

    #[test]
    fn test_autofill_formula_horizontal() {
        let mut app = App::default();

        // Set up cells with values for reference
        app.set_cell_with_undo(0, 1, CellData { value: "10".to_string(), formula: None, format: None, comment: None }); // B1 = 10
        app.set_cell_with_undo(1, 1, CellData { value: "20".to_string(), formula: None, format: None, comment: None }); // B2 = 20
        app.set_cell_with_undo(0, 2, CellData { value: "30".to_string(), formula: None, format: None, comment: None }); // C1 = 30
        app.set_cell_with_undo(1, 2, CellData { value: "40".to_string(), formula: None, format: None, comment: None }); // C2 = 40

        // Set up a formula in A1 that references B1:B2
        app.set_cell_with_undo(0, 0, CellData {
            value: "30".to_string(),
            formula: Some("=SUM(B1:B2)".to_string()),
            format: None,
            comment: None,
        });

        // Select A1:D1 (horizontal autofill, A1 has formula, B1 has value, C1 has value, D1 is empty)
        app.selection_start = Some((0, 0));
        app.selection_end = Some((0, 3));

        // Autofill - should only fill D1 since it's the only empty cell
        app.autofill_selection();

        // Check that only the empty cell D1 got the adjusted formula
        let d1_cell = app.workbook.current_sheet().get_cell(0, 3);
        // The formula from A1 is adjusted by 3 columns: B->E, so =SUM(E1:E2)
        assert_eq!(d1_cell.formula, Some("=SUM(E1:E2)".to_string()));

        // Verify B1 and C1 still have their original values (not overwritten)
        assert_eq!(app.workbook.current_sheet().get_cell(0, 1).value, "10");
        assert_eq!(app.workbook.current_sheet().get_cell(0, 2).value, "30");
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
        app.set_cell_with_undo(1, 0, CellData { value: "10".to_string(), formula: None, format: None, comment: None }); // A2 = 10
        app.set_cell_with_undo(1, 1, CellData { value: "20".to_string(), formula: None, format: None, comment: None }); // B2 = 20
        app.set_cell_with_undo(2, 0, CellData { value: "30".to_string(), formula: None, format: None, comment: None }); // A3 = 30
        app.set_cell_with_undo(2, 1, CellData { value: "40".to_string(), formula: None, format: None, comment: None }); // B3 = 40

        // Set up a formula in A1 that references A2+B2
        app.set_cell_with_undo(0, 0, CellData {
            value: "30".to_string(),
            formula: Some("=A2+B2".to_string()),
            format: None,
            comment: None,
        });

        // Select A1:A4 (vertical autofill, A1 has formula, A2-A3 have values, A4 is empty)
        app.selection_start = Some((0, 0));
        app.selection_end = Some((3, 0));

        // Autofill - should only fill A4 since it's the only empty cell
        app.autofill_selection();

        // Check that only the empty cell A4 got the adjusted formula
        let a4_cell = app.workbook.current_sheet().get_cell(3, 0);
        // The formula from A1 is adjusted by 3 rows: A2->A5, B2->B5, so =A5+B5
        assert_eq!(a4_cell.formula, Some("=A5+B5".to_string()));

        // Verify A2 and A3 still have their original values (not overwritten)
        assert_eq!(app.workbook.current_sheet().get_cell(1, 0).value, "10");
        assert_eq!(app.workbook.current_sheet().get_cell(2, 0).value, "30");
    }

    #[test]
    fn test_autofill_pattern_arithmetic() {
        let mut app = App::default();

        // Set up arithmetic pattern: 1, 2, 3
        app.set_cell_with_undo(0, 0, CellData { value: "1".to_string(), formula: None, format: None, comment: None });
        app.set_cell_with_undo(1, 0, CellData { value: "2".to_string(), formula: None, format: None, comment: None });
        app.set_cell_with_undo(2, 0, CellData { value: "3".to_string(), formula: None, format: None, comment: None });

        // Select A1:A6 (A1-A3 have values, A4-A6 are empty)
        app.selection_start = Some((0, 0));
        app.selection_end = Some((5, 0));

        // Autofill
        app.autofill_selection();

        // Check pattern continuation
        assert_eq!(app.workbook.current_sheet().get_cell(3, 0).value, "4");
        assert_eq!(app.workbook.current_sheet().get_cell(4, 0).value, "5");
        assert_eq!(app.workbook.current_sheet().get_cell(5, 0).value, "6");
    }

    #[test]
    fn test_autofill_pattern_days() {
        let mut app = App::default();

        // Set up days pattern: Mon, Tue
        app.set_cell_with_undo(0, 0, CellData { value: "Mon".to_string(), formula: None, format: None, comment: None });
        app.set_cell_with_undo(1, 0, CellData { value: "Tue".to_string(), formula: None, format: None, comment: None });

        // Select A1:A5 (A1-A2 have values, A3-A5 are empty)
        app.selection_start = Some((0, 0));
        app.selection_end = Some((4, 0));

        // Autofill
        app.autofill_selection();

        // Check pattern continuation
        assert_eq!(app.workbook.current_sheet().get_cell(2, 0).value, "Wed");
        assert_eq!(app.workbook.current_sheet().get_cell(3, 0).value, "Thu");
        assert_eq!(app.workbook.current_sheet().get_cell(4, 0).value, "Fri");
    }

    #[test]
    fn test_autofill_pattern_prefixed() {
        let mut app = App::default();

        // Set up prefixed pattern: Item1, Item2
        app.set_cell_with_undo(0, 0, CellData { value: "Item1".to_string(), formula: None, format: None, comment: None });
        app.set_cell_with_undo(1, 0, CellData { value: "Item2".to_string(), formula: None, format: None, comment: None });

        // Select A1:A5 (A1-A2 have values, A3-A5 are empty)
        app.selection_start = Some((0, 0));
        app.selection_end = Some((4, 0));

        // Autofill
        app.autofill_selection();

        // Check pattern continuation
        assert_eq!(app.workbook.current_sheet().get_cell(2, 0).value, "Item3");
        assert_eq!(app.workbook.current_sheet().get_cell(3, 0).value, "Item4");
        assert_eq!(app.workbook.current_sheet().get_cell(4, 0).value, "Item5");
    }

    #[test]
    fn test_autofill_pattern_months_short() {
        let mut app = App::default();

        // Set up months pattern: Jan, Feb, Mar
        app.set_cell_with_undo(0, 0, CellData { value: "Jan".to_string(), formula: None, format: None, comment: None });
        app.set_cell_with_undo(1, 0, CellData { value: "Feb".to_string(), formula: None, format: None, comment: None });
        app.set_cell_with_undo(2, 0, CellData { value: "Mar".to_string(), formula: None, format: None, comment: None });

        // Select A1:A7 (A1-A3 have values, A4-A7 are empty)
        app.selection_start = Some((0, 0));
        app.selection_end = Some((6, 0));

        // Autofill
        app.autofill_selection();

        // Check pattern continuation
        assert_eq!(app.workbook.current_sheet().get_cell(3, 0).value, "Apr");
        assert_eq!(app.workbook.current_sheet().get_cell(4, 0).value, "May");
        assert_eq!(app.workbook.current_sheet().get_cell(5, 0).value, "Jun");
        assert_eq!(app.workbook.current_sheet().get_cell(6, 0).value, "Jul");
    }

    #[test]
    fn test_autofill_pattern_months_full() {
        let mut app = App::default();

        // Set up full months pattern: January, February
        app.set_cell_with_undo(0, 0, CellData { value: "January".to_string(), formula: None, format: None, comment: None });
        app.set_cell_with_undo(1, 0, CellData { value: "February".to_string(), formula: None, format: None, comment: None });

        // Select A1:A5 (A1-A2 have values, A3-A5 are empty)
        app.selection_start = Some((0, 0));
        app.selection_end = Some((4, 0));

        // Autofill
        app.autofill_selection();

        // Check pattern continuation
        assert_eq!(app.workbook.current_sheet().get_cell(2, 0).value, "March");
        assert_eq!(app.workbook.current_sheet().get_cell(3, 0).value, "April");
        assert_eq!(app.workbook.current_sheet().get_cell(4, 0).value, "May");
    }

    #[test]
    fn test_autofill_pattern_quarters() {
        let mut app = App::default();

        // Set up quarters pattern: Q1, Q2
        app.set_cell_with_undo(0, 0, CellData { value: "Q1".to_string(), formula: None, format: None, comment: None });
        app.set_cell_with_undo(1, 0, CellData { value: "Q2".to_string(), formula: None, format: None, comment: None });

        // Select A1:A6 (A1-A2 have values, A3-A6 are empty)
        app.selection_start = Some((0, 0));
        app.selection_end = Some((5, 0));

        // Autofill
        app.autofill_selection();

        // Check pattern continuation with wrap-around
        assert_eq!(app.workbook.current_sheet().get_cell(2, 0).value, "Q3");
        assert_eq!(app.workbook.current_sheet().get_cell(3, 0).value, "Q4");
        assert_eq!(app.workbook.current_sheet().get_cell(4, 0).value, "Q1"); // Wraps
        assert_eq!(app.workbook.current_sheet().get_cell(5, 0).value, "Q2"); // Wraps
    }

    #[test]
    fn test_autofill_pattern_months_wrap() {
        let mut app = App::default();

        // Set up months pattern starting near end: Oct, Nov, Dec
        app.set_cell_with_undo(0, 0, CellData { value: "Oct".to_string(), formula: None, format: None, comment: None });
        app.set_cell_with_undo(1, 0, CellData { value: "Nov".to_string(), formula: None, format: None, comment: None });
        app.set_cell_with_undo(2, 0, CellData { value: "Dec".to_string(), formula: None, format: None, comment: None });

        // Select A1:A6 (A1-A3 have values, A4-A6 are empty)
        app.selection_start = Some((0, 0));
        app.selection_end = Some((5, 0));

        // Autofill
        app.autofill_selection();

        // Check pattern continuation with wrap-around
        assert_eq!(app.workbook.current_sheet().get_cell(3, 0).value, "Jan"); // Wraps
        assert_eq!(app.workbook.current_sheet().get_cell(4, 0).value, "Feb");
        assert_eq!(app.workbook.current_sheet().get_cell(5, 0).value, "Mar");
    }

    #[test]
    fn test_autofill_horizontal_pattern() {
        let mut app = App::default();

        // Set up arithmetic pattern horizontally: 10, 20 in A1, B1
        app.set_cell_with_undo(0, 0, CellData { value: "10".to_string(), formula: None, format: None, comment: None });
        app.set_cell_with_undo(0, 1, CellData { value: "20".to_string(), formula: None, format: None, comment: None });

        // Select A1:E1 (A1-B1 have values, C1-E1 are empty) - wide selection = fill right
        app.selection_start = Some((0, 0));
        app.selection_end = Some((0, 4));

        // Autofill
        app.autofill_selection();

        // Check pattern continuation
        assert_eq!(app.workbook.current_sheet().get_cell(0, 2).value, "30");
        assert_eq!(app.workbook.current_sheet().get_cell(0, 3).value, "40");
        assert_eq!(app.workbook.current_sheet().get_cell(0, 4).value, "50");
    }

    // === Copy/Paste Tests ===

    #[test]
    fn test_copy_paste_single_cell() {
        let mut app = App::default();
        app.set_cell_with_undo(0, 0, CellData { value: "Hello".to_string(), formula: None, format: None, comment: None });

        // Copy A1
        app.selected_row = 0;
        app.selected_col = 0;
        app.copy_selection();
        assert!(app.clipboard.is_some());

        // Paste to B2
        app.selected_row = 1;
        app.selected_col = 1;
        app.paste();

        assert_eq!(app.workbook.current_sheet().get_cell(1, 1).value, "Hello");
    }

    #[test]
    fn test_copy_paste_range() {
        let mut app = App::default();
        app.set_cell_with_undo(0, 0, CellData { value: "A".to_string(), formula: None, format: None, comment: None });
        app.set_cell_with_undo(0, 1, CellData { value: "B".to_string(), formula: None, format: None, comment: None });
        app.set_cell_with_undo(1, 0, CellData { value: "C".to_string(), formula: None, format: None, comment: None });
        app.set_cell_with_undo(1, 1, CellData { value: "D".to_string(), formula: None, format: None, comment: None });

        // Select A1:B2
        app.selection_start = Some((0, 0));
        app.selection_end = Some((1, 1));
        app.copy_selection();

        // Paste to C3
        app.selected_row = 2;
        app.selected_col = 2;
        app.paste();

        assert_eq!(app.workbook.current_sheet().get_cell(2, 2).value, "A");
        assert_eq!(app.workbook.current_sheet().get_cell(2, 3).value, "B");
        assert_eq!(app.workbook.current_sheet().get_cell(3, 2).value, "C");
        assert_eq!(app.workbook.current_sheet().get_cell(3, 3).value, "D");
    }

    #[test]
    fn test_cut_paste() {
        let mut app = App::default();
        app.set_cell_with_undo(0, 0, CellData { value: "Move me".to_string(), formula: None, format: None, comment: None });

        app.selected_row = 0;
        app.selected_col = 0;
        app.cut_selection();

        // Original cell should be cleared
        assert!(app.workbook.current_sheet().get_cell(0, 0).value.is_empty());

        // Paste to new location
        app.selected_row = 2;
        app.selected_col = 2;
        app.paste();

        assert_eq!(app.workbook.current_sheet().get_cell(2, 2).value, "Move me");
    }

    #[test]
    fn test_paste_nothing() {
        let mut app = App::default();
        app.paste(); // Should not crash
        assert!(app.status_message.as_ref().unwrap().contains("Nothing to paste"));
    }

    #[test]
    fn test_copy_paste_formula_adjusts_refs() {
        let mut app = App::default();
        app.set_cell_with_undo(0, 0, CellData { value: "10".to_string(), formula: None, format: None, comment: None });
        app.set_cell_with_undo(0, 1, CellData {
            value: "20".to_string(),
            formula: Some("=A1*2".to_string()),
            format: None,
            comment: None,
        });

        // Copy B1 (has formula =A1*2)
        app.selected_row = 0;
        app.selected_col = 1;
        app.copy_selection();

        // Paste to B2 (should adjust to =A2*2)
        app.selected_row = 1;
        app.selected_col = 1;
        app.paste();

        let pasted = app.workbook.current_sheet().get_cell(1, 1);
        assert!(pasted.formula.is_some());
        assert_eq!(pasted.formula.unwrap(), "=A2*2");
    }

    // === Find and Replace Tests ===

    #[test]
    fn test_find_replace_basic() {
        let mut app = App::default();
        app.set_cell_with_undo(0, 0, CellData { value: "hello world".to_string(), formula: None, format: None, comment: None });
        app.set_cell_with_undo(1, 0, CellData { value: "hello there".to_string(), formula: None, format: None, comment: None });
        app.set_cell_with_undo(2, 0, CellData { value: "goodbye".to_string(), formula: None, format: None, comment: None });

        app.start_find_replace();
        assert!(matches!(app.mode, AppMode::FindReplace));

        app.find_replace_search = "hello".to_string();
        app.find_replace_search();

        assert_eq!(app.find_replace_results.len(), 2);
    }

    #[test]
    fn test_replace_current() {
        let mut app = App::default();
        app.set_cell_with_undo(0, 0, CellData { value: "hello world".to_string(), formula: None, format: None, comment: None });

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
        app.set_cell_with_undo(0, 0, CellData { value: "cat".to_string(), formula: None, format: None, comment: None });
        app.set_cell_with_undo(1, 0, CellData { value: "cat food".to_string(), formula: None, format: None, comment: None });
        app.set_cell_with_undo(2, 0, CellData { value: "dog".to_string(), formula: None, format: None, comment: None });

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

    // === Command Palette Tests ===

    #[test]
    fn test_command_palette_insert_row() {
        let mut app = App::default();
        app.set_cell_with_undo(0, 0, CellData { value: "A1".to_string(), formula: None, format: None, comment: None });
        app.set_cell_with_undo(1, 0, CellData { value: "A2".to_string(), formula: None, format: None, comment: None });
        let orig_rows = app.workbook.current_sheet().rows;

        app.selected_row = 1;
        app.start_command_palette();
        app.command_input = "ir".to_string();
        app.execute_command();

        assert_eq!(app.workbook.current_sheet().rows, orig_rows + 1);
        assert!(matches!(app.mode, AppMode::Normal));
    }

    #[test]
    fn test_command_palette_delete_row() {
        let mut app = App::default();
        app.set_cell_with_undo(0, 0, CellData { value: "A1".to_string(), formula: None, format: None, comment: None });
        app.set_cell_with_undo(1, 0, CellData { value: "A2".to_string(), formula: None, format: None, comment: None });
        let orig_rows = app.workbook.current_sheet().rows;

        app.selected_row = 0;
        app.start_command_palette();
        app.command_input = "dr".to_string();
        app.execute_command();

        assert_eq!(app.workbook.current_sheet().rows, orig_rows - 1);
        assert_eq!(app.workbook.current_sheet().get_cell(0, 0).value, "A2"); // Shifted up
    }

    #[test]
    fn test_command_palette_format_currency() {
        let mut app = App::default();
        app.set_cell_with_undo(0, 0, CellData { value: "1234.5".to_string(), formula: None, format: None, comment: None });

        app.selected_row = 0;
        app.selected_col = 0;
        app.start_command_palette();
        app.command_input = "format currency".to_string();
        app.execute_command();

        let cell = app.workbook.current_sheet().get_cell(0, 0);
        assert!(cell.format.is_some());
        assert!(matches!(cell.format.unwrap().number_format, NumberFormat::Currency { .. }));
    }

    #[test]
    fn test_command_palette_format_percentage() {
        let mut app = App::default();
        app.set_cell_with_undo(0, 0, CellData { value: "0.5".to_string(), formula: None, format: None, comment: None });

        app.selected_row = 0;
        app.selected_col = 0;
        app.start_command_palette();
        app.command_input = "format percent 2".to_string();
        app.execute_command();

        let cell = app.workbook.current_sheet().get_cell(0, 0);
        assert!(cell.format.is_some());
        assert!(matches!(cell.format.unwrap().number_format, NumberFormat::Percentage { decimals: 2 }));
    }

    #[test]
    fn test_command_palette_unknown_command() {
        let mut app = App::default();
        app.start_command_palette();
        app.command_input = "foobar".to_string();
        app.execute_command();

        assert!(app.status_message.as_ref().unwrap().contains("Unknown command"));
        assert!(matches!(app.mode, AppMode::Normal));
    }

    // === Sort Tests ===

    #[test]
    fn test_sort_column_ascending() {
        let mut app = App::default();
        app.set_cell_with_undo(0, 0, CellData { value: "30".to_string(), formula: None, format: None, comment: None });
        app.set_cell_with_undo(1, 0, CellData { value: "10".to_string(), formula: None, format: None, comment: None });
        app.set_cell_with_undo(2, 0, CellData { value: "20".to_string(), formula: None, format: None, comment: None });

        app.selected_col = 0;
        app.sort_column_asc();

        assert_eq!(app.workbook.current_sheet().get_cell(0, 0).value, "10");
        assert_eq!(app.workbook.current_sheet().get_cell(1, 0).value, "20");
        assert_eq!(app.workbook.current_sheet().get_cell(2, 0).value, "30");
    }

    #[test]
    fn test_sort_column_descending() {
        let mut app = App::default();
        app.set_cell_with_undo(0, 0, CellData { value: "10".to_string(), formula: None, format: None, comment: None });
        app.set_cell_with_undo(1, 0, CellData { value: "30".to_string(), formula: None, format: None, comment: None });
        app.set_cell_with_undo(2, 0, CellData { value: "20".to_string(), formula: None, format: None, comment: None });

        app.selected_col = 0;
        app.sort_column_desc();

        assert_eq!(app.workbook.current_sheet().get_cell(0, 0).value, "30");
        assert_eq!(app.workbook.current_sheet().get_cell(1, 0).value, "20");
        assert_eq!(app.workbook.current_sheet().get_cell(2, 0).value, "10");
    }

    #[test]
    fn test_sort_preserves_other_columns() {
        let mut app = App::default();
        app.set_cell_with_undo(0, 0, CellData { value: "30".to_string(), formula: None, format: None, comment: None });
        app.set_cell_with_undo(0, 1, CellData { value: "C".to_string(), formula: None, format: None, comment: None });
        app.set_cell_with_undo(1, 0, CellData { value: "10".to_string(), formula: None, format: None, comment: None });
        app.set_cell_with_undo(1, 1, CellData { value: "A".to_string(), formula: None, format: None, comment: None });
        app.set_cell_with_undo(2, 0, CellData { value: "20".to_string(), formula: None, format: None, comment: None });
        app.set_cell_with_undo(2, 1, CellData { value: "B".to_string(), formula: None, format: None, comment: None });

        app.selected_col = 0;
        app.sort_column_asc();

        // Column B should follow the sort
        assert_eq!(app.workbook.current_sheet().get_cell(0, 1).value, "A");
        assert_eq!(app.workbook.current_sheet().get_cell(1, 1).value, "B");
        assert_eq!(app.workbook.current_sheet().get_cell(2, 1).value, "C");
    }

    #[test]
    fn test_sort_undo() {
        let mut app = App::default();
        app.set_cell_with_undo(0, 0, CellData { value: "30".to_string(), formula: None, format: None, comment: None });
        app.set_cell_with_undo(1, 0, CellData { value: "10".to_string(), formula: None, format: None, comment: None });
        app.set_cell_with_undo(2, 0, CellData { value: "20".to_string(), formula: None, format: None, comment: None });

        app.selected_col = 0;
        app.sort_column_asc();

        assert_eq!(app.workbook.current_sheet().get_cell(0, 0).value, "10");

        app.undo();

        assert_eq!(app.workbook.current_sheet().get_cell(0, 0).value, "30");
        assert_eq!(app.workbook.current_sheet().get_cell(1, 0).value, "10");
        assert_eq!(app.workbook.current_sheet().get_cell(2, 0).value, "20");
    }

    // === Freeze Panes Tests ===

    #[test]
    fn test_freeze_panes() {
        let mut app = App::default();
        app.selected_row = 2;
        app.selected_col = 1;

        app.start_command_palette();
        app.command_input = "freeze".to_string();
        app.execute_command();

        assert_eq!(app.frozen_rows, 2);
        assert_eq!(app.frozen_cols, 1);
    }

    #[test]
    fn test_unfreeze_panes() {
        let mut app = App::default();
        app.frozen_rows = 3;
        app.frozen_cols = 2;

        app.start_command_palette();
        app.command_input = "unfreeze".to_string();
        app.execute_command();

        assert_eq!(app.frozen_rows, 0);
        assert_eq!(app.frozen_cols, 0);
    }

    // === Go-to Cell Tests ===

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

    // === Jump to Home/End Tests ===

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
        app.set_cell_with_undo(5, 3, CellData { value: "data".to_string(), formula: None, format: None, comment: None });
        app.set_cell_with_undo(10, 7, CellData { value: "last".to_string(), formula: None, format: None, comment: None });

        app.jump_to_end();

        assert_eq!(app.selected_row, 10);
        assert_eq!(app.selected_col, 7);
    }

    // === Selection Stats Tests ===

    #[test]
    fn test_selection_stats() {
        let mut app = App::default();
        app.set_cell_with_undo(0, 0, CellData { value: "10".to_string(), formula: None, format: None, comment: None });
        app.set_cell_with_undo(1, 0, CellData { value: "20".to_string(), formula: None, format: None, comment: None });
        app.set_cell_with_undo(2, 0, CellData { value: "30".to_string(), formula: None, format: None, comment: None });

        app.selection_start = Some((0, 0));
        app.selection_end = Some((2, 0));

        let stats = app.get_selection_stats();
        assert!(stats.is_some());
        let (sum, avg, count) = stats.unwrap();
        assert_eq!(sum, 60.0);
        assert_eq!(avg, 20.0);
        assert_eq!(count, 3);
    }

    #[test]
    fn test_selection_stats_single_cell() {
        let mut app = App::default();
        app.set_cell_with_undo(0, 0, CellData { value: "10".to_string(), formula: None, format: None, comment: None });

        app.selection_start = Some((0, 0));
        app.selection_end = Some((0, 0));

        // Single cell should return None
        assert!(app.get_selection_stats().is_none());
    }

    #[test]
    fn test_selection_stats_no_numbers() {
        let mut app = App::default();
        app.set_cell_with_undo(0, 0, CellData { value: "hello".to_string(), formula: None, format: None, comment: None });
        app.set_cell_with_undo(1, 0, CellData { value: "world".to_string(), formula: None, format: None, comment: None });

        app.selection_start = Some((0, 0));
        app.selection_end = Some((1, 0));

        // No numeric values should return None
        assert!(app.get_selection_stats().is_none());
    }

    // === Batch Undo Tests ===

    #[test]
    fn test_batch_undo() {
        let mut app = App::default();
        app.set_cell_with_undo(0, 0, CellData { value: "A".to_string(), formula: None, format: None, comment: None });
        app.set_cell_with_undo(1, 0, CellData { value: "B".to_string(), formula: None, format: None, comment: None });

        // Cut = batch undo of clearing cells
        app.selection_start = Some((0, 0));
        app.selection_end = Some((1, 0));
        app.cut_selection();

        assert!(app.workbook.current_sheet().get_cell(0, 0).value.is_empty());
        assert!(app.workbook.current_sheet().get_cell(1, 0).value.is_empty());

        // Single undo should restore both cells
        app.undo();

        assert_eq!(app.workbook.current_sheet().get_cell(0, 0).value, "A");
        assert_eq!(app.workbook.current_sheet().get_cell(1, 0).value, "B");
    }

    // === Format on Selection Tests ===

    #[test]
    fn test_set_format_on_selection() {
        let mut app = App::default();
        app.set_cell_with_undo(0, 0, CellData { value: "100".to_string(), formula: None, format: None, comment: None });
        app.set_cell_with_undo(0, 1, CellData { value: "200".to_string(), formula: None, format: None, comment: None });
        app.set_cell_with_undo(1, 0, CellData { value: "300".to_string(), formula: None, format: None, comment: None });

        app.selection_start = Some((0, 0));
        app.selection_end = Some((1, 1));

        app.set_selection_format(NumberFormat::Currency { symbol: "$".to_string(), decimals: 2 });

        for row in 0..=1 {
            for col in 0..=1 {
                let cell = app.workbook.current_sheet().get_cell(row, col);
                assert!(cell.format.is_some(), "Cell ({},{}) should have format", row, col);
            }
        }
    }

    #[test]
    fn test_set_format_general_clears_format() {
        let mut app = App::default();
        app.set_cell_with_undo(0, 0, CellData {
            value: "100".to_string(),
            formula: None,
            format: Some(CellFormat { number_format: NumberFormat::Currency { symbol: "$".to_string(), decimals: 2 }, ..CellFormat::default() }),
            comment: None,
        });

        app.selected_row = 0;
        app.selected_col = 0;
        app.set_selection_format(NumberFormat::General);

        let cell = app.workbook.current_sheet().get_cell(0, 0);
        assert!(cell.format.is_none()); // General clears format
    }

    // === Cell Styling Tests ===

    #[test]
    fn test_toggle_bold() {
        let mut app = App::default();
        app.set_cell_with_undo(0, 0, CellData { value: "test".to_string(), formula: None, format: None, comment: None });
        app.selected_row = 0;
        app.selected_col = 0;

        // Toggle bold on
        app.toggle_bold();
        let cell = app.workbook.current_sheet().get_cell(0, 0);
        assert!(cell.format.as_ref().unwrap().style.bold);

        // Toggle bold off
        app.toggle_bold();
        let cell = app.workbook.current_sheet().get_cell(0, 0);
        assert!(!cell.format.as_ref().unwrap().style.bold);
    }

    #[test]
    fn test_toggle_underline() {
        let mut app = App::default();
        app.set_cell_with_undo(0, 0, CellData { value: "test".to_string(), formula: None, format: None, comment: None });
        app.selected_row = 0;
        app.selected_col = 0;

        // Toggle underline on
        app.toggle_underline();
        let cell = app.workbook.current_sheet().get_cell(0, 0);
        assert!(cell.format.as_ref().unwrap().style.underline);

        // Toggle underline off
        app.toggle_underline();
        let cell = app.workbook.current_sheet().get_cell(0, 0);
        assert!(!cell.format.as_ref().unwrap().style.underline);
    }

    #[test]
    fn test_toggle_bold_on_selection() {
        let mut app = App::default();
        app.set_cell_with_undo(0, 0, CellData { value: "A".to_string(), formula: None, format: None, comment: None });
        app.set_cell_with_undo(0, 1, CellData { value: "B".to_string(), formula: None, format: None, comment: None });
        app.set_cell_with_undo(1, 0, CellData { value: "C".to_string(), formula: None, format: None, comment: None });

        app.selection_start = Some((0, 0));
        app.selection_end = Some((1, 1));
        app.selecting = true;

        app.toggle_bold();

        for row in 0..=1 {
            for col in 0..=1 {
                let cell = app.workbook.current_sheet().get_cell(row, col);
                assert!(cell.format.as_ref().unwrap().style.bold, "Cell ({},{}) should be bold", row, col);
            }
        }
    }

    #[test]
    fn test_set_fg_color() {
        let mut app = App::default();
        app.set_cell_with_undo(0, 0, CellData { value: "test".to_string(), formula: None, format: None, comment: None });
        app.selected_row = 0;
        app.selected_col = 0;

        app.set_selection_fg_color(Some(TerminalColor::Red));
        let cell = app.workbook.current_sheet().get_cell(0, 0);
        assert_eq!(cell.format.as_ref().unwrap().style.fg_color, Some(TerminalColor::Red));

        // Clear color
        app.set_selection_fg_color(None);
        let cell = app.workbook.current_sheet().get_cell(0, 0);
        assert_eq!(cell.format.as_ref().unwrap().style.fg_color, None);
    }

    #[test]
    fn test_set_bg_color() {
        let mut app = App::default();
        app.set_cell_with_undo(0, 0, CellData { value: "test".to_string(), formula: None, format: None, comment: None });
        app.selected_row = 0;
        app.selected_col = 0;

        app.set_selection_bg_color(Some(TerminalColor::Blue));
        let cell = app.workbook.current_sheet().get_cell(0, 0);
        assert_eq!(cell.format.as_ref().unwrap().style.bg_color, Some(TerminalColor::Blue));
    }

    #[test]
    fn test_command_bold() {
        let mut app = App::default();
        app.set_cell_with_undo(0, 0, CellData { value: "test".to_string(), formula: None, format: None, comment: None });
        app.selected_row = 0;
        app.selected_col = 0;

        app.start_command_palette();
        app.command_input = "bold".to_string();
        app.execute_command();

        let cell = app.workbook.current_sheet().get_cell(0, 0);
        assert!(cell.format.as_ref().unwrap().style.bold);
    }

    #[test]
    fn test_command_underline() {
        let mut app = App::default();
        app.set_cell_with_undo(0, 0, CellData { value: "test".to_string(), formula: None, format: None, comment: None });
        app.selected_row = 0;
        app.selected_col = 0;

        app.start_command_palette();
        app.command_input = "underline".to_string();
        app.execute_command();

        let cell = app.workbook.current_sheet().get_cell(0, 0);
        assert!(cell.format.as_ref().unwrap().style.underline);
    }

    #[test]
    fn test_command_color() {
        let mut app = App::default();
        app.set_cell_with_undo(0, 0, CellData { value: "test".to_string(), formula: None, format: None, comment: None });
        app.selected_row = 0;
        app.selected_col = 0;

        app.start_command_palette();
        app.command_input = "color red".to_string();
        app.execute_command();

        let cell = app.workbook.current_sheet().get_cell(0, 0);
        assert_eq!(cell.format.as_ref().unwrap().style.fg_color, Some(TerminalColor::Red));
    }

    #[test]
    fn test_command_bg_color() {
        let mut app = App::default();
        app.set_cell_with_undo(0, 0, CellData { value: "test".to_string(), formula: None, format: None, comment: None });
        app.selected_row = 0;
        app.selected_col = 0;

        app.start_command_palette();
        app.command_input = "bg blue".to_string();
        app.execute_command();

        let cell = app.workbook.current_sheet().get_cell(0, 0);
        assert_eq!(cell.format.as_ref().unwrap().style.bg_color, Some(TerminalColor::Blue));
    }

    #[test]
    fn test_command_color_none_clears() {
        let mut app = App::default();
        app.set_cell_with_undo(0, 0, CellData { value: "test".to_string(), formula: None, format: None, comment: None });
        app.selected_row = 0;
        app.selected_col = 0;

        // Set color first
        app.set_selection_fg_color(Some(TerminalColor::Red));
        assert_eq!(app.workbook.current_sheet().get_cell(0, 0).format.as_ref().unwrap().style.fg_color, Some(TerminalColor::Red));

        // Clear via command
        app.start_command_palette();
        app.command_input = "color none".to_string();
        app.execute_command();

        let cell = app.workbook.current_sheet().get_cell(0, 0);
        assert_eq!(cell.format.as_ref().unwrap().style.fg_color, None);
    }

    #[test]
    fn test_terminal_color_from_name() {
        assert_eq!(TerminalColor::from_name("red"), Some(TerminalColor::Red));
        assert_eq!(TerminalColor::from_name("Blue"), Some(TerminalColor::Blue));
        assert_eq!(TerminalColor::from_name("lightgreen"), Some(TerminalColor::LightGreen));
        assert_eq!(TerminalColor::from_name("CYAN"), Some(TerminalColor::Cyan));
        assert_eq!(TerminalColor::from_name("invalid"), None);
    }

    // === Multiple Sheets Tests ===

    #[test]
    fn test_default_workbook_has_one_sheet() {
        let app = App::default();
        assert_eq!(app.workbook.sheets.len(), 1);
        assert_eq!(app.workbook.sheet_names[0], "Sheet1");
        assert_eq!(app.workbook.active_sheet, 0);
    }

    #[test]
    fn test_add_sheet_command() {
        let mut app = App::default();
        app.start_command_palette();
        app.command_input = "sheet new".to_string();
        app.execute_command();

        assert_eq!(app.workbook.sheets.len(), 2);
        assert_eq!(app.workbook.active_sheet, 1); // Switched to new sheet
        assert_eq!(app.workbook.sheet_names[1], "Sheet2");
    }

    #[test]
    fn test_delete_sheet_command() {
        let mut app = App::default();
        // Add a second sheet
        app.workbook.add_sheet("Sheet2".to_string());
        app.workbook.active_sheet = 1;

        app.start_command_palette();
        app.command_input = "sheet delete".to_string();
        app.execute_command();

        assert_eq!(app.workbook.sheets.len(), 1);
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
    fn test_rename_sheet_command() {
        let mut app = App::default();
        app.start_command_palette();
        app.command_input = "rename Revenue".to_string();
        app.execute_command();

        assert_eq!(app.workbook.sheet_names[0], "Revenue");
    }

    #[test]
    fn test_switch_sheets() {
        let mut app = App::default();
        app.workbook.add_sheet("Sheet2".to_string());

        // Set data in sheet 1
        app.set_cell_with_undo(0, 0, CellData { value: "Sheet1Data".to_string(), formula: None, format: None, comment: None });

        // Switch to sheet 2
        app.switch_next_sheet();
        assert_eq!(app.workbook.active_sheet, 1);
        assert!(app.workbook.current_sheet().get_cell(0, 0).value.is_empty());

        // Set data in sheet 2
        app.set_cell_with_undo(0, 0, CellData { value: "Sheet2Data".to_string(), formula: None, format: None, comment: None });

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

    // === Phase 9: Filtering & Delight Tests ===

    #[test]
    fn test_set_cell_comment() {
        let mut app = App::default();
        app.set_cell_with_undo(0, 0, CellData { value: "Hello".to_string(), formula: None, format: None, comment: None });

        app.set_cell_comment(Some("This is a comment".to_string()));
        let cell = app.workbook.current_sheet().get_cell(0, 0);
        assert_eq!(cell.comment, Some("This is a comment".to_string()));
        assert_eq!(cell.value, "Hello"); // Value preserved
    }

    #[test]
    fn test_clear_cell_comment() {
        let mut app = App::default();
        app.set_cell_with_undo(0, 0, CellData { value: "Hello".to_string(), formula: None, format: None, comment: Some("old".to_string()) });

        app.set_cell_comment(None);
        let cell = app.workbook.current_sheet().get_cell(0, 0);
        assert_eq!(cell.comment, None);
    }

    #[test]
    fn test_comment_command() {
        let mut app = App::default();
        app.set_cell_with_undo(0, 0, CellData { value: "Data".to_string(), formula: None, format: None, comment: None });

        app.start_command_palette();
        app.command_input = "comment Test note".to_string();
        app.execute_command();

        let cell = app.workbook.current_sheet().get_cell(0, 0);
        assert_eq!(cell.comment, Some("test note".to_string()));
    }

    #[test]
    fn test_comment_clear_command() {
        let mut app = App::default();
        app.set_cell_with_undo(0, 0, CellData { value: "Data".to_string(), formula: None, format: None, comment: Some("note".to_string()) });

        app.start_command_palette();
        app.command_input = "comment clear".to_string();
        app.execute_command();

        let cell = app.workbook.current_sheet().get_cell(0, 0);
        assert_eq!(cell.comment, None);
    }

    #[test]
    fn test_apply_filter() {
        let mut app = App::default();
        app.set_cell_with_undo(0, 0, CellData { value: "Apple".to_string(), formula: None, format: None, comment: None });
        app.set_cell_with_undo(1, 0, CellData { value: "Banana".to_string(), formula: None, format: None, comment: None });
        app.set_cell_with_undo(2, 0, CellData { value: "Apple".to_string(), formula: None, format: None, comment: None });
        app.set_cell_with_undo(3, 0, CellData { value: "Cherry".to_string(), formula: None, format: None, comment: None });

        app.apply_filter(0, Some("Apple".to_string()));

        // Rows 1 and 3 should be hidden (Banana and Cherry)
        assert!(!app.hidden_rows.contains(&0));
        assert!(app.hidden_rows.contains(&1));
        assert!(!app.hidden_rows.contains(&2));
        assert!(app.hidden_rows.contains(&3));
    }

    #[test]
    fn test_clear_filter() {
        let mut app = App::default();
        app.set_cell_with_undo(0, 0, CellData { value: "Apple".to_string(), formula: None, format: None, comment: None });
        app.set_cell_with_undo(1, 0, CellData { value: "Banana".to_string(), formula: None, format: None, comment: None });

        app.apply_filter(0, Some("Apple".to_string()));
        assert!(!app.hidden_rows.is_empty());

        app.clear_filter();
        assert!(app.hidden_rows.is_empty());
        assert_eq!(app.filter_column, None);
        assert_eq!(app.filter_value, None);
    }

    #[test]
    fn test_filter_command() {
        let mut app = App::default();
        app.set_cell_with_undo(0, 0, CellData { value: "Yes".to_string(), formula: None, format: None, comment: None });
        app.set_cell_with_undo(1, 0, CellData { value: "No".to_string(), formula: None, format: None, comment: None });
        app.set_cell_with_undo(2, 0, CellData { value: "Yes".to_string(), formula: None, format: None, comment: None });

        app.start_command_palette();
        app.command_input = "filter a yes".to_string();
        app.execute_command();

        assert!(!app.hidden_rows.contains(&0));
        assert!(app.hidden_rows.contains(&1));
        assert!(!app.hidden_rows.contains(&2));
    }

    #[test]
    fn test_unfilter_command() {
        let mut app = App::default();
        app.hidden_rows.insert(1);
        app.filter_column = Some(0);

        app.start_command_palette();
        app.command_input = "unfilter".to_string();
        app.execute_command();

        assert!(app.hidden_rows.is_empty());
        assert_eq!(app.filter_column, None);
    }

    #[test]
    fn test_parse_column_label() {
        use crate::domain::Spreadsheet;
        assert_eq!(Spreadsheet::parse_column_label("A"), Some(0));
        assert_eq!(Spreadsheet::parse_column_label("B"), Some(1));
        assert_eq!(Spreadsheet::parse_column_label("Z"), Some(25));
        assert_eq!(Spreadsheet::parse_column_label("AA"), Some(26));
        assert_eq!(Spreadsheet::parse_column_label("a"), Some(0));
        assert_eq!(Spreadsheet::parse_column_label(""), None);
        assert_eq!(Spreadsheet::parse_column_label("1"), None);
    }

    #[test]
    fn test_comment_undo() {
        let mut app = App::default();
        app.set_cell_with_undo(0, 0, CellData { value: "Hello".to_string(), formula: None, format: None, comment: None });

        app.set_cell_comment(Some("My comment".to_string()));
        assert_eq!(app.workbook.current_sheet().get_cell(0, 0).comment, Some("My comment".to_string()));

        app.undo();
        assert_eq!(app.workbook.current_sheet().get_cell(0, 0).comment, None);
    }
}