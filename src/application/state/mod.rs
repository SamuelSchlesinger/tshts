//! Application state management for the terminal spreadsheet.
//!
//! This module contains the main application state and mode management
//! for the terminal user interface.

use crate::domain::{Spreadsheet, Workbook, CellData, CellFormat, NumberFormat, TerminalColor, FormulaEvaluator};
use std::collections::{HashMap, HashSet, VecDeque};

/// Performs case-insensitive string replacement, preserving the replacement text as-is.
fn case_insensitive_replace(text: &str, search: &str, replacement: &str) -> String {
    let lower_text = text.to_lowercase();
    let lower_search = search.to_lowercase();
    let mut result = String::new();
    let mut start = 0;
    while let Some(pos) = lower_text[start..].find(&lower_search) {
        result.push_str(&text[start..start + pos]);
        result.push_str(replacement);
        start += pos + search.len();
    }
    result.push_str(&text[start..]);
    result
}

/// Search matcher that honors the regex and case-sensitive flags.
/// Falls back to substring matching if the regex is invalid.
pub struct TextMatcher {
    regex: Option<regex::Regex>,
    needle: String,
    case_sensitive: bool,
}

impl TextMatcher {
    pub fn new(query: &str, use_regex: bool, case_sensitive: bool) -> Self {
        let regex = if use_regex {
            let pattern = if case_sensitive {
                query.to_string()
            } else {
                format!("(?i){}", query)
            };
            regex::Regex::new(&pattern).ok()
        } else {
            None
        };
        Self {
            regex,
            needle: if case_sensitive { query.to_string() } else { query.to_lowercase() },
            case_sensitive,
        }
    }

    pub fn is_match(&self, hay: &str) -> bool {
        if let Some(ref r) = self.regex {
            return r.is_match(hay);
        }
        if self.case_sensitive {
            hay.contains(&self.needle)
        } else {
            hay.to_lowercase().contains(&self.needle)
        }
    }

    /// Replace all matches in `hay` with `replacement`. For regex mode,
    /// captures via `$1` style are supported.
    pub fn replace_all(&self, hay: &str, replacement: &str) -> String {
        if let Some(ref r) = self.regex {
            return r.replace_all(hay, replacement).into_owned();
        }
        if self.case_sensitive {
            hay.replace(&self.needle, replacement)
        } else {
            case_insensitive_replace(hay, &self.needle, replacement)
        }
    }
}

/// The application mode determines how user input is interpreted.
#[derive(Debug)]
pub enum AppMode {
    /// Normal navigation mode - arrow keys move selection, shortcuts available.
    /// Vim's "normal mode": single-letter keys trigger commands; no typing
    /// directly into cells.
    Normal,
    /// Cell editing mode - user is typing into a cell. Vim's "insert mode".
    Editing,
    /// Vim-style visual mode: a range selection that motion keys extend.
    /// `kind` determines whether the selection is per-cell, whole-row, or
    /// rectangular block.
    Visual { kind: VisualKind },
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
    /// Command palette mode (vim's `:` command-line mode)
    CommandPalette,
    /// Confirmation prompt before a destructive action (quit/load/import).
    ConfirmDiscard,
}

/// The granularity of a vim-style visual selection.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum VisualKind {
    /// `v` — character/cell-granularity selection.
    Cell,
    /// `V` — whole-row selection.
    Row,
    /// `Ctrl-V` — rectangular block selection.
    Block,
}

/// A vim operator awaiting a motion (or a repeat of its own key for a line-op).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum VimOperator {
    /// `d` — delete (cut). Operates with undo.
    Delete,
    /// `y` — yank (copy).
    Yank,
    /// `c` — change: delete and enter insert mode.
    Change,
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
    /// True if this clipboard was produced by a vim line-operator (`yy`,
    /// `dd`, `cc`, or a count thereof). Drives `p`/`P` placement: row-shaped
    /// clipboards paste one row below/above, cell-shaped paste one col
    /// right/left.
    pub is_row_op: bool,
}

/// Main application state containing the spreadsheet and UI state.
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
    /// Active help search query (when user has pressed `/` inside the help popup)
    pub help_search: String,
    /// True when user is typing into help_search (vs scrolling).
    pub help_search_active: bool,
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
    /// Hidden columns. `:hide col E` adds; rendered cells skip these.
    pub hidden_cols: HashSet<usize>,
    /// Active filter column and criteria
    pub filter_column: Option<usize>,
    pub filter_value: Option<String>,
    /// Search/find-replace options: regex match instead of substring.
    pub search_regex: bool,
    /// Search/find-replace options: case sensitive (default false).
    pub search_case_sensitive: bool,
    /// Whether the workbook has unsaved changes since last save/load.
    pub dirty: bool,
    /// A pending destructive action awaiting user confirmation (quit/load/import).
    /// While Some, the app is in AppMode::ConfirmDiscard.
    pub pending_action: Option<PendingAction>,
    /// Set to true to signal the main loop to exit.
    pub should_quit: bool,
    /// Cached selection-stats result + the (start, end) range it was computed
    /// for. Invalidated when the selection changes or a cell mutates.
    pub stats_cache: Option<((usize, usize), (usize, usize), Option<(f64, f64, usize)>)>,
    /// Rendered column-X ranges from the last frame, used by mouse hit-test
    /// to map x → column. `(col, x_start_inclusive, x_end_exclusive)`.
    pub last_col_rects: Vec<(usize, u16, u16)>,
    /// First data-row's terminal Y (after grid border + col-letters row).
    pub last_grid_top_y: u16,
    /// Active chart popup, if any. Cleared by Esc.
    pub chart_popup: Option<ChartPopup>,
    /// Iterative-calculation enabled (for circular refs).
    pub iterative_calc: bool,
    /// A1 (default) vs R1C1 reference mode. Affects how cell refs render in
    /// the formula bar and headers.
    pub r1c1_mode: bool,
    /// Per-column data-validation rules: maps col → predicate formula
    /// (with `_` bound to the input). Violators are flagged with `!`.
    pub validations: HashMap<usize, String>,
    /// Vim count prefix being accumulated (e.g. `5` then `j` → move down 5).
    /// `None` means "no count typed yet"; `Some(0)` never appears since `0`
    /// is the start-of-row motion.
    pub vim_count: Option<usize>,
    /// Vim operator awaiting a motion: `d`, `y`, or `c`. Cleared after the
    /// motion fires, or on Esc.
    pub vim_pending_op: Option<VimOperator>,
    /// True after a single `g` press, waiting for the second `g` to make `gg`.
    /// Cleared on any other key.
    pub vim_awaiting_g: bool,
}

/// A popup chart over a range. `kind` controls the shape. We store the
/// source range so the chart re-fetches values at render time; editing a
/// cell in the range updates the chart on the next frame.
#[derive(Debug, Clone)]
pub struct ChartPopup {
    pub title: String,
    pub source: ((usize, usize), (usize, usize)),
    pub kind: ChartKind,
}

#[derive(Debug, Clone, Copy)]
pub enum ChartKind {
    Bar,
    Line,
    Sparkline,
}

/// Destructive action awaiting confirmation when there are unsaved changes.
#[derive(Debug, Clone)]
pub enum PendingAction {
    Quit,
    LoadFile,
}

/// Where the cursor should land after a cell edit completes.
#[derive(Debug, Clone, Copy)]
enum EditExitDir {
    Down,
    Right,
}

/// Parse a literal cell-range like `A1:B10` into ((row, col), (row, col)).
fn parse_range(s: &str) -> Option<((usize, usize), (usize, usize))> {
    let mut parts = s.split(':');
    let start = parts.next()?;
    let end = parts.next()?;
    if parts.next().is_some() {
        return None;
    }
    let s = crate::domain::Spreadsheet::parse_cell_reference(start)?;
    let e = crate::domain::Spreadsheet::parse_cell_reference(end)?;
    Some((s, e))
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
            help_search: String::new(),
            help_search_active: false,
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
            hidden_cols: HashSet::new(),
            filter_column: None,
            filter_value: None,
            search_regex: false,
            search_case_sensitive: false,
            dirty: false,
            pending_action: None,
            should_quit: false,
            stats_cache: None,
            last_col_rects: Vec::new(),
            last_grid_top_y: 0,
            chart_popup: None,
            iterative_calc: false,
            r1c1_mode: false,
            validations: HashMap::new(),
            vim_count: None,
            vim_pending_op: None,
            vim_awaiting_g: false,
        }
    }
}


mod editing;
mod clipboard;
mod autofill;
mod command;
mod formatting;
mod io;
mod search;
mod vim;
mod lifecycle;
mod navigation;

impl App {
    pub fn dismiss_transients(&mut self) {
        self.clear_selection();
        self.search_results.clear();
        self.search_result_index = 0;
        self.status_message = None;
        self.chart_popup = None;
        if self.filter_column.is_some() || !self.hidden_rows.is_empty() {
            self.clear_filter();
        }
    }

    pub fn recalc_all(&mut self) {
        // Clone the workbook so we can give the evaluator a read-only view
        // while mutating the originals. Cheap-ish: cells are Arc-free clones
        // of strings + small metadata.
        let wb_snapshot = self.workbook.clone();
        let names = wb_snapshot.named_ranges.clone();
        for (idx, sheet) in self.workbook.sheets.iter_mut().enumerate() {
            sheet.rebuild_dependencies();
            let cells: Vec<(usize, usize, String)> = sheet
                .cells
                .iter()
                .filter_map(|(&(r, c), cd)| cd.formula.as_ref().map(|f| (r, c, f.clone())))
                .collect();
            let snap_sheet = &wb_snapshot.sheets[idx];
            for (row, col, formula) in cells {
                let evaluator =
                    FormulaEvaluator::with_workbook(&wb_snapshot, snap_sheet, &names);
                let value = evaluator.evaluate_formula(&formula);
                let mut cd = sheet.get_cell(row, col);
                cd.value = value;
                sheet.cells.insert((row, col), cd);
            }
        }
        self.status_message = Some("Recalculated all formulas".to_string());
    }

    fn record_action(&mut self, action: UndoAction) {
        const MAX_UNDO_STACK_SIZE: usize = 1000;
        self.undo_stack.push_back(action);
        if self.undo_stack.len() > MAX_UNDO_STACK_SIZE {
            self.undo_stack.pop_front();
        }
        self.redo_stack.clear();
        self.dirty = true;
        self.invalidate_stats_cache();
        crate::infrastructure::autosave::mark_dirty();
    }

    pub fn undo(&mut self) {
        if let Some(action) = self.undo_stack.pop_back() {
            self.apply_undo(&action);
            self.redo_stack.push_back(action);
            self.dirty = true;
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

    pub fn redo(&mut self) {
        if let Some(action) = self.redo_stack.pop_back() {
            self.apply_redo(&action);
            self.undo_stack.push_back(action);
            self.dirty = true;
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

    fn maybe_extend_table(&mut self, row: usize, col: usize) {
        let sheet_idx = self.workbook.active_sheet;
        let table_idx = {
            let sheet = self.workbook.current_sheet();
            sheet
                .tables
                .iter()
                .position(|t| {
                    row == t.bottom_row + 1
                        && col >= t.left_col
                        && col <= t.right_col
                })
        };
        let Some(ti) = table_idx else { return; };
        let table_clone = self.workbook.sheets[sheet_idx].tables[ti].clone();
        self.workbook.sheets[sheet_idx].tables[ti].bottom_row += 1;
        // Re-register each column's named range with the new bounds.
        let new_bottom = self.workbook.sheets[sheet_idx].tables[ti].bottom_row;
        let body_top = table_clone.top_row + 1;
        for (i, header) in table_clone.headers.iter().enumerate() {
            let key = format!(
                "{}[{}]",
                table_clone.name.to_uppercase(),
                header.to_uppercase()
            );
            let value = format!(
                "{}:{}",
                crate::domain::Spreadsheet::format_cell_reference(
                    body_top,
                    table_clone.left_col + i,
                    false,
                    false
                ),
                crate::domain::Spreadsheet::format_cell_reference(
                    new_bottom,
                    table_clone.left_col + i,
                    false,
                    false
                ),
            );
            self.workbook.named_ranges.insert(key.clone(), value.clone());
            for s in &mut self.workbook.sheets {
                s.named_ranges.insert(key.clone(), value.clone());
            }
        }
    }

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

        // Apply the change, then auto-extend any table that this row sits below.
        self.workbook.current_sheet_mut().set_cell(row, col, new_data);
        self.maybe_extend_table(row, col);

        // Workbook-level cross-sheet dep maintenance: register this cell's
        // (new) dependencies, then propagate changes to anything that
        // depended on its old value.
        let sheet_name = self.workbook.sheet_names[self.workbook.active_sheet].clone();
        self.workbook.register_cross_sheet_deps(&sheet_name, row, col);
        self.workbook.propagate_cross_sheet_changes(&sheet_name, row, col);
    }

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

        // Cross-sheet maintenance.
        let sheet_name = self.workbook.sheet_names[self.workbook.active_sheet].clone();
        self.workbook.register_cross_sheet_deps(&sheet_name, row, col);
        self.workbook.propagate_cross_sheet_changes(&sheet_name, row, col);
    }

    pub fn start_selection(&mut self) {
        self.selection_start = Some((self.selected_row, self.selected_col));
        self.selection_end = Some((self.selected_row, self.selected_col));
        self.selecting = true;
        self.stats_cache = None;
    }

    pub fn update_selection(&mut self, row: usize, col: usize) {
        if self.selecting {
            self.selection_end = Some((row, col));
            self.stats_cache = None;
        }
    }

    pub fn clear_selection(&mut self) {
        self.selection_start = None;
        self.selection_end = None;
        self.selecting = false;
        self.stats_cache = None;
    }

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

    pub fn is_cell_selected(&self, row: usize, col: usize) -> bool {
        if let Some(((min_row, min_col), (max_row, max_col))) = self.get_selection_range() {
            row >= min_row && row <= max_row && col >= min_col && col <= max_col
        } else {
            false
        }
    }

    pub fn update_viewport_size(&mut self, rows: usize, cols: usize) {
        self.viewport_rows = rows;
        self.viewport_cols = cols;
    }

    pub fn ensure_cursor_visible(&mut self) {
        // Frozen rows occupy viewport space at the top; subtract them from
        // the area available for scrolled rows.
        let usable_rows = self.viewport_rows.saturating_sub(self.frozen_rows).max(1);

        // If the cursor is in a frozen row, no scrolling needed — it's always shown.
        if self.selected_row >= self.frozen_rows {
            let body_top = self.scroll_row.max(self.frozen_rows);
            if self.selected_row < body_top {
                self.scroll_row = self.selected_row;
            } else if self.selected_row >= body_top + usable_rows {
                self.scroll_row = (self.selected_row + 1).saturating_sub(usable_rows);
                if self.scroll_row < self.frozen_rows {
                    self.scroll_row = self.frozen_rows;
                }
            }
        }

        // Horizontal scrolling — symmetric with rows.
        let usable_cols = self.viewport_cols.saturating_sub(self.frozen_cols).max(1);
        if self.selected_col >= self.frozen_cols {
            let body_left = self.scroll_col.max(self.frozen_cols);
            if self.selected_col < body_left {
                self.scroll_col = self.selected_col;
            } else if self.selected_col >= body_left + usable_cols {
                self.scroll_col = (self.selected_col + 1).saturating_sub(usable_cols);
                if self.scroll_col < self.frozen_cols {
                    self.scroll_col = self.frozen_cols;
                }
            }
        }
    }

    pub fn get_selection_stats(&mut self) -> Option<(f64, f64, usize)> {
        let (start, end) = self.get_selection_range()?;
        if start == end {
            return None; // Single cell, no stats
        }
        if let Some((cs, ce, v)) = &self.stats_cache {
            if *cs == start && *ce == end {
                return *v;
            }
        }
        let ((start_row, start_col), (end_row, end_col)) = (start, end);
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
        let result = if count > 0 {
            Some((sum, sum / count as f64, count))
        } else {
            None
        };
        self.stats_cache = Some((start, end, result));
        result
    }

    fn invalidate_stats_cache(&mut self) {
        self.stats_cache = None;
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
    fn test_confirm_discard_save_then_quit_with_known_filename() {
        use tempfile::NamedTempFile;
        let tmp = NamedTempFile::new().unwrap();
        let path = tmp.path().to_str().unwrap().to_string();

        let mut app = App::default();
        app.workbook.current_sheet_mut().set_cell(0, 0, crate::domain::CellData {
            value: "x".to_string(), formula: None, format: None, comment: None,
        spill_anchor: None,
        });
        app.dirty = true;
        app.filename = Some(path);

        // Quit with dirty → prompt.
        app.request_quit();
        assert!(matches!(app.mode, AppMode::ConfirmDiscard));
        assert!(!app.should_quit);

        // Simulate "s" (save & quit). save_in_place_or_prompt should succeed
        // because filename is known.
        let pending = app.pending_action.take();
        app.save_in_place_or_prompt();
        assert!(!app.dirty);
        // Trigger the deferred quit.
        if let Some(action) = pending {
            app.pending_action = Some(action);
            app.confirm_pending_action();
        }
        assert!(app.should_quit);
    }

    #[test]
    fn test_cross_sheet_auto_recalc() {
        // Sheet2!A1 = Sheet1!A1 + 10. Editing Sheet1!A1 should auto-update
        // Sheet2!A1 without a manual F5.
        let mut app = App::default();
        app.workbook.add_sheet("Sheet2".to_string());
        // Sheet1!A1 = 5
        app.workbook.active_sheet = 0;
        app.set_cell_with_undo(0, 0, CellData {
            value: "5".to_string(),
            formula: None,
            format: None,
            comment: None,
        spill_anchor: None,
        });
        // Sheet2!A1 = =Sheet1!A1 + 10  (evaluates to 15)
        app.workbook.active_sheet = 1;
        app.set_cell_with_undo(0, 0, CellData {
            value: "15".to_string(),
            formula: Some("=Sheet1!A1 + 10".to_string()),
            format: None,
            comment: None,
        spill_anchor: None,
        });
        assert_eq!(app.workbook.sheets[1].get_cell(0, 0).value, "15");

        // Now change Sheet1!A1 to 20. Sheet2!A1 should auto-update to 30.
        app.workbook.active_sheet = 0;
        app.set_cell_with_undo(0, 0, CellData {
            value: "20".to_string(),
            formula: None,
            format: None,
            comment: None,
        spill_anchor: None,
        });
        assert_eq!(app.workbook.sheets[1].get_cell(0, 0).value, "30");
    }

    #[test]
    fn test_cross_sheet_cycle_rejected() {
        let mut app = App::default();
        app.workbook.add_sheet("Sheet2".to_string());
        // Sheet1!A1 = =Sheet2!A1
        app.workbook.active_sheet = 0;
        app.start_editing();
        app.input = "=Sheet2!A1".to_string();
        app.cursor_position = app.input.chars().count();
        app.finish_editing();
        // Sheet2!A1 = =Sheet1!A1 — should be rejected (cross-sheet cycle).
        app.workbook.active_sheet = 1;
        app.start_editing();
        app.input = "=Sheet1!A1".to_string();
        app.cursor_position = app.input.chars().count();
        app.finish_editing();
        // The reject path returns early without writing the formula, so
        // Sheet2!A1 stays empty/uninitialized.
        let cell = app.workbook.sheets[1].get_cell(0, 0);
        assert!(cell.formula.is_none(), "expected cross-sheet cycle rejected, got formula={:?}", cell.formula);
    }

    #[test]
    fn test_cross_sheet_chain_propagates() {
        // Three-link chain: Sheet1!A1 → Sheet2!A1 → Sheet3!A1.
        let mut app = App::default();
        app.workbook.add_sheet("Sheet2".to_string());
        app.workbook.add_sheet("Sheet3".to_string());

        app.workbook.active_sheet = 0;
        app.set_cell_with_undo(0, 0, CellData {
            value: "1".to_string(),
            formula: None,
            format: None,
            comment: None,
        spill_anchor: None,
        });
        app.workbook.active_sheet = 1;
        app.set_cell_with_undo(0, 0, CellData {
            value: "2".to_string(),
            formula: Some("=Sheet1!A1 + 1".to_string()),
            format: None,
            comment: None,
        spill_anchor: None,
        });
        app.workbook.active_sheet = 2;
        app.set_cell_with_undo(0, 0, CellData {
            value: "3".to_string(),
            formula: Some("=Sheet2!A1 + 1".to_string()),
            format: None,
            comment: None,
        spill_anchor: None,
        });
        assert_eq!(app.workbook.sheets[2].get_cell(0, 0).value, "3");

        // Bump the head of the chain.
        app.workbook.active_sheet = 0;
        app.set_cell_with_undo(0, 0, CellData {
            value: "10".to_string(),
            formula: None,
            format: None,
            comment: None,
        spill_anchor: None,
        });
        assert_eq!(app.workbook.sheets[1].get_cell(0, 0).value, "11");
        assert_eq!(app.workbook.sheets[2].get_cell(0, 0).value, "12");
    }

    #[test]
    fn test_smoke_end_to_end_flow() {
        // High-level sanity check exercising the key flows wired by the
        // recent refactors. Does not touch the terminal; only the App API.
        let mut app = App::default();
        assert!(!app.dirty);
        assert!(!app.should_quit);

        // Start an edit and commit a value via the normal Editing flow.
        app.start_editing();
        app.input = "12".to_string();
        app.cursor_position = 2;
        app.finish_editing();
        assert_eq!(app.workbook.current_sheet().get_cell(0, 0).value, "12");
        assert!(app.dirty);

        // A formula with an absolute reference round-trips through autofill.
        app.selected_row = 0;
        app.selected_col = 1;
        app.start_editing();
        app.input = "=A1*$B$5".to_string();
        app.cursor_position = app.input.chars().count();
        app.finish_editing();
        let evaluator = crate::domain::FormulaEvaluator::new(app.workbook.current_sheet());
        // $-anchored part survives an autofill row shift.
        let shifted = evaluator.adjust_formula_references("=A1*$B$5", 1, 0);
        assert_eq!(shifted, "=A2*$B$5");

        // Dirty-aware quit prompts.
        app.dirty = true;
        app.request_quit();
        assert!(matches!(app.mode, AppMode::ConfirmDiscard));
        assert!(!app.should_quit);
        app.confirm_pending_action();
        assert!(app.should_quit);

        // Esc-dismiss clears transient state.
        let mut app2 = App::default();
        app2.search_results.push((1, 1));
        app2.status_message = Some("noise".to_string());
        app2.dismiss_transients();
        assert!(app2.search_results.is_empty());
        assert!(app2.status_message.is_none());

        // recalc_all is callable and idempotent.
        app2.recalc_all();
        assert_eq!(
            app2.status_message.as_deref(),
            Some("Recalculated all formulas")
        );
    }

    #[test]
    fn test_vlookup_basic() {
        use crate::domain::FormulaEvaluator;
        let mut sheet = crate::domain::Spreadsheet::default();
        // Single-column lookup
        for (i, v) in ["a", "b", "c", "d"].iter().enumerate() {
            sheet.set_cell(i, 0, crate::domain::CellData {
                value: v.to_string(), formula: None, format: None, comment: None,
            spill_anchor: None,
            });
        }
        let evaluator = FormulaEvaluator::new(&sheet);
        assert_eq!(
            evaluator.evaluate_formula("=VLOOKUP(\"c\", A1:A4, 1, 0)"),
            "c"
        );
    }

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
        spill_anchor: None,
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
        spill_anchor: None,
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
        spill_anchor: None,
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
        app.set_cell_with_undo(0, 1, CellData { value: "10".to_string(), formula: None, format: None, comment: None, spill_anchor: None }); // B1 = 10
        app.set_cell_with_undo(1, 1, CellData { value: "20".to_string(), formula: None, format: None, comment: None, spill_anchor: None }); // B2 = 20
        app.set_cell_with_undo(0, 2, CellData { value: "30".to_string(), formula: None, format: None, comment: None, spill_anchor: None }); // C1 = 30
        app.set_cell_with_undo(1, 2, CellData { value: "40".to_string(), formula: None, format: None, comment: None, spill_anchor: None }); // C2 = 40

        // Set up a formula in A1 that references B1:B2
        app.set_cell_with_undo(0, 0, CellData {
            value: "30".to_string(),
            formula: Some("=SUM(B1:B2)".to_string()),
            format: None,
            comment: None,
        spill_anchor: None,
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
        app.set_cell_with_undo(1, 0, CellData { value: "10".to_string(), formula: None, format: None, comment: None, spill_anchor: None }); // A2 = 10
        app.set_cell_with_undo(1, 1, CellData { value: "20".to_string(), formula: None, format: None, comment: None, spill_anchor: None }); // B2 = 20
        app.set_cell_with_undo(2, 0, CellData { value: "30".to_string(), formula: None, format: None, comment: None, spill_anchor: None }); // A3 = 30
        app.set_cell_with_undo(2, 1, CellData { value: "40".to_string(), formula: None, format: None, comment: None, spill_anchor: None }); // B3 = 40

        // Set up a formula in A1 that references A2+B2
        app.set_cell_with_undo(0, 0, CellData {
            value: "30".to_string(),
            formula: Some("=A2+B2".to_string()),
            format: None,
            comment: None,
        spill_anchor: None,
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
        app.set_cell_with_undo(0, 0, CellData { value: "1".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        app.set_cell_with_undo(1, 0, CellData { value: "2".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        app.set_cell_with_undo(2, 0, CellData { value: "3".to_string(), formula: None, format: None, comment: None, spill_anchor: None });

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
        app.set_cell_with_undo(0, 0, CellData { value: "Mon".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        app.set_cell_with_undo(1, 0, CellData { value: "Tue".to_string(), formula: None, format: None, comment: None, spill_anchor: None });

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
        app.set_cell_with_undo(0, 0, CellData { value: "Item1".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        app.set_cell_with_undo(1, 0, CellData { value: "Item2".to_string(), formula: None, format: None, comment: None, spill_anchor: None });

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
        app.set_cell_with_undo(0, 0, CellData { value: "Jan".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        app.set_cell_with_undo(1, 0, CellData { value: "Feb".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        app.set_cell_with_undo(2, 0, CellData { value: "Mar".to_string(), formula: None, format: None, comment: None, spill_anchor: None });

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
        app.set_cell_with_undo(0, 0, CellData { value: "January".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        app.set_cell_with_undo(1, 0, CellData { value: "February".to_string(), formula: None, format: None, comment: None, spill_anchor: None });

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
        app.set_cell_with_undo(0, 0, CellData { value: "Q1".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        app.set_cell_with_undo(1, 0, CellData { value: "Q2".to_string(), formula: None, format: None, comment: None, spill_anchor: None });

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
        app.set_cell_with_undo(0, 0, CellData { value: "Oct".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        app.set_cell_with_undo(1, 0, CellData { value: "Nov".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        app.set_cell_with_undo(2, 0, CellData { value: "Dec".to_string(), formula: None, format: None, comment: None, spill_anchor: None });

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
        app.set_cell_with_undo(0, 0, CellData { value: "10".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        app.set_cell_with_undo(0, 1, CellData { value: "20".to_string(), formula: None, format: None, comment: None, spill_anchor: None });

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
        app.set_cell_with_undo(0, 0, CellData { value: "Hello".to_string(), formula: None, format: None, comment: None, spill_anchor: None });

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
        app.set_cell_with_undo(0, 0, CellData { value: "A".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        app.set_cell_with_undo(0, 1, CellData { value: "B".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        app.set_cell_with_undo(1, 0, CellData { value: "C".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        app.set_cell_with_undo(1, 1, CellData { value: "D".to_string(), formula: None, format: None, comment: None, spill_anchor: None });

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
        app.set_cell_with_undo(0, 0, CellData { value: "Move me".to_string(), formula: None, format: None, comment: None, spill_anchor: None });

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
        app.set_cell_with_undo(0, 0, CellData { value: "10".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        app.set_cell_with_undo(0, 1, CellData {
            value: "20".to_string(),
            formula: Some("=A1*2".to_string()),
            format: None,
            comment: None,
        spill_anchor: None,
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

    // === Command Palette Tests ===

    #[test]
    fn test_command_palette_insert_row() {
        let mut app = App::default();
        app.set_cell_with_undo(0, 0, CellData { value: "A1".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        app.set_cell_with_undo(1, 0, CellData { value: "A2".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
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
        app.set_cell_with_undo(0, 0, CellData { value: "A1".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        app.set_cell_with_undo(1, 0, CellData { value: "A2".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
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
        app.set_cell_with_undo(0, 0, CellData { value: "1234.5".to_string(), formula: None, format: None, comment: None, spill_anchor: None });

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
        app.set_cell_with_undo(0, 0, CellData { value: "0.5".to_string(), formula: None, format: None, comment: None, spill_anchor: None });

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
        app.set_cell_with_undo(0, 0, CellData { value: "30".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        app.set_cell_with_undo(1, 0, CellData { value: "10".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        app.set_cell_with_undo(2, 0, CellData { value: "20".to_string(), formula: None, format: None, comment: None, spill_anchor: None });

        app.selected_col = 0;
        app.sort_column_asc();

        assert_eq!(app.workbook.current_sheet().get_cell(0, 0).value, "10");
        assert_eq!(app.workbook.current_sheet().get_cell(1, 0).value, "20");
        assert_eq!(app.workbook.current_sheet().get_cell(2, 0).value, "30");
    }

    #[test]
    fn test_sort_column_descending() {
        let mut app = App::default();
        app.set_cell_with_undo(0, 0, CellData { value: "10".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        app.set_cell_with_undo(1, 0, CellData { value: "30".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        app.set_cell_with_undo(2, 0, CellData { value: "20".to_string(), formula: None, format: None, comment: None, spill_anchor: None });

        app.selected_col = 0;
        app.sort_column_desc();

        assert_eq!(app.workbook.current_sheet().get_cell(0, 0).value, "30");
        assert_eq!(app.workbook.current_sheet().get_cell(1, 0).value, "20");
        assert_eq!(app.workbook.current_sheet().get_cell(2, 0).value, "10");
    }

    #[test]
    fn test_sort_preserves_other_columns() {
        let mut app = App::default();
        app.set_cell_with_undo(0, 0, CellData { value: "30".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        app.set_cell_with_undo(0, 1, CellData { value: "C".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        app.set_cell_with_undo(1, 0, CellData { value: "10".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        app.set_cell_with_undo(1, 1, CellData { value: "A".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        app.set_cell_with_undo(2, 0, CellData { value: "20".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        app.set_cell_with_undo(2, 1, CellData { value: "B".to_string(), formula: None, format: None, comment: None, spill_anchor: None });

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
        app.set_cell_with_undo(0, 0, CellData { value: "30".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        app.set_cell_with_undo(1, 0, CellData { value: "10".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        app.set_cell_with_undo(2, 0, CellData { value: "20".to_string(), formula: None, format: None, comment: None, spill_anchor: None });

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
        app.set_cell_with_undo(5, 3, CellData { value: "data".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        app.set_cell_with_undo(10, 7, CellData { value: "last".to_string(), formula: None, format: None, comment: None, spill_anchor: None });

        app.jump_to_end();

        assert_eq!(app.selected_row, 10);
        assert_eq!(app.selected_col, 7);
    }

    // === Selection Stats Tests ===

    #[test]
    fn test_selection_stats() {
        let mut app = App::default();
        app.set_cell_with_undo(0, 0, CellData { value: "10".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        app.set_cell_with_undo(1, 0, CellData { value: "20".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        app.set_cell_with_undo(2, 0, CellData { value: "30".to_string(), formula: None, format: None, comment: None, spill_anchor: None });

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
        app.set_cell_with_undo(0, 0, CellData { value: "10".to_string(), formula: None, format: None, comment: None, spill_anchor: None });

        app.selection_start = Some((0, 0));
        app.selection_end = Some((0, 0));

        // Single cell should return None
        assert!(app.get_selection_stats().is_none());
    }

    #[test]
    fn test_selection_stats_no_numbers() {
        let mut app = App::default();
        app.set_cell_with_undo(0, 0, CellData { value: "hello".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        app.set_cell_with_undo(1, 0, CellData { value: "world".to_string(), formula: None, format: None, comment: None, spill_anchor: None });

        app.selection_start = Some((0, 0));
        app.selection_end = Some((1, 0));

        // No numeric values should return None
        assert!(app.get_selection_stats().is_none());
    }

    // === Batch Undo Tests ===

    #[test]
    fn test_batch_undo() {
        let mut app = App::default();
        app.set_cell_with_undo(0, 0, CellData { value: "A".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        app.set_cell_with_undo(1, 0, CellData { value: "B".to_string(), formula: None, format: None, comment: None, spill_anchor: None });

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
        app.set_cell_with_undo(0, 0, CellData { value: "100".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        app.set_cell_with_undo(0, 1, CellData { value: "200".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        app.set_cell_with_undo(1, 0, CellData { value: "300".to_string(), formula: None, format: None, comment: None, spill_anchor: None });

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
        spill_anchor: None,
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
        app.set_cell_with_undo(0, 0, CellData { value: "test".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
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
        app.set_cell_with_undo(0, 0, CellData { value: "test".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
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
        app.set_cell_with_undo(0, 0, CellData { value: "A".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        app.set_cell_with_undo(0, 1, CellData { value: "B".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        app.set_cell_with_undo(1, 0, CellData { value: "C".to_string(), formula: None, format: None, comment: None, spill_anchor: None });

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
        app.set_cell_with_undo(0, 0, CellData { value: "test".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
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
        app.set_cell_with_undo(0, 0, CellData { value: "test".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        app.selected_row = 0;
        app.selected_col = 0;

        app.set_selection_bg_color(Some(TerminalColor::Blue));
        let cell = app.workbook.current_sheet().get_cell(0, 0);
        assert_eq!(cell.format.as_ref().unwrap().style.bg_color, Some(TerminalColor::Blue));
    }

    #[test]
    fn test_command_bold() {
        let mut app = App::default();
        app.set_cell_with_undo(0, 0, CellData { value: "test".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
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
        app.set_cell_with_undo(0, 0, CellData { value: "test".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
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
        app.set_cell_with_undo(0, 0, CellData { value: "test".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
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
        app.set_cell_with_undo(0, 0, CellData { value: "test".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
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
        app.set_cell_with_undo(0, 0, CellData { value: "test".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
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

    // === Phase 9: Filtering & Delight Tests ===

    #[test]
    fn test_set_cell_comment() {
        let mut app = App::default();
        app.set_cell_with_undo(0, 0, CellData { value: "Hello".to_string(), formula: None, format: None, comment: None, spill_anchor: None });

        app.set_cell_comment(Some("This is a comment".to_string()));
        let cell = app.workbook.current_sheet().get_cell(0, 0);
        assert_eq!(cell.comment, Some("This is a comment".to_string()));
        assert_eq!(cell.value, "Hello"); // Value preserved
    }

    #[test]
    fn test_clear_cell_comment() {
        let mut app = App::default();
        app.set_cell_with_undo(0, 0, CellData { value: "Hello".to_string(), formula: None, format: None, comment: Some("old".to_string()), spill_anchor: None });

        app.set_cell_comment(None);
        let cell = app.workbook.current_sheet().get_cell(0, 0);
        assert_eq!(cell.comment, None);
    }

    #[test]
    fn test_comment_command() {
        let mut app = App::default();
        app.set_cell_with_undo(0, 0, CellData { value: "Data".to_string(), formula: None, format: None, comment: None, spill_anchor: None });

        app.start_command_palette();
        app.command_input = "comment Test note".to_string();
        app.execute_command();

        // Comment text preserves case (it's user-facing prose).
        let cell = app.workbook.current_sheet().get_cell(0, 0);
        assert_eq!(cell.comment, Some("Test note".to_string()));
    }

    #[test]
    fn test_comment_clear_command() {
        let mut app = App::default();
        app.set_cell_with_undo(0, 0, CellData { value: "Data".to_string(), formula: None, format: None, comment: Some("note".to_string()), spill_anchor: None });

        app.start_command_palette();
        app.command_input = "comment clear".to_string();
        app.execute_command();

        let cell = app.workbook.current_sheet().get_cell(0, 0);
        assert_eq!(cell.comment, None);
    }

    #[test]
    fn test_apply_filter() {
        let mut app = App::default();
        app.set_cell_with_undo(0, 0, CellData { value: "Apple".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        app.set_cell_with_undo(1, 0, CellData { value: "Banana".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        app.set_cell_with_undo(2, 0, CellData { value: "Apple".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        app.set_cell_with_undo(3, 0, CellData { value: "Cherry".to_string(), formula: None, format: None, comment: None, spill_anchor: None });

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
        app.set_cell_with_undo(0, 0, CellData { value: "Apple".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        app.set_cell_with_undo(1, 0, CellData { value: "Banana".to_string(), formula: None, format: None, comment: None, spill_anchor: None });

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
        app.set_cell_with_undo(0, 0, CellData { value: "Yes".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        app.set_cell_with_undo(1, 0, CellData { value: "No".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        app.set_cell_with_undo(2, 0, CellData { value: "Yes".to_string(), formula: None, format: None, comment: None, spill_anchor: None });

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
        app.set_cell_with_undo(0, 0, CellData { value: "Hello".to_string(), formula: None, format: None, comment: None, spill_anchor: None });

        app.set_cell_comment(Some("My comment".to_string()));
        assert_eq!(app.workbook.current_sheet().get_cell(0, 0).comment, Some("My comment".to_string()));

        app.undo();
        assert_eq!(app.workbook.current_sheet().get_cell(0, 0).comment, None);
    }
}