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

impl App {
    /// Switches to editing mode for the currently selected cell. Spill
    /// ghosts are read-only — attempts to edit them surface a hint instead.
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

    /// Starts editing the current cell, replacing its contents with the given character.
    /// Requests application quit. If the workbook is dirty, transitions to a
    /// confirmation prompt instead of quitting immediately.
    pub fn request_quit(&mut self) {
        if self.dirty {
            self.pending_action = Some(PendingAction::Quit);
            self.mode = AppMode::ConfirmDiscard;
            self.status_message = Some(
                "Unsaved changes — quit anyway? (y=quit, n=cancel, s=save & quit)".to_string(),
            );
        } else {
            self.should_quit = true;
        }
    }

    /// Requests loading a file, prompting for confirmation if dirty.
    pub fn request_load_file(&mut self) {
        if self.dirty {
            self.pending_action = Some(PendingAction::LoadFile);
            self.mode = AppMode::ConfirmDiscard;
            self.status_message = Some(
                "Unsaved changes — load new file anyway? (y=discard, n=cancel)".to_string(),
            );
        } else {
            self.start_load_file();
        }
    }

    /// Confirms the pending destructive action.
    pub fn confirm_pending_action(&mut self) {
        let action = self.pending_action.take();
        self.mode = AppMode::Normal;
        self.status_message = None;
        match action {
            Some(PendingAction::Quit) => self.should_quit = true,
            Some(PendingAction::LoadFile) => self.start_load_file(),
            None => {}
        }
    }

    /// Cancels the pending destructive action.
    pub fn cancel_pending_action(&mut self) {
        self.pending_action = None;
        self.mode = AppMode::Normal;
        self.status_message = Some("Cancelled".to_string());
    }

    // ------------------------------------------------------------------
    // Vim helpers
    // ------------------------------------------------------------------

    /// Clear any pending vim state (count, pending operator, awaiting-g flag).
    pub fn vim_reset_pending(&mut self) {
        self.vim_count = None;
        self.vim_pending_op = None;
        self.vim_awaiting_g = false;
    }

    /// Enter visual mode with the given selection granularity, anchoring the
    /// selection at the current cursor.
    pub fn vim_enter_visual(&mut self, kind: VisualKind) {
        self.start_selection();
        match kind {
            VisualKind::Row => {
                let last_col = self.workbook.current_sheet().cols.saturating_sub(1);
                self.selection_start = Some((self.selected_row, 0));
                self.selection_end = Some((self.selected_row, last_col));
            }
            VisualKind::Cell | VisualKind::Block => {
                self.selection_start = Some((self.selected_row, self.selected_col));
                self.selection_end = Some((self.selected_row, self.selected_col));
            }
        }
        self.mode = AppMode::Visual { kind };
    }

    /// Exit visual mode back to Normal, clearing the selection.
    pub fn vim_exit_visual(&mut self) {
        self.clear_selection();
        self.mode = AppMode::Normal;
        self.vim_reset_pending();
    }

    /// Apply a vim operator (delete/yank/change) to a rectangular range.
    /// Coordinates are inclusive; ordering is normalized internally.
    pub fn vim_apply_operator(
        &mut self,
        op: VimOperator,
        r0: usize,
        c0: usize,
        r1: usize,
        c1: usize,
    ) {
        let (r0, r1) = (r0.min(r1), r0.max(r1));
        let (c0, c1) = (c0.min(c1), c0.max(c1));
        // Yank/Delete both copy first; we temporarily set the selection so
        // copy_selection picks up the right cells without us re-implementing it.
        let prev_start = self.selection_start;
        let prev_end = self.selection_end;
        self.selection_start = Some((r0, c0));
        self.selection_end = Some((r1, c1));
        self.copy_selection();
        self.selection_start = prev_start;
        self.selection_end = prev_end;

        match op {
            VimOperator::Yank => {
                self.status_message =
                    Some(format!("Yanked {} cell(s)", (r1 - r0 + 1) * (c1 - c0 + 1)));
            }
            VimOperator::Delete | VimOperator::Change => {
                let mut batch = Vec::new();
                for row in r0..=r1 {
                    for col in c0..=c1 {
                        let old = self
                            .workbook
                            .current_sheet()
                            .cells
                            .get(&(row, col))
                            .cloned();
                        if old.is_some() {
                            self.workbook.current_sheet_mut().clear_cell(row, col);
                            batch.push(UndoAction::CellModified {
                                row,
                                col,
                                old_cell: old,
                                new_cell: None,
                            });
                        }
                    }
                }
                if !batch.is_empty() {
                    self.record_action(UndoAction::Batch(batch));
                    self.dirty = true;
                }
                self.status_message =
                    Some(format!("Deleted {} cell(s)", (r1 - r0 + 1) * (c1 - c0 + 1)));
                if matches!(op, VimOperator::Change) {
                    self.selected_row = r0;
                    self.selected_col = c0;
                    self.start_editing();
                }
            }
        }
        self.vim_reset_pending();
    }

    /// Apply an operator to N whole rows starting at the cursor row (vim `dd`,
    /// `yy`, `cc` with optional count). Marks the resulting clipboard as
    /// row-shaped so `p`/`P` paste below/above instead of right/left.
    pub fn vim_apply_line_op(&mut self, op: VimOperator, count: usize) {
        let count = count.max(1);
        let last_col = self.workbook.current_sheet().cols.saturating_sub(1);
        let r0 = self.selected_row;
        let r1 = (r0 + count - 1).min(self.workbook.current_sheet().rows.saturating_sub(1));
        self.vim_apply_operator(op, r0, 0, r1, last_col);
        if let Some(cb) = self.clipboard.as_mut() {
            cb.is_row_op = true;
        }
    }

    /// `i` / `I` — vim "insert before cursor" / "insert at line start". For
    /// a spreadsheet cell both collapse to placing the text cursor at
    /// position 0. Use `a`/`A` to append at the end.
    pub fn vim_enter_insert(&mut self) {
        self.start_editing();
        self.cursor_position = 0;
    }

    /// `a` / `A` — append after cursor / at end of line. For a cell these
    /// collapse to the same thing: cursor at end of text.
    pub fn vim_enter_insert_at_end(&mut self) {
        self.start_editing();
        self.cursor_position = self.input.chars().count();
    }

    /// `o` — insert a new blank row below the cursor (shifting subsequent
    /// rows down) and enter editing in the same column on the new row.
    pub fn vim_open_row_below(&mut self) {
        let at = self.selected_row + 1;
        self.workbook.current_sheet_mut().insert_row(at);
        self.selected_row = at;
        self.dirty = true;
        self.ensure_cursor_visible();
        self.start_editing();
    }

    /// `O` — insert a new blank row above the cursor (shifting current and
    /// subsequent rows down) and enter editing in the same column. After the
    /// shift, the cursor's row index points at the freshly-inserted blank row.
    pub fn vim_open_row_above(&mut self) {
        let at = self.selected_row;
        self.workbook.current_sheet_mut().insert_row(at);
        self.dirty = true;
        self.ensure_cursor_visible();
        self.start_editing();
    }

    /// `s` — clear current cell and enter insert mode (vim substitute char).
    pub fn vim_substitute_cell(&mut self) {
        self.clear_cell_with_undo(self.selected_row, self.selected_col);
        self.start_editing();
    }

    /// `S` — clear current row and enter insert at first column (vim substitute line).
    pub fn vim_substitute_row(&mut self) {
        self.vim_apply_line_op(VimOperator::Delete, 1);
        self.selected_col = 0;
        self.start_editing();
    }

    /// `p` — paste below/after the cursor. Row-shaped clipboards (from
    /// `yy`/`dd`/`cc` line operators) paste one row below; cell-shaped
    /// clipboards paste one column to the right.
    pub fn vim_paste_below(&mut self) {
        let Some(cb) = self.clipboard.as_ref() else {
            self.status_message = Some("Nothing to paste".to_string());
            return;
        };
        if cb.is_row_op {
            self.selected_row = (self.selected_row + 1)
                .min(self.workbook.current_sheet().rows.saturating_sub(1));
        } else {
            self.selected_col = (self.selected_col + 1)
                .min(self.workbook.current_sheet().cols.saturating_sub(1));
        }
        self.paste();
    }

    /// `P` — paste above/before the cursor.
    pub fn vim_paste_above(&mut self) {
        let Some(cb) = self.clipboard.as_ref() else {
            self.status_message = Some("Nothing to paste".to_string());
            return;
        };
        if cb.is_row_op {
            self.selected_row = self.selected_row.saturating_sub(1);
        } else {
            self.selected_col = self.selected_col.saturating_sub(1);
        }
        self.paste();
    }

    /// Motion `0` — first column of the current row.
    pub fn vim_motion_row_start(&mut self) {
        self.selected_col = 0;
        self.ensure_cursor_visible();
    }

    /// Motion `$` — last non-empty column in the current row (or last column
    /// if the row is empty).
    pub fn vim_motion_row_end(&mut self) {
        let row = self.selected_row;
        let last_data = self
            .workbook
            .current_sheet()
            .cells
            .iter()
            .filter(|((r, _), c)| *r == row && !c.value.is_empty())
            .map(|((_, c), _)| *c)
            .max();
        self.selected_col = last_data
            .unwrap_or_else(|| self.workbook.current_sheet().cols.saturating_sub(1));
        self.ensure_cursor_visible();
    }

    /// Motion `^` — first non-empty column in the current row.
    pub fn vim_motion_row_first_data(&mut self) {
        let row = self.selected_row;
        let first_data = self
            .workbook
            .current_sheet()
            .cells
            .iter()
            .filter(|((r, _), c)| *r == row && !c.value.is_empty())
            .map(|((_, c), _)| *c)
            .min();
        self.selected_col = first_data.unwrap_or(0);
        self.ensure_cursor_visible();
    }

    /// Motion `gg` — first row, keeps current column.
    pub fn vim_motion_top(&mut self) {
        self.selected_row = 0;
        self.ensure_cursor_visible();
    }

    /// Motion `G` — last row that has any data. If no count is given and the
    /// sheet is empty, lands on the last allocated row.
    pub fn vim_motion_bottom(&mut self) {
        let max_data_row = self
            .workbook
            .current_sheet()
            .cells
            .iter()
            .filter(|(_, c)| !c.value.is_empty())
            .map(|((r, _), _)| *r)
            .max();
        self.selected_row = max_data_row
            .unwrap_or_else(|| self.workbook.current_sheet().rows.saturating_sub(1));
        self.ensure_cursor_visible();
    }

    /// `G` with a count — jump to the given 1-based row number.
    pub fn vim_motion_goto_row(&mut self, n: usize) {
        if n == 0 {
            return;
        }
        let target = (n - 1).min(self.workbook.current_sheet().rows.saturating_sub(1));
        self.selected_row = target;
        self.ensure_cursor_visible();
    }

    /// Saves in place if a filename is known and no Shift modifier is desired;
    /// otherwise prompts via SaveAs. Called by `Ctrl+S`.
    pub fn save_in_place_or_prompt(&mut self) {
        if let Some(filename) = self.filename.clone() {
            let result = if filename.to_lowercase().ends_with(".xlsx") {
                crate::infrastructure::xlsx::save_xlsx(&self.workbook, &filename)
                    .map(|_| filename.clone())
            } else {
                crate::infrastructure::FileRepository::save_workbook(&self.workbook, &filename)
            };
            self.set_save_result(result);
        } else {
            self.start_save_as();
        }
    }

    /// Clears all transient UI state: range selection, search highlights,
    /// active filter, and status message. Bound to `Esc` in normal mode.
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

    /// Recalculates every formula cell in every sheet. Bound to `F5`.
    /// Useful after `:cache clear` or for `RAND`-based formulas.
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

    /// Cancels editing without saving changes.
    pub fn cancel_editing(&mut self) {
        self.mode = AppMode::Normal;
        self.input.clear();
        self.cursor_position = 0;
    }

    /// Switches to save-as mode, seeding the prompt with the current filename.
    pub fn start_save_as(&mut self) {
        self.mode = AppMode::SaveAs;
        self.filename_input = self.filename.clone().unwrap_or_else(|| "spreadsheet.tshts".to_string());
        self.cursor_position = self.filename_input.len();
        self.status_message = None;
    }

    pub fn start_load_file(&mut self) {
        self.mode = AppMode::LoadFile;
        self.filename_input = self
            .filename
            .clone()
            .unwrap_or_else(|| "spreadsheet.tshts".to_string());
        self.cursor_position = self.filename_input.len();
        self.status_message = None;
    }

    /// Cycles through filename completion candidates: matching files in the
    /// current directory plus recent files. Called by Tab in filename modes.
    /// Tab-completion for filename input dialogs. Currently exposed for tests
    /// and future Tab-key wiring; the input handler does not yet call it.
    #[allow(dead_code)]
    pub fn complete_filename(&mut self) {
        let input = self.filename_input.clone();
        // Split into (dir, prefix). `dir` is the directory to read; `prefix`
        // is matched against entry names. With no `/`, both default to `.`
        // and the whole input.
        let (dir_part, name_prefix): (String, String) = match input.rfind('/') {
            Some(i) => (input[..=i].to_string(), input[i + 1..].to_string()),
            None => ("".to_string(), input.clone()),
        };
        let read_root = if dir_part.is_empty() { "." } else { dir_part.as_str() };
        let mut candidates: Vec<String> = Vec::new();
        // Recent files only matter when typing from scratch (no dir prefix).
        if dir_part.is_empty() {
            for r in crate::infrastructure::recent::load() {
                if r.starts_with(&name_prefix) && r != name_prefix {
                    candidates.push(r);
                }
            }
        }
        if let Ok(read_dir) = std::fs::read_dir(read_root) {
            for entry in read_dir.flatten() {
                if let Some(name) = entry.file_name().to_str() {
                    if !name.starts_with(&name_prefix) {
                        continue;
                    }
                    // Append `/` to directory matches so a follow-up Tab
                    // descends into them.
                    let is_dir = entry
                        .file_type()
                        .map(|t| t.is_dir())
                        .unwrap_or(false);
                    let full = if is_dir {
                        format!("{}{}/", dir_part, name)
                    } else {
                        format!("{}{}", dir_part, name)
                    };
                    if full != input && !candidates.iter().any(|c| c == &full) {
                        candidates.push(full);
                    }
                }
            }
        }
        if let Some(first) = candidates.into_iter().next() {
            self.filename_input = first;
            self.cursor_position = self.filename_input.chars().count();
        }
    }

    pub fn cancel_filename_input(&mut self) {
        self.mode = AppMode::Normal;
        self.filename_input.clear();
        self.cursor_position = 0;
    }

    pub fn set_save_result(&mut self, result: Result<String, String>) {
        match result {
            Ok(filename) => {
                self.filename = Some(filename.clone());
                self.status_message = Some(format!("Saved to {}", filename));
                self.dirty = false;
                crate::infrastructure::recent::add(&filename);
            }
            Err(error) => {
                self.status_message = Some(format!("Save failed: {}", error));
            }
        }

        self.mode = AppMode::Normal;
        self.filename_input.clear();
        self.cursor_position = 0;
    }

    pub fn set_load_workbook_result(&mut self, result: Result<(Workbook, String), String>) {
        match result {
            Ok((workbook, filename)) => {
                self.workbook = workbook;
                self.filename = Some(filename.clone());
                self.selected_row = 0;
                self.selected_col = 0;
                self.scroll_row = 0;
                self.scroll_col = 0;
                self.undo_stack.clear();
                self.redo_stack.clear();
                self.dirty = false;
                crate::infrastructure::recent::add(&filename);
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

    pub fn get_save_filename(&self) -> String {
        if self.filename_input.is_empty() {
            "spreadsheet.tshts".to_string()
        } else {
            self.filename_input.clone()
        }
    }

    pub fn get_load_filename(&self) -> String {
        if self.filename_input.is_empty() {
            "spreadsheet.tshts".to_string()
        } else {
            self.filename_input.clone()
        }
    }

    pub fn start_csv_export(&mut self) {
        self.mode = AppMode::ExportCsv;
        self.filename_input = self.filename
            .as_ref()
            .map(|f| f.replace(".tshts", ".csv"))
            .unwrap_or_else(|| "spreadsheet.csv".to_string());
        self.cursor_position = self.filename_input.len();
        self.status_message = None;
    }

    pub fn get_csv_export_filename(&self) -> String {
        if self.filename_input.is_empty() {
            "spreadsheet.csv".to_string()
        } else {
            self.filename_input.clone()
        }
    }

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

    pub fn start_csv_import(&mut self) {
        self.mode = AppMode::ImportCsv;
        self.filename_input = "data.csv".to_string();
        self.cursor_position = self.filename_input.len();
        self.status_message = None;
    }

    pub fn get_csv_import_filename(&self) -> String {
        if self.filename_input.is_empty() {
            "data.csv".to_string()
        } else {
            self.filename_input.clone()
        }
    }

    pub fn set_csv_import_result(&mut self, result: Result<Spreadsheet, String>) {
        match result {
            Ok(spreadsheet) => {
                *self.workbook.current_sheet_mut() = spreadsheet;
                self.selected_row = 0;
                self.selected_col = 0;
                self.scroll_row = 0;
                self.scroll_col = 0;
                self.undo_stack.clear();
                self.redo_stack.clear();
                self.dirty = true; // CSV data is in-memory only until saved
                self.status_message = Some("CSV data imported successfully".to_string());
            }
            Err(error) => {
                self.status_message = Some(format!("Import failed: {}", error));
            }
        }

        self.mode = AppMode::Normal;
        self.filename_input.clear();
        self.cursor_position = 0;
    }

    /// Pushes onto the undo stack (capped at `MAX_UNDO_STACK_SIZE`) and clears redo.
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

    /// Wrapper around `Spreadsheet::set_cell` that records the change for undo.
    /// If `(row, col)` sits in the row immediately below a table, extend
    /// that table's range by one row and re-register its column named
    /// ranges. Called from `set_cell_with_undo` so typing under a table
    /// grows it without a manual `:table extend`.
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

    /// Wrapper around `Spreadsheet::clear_cell` that records the change for undo.
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
        self.stats_cache = None;
    }

    /// Updates the selection end position
    pub fn update_selection(&mut self, row: usize, col: usize) {
        if self.selecting {
            self.selection_end = Some((row, col));
            self.stats_cache = None;
        }
    }

    /// Clears the current selection
    pub fn clear_selection(&mut self) {
        self.selection_start = None;
        self.selection_end = None;
        self.selecting = false;
        self.stats_cache = None;
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

    /// Copies the current selection to internal + system + sidecar clipboards.
    /// The system clipboard gets sentinel-prefixed TSV so other apps see
    /// values and tshts can detect rich copies. The sidecar JSON carries
    /// full formula/format/comment data.
    pub fn copy_selection(&mut self) {
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

        // Sentinel-prefixed TSV for system clipboard.
        let mut tsv = String::from(crate::infrastructure::sidecar::SENTINEL);
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
        if let Ok(mut board) = arboard::Clipboard::new() {
            let _ = board.set_text(tsv);
        }
        // Sidecar JSON for formula round-trip. Skipped in tests for the same
        // state-isolation reason as the read path.
        if !cfg!(test) {
            crate::infrastructure::sidecar::write(cells.clone(), start_row, start_col);
        }

        self.clipboard = Some(ClipboardData {
            cells,
            source_row: start_row,
            source_col: start_col,
            is_row_op: false,
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
            is_row_op: false,
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
        // If the system clipboard has our sentinel header, load the matching
        // sidecar JSON (preserves formulas/formats/comments). Otherwise fall
        // back to internal clipboard, then plain-text TSV.
        // Skipped in tests because $HOME and ~/.cache/tshts persist between
        // test runs and would otherwise leak state between cases.
        if !cfg!(test) {
            if let Ok(mut board) = arboard::Clipboard::new() {
                if let Ok(text) = board.get_text() {
                    if crate::infrastructure::sidecar::strip_sentinel(&text).is_some() {
                        if let Some(payload) = crate::infrastructure::sidecar::read() {
                            let cb = ClipboardData {
                                cells: payload.cells,
                                source_row: payload.source_row,
                                source_col: payload.source_col,
                                is_row_op: false,
                            };
                            self.clipboard = Some(cb);
                        }
                    }
                }
            }
        }
        let clipboard = if let Some(ref cb) = self.clipboard {
            cb.clone()
        } else {
            // Tests don't touch the system clipboard (cross-test contamination).
            if !cfg!(test) {
                if let Ok(mut board) = arboard::Clipboard::new() {
                    if let Ok(text) = board.get_text() {
                        if !text.is_empty() {
                            let body = crate::infrastructure::sidecar::strip_sentinel(&text)
                                .unwrap_or(&text);
                            self.paste_tsv(body);
                            return;
                        }
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
                    CellData { value, formula: Some(adjusted), format: cell.format.clone(), comment: cell.comment.clone(), spill_anchor: None }
                } else {
                    cell.clone()
                };
                Some((target_row, target_col, new_cell))
            }).collect()
        };

        // Now apply changes (mutably borrows spreadsheet) — collect the
        // undo batch, then push all writes through `set_many` in one shot
        // so dependents recalc just once.
        let mut batch = Vec::new();
        let mut writes: Vec<(usize, usize, CellData)> = Vec::with_capacity(new_cells.len());
        for (target_row, target_col, new_cell) in &new_cells {
            let old = if self
                .workbook
                .current_sheet()
                .cells
                .contains_key(&(*target_row, *target_col))
            {
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
            writes.push((*target_row, *target_col, new_cell.clone()));
        }
        if !writes.is_empty() {
            self.workbook.current_sheet_mut().set_many(writes);
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
                // If the pasted cell starts with `=`, treat it as a formula
                // and evaluate it relative to the destination cell.
                let new_cell = if value.starts_with('=') {
                    let evaluator = FormulaEvaluator::new(self.workbook.current_sheet());
                    let evaluated = evaluator.evaluate_formula(value);
                    CellData {
                        value: evaluated,
                        formula: Some(value.to_string()),
                        format: None,
                        comment: None,
                    spill_anchor: None,
                    }
                } else {
                    CellData {
                        value: value.to_string(),
                        formula: None,
                        format: None,
                        comment: None,
                    spill_anchor: None,
                    }
                };
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
        self.dirty = true;
        self.status_message = Some(format!("Inserted row at {}", insert_at + 1));
    }

    pub fn delete_row(&mut self) {
        let delete_at = self.selected_row;
        self.workbook.current_sheet_mut().delete_row(delete_at);
        if self.selected_row >= self.workbook.current_sheet().rows {
            self.selected_row = self.workbook.current_sheet().rows.saturating_sub(1);
        }
        self.dirty = true;
        self.status_message = Some(format!("Deleted row {}", delete_at + 1));
    }

    pub fn insert_col(&mut self) {
        let insert_at = self.selected_col;
        self.workbook.current_sheet_mut().insert_col(insert_at);
        self.dirty = true;
        self.status_message = Some(format!("Inserted column at {}", crate::domain::Spreadsheet::column_label(insert_at)));
    }

    pub fn delete_col(&mut self) {
        let delete_at = self.selected_col;
        self.workbook.current_sheet_mut().delete_col(delete_at);
        if self.selected_col >= self.workbook.current_sheet().cols {
            self.selected_col = self.workbook.current_sheet().cols.saturating_sub(1);
        }
        self.dirty = true;
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
        if self.find_replace_search.is_empty() {
            return;
        }
        let matcher = TextMatcher::new(
            &self.find_replace_search,
            self.search_regex,
            self.search_case_sensitive,
        );
        for row in 0..self.workbook.current_sheet().rows {
            for col in 0..self.workbook.current_sheet().cols {
                let cell = self.workbook.current_sheet().get_cell(row, col);
                if matcher.is_match(&cell.value) {
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
        self.find_replace_search();
    }

    pub fn replace_all(&mut self) {
        if self.find_replace_results.is_empty() {
            return;
        }
        let matcher = TextMatcher::new(
            &self.find_replace_search,
            self.search_regex,
            self.search_case_sensitive,
        );
        let mut batch = Vec::new();
        let results = self.find_replace_results.clone();
        for (row, col) in results {
            let cell = self.workbook.current_sheet().get_cell(row, col);
            if cell.formula.is_some() {
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
        // Handle vim-style ex commands that take filenames first — they need
        // case-preserved arguments. Note `:wq` / `:x` / `:wq!` / `:x!` save +
        // quit, and `:q!` skips the dirty check.
        let trimmed = self.command_input.trim().to_string();
        let lower = trimmed.to_lowercase();
        let close_palette = |s: &mut Self| {
            s.mode = AppMode::Normal;
            s.command_input.clear();
            s.cursor_position = 0;
        };
        // Empty input: just close the palette silently (vim convention).
        if trimmed.is_empty() {
            close_palette(self);
            return;
        }
        // Quit / save-quit variants — exact-match on lowercased string.
        match lower.as_str() {
            "q" => {
                close_palette(self);
                self.request_quit();
                return;
            }
            "q!" | "quit!" => {
                close_palette(self);
                self.should_quit = true;
                return;
            }
            "w" => {
                close_palette(self);
                self.save_in_place_or_prompt();
                return;
            }
            "wq" | "x" => {
                close_palette(self);
                self.save_in_place_or_prompt();
                if !self.dirty {
                    self.should_quit = true;
                }
                return;
            }
            "wq!" | "x!" => {
                close_palette(self);
                self.save_in_place_or_prompt();
                self.should_quit = true;
                return;
            }
            "e" | "edit" => {
                close_palette(self);
                self.request_load_file();
                return;
            }
            _ => {}
        }
        // `:w <filename>` and `:e <filename>` preserve case.
        if let Some(rest) = trimmed.strip_prefix("w ").or_else(|| trimmed.strip_prefix("W ")) {
            let name = rest.trim().to_string();
            if !name.is_empty() {
                let result = if name.to_lowercase().ends_with(".xlsx") {
                    crate::infrastructure::xlsx::save_xlsx(&self.workbook, &name)
                        .map(|_| name.clone())
                } else {
                    crate::infrastructure::FileRepository::save_workbook(&self.workbook, &name)
                };
                self.set_save_result(result);
            } else {
                self.status_message = Some("Usage: w <filename>".to_string());
            }
            close_palette(self);
            return;
        }
        if let Some(rest) = trimmed.strip_prefix("e ").or_else(|| trimmed.strip_prefix("E ")) {
            let name = rest.trim().to_string();
            if name.is_empty() {
                self.status_message = Some("Usage: e <filename>".to_string());
                close_palette(self);
                return;
            }
            // `:e <file>` loads the typed file directly. Going through
            // request_load_file/start_load_file would clobber `filename_input`
            // with the currently-open file's name (or the "spreadsheet.tshts"
            // default), defeating the purpose. Vim's `:e` is also a direct
            // load; users who want a dirty-guard can save first.
            let result = if name.to_lowercase().ends_with(".xlsx") {
                crate::infrastructure::xlsx::load_xlsx(&name)
                    .map(|wb| (wb, name.clone()))
            } else {
                crate::infrastructure::FileRepository::load_workbook(&name)
            };
            self.set_load_workbook_result(result);
            close_palette(self);
            return;
        }
        // `:export <file>` writes the current sheet to CSV.
        if let Some(rest) = trimmed
            .strip_prefix("export ")
            .or_else(|| trimmed.strip_prefix("EXPORT "))
        {
            let name = rest.trim().to_string();
            if name.is_empty() {
                self.status_message = Some("Usage: export <filename>".to_string());
            } else {
                let result = crate::domain::CsvExporter::export_to_csv(
                    self.workbook.current_sheet(),
                    &name,
                );
                self.set_csv_export_result(result);
            }
            close_palette(self);
            return;
        }
        // `:import <file>` replaces the current sheet from CSV. We could
        // wire this through request_csv_import for a dirty-check prompt, but
        // start_csv_import would then clobber the typed filename — so for
        // now, mirror `:e <file>`'s pattern and import directly. Users who
        // want a dirty-guard can `:w` first.
        if let Some(rest) = trimmed
            .strip_prefix("import ")
            .or_else(|| trimmed.strip_prefix("IMPORT "))
        {
            let name = rest.trim().to_string();
            if name.is_empty() {
                self.status_message = Some("Usage: import <filename>".to_string());
            } else {
                let result = crate::domain::CsvExporter::import_from_csv(&name);
                self.set_csv_import_result(result);
            }
            close_palette(self);
            return;
        }
        if trimmed.starts_with("rename ") || trimmed.starts_with("RENAME ") {
            let name = trimmed[7..].trim().to_string();
            if !name.is_empty() {
                if self.workbook.rename_sheet(name.clone()) {
                    self.status_message = Some(format!("Renamed sheet to '{}'", name));
                } else {
                    self.status_message = Some(format!(
                        "Rename rejected: '{}' is empty or a duplicate",
                        name
                    ));
                }
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
            ["replace"] | ["find-replace"] | ["find", "replace"] => {
                self.start_find_replace();
            }
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
                // Preserve case for the comment text — the user wrote it for
                // a human to read later. `parts` is lowercased so we recover
                // from `trimmed`.
                let preserved = trimmed
                    .strip_prefix("comment ")
                    .or_else(|| trimmed.strip_prefix("COMMENT "))
                    .or_else(|| trimmed.strip_prefix("Comment "))
                    .unwrap_or("")
                    .trim()
                    .to_string();
                let text_lc = parts[1..].join(" ");
                if text_lc == "clear" || text_lc == "none" {
                    self.set_cell_comment(None);
                } else if !preserved.is_empty() {
                    self.set_cell_comment(Some(preserved));
                } else {
                    self.set_cell_comment(Some(text_lc));
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
            ["recalc"] | ["refresh"] => self.recalc_all(),
            ["cache", "clear"] => {
                crate::infrastructure::fetcher::clear_cache();
                self.status_message = Some("GET cache cleared".to_string());
            }
            ["clipboard", "clear"] => {
                crate::infrastructure::sidecar::clear();
                self.clipboard = None;
                self.status_message = Some("Clipboard cleared".to_string());
            }
            ["import-append", path] | ["append", path] => {
                let path = path.to_string();
                match crate::domain::CsvExporter::append_from_csv(
                    self.workbook.current_sheet_mut(),
                    &path,
                ) {
                    Ok(n) => {
                        self.dirty = true;
                        self.status_message =
                            Some(format!("Appended {} row(s) from {}", n, path));
                    }
                    Err(e) => {
                        self.status_message = Some(format!("Append failed: {}", e));
                    }
                }
            }
            ["hide", "row"] => {
                let hidden = self.selected_row;
                self.hidden_rows.insert(hidden);
                self.status_message = Some(format!("Hid row {}", hidden + 1));
                // Move cursor off the hidden row — otherwise the formula bar
                // keeps showing hidden content and the user can't escape.
                let max_row = self.workbook.current_sheet().rows.saturating_sub(1);
                let mut next = self.selected_row + 1;
                while next <= max_row && self.hidden_rows.contains(&next) {
                    next += 1;
                }
                if next <= max_row {
                    self.selected_row = next;
                } else {
                    // No visible row below — try above.
                    let mut prev = hidden;
                    while prev > 0 && self.hidden_rows.contains(&prev) {
                        prev -= 1;
                    }
                    if !self.hidden_rows.contains(&prev) {
                        self.selected_row = prev;
                    }
                }
                self.ensure_cursor_visible();
            }
            ["show", "rows"] | ["unhide", "rows"] => {
                self.hidden_rows.clear();
                self.status_message = Some("All rows shown".to_string());
            }
            ["hide", "col"] => {
                let hidden = self.selected_col;
                self.hidden_cols.insert(hidden);
                self.status_message = Some(format!(
                    "Hid column {}",
                    crate::domain::Spreadsheet::column_label(hidden)
                ));
                // Move cursor off the hidden col — same reason as hide row.
                let max_col = self.workbook.current_sheet().cols.saturating_sub(1);
                let mut next = self.selected_col + 1;
                while next <= max_col && self.hidden_cols.contains(&next) {
                    next += 1;
                }
                if next <= max_col {
                    self.selected_col = next;
                } else {
                    let mut prev = hidden;
                    while prev > 0 && self.hidden_cols.contains(&prev) {
                        prev -= 1;
                    }
                    if !self.hidden_cols.contains(&prev) {
                        self.selected_col = prev;
                    }
                }
                self.ensure_cursor_visible();
            }
            ["hide", "col", col_name] => {
                if let Some(c) = Spreadsheet::parse_column_label(col_name) {
                    self.hidden_cols.insert(c);
                    self.status_message = Some(format!("Hid column {}", col_name));
                    // If the cursor sits on the just-hidden column, advance it.
                    if self.selected_col == c {
                        let max_col = self.workbook.current_sheet().cols.saturating_sub(1);
                        let mut next = c + 1;
                        while next <= max_col && self.hidden_cols.contains(&next) {
                            next += 1;
                        }
                        if next <= max_col {
                            self.selected_col = next;
                        } else if c > 0 {
                            let mut prev = c - 1;
                            while prev > 0 && self.hidden_cols.contains(&prev) {
                                prev -= 1;
                            }
                            if !self.hidden_cols.contains(&prev) {
                                self.selected_col = prev;
                            }
                        }
                        self.ensure_cursor_visible();
                    }
                } else {
                    self.status_message = Some(format!("Invalid column: {}", col_name));
                }
            }
            ["show", "cols"] | ["unhide", "cols"] => {
                self.hidden_cols.clear();
                self.status_message = Some("All columns shown".to_string());
            }
            ["show", "all"] => {
                self.hidden_rows.clear();
                self.hidden_cols.clear();
                self.status_message = Some("All rows and columns shown".to_string());
            }
            // `:table create A1:D10 name=Sales`
            ["table", "create", range_str, opts @ ..] => {
                if let Some((start, end)) = parse_range(range_str) {
                    let mut name = format!("Table{}", self.workbook.current_sheet().tables.len() + 1);
                    // Recover case-preserved opts from the original command so
                    // `name=Mine` is stored as "Mine", not "mine".
                    let original_opts: Vec<&str> = trimmed
                        .split_whitespace()
                        .skip(3)
                        .collect();
                    for (i, opt) in opts.iter().enumerate() {
                        if opt.starts_with("name=") {
                            // Pull the same position from the original (case-preserved) tokens.
                            if let Some(orig) = original_opts.get(i) {
                                if let Some(n) = orig.strip_prefix("name=") {
                                    name = n.to_string();
                                    continue;
                                }
                            }
                            // Fallback: use lowercased opt.
                            if let Some(n) = opt.strip_prefix("name=") {
                                name = n.to_string();
                            }
                        }
                    }
                    let sheet = self.workbook.current_sheet();
                    let cols = end.1 - start.1 + 1;
                    let mut headers = Vec::with_capacity(cols);
                    for c in start.1..=end.1 {
                        let h = sheet.get_cell(start.0, c).value;
                        headers.push(if h.is_empty() {
                            crate::domain::Spreadsheet::column_label(c)
                        } else {
                            h
                        });
                    }
                    let table = crate::domain::Table {
                        name: name.clone(),
                        top_row: start.0,
                        left_col: start.1,
                        bottom_row: end.0,
                        right_col: end.1,
                        headers,
                    };
                    self.workbook.current_sheet_mut().tables.push(table);
                    // Also register each column as a named range so
                    // `Table1[Col1]` works via existing name resolution.
                    let sheet = self.workbook.current_sheet();
                    let table = sheet.tables.last().unwrap().clone();
                    let body_top = table.top_row + 1; // skip header row
                    for (i, header) in table.headers.iter().enumerate() {
                        let key = format!("{}[{}]", table.name.to_uppercase(), header.to_uppercase());
                        let value = format!(
                            "{}:{}",
                            crate::domain::Spreadsheet::format_cell_reference(body_top, table.left_col + i, false, false),
                            crate::domain::Spreadsheet::format_cell_reference(table.bottom_row, table.left_col + i, false, false),
                        );
                        self.workbook.named_ranges.insert(key.clone(), value.clone());
                        for s in &mut self.workbook.sheets {
                            s.named_ranges.insert(key.clone(), value.clone());
                        }
                    }
                    self.dirty = true;
                    self.status_message = Some(format!("Created table '{}'", name));
                } else {
                    self.status_message = Some(format!("Invalid range: {}", range_str));
                }
            }
            // `:pivot RANGE TARGET row=COL value=COL agg=sum|count|avg|min|max`
            // `:chart bar A1:A10 [title=Sales]`
            ["chart", kind, range_str, opts @ ..] => {
                if let Some((s, e)) = parse_range(range_str) {
                    let mut title = format!("{} chart", kind);
                    for o in opts {
                        if let Some(t) = o.strip_prefix("title=") {
                            title = t.to_string();
                        }
                    }
                    let chart_kind = match *kind {
                        "bar" => ChartKind::Bar,
                        "line" => ChartKind::Line,
                        _ => ChartKind::Sparkline,
                    };
                    self.chart_popup = Some(ChartPopup {
                        title,
                        source: (s, e),
                        kind: chart_kind,
                    });
                    self.status_message = Some("Chart shown (Esc to close)".to_string());
                } else {
                    self.status_message = Some("chart: bad range".to_string());
                }
            }
            ["iterative", "on"] => {
                self.iterative_calc = true;
                for s in &mut self.workbook.sheets {
                    s.iterative_calc = true;
                }
                self.status_message = Some(format!(
                    "Iterative calc: on (max {} iters, eps {})",
                    self.workbook.current_sheet().iter_max,
                    self.workbook.current_sheet().iter_epsilon
                ));
            }
            ["iterative", "off"] => {
                self.iterative_calc = false;
                for s in &mut self.workbook.sheets {
                    s.iterative_calc = false;
                }
                self.status_message = Some("Iterative calc: off".to_string());
            }
            ["iterative", "max", n] => {
                if let Ok(v) = n.parse::<usize>() {
                    for s in &mut self.workbook.sheets {
                        s.iter_max = v;
                    }
                    self.status_message = Some(format!("Iterative max = {}", v));
                } else {
                    self.status_message = Some("iterative max: bad number".to_string());
                }
            }
            ["iterative", "epsilon", n] => {
                if let Ok(v) = n.parse::<f64>() {
                    for s in &mut self.workbook.sheets {
                        s.iter_epsilon = v;
                    }
                    self.status_message = Some(format!("Iterative epsilon = {}", v));
                } else {
                    self.status_message = Some("iterative epsilon: bad number".to_string());
                }
            }
            ["r1c1", "on"] => {
                self.r1c1_mode = true;
                self.status_message = Some("R1C1 reference mode: on".to_string());
            }
            ["r1c1", "off"] => {
                self.r1c1_mode = false;
                self.status_message = Some("R1C1 reference mode: off".to_string());
            }
            // `:validate A "_ > 0"` — predicate with `_` bound to the cell value.
            ["validate", col_name, rest @ ..] if !rest.is_empty() => {
                if let Some(col) = Spreadsheet::parse_column_label(col_name) {
                    let predicate = rest.join(" ");
                    self.validations.insert(col, predicate.clone());
                    self.status_message = Some(format!("Validation set on col {}", col_name));
                } else {
                    self.status_message = Some(format!("Invalid column: {}", col_name));
                }
            }
            ["validate", "clear"] => {
                self.validations.clear();
                self.status_message = Some("Validations cleared".to_string());
            }
            ["pivot", source, target, opts @ ..] => {
                if let (Some((s, e)), Some(t)) = (
                    parse_range(source),
                    crate::domain::Spreadsheet::parse_cell_reference(target),
                ) {
                    let mut row_col: Option<usize> = None;
                    let mut val_col: Option<usize> = None;
                    let mut agg = "sum".to_string();
                    for o in opts {
                        if let Some(c) = o.strip_prefix("row=") {
                            row_col = crate::domain::Spreadsheet::parse_column_label(c);
                        } else if let Some(c) = o.strip_prefix("value=") {
                            val_col = crate::domain::Spreadsheet::parse_column_label(c);
                        } else if let Some(a) = o.strip_prefix("agg=") {
                            agg = a.to_string();
                        }
                    }
                    let row_col = match row_col {
                        Some(c) => c,
                        None => {
                            self.status_message = Some("pivot: row=COL required".to_string());
                            return;
                        }
                    };
                    let val_col = val_col.unwrap_or(row_col);
                    // Collect distinct row-keys preserving sort order. The
                    // pivot values are written as formulas (`=SUMIF(...)`)
                    // so they auto-update when source data changes.
                    let mut keys: std::collections::BTreeSet<String> =
                        std::collections::BTreeSet::new();
                    {
                        let sheet = self.workbook.current_sheet();
                        for r in s.0..=e.0 {
                            let k = sheet.get_cell(r, row_col).value;
                            if !k.is_empty() {
                                keys.insert(k);
                            }
                        }
                    }
                    let row_col_label = crate::domain::Spreadsheet::column_label(row_col);
                    let val_col_label = crate::domain::Spreadsheet::column_label(val_col);
                    let source_range = format!(
                        "{}{}:{}{}",
                        row_col_label,
                        s.0 + 1,
                        row_col_label,
                        e.0 + 1
                    );
                    let sum_range = format!(
                        "{}{}:{}{}",
                        val_col_label,
                        s.0 + 1,
                        val_col_label,
                        e.0 + 1
                    );
                    let agg_formula = |key: &str| -> String {
                        let key_esc = key.replace('"', "\"\"");
                        match agg.as_str() {
                            "count" => format!(
                                "=COUNTIF({}, \"{}\")",
                                source_range, key_esc
                            ),
                            "avg" | "average" => format!(
                                "=AVERAGEIF({}, \"{}\", {})",
                                source_range, key_esc, sum_range
                            ),
                            "min" => format!(
                                "=MIN(IF({}=\"{}\", {}))",
                                source_range, key_esc, sum_range
                            ),
                            "max" => format!(
                                "=MAX(IF({}=\"{}\", {}))",
                                source_range, key_esc, sum_range
                            ),
                            _ => format!(
                                "=SUMIF({}, \"{}\", {})",
                                source_range, key_esc, sum_range
                            ),
                        }
                    };
                    let mut rows: Vec<(usize, usize, CellData)> = Vec::new();
                    rows.push((
                        t.0,
                        t.1,
                        CellData {
                            value: "Key".to_string(),
                            formula: None,
                            format: None,
                            comment: None,
                        spill_anchor: None,
                        },
                    ));
                    rows.push((
                        t.0,
                        t.1 + 1,
                        CellData {
                            value: agg.clone(),
                            formula: None,
                            format: None,
                            comment: None,
                        spill_anchor: None,
                        },
                    ));
                    for (i, k) in keys.iter().enumerate() {
                        rows.push((
                            t.0 + 1 + i,
                            t.1,
                            CellData {
                                value: k.clone(),
                                formula: None,
                                format: None,
                                comment: None,
                            spill_anchor: None,
                            },
                        ));
                        let formula = agg_formula(k);
                        // Evaluate immediately so the cell shows its initial
                        // value; set_many will rebuild deps so it tracks.
                        let evaluator = FormulaEvaluator::with_workbook(
                            &self.workbook,
                            self.workbook.current_sheet(),
                            &self.workbook.named_ranges,
                        );
                        let initial = evaluator.evaluate_formula(&formula);
                        rows.push((
                            t.0 + 1 + i,
                            t.1 + 1,
                            CellData {
                                value: initial,
                                formula: Some(formula),
                                format: None,
                                comment: None,
                            spill_anchor: None,
                            },
                        ));
                    }
                    self.workbook.current_sheet_mut().set_many(rows);
                    self.dirty = true;
                    self.status_message = Some(format!(
                        "Pivot written to {}{}: {} groups (auto-refreshes via formulas)",
                        crate::domain::Spreadsheet::column_label(t.1),
                        t.0 + 1,
                        keys.len()
                    ));
                } else {
                    self.status_message = Some("pivot: bad source range or target".to_string());
                }
            }
            // `:goalseek TARGET_CELL EXPECTED INPUT_CELL` — bisect input until target = expected.
            ["goalseek", target, expected, input] => {
                let target_pos = crate::domain::Spreadsheet::parse_cell_reference(target);
                let input_pos = crate::domain::Spreadsheet::parse_cell_reference(input);
                let expected_v: f64 = expected.parse().unwrap_or(0.0);
                if let (Some(t), Some(i)) = (target_pos, input_pos) {
                    let mut lo = -1e9_f64;
                    let mut hi = 1e9_f64;
                    let original_input = self.workbook.current_sheet().get_cell(i.0, i.1).value.clone();
                    let mut result: Option<f64> = None;
                    for _ in 0..80 {
                        let mid = (lo + hi) / 2.0;
                        let cell = CellData {
                            value: mid.to_string(),
                            formula: None,
                            format: None,
                            comment: None,
                        spill_anchor: None,
                        };
                        self.workbook.current_sheet_mut().set_cell(i.0, i.1, cell);
                        let cur: f64 = self.workbook.current_sheet().get_cell(t.0, t.1).value.parse().unwrap_or(0.0);
                        if (cur - expected_v).abs() < 1e-6 {
                            result = Some(mid);
                            break;
                        }
                        if cur < expected_v {
                            lo = mid;
                        } else {
                            hi = mid;
                        }
                    }
                    if let Some(v) = result {
                        self.dirty = true;
                        self.status_message = Some(format!("Goal seek: {} = {:.6}", input, v));
                    } else {
                        // Restore original input on failure.
                        let cell = CellData {
                            value: original_input,
                            formula: None,
                            format: None,
                            comment: None,
                        spill_anchor: None,
                        };
                        self.workbook.current_sheet_mut().set_cell(i.0, i.1, cell);
                        self.status_message = Some("Goal seek did not converge".to_string());
                    }
                } else {
                    self.status_message = Some("goalseek: bad cell reference".to_string());
                }
            }
            // `:trace` — show what cells the current cell's formula depends on.
            ["trace"] | ["trace", "precedents"] => {
                let cell = self
                    .workbook
                    .current_sheet()
                    .get_cell(self.selected_row, self.selected_col);
                if let Some(formula) = cell.formula {
                    let evaluator = FormulaEvaluator::new(self.workbook.current_sheet());
                    let refs = evaluator.extract_cell_references(&formula);
                    if refs.is_empty() {
                        self.status_message = Some("(no precedents)".to_string());
                    } else {
                        let refs_str: Vec<String> = refs
                            .iter()
                            .map(|(r, c)| {
                                format!(
                                    "{}{}",
                                    crate::domain::Spreadsheet::column_label(*c),
                                    r + 1
                                )
                            })
                            .collect();
                        self.status_message = Some(format!("Precedents: {}", refs_str.join(", ")));
                    }
                } else {
                    self.status_message = Some("(cell has no formula)".to_string());
                }
            }
            ["trace", "dependents"] => {
                let pos = (self.selected_row, self.selected_col);
                let deps = self.workbook.current_sheet().dependents.get(&pos).cloned();
                match deps {
                    Some(set) if !set.is_empty() => {
                        let s: Vec<String> = set
                            .iter()
                            .map(|(r, c)| {
                                format!(
                                    "{}{}",
                                    crate::domain::Spreadsheet::column_label(*c),
                                    r + 1
                                )
                            })
                            .collect();
                        self.status_message = Some(format!("Dependents: {}", s.join(", ")));
                    }
                    _ => self.status_message = Some("(no dependents)".to_string()),
                }
            }
            ["table", "list"] => {
                let sheet = self.workbook.current_sheet();
                if sheet.tables.is_empty() {
                    self.status_message = Some("(no tables on this sheet)".to_string());
                } else {
                    let names: Vec<String> = sheet
                        .tables
                        .iter()
                        .map(|t| format!("{}({}x{})", t.name, t.bottom_row - t.top_row + 1, t.headers.len()))
                        .collect();
                    self.status_message = Some(names.join(", "));
                }
            }
            ["cf", "clear"] => {
                let n = self.workbook.current_sheet().conditional_formats.len();
                self.workbook.current_sheet_mut().conditional_formats.clear();
                self.workbook.current_sheet_mut().cf_cache.borrow_mut().clear();
                self.dirty = true;
                self.status_message = Some(format!("Cleared {} conditional format rule(s)", n));
            }
            ["cf", "list"] => {
                let rules = &self.workbook.current_sheet().conditional_formats;
                if rules.is_empty() {
                    self.status_message = Some("(no conditional formats)".to_string());
                } else {
                    let entries: Vec<String> = rules.iter().enumerate().map(|(i, r)| {
                        format!("#{} col {} when {}", i, crate::domain::Spreadsheet::column_label(r.column), r.predicate)
                    }).collect();
                    self.status_message = Some(entries.join("  |  "));
                }
            }
            // `:cf <col> <predicate> [bg=COLOR] [fg=COLOR] [bold] [underline]`
            // The predicate may use `_` for the cell value. To pass a spaced
            // predicate, wrap in quotes — but the palette splits on whitespace,
            // so we accept the predicate as a single token. Workaround: join
            // all middle tokens that don't look like style keys.
            ["cf", col_name, rest @ ..] if !rest.is_empty() => {
                if let Some(col) = Spreadsheet::parse_column_label(col_name) {
                    let mut predicate_parts = Vec::new();
                    let mut style = crate::domain::CellStyle::default();
                    for tok in rest {
                        if let Some(c) = tok.strip_prefix("bg=") {
                            style.bg_color = TerminalColor::from_name(c);
                        } else if let Some(c) = tok.strip_prefix("fg=") {
                            style.fg_color = TerminalColor::from_name(c);
                        } else if *tok == "bold" {
                            style.bold = true;
                        } else if *tok == "underline" {
                            style.underline = true;
                        } else {
                            predicate_parts.push(*tok);
                        }
                    }
                    if predicate_parts.is_empty() {
                        self.status_message = Some(
                            "Usage: cf <col> <predicate> [bg=color] [fg=color] [bold] [underline]"
                                .to_string(),
                        );
                    } else {
                        let predicate = predicate_parts.join(" ");
                        {
                            let sheet = self.workbook.current_sheet_mut();
                            sheet.conditional_formats.push(crate::domain::ConditionalFormat {
                                column: col,
                                predicate: predicate.clone(),
                                style,
                            });
                            sheet.cf_cache.borrow_mut().clear();
                        }
                        self.dirty = true;
                        self.status_message = Some(format!(
                            "Added cf for col {}: {}",
                            crate::domain::Spreadsheet::column_label(col),
                            predicate
                        ));
                    }
                } else {
                    self.status_message = Some(format!("Invalid column: {}", col_name));
                }
            }
            ["name", name, rest @ ..] if !rest.is_empty() => {
                // Join the remaining tokens so values with spaces (e.g.
                // `LAMBDA(x, x*2)`) survive intact.
                let value = rest.join(" ");
                self.workbook.set_name(name, &value);
                self.dirty = true;
                self.status_message = Some(format!("Named '{}' = {}", name, value));
            }
            ["unname", name] => {
                if self.workbook.remove_name(name) {
                    self.dirty = true;
                    self.status_message = Some(format!("Removed name '{}'", name));
                } else {
                    self.status_message = Some(format!("No such name: {}", name));
                }
            }
            ["names"] => {
                if self.workbook.named_ranges.is_empty() {
                    self.status_message = Some("(no named ranges)".to_string());
                } else {
                    let mut entries: Vec<String> = self
                        .workbook
                        .named_ranges
                        .iter()
                        .map(|(k, v)| format!("{}={}", k, v))
                        .collect();
                    entries.sort();
                    self.status_message = Some(entries.join("  "));
                }
            }
            ["autosave", "on"] => {
                crate::infrastructure::autosave::enable();
                self.status_message =
                    Some("Auto-save enabled (30s idle, writes to current filename)".to_string());
            }
            ["autosave", "off"] => {
                crate::infrastructure::autosave::disable();
                self.status_message = Some("Auto-save disabled".to_string());
            }
            ["regex", "on"] => {
                self.search_regex = true;
                self.status_message = Some("Regex search: on".to_string());
            }
            ["regex", "off"] => {
                self.search_regex = false;
                self.status_message = Some("Regex search: off".to_string());
            }
            ["case", "on"] | ["case-sensitive", "on"] => {
                self.search_case_sensitive = true;
                self.status_message = Some("Case-sensitive search: on".to_string());
            }
            ["case", "off"] | ["case-sensitive", "off"] => {
                self.search_case_sensitive = false;
                self.status_message = Some("Case-sensitive search: off".to_string());
            }
            _ => {
                self.status_message = Some(format!("Unknown command: {}", self.command_input));
            }
        }
        self.mode = AppMode::Normal;
        self.command_input.clear();
        self.cursor_position = 0;
    }

    /// Returns up to `max` command suggestions matching the current
    /// `command_input` prefix. Used by the command palette autocomplete.
    pub fn command_suggestions(&self, max: usize) -> Vec<&'static str> {
        const ALL: &[&str] = &[
            "q", "q!", "w", "wq", "x", "e ",
            "export ", "import ", "import-append ",
            "ir", "dr", "ic", "dc",
            "insert row", "delete row", "insert col", "delete col",
            "sort asc", "sort desc",
            "freeze", "unfreeze",
            "format general", "format number", "format currency", "format percent",
            "bold", "underline",
            "color red", "color green", "color blue", "color yellow", "color none",
            "bg red", "bg green", "bg blue", "bg yellow", "bg none",
            "sheet new", "sheet delete", "sheet next", "sheet prev",
            "rename ",
            "comment ",
            "filter ", "unfilter",
            "hide row", "show rows",
            "recalc", "cache clear",
            "regex on", "regex off",
            "case on", "case off",
            "autosave on", "autosave off",
            "import-append ",
            "name ", "unname ", "names",
            "cf ", "cf clear", "cf list",
            "table create ", "table list",
            "pivot ", "goalseek ",
            "trace", "trace dependents",
        ];
        let q = self.command_input.trim().to_lowercase();
        if q.is_empty() {
            return ALL.iter().take(max).copied().collect();
        }
        ALL.iter()
            .filter(|c| c.starts_with(&q))
            .take(max)
            .copied()
            .collect()
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
        let mut max_row = 0;
        let mut max_col = 0;
        for &(r, c) in self.workbook.current_sheet().cells.keys() {
            max_row = max_row.max(r);
            max_col = max_col.max(c);
        }
        if max_row == 0 {
            self.status_message = Some("Nothing to sort".to_string());
            return;
        }

        // Capture each existing row as (original_row_index, Vec<Option<CellData>>).
        let mut rows: Vec<(usize, Vec<Option<CellData>>)> = Vec::with_capacity(max_row + 1);
        for row in 0..=max_row {
            let mut row_data = Vec::with_capacity(max_col + 1);
            for c in 0..=max_col {
                row_data.push(
                    self.workbook
                        .current_sheet()
                        .cells
                        .contains_key(&(row, c))
                        .then(|| self.workbook.current_sheet().get_cell(row, c)),
                );
            }
            rows.push((row, row_data));
        }

        rows.sort_by(|(_, a), (_, b)| {
            let a_val = a.get(col).and_then(|c| c.as_ref()).map(|c| &c.value);
            let b_val = b.get(col).and_then(|c| c.as_ref()).map(|c| &c.value);
            let cmp = match (a_val, b_val) {
                (None, None) => std::cmp::Ordering::Equal,
                (None, Some(_)) => std::cmp::Ordering::Greater,
                (Some(_), None) => std::cmp::Ordering::Less,
                (Some(a), Some(b)) => match (a.parse::<f64>(), b.parse::<f64>()) {
                    (Ok(an), Ok(bn)) => an.partial_cmp(&bn).unwrap_or(std::cmp::Ordering::Equal),
                    _ => a.cmp(b),
                },
            };
            if ascending { cmp } else { cmp.reverse() }
        });

        // Build old->new row mapping so we can rewrite intra-sort formula refs.
        let mut row_map: std::collections::HashMap<usize, usize> =
            std::collections::HashMap::with_capacity(rows.len());
        for (new_row, (old_row, _)) in rows.iter().enumerate() {
            row_map.insert(*old_row, new_row);
        }

        // Rewrite formula references that fall inside the sort range.
        let evaluator = FormulaEvaluator::new(self.workbook.current_sheet());
        let max_row_bound = max_row;
        for (_, row_cells) in rows.iter_mut() {
            for cell_opt in row_cells.iter_mut() {
                if let Some(cell) = cell_opt.as_mut() {
                    if let Some(formula) = cell.formula.clone() {
                        let adjusted = evaluator.remap_row_references(&formula, &row_map, max_row_bound);
                        if adjusted != formula {
                            cell.formula = Some(adjusted);
                        }
                    }
                }
            }
        }

        // Apply, batched for undo + single recalc via set_many.
        let mut batch = Vec::new();
        let mut writes: Vec<(usize, usize, CellData)> = Vec::new();
        let mut clears: Vec<(usize, usize)> = Vec::new();
        for (new_row, (_old, row_data)) in rows.iter().enumerate() {
            for (col_idx, cell_opt) in row_data.iter().enumerate() {
                let old = if self
                    .workbook
                    .current_sheet()
                    .cells
                    .contains_key(&(new_row, col_idx))
                {
                    Some(self.workbook.current_sheet().get_cell(new_row, col_idx))
                } else {
                    None
                };
                let new = cell_opt.clone();
                if old != new {
                    batch.push(UndoAction::CellModified {
                        row: new_row,
                        col: col_idx,
                        old_cell: old,
                        new_cell: new.clone(),
                    });
                    match new {
                        Some(cell) => writes.push((new_row, col_idx, cell)),
                        None => clears.push((new_row, col_idx)),
                    }
                }
            }
        }
        // Clears first (so dependent invalidation happens), then bulk writes.
        for (r, c) in clears {
            self.workbook.current_sheet_mut().clear_cell(r, c);
        }
        if !writes.is_empty() {
            self.workbook.current_sheet_mut().set_many(writes);
        }
        if !batch.is_empty() {
            self.record_action(UndoAction::Batch(batch));
        }
        let dir = if ascending { "ascending" } else { "descending" };
        self.status_message = Some(format!(
            "Sorted by column {} {}",
            crate::domain::Spreadsheet::column_label(col),
            dir
        ));
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
                    _ => {
                        let existing_style = cell.format.as_ref().map(|f| f.style.clone()).unwrap_or_default();
                        Some(CellFormat { number_format: number_format.clone(), style: existing_style })
                    }
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
        let exists = self.workbook.current_sheet().cells.contains_key(&(row, col));
        let mut cell = self.workbook.current_sheet().get_cell(row, col);
        let old_cell = if exists { Some(cell.clone()) } else { None };
        if cell.comment == comment {
            // No change, nothing to record.
            self.status_message = Some("Comment unchanged".to_string());
            return;
        }
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
    /// Uses a cache keyed by the (start, end) range; invalidated on any
    /// cell mutation via `invalidate_stats_cache`.
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

    /// Drop the stats cache. Called when cells change so the next read
    /// recomputes against fresh data.
    fn invalidate_stats_cache(&mut self) {
        self.stats_cache = None;
    }

    /// Detects a pattern (arithmetic sequence, text+number, days/months, formula)
    /// from non-empty cells in the selection and fills the empty cells.
    /// Fill axis is the long axis of the selection; ties fill down.
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
                    spill_anchor: None,
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
            spill_anchor: None,
            }));
        }

        (changes, pattern_desc)
    }

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
                    spill_anchor: None,
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
            spill_anchor: None,
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