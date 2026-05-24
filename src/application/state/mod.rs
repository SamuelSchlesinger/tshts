//! Application state management for the terminal spreadsheet.
//!
//! This module contains the main application state and mode management
//! for the terminal user interface.

use crate::domain::{Spreadsheet, Workbook, CellData, CellFormat, NumberFormat, TerminalColor, FormulaEvaluator};
use std::collections::{HashMap, HashSet, VecDeque};


mod matcher;
mod undo;
pub use matcher::TextMatcher;
pub use undo::UndoAction;

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

/// Selection-stats cache payload: `(start_pos, end_pos, stats)` where
/// `stats` is `(sum, average, count)` once computed, or `None` if the
/// selected range contains no numeric cells.
pub type StatsCache = ((usize, usize), (usize, usize), Option<(f64, f64, usize)>);

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
    /// Search results as (row, col) coordinates, in row-major order so
    /// next/prev navigation is deterministic.
    pub search_results: Vec<(usize, usize)>,
    /// HashSet mirror of `search_results` for O(1) lookup during render.
    /// A 1k-hit search rendered to a 900-cell viewport at 10fps used to
    /// burn ~9M Vec::contains comparisons per second; this drops it to
    /// hash-table lookups. Kept in sync wherever `search_results` is
    /// mutated.
    pub search_results_set: std::collections::HashSet<(usize, usize)>,
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
    pub stats_cache: Option<StatsCache>,
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
    /// True when the current Editing session was entered via `o` / `O`
    /// (which inserts a fresh row) and the user hasn't committed yet. If
    /// they Esc-cancel from this state, the row insertion is rolled back
    /// too — otherwise `o<Esc>` would leave a phantom empty row. Cleared
    /// on every fresh `start_editing` and on every Editing-mode exit.
    pub pending_open_row: bool,
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
    Up,
    Right,
    Left,
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
            search_results_set: std::collections::HashSet::new(),
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
            pending_open_row: false,
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
        self.search_results_set.clear();
        self.search_result_index = 0;
        self.status_message = None;
        self.chart_popup = None;
        // Vim pending operator/count/awaiting-g would otherwise leak across
        // Esc and cause the next motion key to act as `<op><motion>`.
        self.vim_reset_pending();
        // Filters are long-lived view state — Esc shouldn't kill them.
        // Use `:filter clear` or the relevant command to clear explicitly.
        // Hidden rows persist alongside filter; same reasoning.
    }

    pub fn recalc_all(&mut self) {
        // PR 3/4 path: mark every formula cell dirty and let the
        // graph-driven executor walk the topo levels. The dirty mark
        // covers the case where the user manually invoked :recalc — we
        // want to recompute everything even if mutation paths didn't
        // touch a cell since last recalc (e.g. RAND should re-roll;
        // NOW should pick up the current time).
        for (idx, sheet) in self.workbook.sheets.iter_mut().enumerate() {
            // Keep the per-sheet dep graph in sync — the legacy cross-
            // sheet engine still reads it, and `build_dep_graph_from_scratch`
            // uses the per-sheet refs as input.
            // Mark every formula cell on this sheet dirty.
            let name = self.workbook.sheet_names[idx].clone();
            for (&(r, c), cd) in &sheet.cells {
                if cd.formula.is_some() {
                    self.workbook.dirty.insert((name.clone(), r, c));
                }
            }
        }
        // Rebuild the workbook dep graph in case it's stale (e.g. a
        // load path that didn't pre-build, or a structural edit that
        // invalidated entries).
        self.workbook.build_dep_graph_from_scratch();
        self.workbook.rebuild_cross_sheet_deps();
        // The unified executor drains dirty and walks topo levels.
        // Surface pass-level errors (e.g. iterative-calc non-convergence)
        // via status message — values are still committed best-effort.
        match self.workbook.recalc_via_graph_result() {
            Ok(()) => {
                self.status_message = Some("Recalculated all formulas".to_string());
            }
            Err(e) => {
                self.status_message = Some(format!("Recalc: {}", e));
            }
        }
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

    /// Run a structural mutation with a coarse snapshot-based undo entry.
    /// `description` shows in :history-style listings (currently only used
    /// for debugging). The snapshot pair is captured around the closure;
    /// no-op mutations don't push an undo entry.
    ///
    /// Use for ops that have no cheap fine-grained inverse: :sheet add/
    /// delete, :name/:unname, :freeze, :filter, :table create, :iterative
    /// toggles. The clone is bounded by the workbook size (~10 MB for
    /// 100k-cell workbooks); these commands run at most a few times per
    /// session so the cost is acceptable.
    pub(crate) fn with_snapshot_undo<F: FnOnce(&mut App)>(&mut self, description: &str, f: F) {
        let pre = Box::new(self.workbook.clone());
        f(self);
        let post = Box::new(self.workbook.clone());
        self.record_action(UndoAction::WorkbookSnapshot {
            description: description.to_string(),
            pre,
            post,
        });
    }

    pub fn undo(&mut self) {
        if let Some(action) = self.undo_stack.pop_back() {
            let label = action.description();
            match action.revert(&mut self.workbook) {
                Ok(()) => {
                    self.status_message = Some(format!("Undo: {}", label));
                }
                Err(e) => {
                    self.status_message = Some(format!("Undo {}: {}", label, e));
                }
            }
            self.redo_stack.push_back(action);
            self.dirty = true;
        }
    }

    /// Propagation hook for cell mutations that don't go through
    /// `Workbook::write_cells_on_active`. With the unified graph executor
    /// as the single source of truth, propagation is just "mark dirty
    /// and recalc," which `set_cell_on_active` already does. This helper
    /// stays for callers that wrote a cell via some bespoke path and
    /// then explicitly want to flush; today it just runs a recalc.
    pub(crate) fn propagate_cell_change(&mut self, row: usize, col: usize) {
        // Mark the cell dirty (the bespoke caller may not have done it)
        // and run the unified recalc.
        let sheet_name = self.workbook.sheet_names[self.workbook.active_sheet].clone();
        self.workbook.mark_dirty(&sheet_name, row, col);
        let _ = self.workbook.recalc_via_graph_result();
    }

    pub fn redo(&mut self) {
        if let Some(action) = self.redo_stack.pop_back() {
            let label = action.description();
            match action.apply(&mut self.workbook) {
                Ok(()) => {
                    self.status_message = Some(format!("Redo: {}", label));
                }
                Err(e) => {
                    self.status_message = Some(format!("Redo {}: {}", label, e));
                }
            }
            self.undo_stack.push_back(action);
            self.dirty = true;
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
        // Named-range bounds shifted; any formula referencing this table
        // by Table[Col] sees a different range. Conservative dirty mark
        // — table-extends are rare (only fires when typing into the row
        // immediately below an existing table).
        self.workbook.mark_all_formula_cells_dirty();
    }

    pub fn set_cell_with_undo(&mut self, row: usize, col: usize, new_data: CellData) {
        // Get the old cell data
        let old_cell = if self.workbook.current_sheet().cells.contains_key(&(row, col)) {
            Some(self.workbook.current_sheet().get_cell(row, col))
        } else {
            None
        };

        // No-op suppression: if the new cell is byte-identical to the
        // existing one (same value/formula/format/comment/spill_anchor),
        // there's nothing to record. Without this, pressing Enter on an
        // unchanged cell grew the undo stack and (more importantly) marked
        // the workbook dirty, triggering spurious autosave activity.
        if let Some(ref existing) = old_cell
            && *existing == new_data
        {
            return;
        }

        // Record the action
        let action = UndoAction::CellModified {
            row,
            col,
            old_cell,
            new_cell: Some(new_data.clone()),
        };
        self.record_action(action);

        // Apply the change via the workbook-aware mutation path so the
        // same-sheet recalc cascade sees cross-sheet refs correctly; then
        // auto-extend any table that this row sits below.
        self.workbook.set_cell_on_active(row, col, new_data);
        self.maybe_extend_table(row, col);

        // Workbook-level cross-sheet dep maintenance: register this cell's
        // (new) dependencies, then propagate changes to anything that
        // depended on its old value.
        self.propagate_cell_change(row, col);
    }

    /// Write many cells as a single undo-able action. Use for bulk writes
    /// like pivot table generation where per-cell undo entries would force
    /// the user to press `u` dozens of times to back out one logical action.
    pub fn set_many_with_undo(&mut self, cells: Vec<(usize, usize, CellData)>) {
        if cells.is_empty() {
            return;
        }
        // Snapshot pre-images for undo before any write lands.
        let batch: Vec<UndoAction> = cells
            .iter()
            .map(|(row, col, new_data)| {
                let old_cell = if self.workbook.current_sheet().cells.contains_key(&(*row, *col)) {
                    Some(self.workbook.current_sheet().get_cell(*row, *col))
                } else {
                    None
                };
                UndoAction::CellModified {
                    row: *row,
                    col: *col,
                    old_cell,
                    new_cell: Some(new_data.clone()),
                }
            })
            .collect();
        // Single workbook API call handles same-sheet recalc + cross-sheet
        // propagation for the whole batch.
        self.workbook.write_cells_on_active(cells);
        self.record_action(UndoAction::Batch(batch));
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
        
        // Apply the change via the workbook-aware mutation path.
        self.workbook.clear_cell_on_active(row, col);

        // Cross-sheet maintenance.
        self.propagate_cell_change(row, col);
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
        // Always compute stats — even a 1x1 selection benefits from
        // surfacing the cell's numeric value (the value the user is
        // standing on, which the formula bar otherwise hides behind the
        // formula text). The scenario test framework also relies on
        // SUM= being published for single-cell selections to read
        // computed values back out of tshts.
        if let Some((cs, ce, v)) = &self.stats_cache
            && *cs == start && *ce == end {
                return *v;
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
mod tests;
