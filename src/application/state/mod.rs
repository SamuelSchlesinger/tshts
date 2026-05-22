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
///
/// Each variant co-locates its data with its `apply` / `revert` logic via
/// the impl block below — adding a new mutation type means adding a variant
/// AND its two methods in one place, rather than touching three sites
/// (variant + apply_undo arm + apply_redo arm in App).
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
    /// Sheet-level conditional-format list replaced (sheet_idx, old, new).
    /// Lets `:cf <col> ...` and `:cf clear` participate in undo/redo.
    ConditionalFormatsReplaced {
        sheet_idx: usize,
        old: Vec<crate::domain::ConditionalFormat>,
        new: Vec<crate::domain::ConditionalFormat>,
    },
    /// Row was inserted at `at` on `sheet_idx`. Undo = delete the row.
    /// Used by vim `o`/`O` so an unwanted opened row can be rolled back. We
    /// don't carry cell data because the row is always inserted empty —
    /// any content typed in after gets its own CellModified entry.
    RowInserted {
        sheet_idx: usize,
        at: usize,
    },
    /// Row was deleted at `at` on `sheet_idx`. Undo restores `pre`;
    /// redo re-runs the delete. We snapshot the entire workbook because
    /// cross-sheet refs to the deleted row become `#REF!` and per-sheet
    /// formula shifts can't be undone purely from the deleted row's contents.
    RowDeleted {
        sheet_idx: usize,
        at: usize,
        pre: Box<Workbook>,
    },
    /// Column was inserted at `at` on `sheet_idx`. Undo = delete the column.
    ColInserted {
        sheet_idx: usize,
        at: usize,
    },
    /// Column was deleted. Same reasoning as RowDeleted.
    ColDeleted {
        sheet_idx: usize,
        at: usize,
        pre: Box<Workbook>,
    },
    /// Coarse workbook-level snapshot. Used as an escape hatch for
    /// structural operations (sheet add/delete/rename, freeze, filter,
    /// table create, iterative-calc toggles, named-range edits) where
    /// fine-grained reversal would require its own variant and the
    /// command is rare enough that round-tripping a whole workbook is
    /// acceptable. `pre`/`post` are the workbook state before/after the
    /// command; revert and apply just swap them in.
    WorkbookSnapshot {
        description: String,
        pre: Box<Workbook>,
        post: Box<Workbook>,
    },
}

impl UndoAction {
    /// Roll the workbook back to the state before this action was applied.
    /// Used by `App::undo`.
    pub fn revert(&self, workbook: &mut Workbook) {
        match self {
            UndoAction::CellModified { row, col, old_cell, new_cell: _ } => {
                restore_cell(workbook, *row, *col, old_cell.as_ref());
            }
            UndoAction::Batch(actions) => {
                for a in actions.iter().rev() {
                    a.revert(workbook);
                }
            }
            UndoAction::ConditionalFormatsReplaced { sheet_idx, old, new: _ } => {
                restore_cf(workbook, *sheet_idx, old);
            }
            UndoAction::RowInserted { sheet_idx, at } => {
                if *sheet_idx < workbook.sheets.len() {
                    with_active_sheet(workbook, *sheet_idx, |wb| wb.delete_row_on_active(*at));
                }
            }
            UndoAction::RowDeleted { pre, .. } => {
                restore_workbook(workbook, pre);
            }
            UndoAction::ColInserted { sheet_idx, at } => {
                if *sheet_idx < workbook.sheets.len() {
                    with_active_sheet(workbook, *sheet_idx, |wb| wb.delete_col_on_active(*at));
                }
            }
            UndoAction::ColDeleted { pre, .. } => {
                restore_workbook(workbook, pre);
            }
            UndoAction::WorkbookSnapshot { pre, .. } => {
                restore_workbook(workbook, pre);
            }
        }
    }

    /// Re-apply this action to the workbook. Used by `App::redo`.
    pub fn apply(&self, workbook: &mut Workbook) {
        match self {
            UndoAction::CellModified { row, col, old_cell: _, new_cell } => {
                restore_cell(workbook, *row, *col, new_cell.as_ref());
            }
            UndoAction::Batch(actions) => {
                for a in actions {
                    a.apply(workbook);
                }
            }
            UndoAction::ConditionalFormatsReplaced { sheet_idx, old: _, new } => {
                restore_cf(workbook, *sheet_idx, new);
            }
            UndoAction::RowInserted { sheet_idx, at } => {
                if *sheet_idx < workbook.sheets.len() {
                    with_active_sheet(workbook, *sheet_idx, |wb| wb.insert_row_on_active(*at));
                }
            }
            UndoAction::RowDeleted { sheet_idx, at, .. } => {
                if *sheet_idx < workbook.sheets.len() {
                    with_active_sheet(workbook, *sheet_idx, |wb| wb.delete_row_on_active(*at));
                }
            }
            UndoAction::ColInserted { sheet_idx, at } => {
                if *sheet_idx < workbook.sheets.len() {
                    with_active_sheet(workbook, *sheet_idx, |wb| wb.insert_col_on_active(*at));
                }
            }
            UndoAction::ColDeleted { sheet_idx, at, .. } => {
                if *sheet_idx < workbook.sheets.len() {
                    with_active_sheet(workbook, *sheet_idx, |wb| wb.delete_col_on_active(*at));
                }
            }
            UndoAction::WorkbookSnapshot { post, .. } => {
                restore_workbook(workbook, post);
            }
        }
    }
}

/// Replace `workbook` with `pre`'s contents (deep clone). Used by
/// RowDeleted/ColDeleted undo where structural snapshots are the most
/// reliable way to roll back.
fn restore_workbook(workbook: &mut Workbook, pre: &Workbook) {
    *workbook = (*pre).clone();
    workbook.rebuild_cross_sheet_deps();
    for sheet in &mut workbook.sheets {
        sheet.resweep_all_spills();
    }
}

/// Restore `(row, col)` to `data` (or clear if `None`), then propagate.
/// Shared by `apply` and `revert` for `CellModified`.
fn restore_cell(workbook: &mut Workbook, row: usize, col: usize, data: Option<&CellData>) {
    match data {
        Some(d) => workbook.current_sheet_mut().set_cell(row, col, d.clone()),
        None => workbook.current_sheet_mut().clear_cell(row, col),
    }
    workbook.propagate_active_cell(row, col);
}

fn restore_cf(workbook: &mut Workbook, sheet_idx: usize, rules: &[crate::domain::ConditionalFormat]) {
    if let Some(sheet) = workbook.sheets.get_mut(sheet_idx) {
        sheet.conditional_formats = rules.to_vec();
        sheet.cf_cache.borrow_mut().clear();
    }
}

/// Run `f` with `sheet_idx` as the active sheet, restoring the prior active
/// sheet afterward. Lets undo/redo operations that take an explicit
/// sheet_idx reuse the workbook's `*_on_active` family without a permanent
/// active-sheet switch surfacing to the UI.
fn with_active_sheet<F: FnOnce(&mut Workbook)>(workbook: &mut Workbook, sheet_idx: usize, f: F) {
    let prior = workbook.active_sheet;
    workbook.active_sheet = sheet_idx;
    f(workbook);
    workbook.active_sheet = prior;
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
                    FormulaEvaluator::for_workbook(&wb_snapshot, snap_sheet, &names);
                let value = evaluator.evaluate_formula(&formula);
                let mut cd = sheet.get_cell(row, col);
                cd.value = value;
                sheet.cells.insert((row, col), cd);
            }
        }
        // The cross-sheet dep graph is built off formulas; rebuild it after
        // a recalc so any prior drift (e.g. graph populated before some
        // sheets were loaded) is corrected.
        self.workbook.rebuild_cross_sheet_deps();
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
            action.revert(&mut self.workbook);
            self.redo_stack.push_back(action);
            self.dirty = true;
        }
    }

    /// Cross-sheet propagation hook for cell mutations that don't go through
    /// `Workbook::write_cells_on_active`. Forwards to the workbook, which
    /// owns the actual dep registration + propagation logic.
    pub(crate) fn propagate_cell_change(&mut self, row: usize, col: usize) {
        self.workbook.propagate_active_cell(row, col);
    }

    pub fn redo(&mut self) {
        if let Some(action) = self.redo_stack.pop_back() {
            action.apply(&mut self.workbook);
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
        if start == end {
            return None; // Single cell, no stats
        }
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

    #[test]
    fn test_terminal_color_from_name() {
        assert_eq!(TerminalColor::from_name("red"), Some(TerminalColor::Red));
        assert_eq!(TerminalColor::from_name("Blue"), Some(TerminalColor::Blue));
        assert_eq!(TerminalColor::from_name("lightgreen"), Some(TerminalColor::LightGreen));
        assert_eq!(TerminalColor::from_name("CYAN"), Some(TerminalColor::Cyan));
        assert_eq!(TerminalColor::from_name("invalid"), None);
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

}
