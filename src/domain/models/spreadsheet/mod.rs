//! Submodule of `models` — see models/mod.rs.

#![allow(unused_imports)]
use serde::{Deserialize, Serialize};
use std::collections::HashMap;
use super::*;

#[derive(Debug, Serialize, Deserialize)]
pub struct Spreadsheet {
    /// Cell data stored as a sparse matrix using (row, col) coordinates
    #[serde(serialize_with = "serialize_cells", deserialize_with = "deserialize_cells")]
    pub cells: HashMap<(usize, usize), CellData>,
    /// Maximum number of rows in the spreadsheet
    pub rows: usize,
    /// Maximum number of columns in the spreadsheet
    pub cols: usize,
    /// Custom column widths for specific columns
    pub column_widths: HashMap<usize, usize>,
    /// Default width for columns without custom widths
    pub default_column_width: usize,
    /// Named ranges resolvable inside formulas. Map keys are uppercase by
    /// convention. Synced from `Workbook::named_ranges` so per-sheet recalc
    /// can resolve names without needing workbook access.
    #[serde(default)]
    pub named_ranges: HashMap<String, String>,
    /// Conditional-formatting rules applied at render time. A rule fires for
    /// a cell when the cell is in `column_range` and the `predicate` formula
    /// (with `_` bound to the cell's value) evaluates truthy.
    #[serde(default)]
    pub conditional_formats: Vec<ConditionalFormat>,
    /// Tables defined on this sheet. Structured refs like `Table1[Col1]`
    /// resolve via this list.
    #[serde(default)]
    pub tables: Vec<Table>,
    /// Conditional-format style cache, keyed by (row, col). Populated lazily
    /// on first lookup; invalidated wholesale on any cell mutation or rule
    /// change. Refcell so `conditional_style_for(&self)` can write.
    #[serde(skip)]
    pub cf_cache: std::sync::Mutex<HashMap<(usize, usize), Option<CellStyle>>>,
    /// Persistent view state for this sheet. Save/load round-trips freezes,
    /// hidden rows/cols, filter criteria, and per-column data-validation
    /// rules so reopening a workbook restores the user's full workspace.
    #[serde(default)]
    pub view_state: SheetViewState,
}

impl Clone for Spreadsheet {
    /// `Mutex<HashMap>` doesn't implement `Clone`. We hand-roll the impl
    /// so cloning a Spreadsheet snapshots the cells and other persistent
    /// state but DROPS the conditional-format cache — the cache is a
    /// render-time memoization that the clone will re-populate as
    /// needed. Dropping it also keeps the Sync property: the snapshot
    /// can be shared across rayon workers without contending on the
    /// cache's mutex (workers don't render).
    fn clone(&self) -> Self {
        Self {
            cells: self.cells.clone(),
            rows: self.rows,
            cols: self.cols,
            column_widths: self.column_widths.clone(),
            default_column_width: self.default_column_width,
            named_ranges: self.named_ranges.clone(),
            conditional_formats: self.conditional_formats.clone(),
            tables: self.tables.clone(),
            cf_cache: std::sync::Mutex::new(HashMap::new()),
            view_state: self.view_state.clone(),
        }
    }
}

/// Persistent per-sheet view state. The App keeps these on its struct for
/// runtime convenience, but they're synced to/from this on save and load.
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct SheetViewState {
    /// Number of rows frozen at the top.
    #[serde(default)]
    pub frozen_rows: usize,
    /// Number of columns frozen on the left.
    #[serde(default)]
    pub frozen_cols: usize,
    /// Row indices hidden via `:hide` / filter. Stored as a sorted Vec so the
    /// serialized form is stable across runs (HashSet iteration order is not).
    #[serde(default)]
    pub hidden_rows: Vec<usize>,
    /// Column indices hidden via `:hide col E`.
    #[serde(default)]
    pub hidden_cols: Vec<usize>,
    /// Active filter (column, criteria). Pair of None when no filter set.
    #[serde(default)]
    pub filter_column: Option<usize>,
    #[serde(default)]
    pub filter_value: Option<String>,
    /// Per-column data validation predicates.
    #[serde(default)]
    pub validations: HashMap<usize, String>,
}

/// A named rectangular region with column headers. Auto-expands when rows
/// are added immediately below.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct Table {
    pub name: String,
    /// Top-left and bottom-right cells (inclusive).
    pub top_row: usize,
    pub left_col: usize,
    pub bottom_row: usize,
    pub right_col: usize,
    /// Column header names (one per column in the range).
    pub headers: Vec<String>,
}

/// Conditional-formatting rule. The predicate is evaluated per cell with the
/// cell's value substituted for the literal token `_` (e.g. `_ > 100` or
/// `LEN(_) = 0`). Only the style is applied — number format is untouched.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct ConditionalFormat {
    /// Column index this rule applies to.
    pub column: usize,
    /// Predicate formula, sans leading `=`. Use `_` to refer to the cell value.
    pub predicate: String,
    /// Style applied when predicate is truthy.
    pub style: CellStyle,
}

impl Default for Spreadsheet {
    fn default() -> Self {
        Self {
            cells: HashMap::new(),
            rows: 100,
            cols: 26,
            column_widths: HashMap::new(),
            default_column_width: 8,
            named_ranges: HashMap::new(),
            conditional_formats: Vec::new(),
            tables: Vec::new(),
            cf_cache: std::sync::Mutex::new(HashMap::new()),
            view_state: SheetViewState::default(),
        }
    }
}

impl Spreadsheet {
    /// Returns the cell at `(row, col)`, or a default empty cell if unset.
    ///
    /// ```
    /// use tshts::domain::Spreadsheet;
    ///
    /// let sheet = Spreadsheet::default();
    /// assert!(sheet.get_cell(0, 0).value.is_empty());
    /// ```
    pub fn get_cell(&self, row: usize, col: usize) -> CellData {
        self.cells.get(&(row, col)).cloned().unwrap_or_default()
    }

    /// Low-level cell write without dependency tracking or recalc. Use `set_cell`.
    fn set_cell_internal(&mut self, row: usize, col: usize, data: CellData) {
        self.cells.insert((row, col), data.clone());
        // Any cell mutation could change a CF predicate's evaluation
        // (because predicates can reference other cells too). Drop the
        // whole cache; reads will repopulate lazily.
        self.cf_cache.lock().unwrap().clear();
    }

    /// Writes a cell. Sweeps any spill ghosts owned by the previous version
    /// of this cell, writes `data` as-is (caller is responsible for the
    /// `value` field — typically pre-evaluated for formulas), and re-spills
    /// if the new formula produces an array.
    ///
    /// **Does not propagate to dependents.** That's the workbook executor's
    /// job — call `Workbook::set_cell_on_active` (or any other workbook-level
    /// mutation API) which marks the cell dirty and routes recalc through
    /// `Workbook::recalc_via_graph_result()`. Calling `Spreadsheet::set_cell`
    /// directly is appropriate for low-level test fixtures and intra-workbook
    /// machinery (e.g. cyclic-pass loops that have their own propagation
    /// discipline) but bypasses the normal recompute pipeline.
    ///
    /// `pub(crate)` because external consumers should use the workbook
    /// mutators (`Workbook::set_cell_on_active` etc.) which run a full
    /// recalc; calling Spreadsheet::set_cell directly leaves dependents
    /// stale and is only safe from inside the calc engine.
    pub(crate) fn set_cell(&mut self, row: usize, col: usize, data: CellData) {
        self.sweep_spill_ghosts_for(row, col);
        self.set_cell_internal(row, col, data);
        self.maybe_spill(row, col);
    }

    /// Drop every spill ghost on the sheet, then re-evaluate `maybe_spill`
    /// for each cell that has a formula. Called after row/col insert/delete
    /// because structural shifts can leave ghosts pointing to a moved or
    /// vanished anchor, or leave gaps between an anchor and its surviving
    /// ghosts. Cheap on TUI-scale sheets; insert/delete are uncommon.
    pub(crate) fn resweep_all_spills(&mut self) {
        let ghosts: Vec<(usize, usize)> = self
            .cells
            .iter()
            .filter(|(_, cd)| cd.spill_anchor.is_some())
            .map(|(&pos, _)| pos)
            .collect();
        for pos in ghosts {
            self.cells.remove(&pos);
        }
        // Sort anchors so two formulas competing for the same spill region
        // resolve identically across runs. HashMap iteration order is not
        // stable, which would otherwise let `=ARRAY()` at A1 vs B1 produce
        // different #SPILL!/winner outcomes between sessions on the same
        // saved file.
        let mut anchors: Vec<(usize, usize)> = self
            .cells
            .iter()
            .filter(|(_, cd)| cd.formula.is_some())
            .map(|(&pos, _)| pos)
            .collect();
        anchors.sort();
        for (r, c) in anchors {
            self.maybe_spill(r, c);
        }
    }

    /// Clear any spill ghosts whose anchor is the given cell. Called before
    /// the cell is rewritten so we don't leak stale ghosts when the
    /// formula's array shape shrinks (or disappears entirely).
    pub(crate) fn sweep_spill_ghosts_for(&mut self, anchor_row: usize, anchor_col: usize) {
        let targets: Vec<(usize, usize)> = self
            .cells
            .iter()
            .filter_map(|(&pos, cd)| {
                if cd.spill_anchor == Some((anchor_row, anchor_col)) {
                    Some(pos)
                } else {
                    None
                }
            })
            .collect();
        for pos in targets {
            self.cells.remove(&pos);
        }
    }

    /// After a cell write, if its formula evaluated to a multi-cell Array,
    /// expand the array into ghost cells. If any target overlaps a
    /// non-empty cell, set the anchor to `#SPILL!`.
    pub(crate) fn maybe_spill(&mut self, row: usize, col: usize) {
        use crate::domain::parser::{Parser, Value, ExpressionEvaluator, FunctionRegistry};
        let formula = match self
            .cells
            .get(&(row, col))
            .and_then(|cd| cd.formula.clone())
        {
            Some(f) if f.starts_with('=') => f,
            _ => return,
        };
        let expr_src = &formula[1..];
        let ast = match Parser::new(expr_src).and_then(|mut p| p.parse()) {
            Ok(a) => a,
            Err(_) => return,
        };
        let registry = FunctionRegistry::shared_builtin();
        let names = if self.named_ranges.is_empty() {
            None
        } else {
            Some(&self.named_ranges)
        };
        let value = super::workbook::with_workbook_context(|wb| {
            let evaluator = ExpressionEvaluator::new(self, &registry, names, wb);
            evaluator.evaluate(&ast).ok()
        });
        let value = match value {
            Some(v) => v,
            None => return,
        };
        // Determine shape. List is 1×N (single row); Array carries explicit
        // shape; scalars don't spill.
        let (rows, cols, data): (usize, usize, Vec<Value>) = match value {
            Value::Array { rows, cols, data } => (rows, cols, data),
            Value::List(items) if items.len() > 1 => (items.len(), 1, items),
            _ => return,
        };
        if rows == 1 && cols == 1 {
            return;
        }
        // Bounds check.
        if row + rows > self.rows || col + cols > self.cols {
            // Doesn't fit — write #SPILL! at the anchor.
            if let Some(cd) = self.cells.get_mut(&(row, col)) {
                cd.value = "#SPILL!".to_string();
            }
            return;
        }
        // Collision check: any target (other than the anchor itself) that
        // already has a non-ghost value blocks the spill.
        for r in row..row + rows {
            for c in col..col + cols {
                if r == row && c == col {
                    continue;
                }
                if let Some(existing) = self.cells.get(&(r, c))
                    && existing.spill_anchor != Some((row, col))
                        && (!existing.value.is_empty() || existing.formula.is_some())
                    {
                        if let Some(cd) = self.cells.get_mut(&(row, col)) {
                            cd.value = "#SPILL!".to_string();
                        }
                        return;
                    }
            }
        }
        // Write the anchor's value (the first element of the array) and the
        // ghosts for everything else.
        for r in row..row + rows {
            for c in col..col + cols {
                let idx = (r - row) * cols + (c - col);
                let v = data.get(idx).map(|v| v.to_string()).unwrap_or_default();
                if r == row && c == col {
                    if let Some(cd) = self.cells.get_mut(&(r, c)) {
                        cd.value = v;
                    }
                } else {
                    let ghost = CellData {
                        value: v,
                        formula: None,
                        format: None,
                        comment: None,
                        spill_anchor: Some((row, col)),
                    };
                    self.cells.insert((r, c), ghost);
                }
            }
        }
    }

    /// Clears the cell. Sweeps any owned spill ghosts. Does not propagate
    /// to dependents — that's the workbook executor's job (use
    /// `Workbook::clear_cell_on_active`).
    pub fn clear_cell(&mut self, row: usize, col: usize) {
        self.sweep_spill_ghosts_for(row, col);
        self.cells.remove(&(row, col));
    }

    /// Bulk-clear many cells. Pure write operation; the workbook executor
    /// handles dep propagation. Symmetric counterpart of `set_many`.
    pub fn clear_many(&mut self, positions: Vec<(usize, usize)>) {
        for (row, col) in positions {
            self.sweep_spill_ghosts_for(row, col);
            self.cells.remove(&(row, col));
        }
    }

    /// Bulk-set many cells. Pure write operation; the workbook executor
    /// handles dep propagation. Used by sort/autofill/paste — write the
    /// batch, mark dirty at the workbook level, then run a single
    /// `recalc_via_graph_result()` over the union of dependents.
    pub fn set_many(&mut self, updates: Vec<(usize, usize, CellData)>) {
        for (row, col, data) in updates {
            self.sweep_spill_ghosts_for(row, col);
            self.set_cell_internal(row, col, data);
            self.maybe_spill(row, col);
        }
    }

    /// Force a single cell to re-evaluate its formula and refresh `value`.
    /// Used by the workbook executor (per-cell eval inside a recalc pass)
    /// and by cross-sheet ref-rewrite paths (`Workbook::remove_sheet` etc.)
    /// where the formula text changed without going through `set_cell`.
    /// Honors the thread-local workbook context set by
    /// `Workbook::with_recalc_context` so cross-sheet refs resolve.
    pub fn refresh_cell_value(&mut self, row: usize, col: usize) {
        let cell_pos = (row, col);
        if let Some(cell) = self.cells.get(&cell_pos).cloned()
            && let Some(ref formula) = cell.formula
        {
            use crate::domain::services::FormulaEvaluator;
            let new_value = super::workbook::with_workbook_context(|wb_opt| {
                let mut ev = FormulaEvaluator::new(self);
                if !self.named_ranges.is_empty() {
                    ev = ev.with_names(&self.named_ranges);
                }
                if let Some(wb) = wb_opt {
                    ev = ev.with_workbook(wb);
                }
                ev.evaluate_formula(formula)
            });
            let mut updated_cell = cell;
            updated_cell.value = new_value;
            self.set_cell_internal(row, col, updated_cell);
        }
    }

    /// Converts a zero-based column index to an Excel-style label (A, Z, AA, ...).
    ///
    /// ```
    /// use tshts::domain::Spreadsheet;
    /// assert_eq!(Spreadsheet::column_label(0), "A");
    /// assert_eq!(Spreadsheet::column_label(26), "AA");
    /// ```
    pub fn column_label(col: usize) -> String {
        let mut result = String::new();
        let mut c = col;
        loop {
            result = char::from(b'A' + (c % 26) as u8).to_string() + &result;
            if c < 26 {
                break;
            }
            c = c / 26 - 1;
        }
        result
    }

    /// Parses a column label (like "A", "B", "AA") into a zero-based column index.
    pub fn parse_column_label(label: &str) -> Option<usize> {
        let label = label.to_uppercase();
        if label.is_empty() || !label.chars().all(|c| c.is_ascii_alphabetic()) {
            return None;
        }
        let mut col = 0usize;
        for ch in label.chars() {
            col = col * 26 + (ch as usize - 'A' as usize + 1);
        }
        Some(col - 1)
    }

    /// Parses Excel-style references like "A1" or "AA123" into `(row, col)`.
    ///
    /// ```
    /// use tshts::domain::Spreadsheet;
    /// assert_eq!(Spreadsheet::parse_cell_reference("A1"), Some((0, 0)));
    /// assert_eq!(Spreadsheet::parse_cell_reference("invalid"), None);
    /// ```
    pub fn parse_cell_reference(cell_ref: &str) -> Option<(usize, usize)> {
        Self::parse_cell_reference_with_flags(cell_ref).map(|(r, c, _, _)| (r, c))
    }

    /// Parse a possibly sheet-qualified cell reference: `Sheet2!A1`,
    /// `'Some Sheet'!B5`, `$A$1`, etc. Returns (sheet?, row, col, abs_row, abs_col).
    pub fn parse_qualified_reference(cell_ref: &str) -> Option<(Option<String>, usize, usize, bool, bool)> {
        // Find the `!` that separates the sheet name from the cell ref,
        // skipping any `!` inside a `'...'` quoted sheet name. Previously
        // `rfind('!')` would slice the wrong split point for names like
        // `'oh!no'!A1`.
        let bang_pos = if let Some(b) = cell_ref.strip_prefix('\'') {
            // Find the closing quote, skipping `''` escapes.
            let mut end = None;
            let bytes = b.as_bytes();
            let mut i = 0;
            while i < bytes.len() {
                if bytes[i] == b'\'' {
                    if i + 1 < bytes.len() && bytes[i + 1] == b'\'' {
                        i += 2;
                        continue;
                    }
                    end = Some(i);
                    break;
                }
                i += 1;
            }
            // Convert relative position back to absolute (1 = opening quote).
            end.map(|e| 1 + e + 1)
                .filter(|&after_quote| cell_ref[after_quote..].starts_with('!'))
        } else {
            cell_ref.find('!')
        };
        if let Some(idx) = bang_pos {
            let (sheet, rest) = cell_ref.split_at(idx);
            let cell_part = &rest[1..];
            let sheet_name = if sheet.starts_with('\'') && sheet.ends_with('\'') && sheet.len() >= 2 {
                sheet[1..sheet.len() - 1].replace("''", "'")
            } else {
                sheet.to_string()
            };
            let (row, col, ar, ac) = Self::parse_cell_reference_with_flags(cell_part)?;
            Some((Some(sheet_name), row, col, ar, ac))
        } else {
            let (row, col, ar, ac) = Self::parse_cell_reference_with_flags(cell_ref)?;
            Some((None, row, col, ar, ac))
        }
    }

    /// Detect a parser-emitted 3-D range marker `Sheet1..Sheet3!A1`.
    /// Returns (start_sheet, end_sheet, cell_part).
    pub fn parse_three_d_marker(s: &str) -> Option<(String, String, String)> {
        let bang = s.rfind('!')?;
        let (sheets, rest) = s.split_at(bang);
        let cell = &rest[1..];
        let mid = sheets.find("..")?;
        let (s1, s2) = sheets.split_at(mid);
        let s2 = &s2[2..];
        Some((s1.to_string(), s2.to_string(), cell.to_string()))
    }

    /// Like `parse_cell_reference` but also returns whether the row and column
    /// are absolute (`$`-prefixed). Supports `$A$1`, `$A1`, `A$1`, `A1`.
    pub fn parse_cell_reference_with_flags(cell_ref: &str) -> Option<(usize, usize, bool, bool)> {
        if cell_ref.is_empty() {
            return None;
        }
        let mut chars = cell_ref.chars().peekable();
        let abs_col = if chars.peek() == Some(&'$') {
            chars.next();
            true
        } else {
            false
        };
        let mut col_str = String::new();
        while let Some(&ch) = chars.peek() {
            if ch.is_ascii_alphabetic() {
                col_str.push(ch.to_ascii_uppercase());
                chars.next();
            } else {
                break;
            }
        }
        let abs_row = if chars.peek() == Some(&'$') {
            chars.next();
            true
        } else {
            false
        };
        let mut row_str = String::new();
        while let Some(&ch) = chars.peek() {
            if ch.is_ascii_digit() {
                row_str.push(ch);
                chars.next();
            } else {
                return None; // trailing garbage
            }
        }
        if col_str.is_empty() || row_str.is_empty() {
            return None;
        }
        let col = Self::column_str_to_index(&col_str)?;
        let row = row_str.parse::<usize>().ok()?.checked_sub(1)?;
        Some((row, col, abs_row, abs_col))
    }

    /// Format a cell reference with optional `$` markers.
    pub fn format_cell_reference(row: usize, col: usize, abs_row: bool, abs_col: bool) -> String {
        let mut s = String::new();
        if abs_col {
            s.push('$');
        }
        s.push_str(&Self::column_label(col));
        if abs_row {
            s.push('$');
        }
        s.push_str(&(row + 1).to_string());
        s
    }
    
    fn column_str_to_index(col_str: &str) -> Option<usize> {
        if col_str.is_empty() {
            return None;
        }
        
        let mut result = 0;
        for ch in col_str.chars() {
            if !ch.is_ascii_alphabetic() {
                return None;
            }
            result = result * 26 + (ch as usize - 'A' as usize + 1);
        }
        Some(result - 1)
    }

    /// Custom width if set, else the spreadsheet's default.
    pub fn get_column_width(&self, col: usize) -> usize {
        self.column_widths.get(&col).copied().unwrap_or(self.default_column_width)
    }

    pub fn set_column_width(&mut self, col: usize, width: usize) {
        self.column_widths.insert(col, width);
    }

    /// Resizes a column to fit its content (clamped to 3..=50 chars).
    #[allow(dead_code)]
    pub fn auto_resize_column(&mut self, col: usize) {
        self.auto_resize_column_with_cap(col, 50);
    }

    /// Auto-resize with a caller-supplied maximum (e.g. derived from the
    /// viewport so a single wide cell can't push other columns off-screen).
    pub fn auto_resize_column_with_cap(&mut self, col: usize, max_cap: usize) {
        use unicode_width::UnicodeWidthStr;
        use super::style::format_cell_value;
        let mut max_width = Self::column_label(col).width();

        for (&(_, c), cell) in &self.cells {
            if c == col {
                // The rendered width drives the auto-fit, not the raw
                // value's width. A cell with raw "9876.54" and currency
                // format displays as "$9,876.54" (9 chars vs 7); sizing
                // by raw width would clip the displayed text.
                let displayed = if let Some(ref fmt) = cell.format {
                    format_cell_value(&cell.value, fmt)
                } else {
                    cell.value.clone()
                };
                let value_width = displayed.width();
                let formula_width = cell.formula.as_ref().map(|f| f.width()).unwrap_or(0);
                let content_width = value_width.max(formula_width);
                max_width = max_width.max(content_width);
            }
        }

        max_width = max_width.max(3).min(max_cap.max(3));
        self.set_column_width(col, max_width);
    }

    /// Automatically resizes all columns to fit their content.
    ///
    /// Single pass over the sparse `cells` map (O(N) total) instead of
    /// O(cols × cells) — the per-column loop iterated the entire HashMap.
    pub fn auto_resize_all_columns(&mut self) {
        use unicode_width::UnicodeWidthStr;
        use super::style::format_cell_value;
        let mut widths: HashMap<usize, usize> = HashMap::with_capacity(self.cols);
        for col in 0..self.cols {
            widths.insert(col, Self::column_label(col).width());
        }
        for (&(_, c), cell) in &self.cells {
            // Width the user sees, not raw — see auto_resize_column_with_cap.
            let displayed = if let Some(ref fmt) = cell.format {
                format_cell_value(&cell.value, fmt)
            } else {
                cell.value.clone()
            };
            let value_width = displayed.width();
            let formula_width = cell.formula.as_ref().map(|f| f.width()).unwrap_or(0);
            let content_width = value_width.max(formula_width);
            let entry = widths.entry(c).or_insert(3);
            *entry = (*entry).max(content_width);
        }
        for (col, mut w) in widths {
            w = w.clamp(3, 50);
            self.set_column_width(col, w);
        }
    }

    /// Inserts a row at the given index, shifting all rows at or below down by 1.
    /// Updates formula references accordingly.
    pub fn insert_row(&mut self, at: usize) {
        use crate::domain::services::FormulaEvaluator;

        // Collect all cells, sorted by row descending so we shift from bottom up
        let mut entries: Vec<((usize, usize), CellData)> = self.cells.drain().collect();
        entries.sort_by_key(|((row, _), _)| std::cmp::Reverse(*row));

        let mut new_cells = std::collections::HashMap::new();
        for ((row, col), cell) in entries {
            if row >= at {
                new_cells.insert((row + 1, col), cell);
            } else {
                new_cells.insert((row, col), cell);
            }
        }

        self.cells = new_cells;
        self.rows += 1;

        // Adjust formula references in all cells
        let updates = {
            let cells_with_formulas: Vec<_> = self.cells.iter()
                .filter_map(|(&(r, c), cell)| cell.formula.as_ref().map(|f| (r, c, f.clone())))
                .collect();
            let evaluator = FormulaEvaluator::new(self);
            let mut updates = Vec::new();
            for (row, col, formula) in cells_with_formulas {
                let adjusted = evaluator.adjust_formula_for_row_insert(&formula, at);
                if adjusted != formula {
                    let value = evaluator.evaluate_formula(&adjusted);
                    let existing = self.cells.get(&(row, col));
                    updates.push((row, col, CellData {
                        value,
                        formula: Some(adjusted),
                        format: existing.and_then(|c| c.format.clone()),
                        comment: existing.and_then(|c| c.comment.clone()),
                    spill_anchor: None,
                    }));
                }
            }
            updates
        };
        for (row, col, cell) in updates {
            self.cells.insert((row, col), cell);
        }

        self.resweep_all_spills();
        // Per-sheet dep graph is gone — the workbook executor rebuilds
        // the unified graph via `rebuild_cross_sheet_deps` after every
        // structural edit (see Workbook::insert_row_on_active etc).
    }

    /// Deletes the row at the given index, shifting all rows below up by 1.
    /// Updates formula references accordingly.
    pub fn delete_row(&mut self, at: usize) {
        use crate::domain::services::FormulaEvaluator;

        let mut entries: Vec<((usize, usize), CellData)> = self.cells.drain().collect();
        entries.sort_by_key(|a| a.0.0);

        let mut new_cells = std::collections::HashMap::new();
        for ((row, col), cell) in entries {
            if row == at {
                // Skip deleted row
            } else if row > at {
                new_cells.insert((row - 1, col), cell);
            } else {
                new_cells.insert((row, col), cell);
            }
        }

        self.cells = new_cells;
        if self.rows > 1 { self.rows -= 1; }

        // Adjust formula references
        let updates = {
            let cells_with_formulas: Vec<_> = self.cells.iter()
                .filter_map(|(&(r, c), cell)| cell.formula.as_ref().map(|f| (r, c, f.clone())))
                .collect();
            let evaluator = FormulaEvaluator::new(self);
            let mut updates = Vec::new();
            for (row, col, formula) in cells_with_formulas {
                let adjusted = evaluator.adjust_formula_for_row_delete(&formula, at);
                if adjusted != formula {
                    let value = evaluator.evaluate_formula(&adjusted);
                    let existing = self.cells.get(&(row, col));
                    updates.push((row, col, CellData {
                        value,
                        formula: Some(adjusted),
                        format: existing.and_then(|c| c.format.clone()),
                        comment: existing.and_then(|c| c.comment.clone()),
                    spill_anchor: None,
                    }));
                }
            }
            updates
        };
        for (row, col, cell) in updates {
            self.cells.insert((row, col), cell);
        }

        self.resweep_all_spills();
        // Per-sheet dep graph is gone — the workbook executor rebuilds
        // the unified graph via `rebuild_cross_sheet_deps` after every
        // structural edit (see Workbook::insert_row_on_active etc).
    }

    /// Inserts a column at the given index, shifting all columns at or to the right by 1.
    /// Updates formula references accordingly.
    pub fn insert_col(&mut self, at: usize) {
        use crate::domain::services::FormulaEvaluator;

        let mut entries: Vec<((usize, usize), CellData)> = self.cells.drain().collect();
        entries.sort_by_key(|((_, col), _)| std::cmp::Reverse(*col));

        let mut new_cells = std::collections::HashMap::new();
        for ((row, col), cell) in entries {
            if col >= at {
                new_cells.insert((row, col + 1), cell);
            } else {
                new_cells.insert((row, col), cell);
            }
        }

        self.cells = new_cells;
        self.cols += 1;

        // Shift column widths
        let mut new_widths = std::collections::HashMap::new();
        for (&c, &w) in &self.column_widths {
            if c >= at {
                new_widths.insert(c + 1, w);
            } else {
                new_widths.insert(c, w);
            }
        }
        self.column_widths = new_widths;

        // Adjust formula references
        let updates = {
            let cells_with_formulas: Vec<_> = self.cells.iter()
                .filter_map(|(&(r, c), cell)| cell.formula.as_ref().map(|f| (r, c, f.clone())))
                .collect();
            let evaluator = FormulaEvaluator::new(self);
            let mut updates = Vec::new();
            for (row, col, formula) in cells_with_formulas {
                let adjusted = evaluator.adjust_formula_for_col_insert(&formula, at);
                if adjusted != formula {
                    let value = evaluator.evaluate_formula(&adjusted);
                    let existing = self.cells.get(&(row, col));
                    updates.push((row, col, CellData {
                        value,
                        formula: Some(adjusted),
                        format: existing.and_then(|c| c.format.clone()),
                        comment: existing.and_then(|c| c.comment.clone()),
                    spill_anchor: None,
                    }));
                }
            }
            updates
        };
        for (row, col, cell) in updates {
            self.cells.insert((row, col), cell);
        }

        self.resweep_all_spills();
        // Per-sheet dep graph is gone — the workbook executor rebuilds
        // the unified graph via `rebuild_cross_sheet_deps` after every
        // structural edit (see Workbook::insert_row_on_active etc).
    }

    /// Deletes the column at the given index, shifting all columns to the right left by 1.
    /// Updates formula references accordingly.
    pub fn delete_col(&mut self, at: usize) {
        use crate::domain::services::FormulaEvaluator;

        let mut entries: Vec<((usize, usize), CellData)> = self.cells.drain().collect();
        entries.sort_by_key(|a| a.0.1);

        let mut new_cells = std::collections::HashMap::new();
        for ((row, col), cell) in entries {
            if col == at {
                // Skip deleted column
            } else if col > at {
                new_cells.insert((row, col - 1), cell);
            } else {
                new_cells.insert((row, col), cell);
            }
        }

        self.cells = new_cells;
        if self.cols > 1 { self.cols -= 1; }

        // Shift column widths
        let mut new_widths = std::collections::HashMap::new();
        for (&c, &w) in &self.column_widths {
            if c == at {
                // Skip
            } else if c > at {
                new_widths.insert(c - 1, w);
            } else {
                new_widths.insert(c, w);
            }
        }
        self.column_widths = new_widths;

        // Adjust formula references
        let updates = {
            let cells_with_formulas: Vec<_> = self.cells.iter()
                .filter_map(|(&(r, c), cell)| cell.formula.as_ref().map(|f| (r, c, f.clone())))
                .collect();
            let evaluator = FormulaEvaluator::new(self);
            let mut updates = Vec::new();
            for (row, col, formula) in cells_with_formulas {
                let adjusted = evaluator.adjust_formula_for_col_delete(&formula, at);
                if adjusted != formula {
                    let value = evaluator.evaluate_formula(&adjusted);
                    let existing = self.cells.get(&(row, col));
                    updates.push((row, col, CellData {
                        value,
                        formula: Some(adjusted),
                        format: existing.and_then(|c| c.format.clone()),
                        comment: existing.and_then(|c| c.comment.clone()),
                    spill_anchor: None,
                    }));
                }
            }
            updates
        };
        for (row, col, cell) in updates {
            self.cells.insert((row, col), cell);
        }

        self.resweep_all_spills();
        // Per-sheet dep graph is gone — the workbook executor rebuilds
        // the unified graph via `rebuild_cross_sheet_deps` after every
        // structural edit (see Workbook::insert_row_on_active etc).
    }

    /// Compute the conditional-format style for the given cell, if any rule
    /// matches. Walks rules in declaration order; later rules layer on top.
    /// True if any conditional-format rule on this sheet uses a volatile
    /// function (NOW, TODAY, RAND, OFFSET, INDIRECT, GET, …) in its
    /// predicate. The recalc engine consults this at pass end to decide
    /// whether to drop this sheet's `cf_cache` — without that, a predicate
    /// like `=NOW() > _` would cache its first truth value forever even
    /// though the underlying clock advances on every recalc.
    pub fn has_volatile_cf_predicate(&self) -> bool {
        use crate::domain::parser::{
            formula_purity, FunctionRegistry, Parser,
        };
        if self.conditional_formats.is_empty() {
            return false;
        }
        let registry = FunctionRegistry::shared_builtin();
        self.conditional_formats.iter().any(|rule| {
            // Substitute the `_` placeholder with a harmless literal so
            // the predicate parses; we only care about which functions
            // it calls, not the value it produces.
            let probe = rule.predicate.replace('_', "0");
            match Parser::new(&probe) {
                Ok(mut p) => match p.parse() {
                    Ok(ast) => formula_purity(&ast, &registry).is_volatile(),
                    Err(_) => false,
                },
                Err(_) => false,
            }
        })
    }

    /// `_` in the predicate is replaced with the cell's value (quoted if
    /// non-numeric) before evaluation.
    pub fn conditional_style_for(&self, row: usize, col: usize) -> Option<CellStyle> {
        use crate::domain::services::FormulaEvaluator;
        if self.conditional_formats.is_empty() {
            return None;
        }
        // Cache hit?
        if let Some(cached) = self.cf_cache.lock().unwrap().get(&(row, col)) {
            return cached.clone();
        }
        let cell = self.get_cell(row, col);
        let value = &cell.value;
        let mut result: Option<CellStyle> = None;
        for rule in &self.conditional_formats {
            if rule.column != col {
                continue;
            }
            let token_value = if value.parse::<f64>().is_ok() {
                value.clone()
            } else {
                format!("\"{}\"", value.replace('"', "\"\""))
            };
            let bound = rule.predicate.replace('_', &token_value);
            let formula = format!("={}", bound);
            let evaluator = if self.named_ranges.is_empty() {
                FormulaEvaluator::new(self)
            } else {
                FormulaEvaluator::new(self).with_names(&self.named_ranges)
            };
            let v = evaluator.evaluate_formula(&formula);
            let truthy = match v.as_str() {
                "" | "0" | "FALSE" | "#ERROR" => false,
                _ => v.parse::<f64>().map(|n| n != 0.0).unwrap_or(true),
            };
            if truthy {
                result = Some(layer_style(result.unwrap_or_default(), &rule.style));
            }
        }
        self.cf_cache.lock().unwrap().insert((row, col), result.clone());
        result
    }

}

fn serialize_cells<S>(cells: &HashMap<(usize, usize), CellData>, serializer: S) -> Result<S::Ok, S::Error>
where
    S: serde::Serializer,
{
    use serde::ser::SerializeSeq;
    let mut seq = serializer.serialize_seq(Some(cells.len()))?;
    for (key, value) in cells {
        seq.serialize_element(&(key.0, key.1, value))?;
    }
    seq.end()
}

fn deserialize_cells<'de, D>(deserializer: D) -> Result<HashMap<(usize, usize), CellData>, D::Error>
where
    D: serde::Deserializer<'de>,
{
    use serde::de::{SeqAccess, Visitor};
    use std::fmt;

    struct CellsVisitor;

    impl<'de> Visitor<'de> for CellsVisitor {
        type Value = HashMap<(usize, usize), CellData>;

        fn expecting(&self, formatter: &mut fmt::Formatter) -> fmt::Result {
            formatter.write_str("a sequence of cell data")
        }

        fn visit_seq<A>(self, mut seq: A) -> Result<Self::Value, A::Error>
        where
            A: SeqAccess<'de>,
        {
            let mut cells = HashMap::new();
            while let Some((row, col, data)) = seq.next_element::<(usize, usize, CellData)>()? {
                cells.insert((row, col), data);
            }
            Ok(cells)
        }
    }

    deserializer.deserialize_seq(CellsVisitor)
}

fn layer_style(base: CellStyle, over: &CellStyle) -> CellStyle {
    CellStyle {
        bold: base.bold || over.bold,
        underline: base.underline || over.underline,
        fg_color: over.fg_color.clone().or(base.fg_color),
        bg_color: over.bg_color.clone().or(base.bg_color),
    }
}


#[cfg(test)]
mod tests;
