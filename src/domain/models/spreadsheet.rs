//! Submodule of `models` — see models/mod.rs.

#![allow(unused_imports)]
use serde::{Deserialize, Serialize};
use std::collections::{HashMap, HashSet, VecDeque};
use super::*;

#[derive(Debug, Clone, Serialize, Deserialize)]
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
    /// Dependency graph: cell -> set of cells that depend on it
    #[serde(skip)]
    pub dependents: HashMap<(usize, usize), HashSet<(usize, usize)>>,
    /// Dependencies: cell -> set of cells it depends on
    #[serde(skip)]
    pub dependencies: HashMap<(usize, usize), HashSet<(usize, usize)>>,
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
    /// When true, `recalculate_dependents` iterates rather than failing on
    /// cycles. Synced from `App::iterative_calc` whenever it's toggled.
    #[serde(skip)]
    pub iterative_calc: bool,
    /// Max passes for iterative recalc. Default 100.
    #[serde(skip, default = "default_iter_max")]
    pub iter_max: usize,
    /// Convergence epsilon (per-cell absolute delta). Default 1e-6.
    #[serde(skip, default = "default_iter_epsilon")]
    pub iter_epsilon: f64,
    /// Conditional-format style cache, keyed by (row, col). Populated lazily
    /// on first lookup; invalidated wholesale on any cell mutation or rule
    /// change. Refcell so `conditional_style_for(&self)` can write.
    #[serde(skip)]
    pub cf_cache: std::cell::RefCell<HashMap<(usize, usize), Option<CellStyle>>>,
}

fn default_iter_max() -> usize { 100 }
fn default_iter_epsilon() -> f64 { 1e-6 }

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
            dependents: HashMap::new(),
            dependencies: HashMap::new(),
            named_ranges: HashMap::new(),
            conditional_formats: Vec::new(),
            tables: Vec::new(),
            iterative_calc: false,
            iter_max: default_iter_max(),
            iter_epsilon: default_iter_epsilon(),
            cf_cache: std::cell::RefCell::new(HashMap::new()),
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
        self.cf_cache.borrow_mut().clear();
    }

    /// Writes a cell, updates the dependency graph, and recalculates dependents.
    ///
    /// ```
    /// use tshts::domain::{Spreadsheet, CellData};
    ///
    /// let mut sheet = Spreadsheet::default();
    /// sheet.set_cell(0, 0, CellData { value: "Hello".to_string(), ..Default::default() });
    /// ```
    pub fn set_cell(&mut self, row: usize, col: usize, data: CellData) {
        // Remove old dependencies for this cell
        self.remove_cell_dependencies(row, col);
        // Sweep any spill ghosts owned by the previous version of this cell.
        self.sweep_spill_ghosts_for(row, col);

        // Set the cell data
        self.set_cell_internal(row, col, data.clone());

        // Add new dependencies if this cell has a formula
        if let Some(ref formula) = data.formula {
            self.add_cell_dependencies(row, col, formula);
        }

        // If the formula's result is an Array, expand it into ghost cells.
        self.maybe_spill(row, col);

        // Recalculate all cells that depend on this cell
        self.recalculate_dependents(row, col);
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
    fn sweep_spill_ghosts_for(&mut self, anchor_row: usize, anchor_col: usize) {
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
    fn maybe_spill(&mut self, row: usize, col: usize) {
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
        let registry = FunctionRegistry::new();
        let names = if self.named_ranges.is_empty() {
            None
        } else {
            Some(&self.named_ranges)
        };
        let evaluator = ExpressionEvaluator::new(self, &registry, names, None);
        let value = match evaluator.evaluate(&ast) {
            Ok(v) => v,
            Err(_) => return,
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

    /// Removes all dependencies for a cell.
    fn remove_cell_dependencies(&mut self, row: usize, col: usize) {
        let cell_pos = (row, col);
        
        // Remove this cell from the dependents of cells it depends on
        if let Some(deps) = self.dependencies.get(&cell_pos).cloned() {
            for dep in deps {
                if let Some(dependents) = self.dependents.get_mut(&dep) {
                    dependents.remove(&cell_pos);
                    if dependents.is_empty() {
                        self.dependents.remove(&dep);
                    }
                }
            }
        }
        
        // Clear this cell's dependencies
        self.dependencies.remove(&cell_pos);
    }

    /// Adds dependencies for a cell based on its formula.
    /// Uses the parse cache to avoid re-tokenizing identical formulas
    /// (common during autofill/paste).
    fn add_cell_dependencies(&mut self, row: usize, col: usize, formula: &str) {
        use crate::domain::services::FormulaEvaluator;
        let evaluator = if self.named_ranges.is_empty() {
            FormulaEvaluator::new(self)
        } else {
            FormulaEvaluator::new(self).with_names(&self.named_ranges)
        };
        let dependencies = evaluator.extract_cell_references(formula);
        let cell_pos = (row, col);

        if !dependencies.is_empty() {
            self.dependencies.insert(cell_pos, dependencies.iter().cloned().collect());
            for dep in dependencies {
                self.dependents.entry(dep).or_default().insert(cell_pos);
            }
        }
    }

    /// Clears the cell and recalculates dependents.
    pub fn clear_cell(&mut self, row: usize, col: usize) {
        // Remove dependencies for this cell
        self.remove_cell_dependencies(row, col);
        // If this cell was a spill anchor, sweep its ghosts away too.
        self.sweep_spill_ghosts_for(row, col);
        // Remove the cell from the cells map
        self.cells.remove(&(row, col));
        // Recalculate cells that depend on this cell
        self.recalculate_dependents(row, col);
    }

    /// Bulk-set many cells with a single topological recalc at the end.
    /// Used by sort/autofill/paste to avoid N full recalc cascades.
    pub fn set_many(&mut self, updates: Vec<(usize, usize, CellData)>) {
        // Remove old deps and write the new values, but defer recalc.
        for (row, col, data) in &updates {
            self.remove_cell_dependencies(*row, *col);
            self.set_cell_internal(*row, *col, data.clone());
            if let Some(formula) = data.formula.as_ref() {
                self.add_cell_dependencies(*row, *col, formula);
            }
        }
        // Collect all transitive dependents of every written cell, then run
        // one topological pass over that combined set.
        let mut to_recalc = HashSet::new();
        let mut queue = VecDeque::new();
        for (row, col, _) in &updates {
            if let Some(deps) = self.dependents.get(&(*row, *col)).cloned() {
                for dep in deps {
                    queue.push_back(dep);
                }
            }
            // Also recompute the cell itself (its formula may need re-eval if
            // its inputs were among the updates).
            if self
                .cells
                .get(&(*row, *col))
                .and_then(|c| c.formula.as_ref())
                .is_some()
            {
                to_recalc.insert((*row, *col));
            }
        }
        while let Some(dep) = queue.pop_front() {
            if to_recalc.insert(dep)
                && let Some(next) = self.dependents.get(&dep).cloned() {
                    for n in next {
                        queue.push_back(n);
                    }
                }
        }
        if to_recalc.is_empty() {
            return;
        }
        let mut in_degree: HashMap<(usize, usize), usize> =
            to_recalc.iter().map(|&c| (c, 0)).collect();
        for &cell in &to_recalc {
            if let Some(deps) = self.dependencies.get(&cell) {
                for dep in deps {
                    if to_recalc.contains(dep) {
                        *in_degree.entry(cell).or_insert(0) += 1;
                    }
                }
            }
        }
        let mut ready: VecDeque<_> = in_degree
            .iter()
            .filter(|&(_, d)| *d == 0)
            .map(|(&c, _)| c)
            .collect();
        while let Some(cell) = ready.pop_front() {
            self.recalculate_cell(cell.0, cell.1);
            if let Some(deps) = self.dependents.get(&cell).cloned() {
                for dep in deps {
                    if let Some(d) = in_degree.get_mut(&dep) {
                        *d -= 1;
                        if *d == 0 {
                            ready.push_back(dep);
                        }
                    }
                }
            }
        }
    }

    /// Recalculates all cells that depend on the given cell using topological ordering.
    fn recalculate_dependents(&mut self, row: usize, col: usize) {
        let cell_pos = (row, col);

        // 1. Collect all transitive dependents
        let mut to_recalc = HashSet::new();
        let mut queue = VecDeque::new();
        if let Some(deps) = self.dependents.get(&cell_pos).cloned() {
            for dep in deps {
                queue.push_back(dep);
            }
        }
        while let Some(dep) = queue.pop_front() {
            if to_recalc.insert(dep)
                && let Some(next) = self.dependents.get(&dep).cloned() {
                    for n in next {
                        queue.push_back(n);
                    }
                }
        }

        if to_recalc.is_empty() {
            return;
        }

        // Iterative-calc mode: just sweep all cells in `to_recalc` up to
        // `iter_max` times, stopping early when no cell's numeric value
        // changes by more than `iter_epsilon` in a full pass. This is the
        // standard Excel approach to intentional circular references.
        if self.iterative_calc {
            let max = self.iter_max;
            let eps = self.iter_epsilon;
            for _ in 0..max {
                let mut changed = false;
                for &(r, c) in &to_recalc {
                    let prev: f64 = self
                        .cells
                        .get(&(r, c))
                        .map(|cd| cd.value.parse::<f64>().unwrap_or(0.0))
                        .unwrap_or(0.0);
                    self.recalculate_cell(r, c);
                    let next: f64 = self
                        .cells
                        .get(&(r, c))
                        .map(|cd| cd.value.parse::<f64>().unwrap_or(0.0))
                        .unwrap_or(0.0);
                    if (next - prev).abs() > eps {
                        changed = true;
                    }
                }
                if !changed {
                    break;
                }
            }
            return;
        }

        // 2. Compute in-degrees within recalc set
        let mut in_degree: HashMap<(usize, usize), usize> = to_recalc.iter().map(|&c| (c, 0)).collect();
        for &cell in &to_recalc {
            if let Some(deps) = self.dependencies.get(&cell) {
                for dep in deps {
                    if to_recalc.contains(dep) {
                        *in_degree.entry(cell).or_insert(0) += 1;
                    }
                }
            }
        }

        // 3. Process in topological order (Kahn's algorithm)
        let mut ready: VecDeque<_> = in_degree.iter()
            .filter(|&(_, d)| *d == 0)
            .map(|(&c, _)| c)
            .collect();
        while let Some(cell) = ready.pop_front() {
            self.recalculate_cell(cell.0, cell.1);
            in_degree.remove(&cell);
            if let Some(deps) = self.dependents.get(&cell).cloned() {
                for dep in deps {
                    if let Some(d) = in_degree.get_mut(&dep) {
                        *d -= 1;
                        if *d == 0 {
                            ready.push_back(dep);
                        }
                    }
                }
            }
        }

        // 4. Anything left has in_degree > 0 — i.e. participates in a cycle
        // that iterative_calc wasn't enabled to resolve. Surface as `#REF!`
        // (Excel uses `#CIRC!`; we reuse Ref to avoid expanding ErrorKind)
        // instead of silently leaving stale values in place.
        for (row, col) in in_degree.into_keys() {
            if let Some(cell) = self.cells.get(&(row, col)).cloned() {
                let mut updated = cell;
                updated.value = "#REF!".to_string();
                self.cells.insert((row, col), updated);
            }
        }
    }

    /// Force a single cell to re-evaluate its formula and refresh `value`.
    /// Public wrapper around `recalculate_cell` for use after non-`set_cell`
    /// mutations (e.g. cross-sheet ref rewrites in `Workbook::remove_sheet`)
    /// where the formula text changed without going through `set_cell`.
    pub fn refresh_cell_value(&mut self, row: usize, col: usize) {
        self.recalculate_cell(row, col);
    }

    /// Recalculates a single cell's value based on its formula.
    fn recalculate_cell(&mut self, row: usize, col: usize) {
        let cell_pos = (row, col);
        if let Some(cell) = self.cells.get(&cell_pos).cloned()
            && let Some(ref formula) = cell.formula {
                use crate::domain::services::FormulaEvaluator;
                let evaluator = if self.named_ranges.is_empty() {
                    FormulaEvaluator::new(self)
                } else {
                    FormulaEvaluator::new(self).with_names(&self.named_ranges)
                };
                let new_value = evaluator.evaluate_formula(formula);
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
        if let Some(idx) = cell_ref.rfind('!') {
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
        let mut max_width = Self::column_label(col).width();

        for (&(_, c), cell) in &self.cells {
            if c == col {
                let value_width = cell.value.width();
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
        let mut widths: HashMap<usize, usize> = HashMap::with_capacity(self.cols);
        for col in 0..self.cols {
            widths.insert(col, Self::column_label(col).width());
        }
        for (&(_, c), cell) in &self.cells {
            let value_width = cell.value.width();
            let formula_width = cell.formula.as_ref().map(|f| f.width()).unwrap_or(0);
            let content_width = value_width.max(formula_width);
            let entry = widths.entry(c).or_insert(3);
            *entry = (*entry).max(content_width);
        }
        for (col, mut w) in widths {
            w = w.max(3).min(50);
            self.set_column_width(col, w);
        }
    }

    /// Inserts a row at the given index, shifting all rows at or below down by 1.
    /// Updates formula references accordingly.
    pub fn insert_row(&mut self, at: usize) {
        use crate::domain::services::FormulaEvaluator;

        // Collect all cells, sorted by row descending so we shift from bottom up
        let mut entries: Vec<((usize, usize), CellData)> = self.cells.drain().collect();
        entries.sort_by(|a, b| b.0.0.cmp(&a.0.0));

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

        self.rebuild_dependencies();
        self.resweep_all_spills();
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

        self.rebuild_dependencies();
        self.resweep_all_spills();
    }

    /// Inserts a column at the given index, shifting all columns at or to the right by 1.
    /// Updates formula references accordingly.
    pub fn insert_col(&mut self, at: usize) {
        use crate::domain::services::FormulaEvaluator;

        let mut entries: Vec<((usize, usize), CellData)> = self.cells.drain().collect();
        entries.sort_by(|a, b| b.0.1.cmp(&a.0.1));

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

        self.rebuild_dependencies();
        self.resweep_all_spills();
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

        self.rebuild_dependencies();
        self.resweep_all_spills();
    }

    /// Compute the conditional-format style for the given cell, if any rule
    /// matches. Walks rules in declaration order; later rules layer on top.
    /// `_` in the predicate is replaced with the cell's value (quoted if
    /// non-numeric) before evaluation.
    pub fn conditional_style_for(&self, row: usize, col: usize) -> Option<CellStyle> {
        use crate::domain::services::FormulaEvaluator;
        if self.conditional_formats.is_empty() {
            return None;
        }
        // Cache hit?
        if let Some(cached) = self.cf_cache.borrow().get(&(row, col)) {
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
        self.cf_cache.borrow_mut().insert((row, col), result.clone());
        result
    }

    /// Rebuilds the dependency graph from formulas. Dependencies aren't
    /// serialized, so this must run after deserializing a spreadsheet.
    pub fn rebuild_dependencies(&mut self) {
        // Clear existing dependencies
        self.dependencies.clear();
        self.dependents.clear();
        
        // Rebuild dependencies for all cells with formulas
        let cells_with_formulas: Vec<_> = self.cells
            .iter()
            .filter_map(|((row, col), cell)| {
                cell.formula.as_ref().map(|formula| (*row, *col, formula.clone()))
            })
            .collect();
        
        for (row, col, formula) in cells_with_formulas {
            self.add_cell_dependencies(row, col, &formula);
        }
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
mod tests {
    use super::*;
    use crate::domain::{CellData};
    #[test]
    fn test_spreadsheet_default() {
        let sheet = Spreadsheet::default();
        assert_eq!(sheet.rows, 100);
        assert_eq!(sheet.cols, 26);
        assert_eq!(sheet.default_column_width, 8);
        assert!(sheet.cells.is_empty());
        assert!(sheet.column_widths.is_empty());
    }

    #[test]
    fn test_get_cell_empty() {
        let sheet = Spreadsheet::default();
        let cell = sheet.get_cell(0, 0);
        assert!(cell.value.is_empty());
        assert!(cell.formula.is_none());
    }

    #[test]
    fn test_set_and_get_cell() {
        let mut sheet = Spreadsheet::default();
        let cell_data = CellData {
            value: "Hello".to_string(),
            formula: None,
            format: None,
            comment: None,
        spill_anchor: None,
        };
        sheet.set_cell(0, 0, cell_data.clone());
        
        let retrieved = sheet.get_cell(0, 0);
        assert_eq!(retrieved.value, "Hello");
        assert!(retrieved.formula.is_none());
    }

    #[test]
    fn test_set_cell_no_auto_resize() {
        let mut sheet = Spreadsheet::default();
        let initial_width = sheet.get_column_width(0);
        
        let long_cell = CellData {
            value: "This is a very long cell value".to_string(),
            formula: None,
            format: None,
            comment: None,
        spill_anchor: None,
        };
        sheet.set_cell(0, 0, long_cell);
        
        let new_width = sheet.get_column_width(0);
        assert_eq!(new_width, initial_width); // No automatic resizing
    }

    #[test]
    fn test_column_label() {
        assert_eq!(Spreadsheet::column_label(0), "A");
        assert_eq!(Spreadsheet::column_label(1), "B");
        assert_eq!(Spreadsheet::column_label(25), "Z");
        assert_eq!(Spreadsheet::column_label(26), "AA");
        assert_eq!(Spreadsheet::column_label(27), "AB");
        assert_eq!(Spreadsheet::column_label(51), "AZ");
        assert_eq!(Spreadsheet::column_label(52), "BA");
        assert_eq!(Spreadsheet::column_label(701), "ZZ");
        assert_eq!(Spreadsheet::column_label(702), "AAA");
    }

    #[test]
    fn test_parse_cell_reference() {
        // Valid references
        assert_eq!(Spreadsheet::parse_cell_reference("A1"), Some((0, 0)));
        assert_eq!(Spreadsheet::parse_cell_reference("B2"), Some((1, 1)));
        assert_eq!(Spreadsheet::parse_cell_reference("Z26"), Some((25, 25)));
        assert_eq!(Spreadsheet::parse_cell_reference("AA1"), Some((0, 26)));
        assert_eq!(Spreadsheet::parse_cell_reference("AB100"), Some((99, 27)));
        
        // Case insensitive
        assert_eq!(Spreadsheet::parse_cell_reference("a1"), Some((0, 0)));
        assert_eq!(Spreadsheet::parse_cell_reference("b2"), Some((1, 1)));
        
        // Invalid references
        assert_eq!(Spreadsheet::parse_cell_reference(""), None);
        assert_eq!(Spreadsheet::parse_cell_reference("A"), None);
        assert_eq!(Spreadsheet::parse_cell_reference("1"), None);
        assert_eq!(Spreadsheet::parse_cell_reference("A0"), None); // Row 0 doesn't exist in Excel notation
        assert_eq!(Spreadsheet::parse_cell_reference("1A"), None);
        assert_eq!(Spreadsheet::parse_cell_reference("A1B"), None);
        assert_eq!(Spreadsheet::parse_cell_reference("A-1"), None);
    }

    #[test]
    fn test_column_width_management() {
        let mut sheet = Spreadsheet::default();
        
        // Test default width
        assert_eq!(sheet.get_column_width(0), 8);
        
        // Test setting custom width
        sheet.set_column_width(0, 15);
        assert_eq!(sheet.get_column_width(0), 15);
        
        // Test other columns still use default
        assert_eq!(sheet.get_column_width(1), 8);
    }

    #[test]
    fn test_auto_resize_column() {
        let mut sheet = Spreadsheet::default();
        
        // Add cells with varying lengths
        sheet.set_cell(0, 0, CellData { value: "Hi".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        sheet.set_cell(1, 0, CellData { value: "Medium length".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        sheet.set_cell(2, 0, CellData { value: "Very long content here".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        
        sheet.auto_resize_column(0);
        let width = sheet.get_column_width(0);
        
        // Should be at least as wide as the longest content
        assert!(width >= "Very long content here".len());
        // But not more than the maximum of 50
        assert!(width <= 50);
    }

    #[test]
    fn test_auto_resize_all_columns() {
        let mut sheet = Spreadsheet::default();
        
        // Add content to multiple columns
        sheet.set_cell(0, 0, CellData { value: "Short".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        sheet.set_cell(0, 1, CellData { value: "Much longer content".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        sheet.set_cell(0, 2, CellData { value: "X".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        
        sheet.auto_resize_all_columns();
        
        // Each column should be sized appropriately
        assert!(sheet.get_column_width(0) >= 5); // "Short".len()
        assert!(sheet.get_column_width(1) >= 19); // "Much longer content".len()
        assert!(sheet.get_column_width(2) >= 3); // Minimum width
    }

    #[test]
    fn test_formula_cell_no_auto_resize() {
        let mut sheet = Spreadsheet::default();
        let initial_width = sheet.get_column_width(0);
        
        let formula_cell = CellData {
            value: "42".to_string(),
            formula: Some("=SUM(A1:A10)".to_string()),
            format: None,
            comment: None,
        spill_anchor: None,
        };
        
        sheet.set_cell(0, 0, formula_cell);
        let width = sheet.get_column_width(0);
        
        // Width should remain unchanged (no automatic resizing)
        assert_eq!(width, initial_width);
    }

    #[test]
    fn test_serialization_roundtrip() {
        let mut original = Spreadsheet::default();
        original.set_cell(0, 0, CellData {
            value: "test".to_string(),
            formula: Some("=1+1".to_string()),
            format: None,
            comment: None,
        spill_anchor: None,
        });
        original.set_cell(1, 1, CellData {
            value: "42".to_string(),
            formula: None,
            format: None,
            comment: None,
        spill_anchor: None,
        });
        original.set_column_width(0, 15);
        
        // Serialize to JSON
        let json = serde_json::to_string(&original).expect("Serialization failed");
        
        // Deserialize back
        let deserialized: Spreadsheet = serde_json::from_str(&json).expect("Deserialization failed");
        
        // Verify data integrity
        assert_eq!(deserialized.rows, original.rows);
        assert_eq!(deserialized.cols, original.cols);
        assert_eq!(deserialized.default_column_width, original.default_column_width);
        
        let cell_0_0 = deserialized.get_cell(0, 0);
        assert_eq!(cell_0_0.value, "test");
        assert_eq!(cell_0_0.formula.unwrap(), "=1+1");
        
        let cell_1_1 = deserialized.get_cell(1, 1);
        assert_eq!(cell_1_1.value, "42");
        assert!(cell_1_1.formula.is_none());
        
        assert_eq!(deserialized.get_column_width(0), 15);
    }

    #[test]
    fn test_automatic_recalculation() {
        let mut sheet = Spreadsheet::default();
        
        // Set up a simple dependency chain: C1 = A1 + B1
        sheet.set_cell(0, 0, CellData { value: "10".to_string(), formula: None, format: None, comment: None, spill_anchor: None }); // A1 = 10
        sheet.set_cell(0, 1, CellData { value: "20".to_string(), formula: None, format: None, comment: None, spill_anchor: None }); // B1 = 20
        sheet.set_cell(0, 2, CellData { 
            value: "30".to_string(), 
            formula: Some("=A1+B1".to_string()),
            format: None,
            comment: None,
        spill_anchor: None,
        }); // C1 = A1+B1 = 30
        
        // Verify initial state
        assert_eq!(sheet.get_cell(0, 2).value, "30");
        
        // Change A1 and verify C1 updates automatically
        sheet.set_cell(0, 0, CellData { value: "15".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        assert_eq!(sheet.get_cell(0, 2).value, "35"); // Should be 15+20=35
        
        // Change B1 and verify C1 updates automatically
        sheet.set_cell(0, 1, CellData { value: "25".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        assert_eq!(sheet.get_cell(0, 2).value, "40"); // Should be 15+25=40
    }

    #[test]
    fn test_dependency_chain_recalculation() {
        let mut sheet = Spreadsheet::default();
        
        // Set up a dependency chain: A1 -> B1 -> C1
        sheet.set_cell(0, 0, CellData { value: "5".to_string(), formula: None, format: None, comment: None, spill_anchor: None }); // A1 = 5
        sheet.set_cell(0, 1, CellData { 
            value: "10".to_string(), 
            formula: Some("=A1*2".to_string()),
            format: None,
            comment: None,
        spill_anchor: None,
        }); // B1 = A1*2 = 10
        sheet.set_cell(0, 2, CellData { 
            value: "20".to_string(), 
            formula: Some("=B1*2".to_string()),
            format: None,
            comment: None,
        spill_anchor: None,
        }); // C1 = B1*2 = 20
        
        // Verify initial state
        assert_eq!(sheet.get_cell(0, 0).value, "5");
        assert_eq!(sheet.get_cell(0, 1).value, "10");
        assert_eq!(sheet.get_cell(0, 2).value, "20");
        
        // Change A1 and verify the entire chain updates
        sheet.set_cell(0, 0, CellData { value: "10".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        assert_eq!(sheet.get_cell(0, 0).value, "10");
        assert_eq!(sheet.get_cell(0, 1).value, "20"); // 10*2=20
        assert_eq!(sheet.get_cell(0, 2).value, "40"); // 20*2=40
    }

    #[test]
    fn test_multiple_dependents() {
        let mut sheet = Spreadsheet::default();
        
        // Set up multiple cells depending on A1
        sheet.set_cell(0, 0, CellData { value: "10".to_string(), formula: None, format: None, comment: None, spill_anchor: None }); // A1 = 10
        sheet.set_cell(0, 1, CellData { 
            value: "11".to_string(), 
            formula: Some("=A1+1".to_string()),
            format: None,
            comment: None,
        spill_anchor: None,
        }); // B1 = A1+1 = 11
        sheet.set_cell(0, 2, CellData { 
            value: "20".to_string(), 
            formula: Some("=A1*2".to_string()),
            format: None,
            comment: None,
        spill_anchor: None,
        }); // C1 = A1*2 = 20
        sheet.set_cell(0, 3, CellData { 
            value: "100".to_string(), 
            formula: Some("=A1*A1".to_string()),
            format: None,
            comment: None,
        spill_anchor: None,
        }); // D1 = A1*A1 = 100
        
        // Verify initial state
        assert_eq!(sheet.get_cell(0, 1).value, "11");
        assert_eq!(sheet.get_cell(0, 2).value, "20");
        assert_eq!(sheet.get_cell(0, 3).value, "100");
        
        // Change A1 and verify all dependents update
        sheet.set_cell(0, 0, CellData { value: "5".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        assert_eq!(sheet.get_cell(0, 1).value, "6");   // 5+1=6
        assert_eq!(sheet.get_cell(0, 2).value, "10");  // 5*2=10
        assert_eq!(sheet.get_cell(0, 3).value, "25");  // 5*5=25
    }

    #[test]
    fn test_dependency_removal() {
        let mut sheet = Spreadsheet::default();
        
        // Set up a dependency
        sheet.set_cell(0, 0, CellData { value: "10".to_string(), formula: None, format: None, comment: None, spill_anchor: None }); // A1 = 10
        sheet.set_cell(0, 1, CellData { 
            value: "20".to_string(), 
            formula: Some("=A1*2".to_string()),
            format: None,
            comment: None,
        spill_anchor: None,
        }); // B1 = A1*2 = 20
        
        // Verify dependency exists
        assert_eq!(sheet.get_cell(0, 1).value, "20");
        
        // Change A1 and verify B1 updates
        sheet.set_cell(0, 0, CellData { value: "15".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        assert_eq!(sheet.get_cell(0, 1).value, "30");
        
        // Replace B1 with a constant value (remove dependency)
        sheet.set_cell(0, 1, CellData { value: "42".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        assert_eq!(sheet.get_cell(0, 1).value, "42");
        
        // Change A1 again - B1 should NOT update since dependency is removed
        sheet.set_cell(0, 0, CellData { value: "100".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        assert_eq!(sheet.get_cell(0, 1).value, "42"); // Should remain 42, not recalculate
    }

    #[test]
    fn test_rebuild_dependencies() {
        let mut sheet = Spreadsheet::default();
        
        // Manually insert cells with formulas (simulating loading from file)
        sheet.cells.insert((0, 0), CellData { value: "10".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        sheet.cells.insert((0, 1), CellData { 
            value: "20".to_string(), 
            formula: Some("=A1*2".to_string()),
            format: None,
            comment: None,
        spill_anchor: None,
        });
        sheet.cells.insert((0, 2), CellData { 
            value: "40".to_string(), 
            formula: Some("=B1*2".to_string()),
            format: None,
            comment: None,
        spill_anchor: None,
        });
        
        // At this point, dependencies are not tracked
        assert!(sheet.dependencies.is_empty());
        assert!(sheet.dependents.is_empty());
        
        // Rebuild dependencies
        sheet.rebuild_dependencies();
        
        // Verify dependencies are now tracked
        assert!(!sheet.dependencies.is_empty());
        assert!(!sheet.dependents.is_empty());
        
        // Test that recalculation works after rebuilding
        sheet.set_cell(0, 0, CellData { value: "5".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        assert_eq!(sheet.get_cell(0, 1).value, "10"); // 5*2=10
        assert_eq!(sheet.get_cell(0, 2).value, "20"); // 10*2=20
    }

    #[test]
    fn test_range_dependency_recalculation() {
        let mut sheet = Spreadsheet::default();
        
        // Set up cells A1:A3 and a SUM formula
        sheet.set_cell(0, 0, CellData { value: "1".to_string(), formula: None, format: None, comment: None, spill_anchor: None }); // A1 = 1
        sheet.set_cell(1, 0, CellData { value: "2".to_string(), formula: None, format: None, comment: None, spill_anchor: None }); // A2 = 2
        sheet.set_cell(2, 0, CellData { value: "3".to_string(), formula: None, format: None, comment: None, spill_anchor: None }); // A3 = 3
        sheet.set_cell(0, 1, CellData { 
            value: "6".to_string(), 
            formula: Some("=SUM(A1:A3)".to_string()),
            format: None,
            comment: None,
        spill_anchor: None,
        }); // B1 = SUM(A1:A3) = 6
        
        // Verify initial state
        assert_eq!(sheet.get_cell(0, 1).value, "6");
        
        // Change one cell in the range
        sheet.set_cell(1, 0, CellData { value: "5".to_string(), formula: None, format: None, comment: None, spill_anchor: None }); // A2 = 5
        assert_eq!(sheet.get_cell(0, 1).value, "9"); // Should be 1+5+3=9
        
        // Change another cell in the range
        sheet.set_cell(2, 0, CellData { value: "10".to_string(), formula: None, format: None, comment: None, spill_anchor: None }); // A3 = 10
        assert_eq!(sheet.get_cell(0, 1).value, "16"); // Should be 1+5+10=16
    }

    #[test]
    fn test_circular_dependency_handling() {
        let mut sheet = Spreadsheet::default();
        
        // Set up a potential circular dependency scenario
        sheet.set_cell(0, 0, CellData { value: "10".to_string(), formula: None, format: None, comment: None, spill_anchor: None }); // A1 = 10
        sheet.set_cell(0, 1, CellData { 
            value: "20".to_string(), 
            formula: Some("=A1*2".to_string()),
            format: None,
            comment: None,
        spill_anchor: None,
        }); // B1 = A1*2 = 20
        
        // Now try to create a circular dependency A1 = B1 + 1
        // This should be prevented by the circular reference check
        use crate::domain::services::FormulaEvaluator;
        let evaluator = FormulaEvaluator::new(&sheet);
        let would_be_circular = evaluator.would_create_circular_reference("=B1+1", (0, 0));
        assert!(would_be_circular); // Should detect the circular reference
        
        // The dependency system should also handle this gracefully
        // Even if somehow a circular dependency got through, recalculation should not hang
    }

    #[test]
    fn test_extract_cell_references_from_formula() {
        use crate::domain::services::FormulaEvaluator;
        
        let sheet = Spreadsheet::default();
        let evaluator = FormulaEvaluator::new(&sheet);
        
        // Test simple cell reference
        let refs = evaluator.extract_cell_references("=A1");
        assert_eq!(refs, vec![(0, 0)]);
        
        // Test multiple cell references
        let refs = evaluator.extract_cell_references("=A1+B2*C3");
        assert_eq!(refs.len(), 3);
        assert!(refs.contains(&(0, 0))); // A1
        assert!(refs.contains(&(1, 1))); // B2
        assert!(refs.contains(&(2, 2))); // C3
        
        // Test range reference
        let refs = evaluator.extract_cell_references("=SUM(A1:A3)");
        assert_eq!(refs.len(), 3);
        assert!(refs.contains(&(0, 0))); // A1
        assert!(refs.contains(&(1, 0))); // A2
        assert!(refs.contains(&(2, 0))); // A3
        
        // Test no references
        let refs = evaluator.extract_cell_references("=5+10");
        assert!(refs.is_empty());
        
        // Test non-formula
        let refs = evaluator.extract_cell_references("Hello World");
        assert!(refs.is_empty());
    }

    #[test]
    fn test_dependency_tracking_persistence() {
        use crate::infrastructure::FileRepository;
        use tempfile::NamedTempFile;
        
        let mut original = Spreadsheet::default();
        
        // Set up dependencies
        original.set_cell(0, 0, CellData { value: "10".to_string(), formula: None, format: None, comment: None, spill_anchor: None }); // A1 = 10
        original.set_cell(0, 1, CellData { 
            value: "20".to_string(), 
            formula: Some("=A1*2".to_string()),
            format: None,
            comment: None,
        spill_anchor: None,
        }); // B1 = A1*2 = 20
        original.set_cell(0, 2, CellData { 
            value: "40".to_string(), 
            formula: Some("=B1*2".to_string()),
            format: None,
            comment: None,
        spill_anchor: None,
        }); // C1 = B1*2 = 40
        
        // Save to file
        let temp_file = NamedTempFile::new().expect("Failed to create temp file");
        let file_path = temp_file.path().to_str().unwrap();
        FileRepository::save_spreadsheet(&original, file_path).expect("Save failed");
        
        // Load from file
        let (mut loaded, _) = FileRepository::load_spreadsheet(file_path).expect("Load failed");
        
        // Dependencies should be rebuilt and functional
        loaded.set_cell(0, 0, CellData { value: "5".to_string(), formula: None, format: None, comment: None, spill_anchor: None }); // Change A1 to 5
        
        // Verify that dependent cells were recalculated
        assert_eq!(loaded.get_cell(0, 1).value, "10"); // B1 = 5*2 = 10
        assert_eq!(loaded.get_cell(0, 2).value, "20"); // C1 = 10*2 = 20
    }

    #[test]
    fn test_diamond_dependency_recalculation() {
        let mut sheet = Spreadsheet::default();

        // Diamond pattern: A1 -> B1, A1 -> C1, B1 -> C1
        // A1 = 10
        sheet.set_cell(0, 0, CellData { value: "10".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        // B1 = A1 * 2
        sheet.set_cell(0, 1, CellData {
            value: "20".to_string(),
            formula: Some("=A1*2".to_string()),
            format: None,
            comment: None,
        spill_anchor: None,
        });
        // C1 = A1 + B1 (depends on both A1 and B1)
        sheet.set_cell(0, 2, CellData {
            value: "30".to_string(),
            formula: Some("=A1+B1".to_string()),
            format: None,
            comment: None,
        spill_anchor: None,
        });

        // Verify initial state
        assert_eq!(sheet.get_cell(0, 0).value, "10");
        assert_eq!(sheet.get_cell(0, 1).value, "20");
        assert_eq!(sheet.get_cell(0, 2).value, "30"); // 10 + 20

        // Change A1 — B1 must update before C1 for correct result
        sheet.set_cell(0, 0, CellData { value: "5".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        assert_eq!(sheet.get_cell(0, 1).value, "10"); // 5*2 = 10
        assert_eq!(sheet.get_cell(0, 2).value, "15"); // 5 + 10 = 15 (not 5 + 20 = 25)
    }

    #[test]
    fn test_auto_resize_column_shrinks() {
        let mut sheet = Spreadsheet::default();

        // Add wide content and auto-resize
        sheet.set_cell(0, 0, CellData {
            value: "This is very wide content".to_string(),
            formula: None,
            format: None,
            comment: None,
        spill_anchor: None,
        });
        sheet.auto_resize_column(0);
        let wide_width = sheet.get_column_width(0);
        assert!(wide_width >= "This is very wide content".len());

        // Replace with short content
        sheet.set_cell(0, 0, CellData {
            value: "Hi".to_string(),
            formula: None,
            format: None,
            comment: None,
        spill_anchor: None,
        });
        sheet.auto_resize_column(0);
        let narrow_width = sheet.get_column_width(0);

        // Column should have shrunk
        assert!(narrow_width < wide_width);
        assert!(narrow_width >= 3); // minimum width
    }

    #[test]
    fn test_insert_row() {
        let mut sheet = Spreadsheet::default();
        sheet.set_cell(0, 0, CellData { value: "A1".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        sheet.set_cell(1, 0, CellData { value: "A2".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        sheet.set_cell(2, 0, CellData { value: "A3".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        let orig_rows = sheet.rows;

        sheet.insert_row(1); // Insert above row 1 (A2)

        assert_eq!(sheet.rows, orig_rows + 1);
        assert_eq!(sheet.get_cell(0, 0).value, "A1"); // Row 0 unchanged
        assert!(sheet.get_cell(1, 0).value.is_empty()); // New empty row
        assert_eq!(sheet.get_cell(2, 0).value, "A2"); // Shifted down
        assert_eq!(sheet.get_cell(3, 0).value, "A3"); // Shifted down
    }

    #[test]
    fn test_delete_row() {
        let mut sheet = Spreadsheet::default();
        sheet.set_cell(0, 0, CellData { value: "A1".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        sheet.set_cell(1, 0, CellData { value: "A2".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        sheet.set_cell(2, 0, CellData { value: "A3".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        let orig_rows = sheet.rows;

        sheet.delete_row(1); // Delete row 1 (A2)

        assert_eq!(sheet.rows, orig_rows - 1);
        assert_eq!(sheet.get_cell(0, 0).value, "A1"); // Row 0 unchanged
        assert_eq!(sheet.get_cell(1, 0).value, "A3"); // Shifted up
    }

    #[test]
    fn test_insert_col() {
        let mut sheet = Spreadsheet::default();
        sheet.set_cell(0, 0, CellData { value: "A1".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        sheet.set_cell(0, 1, CellData { value: "B1".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        sheet.set_cell(0, 2, CellData { value: "C1".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        let orig_cols = sheet.cols;

        sheet.insert_col(1); // Insert before column B

        assert_eq!(sheet.cols, orig_cols + 1);
        assert_eq!(sheet.get_cell(0, 0).value, "A1");
        assert!(sheet.get_cell(0, 1).value.is_empty()); // New empty column
        assert_eq!(sheet.get_cell(0, 2).value, "B1"); // Shifted right
        assert_eq!(sheet.get_cell(0, 3).value, "C1"); // Shifted right
    }

    #[test]
    fn test_delete_col() {
        let mut sheet = Spreadsheet::default();
        sheet.set_cell(0, 0, CellData { value: "A1".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        sheet.set_cell(0, 1, CellData { value: "B1".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        sheet.set_cell(0, 2, CellData { value: "C1".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        let orig_cols = sheet.cols;

        sheet.delete_col(1); // Delete column B

        assert_eq!(sheet.cols, orig_cols - 1);
        assert_eq!(sheet.get_cell(0, 0).value, "A1");
        assert_eq!(sheet.get_cell(0, 1).value, "C1"); // Shifted left
    }

    #[test]
    fn test_insert_row_adjusts_formulas() {
        let mut sheet = Spreadsheet::default();
        sheet.set_cell(0, 0, CellData { value: "10".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        sheet.set_cell(1, 0, CellData { value: "20".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        sheet.set_cell(2, 0, CellData {
            value: "30".to_string(),
            formula: Some("=A1+A2".to_string()),
            format: None,
            comment: None,
        spill_anchor: None,
        });

        sheet.insert_row(1); // Insert before row 1

        // Formula should now reference A1+A3 (A2 shifted to A3)
        let cell = sheet.get_cell(3, 0); // Original row 2 moved to row 3
        assert!(cell.formula.is_some());
        let formula = cell.formula.unwrap();
        assert!(formula.contains("A3"), "Formula should reference A3, got: {}", formula);
    }

    #[test]
    fn test_insert_col_adjusts_formulas() {
        let mut sheet = Spreadsheet::default();
        sheet.set_cell(0, 0, CellData { value: "10".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        sheet.set_cell(0, 1, CellData { value: "20".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        sheet.set_cell(0, 2, CellData {
            value: "30".to_string(),
            formula: Some("=A1+B1".to_string()),
            format: None,
            comment: None,
        spill_anchor: None,
        });

        sheet.insert_col(1); // Insert before column B

        // Formula should adjust: B1 -> C1
        let cell = sheet.get_cell(0, 3); // Original col 2 moved to col 3
        assert!(cell.formula.is_some());
        let formula = cell.formula.unwrap();
        assert!(formula.contains("C1"), "Formula should reference C1, got: {}", formula);
    }

    #[test]
    fn test_cell_data_with_format() {
        let cell = CellData {
            value: "42".to_string(),
            formula: None,
            format: Some(CellFormat {
                number_format: NumberFormat::Currency { symbol: "$".to_string(), decimals: 2 },
                ..CellFormat::default()
            }),
            comment: None,
        spill_anchor: None,
        };
        assert!(cell.format.is_some());
        let fmt = cell.format.unwrap();
        assert!(matches!(fmt.number_format, NumberFormat::Currency { .. }));
    }

}
