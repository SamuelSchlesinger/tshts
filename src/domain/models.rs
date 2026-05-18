//! Domain models for the terminal spreadsheet application.
//!
//! This module contains the core data structures that represent
//! spreadsheet cells and the spreadsheet itself.

use std::collections::{HashMap, HashSet, VecDeque};
use serde::{Deserialize, Serialize};

/// Number format for cell display.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub enum NumberFormat {
    /// Default rendering (no formatting)
    General,
    /// Fixed decimal places with optional thousands separator
    Number { decimals: u32, thousands_sep: bool },
    /// Currency with symbol and decimal places
    Currency { symbol: String, decimals: u32 },
    /// Percentage (multiply by 100 and add %)
    Percentage { decimals: u32 },
}

/// Terminal color for cell styling.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub enum TerminalColor {
    Black,
    Red,
    Green,
    Yellow,
    Blue,
    Magenta,
    Cyan,
    White,
    DarkGray,
    LightRed,
    LightGreen,
    LightYellow,
    LightBlue,
    LightMagenta,
    LightCyan,
}

impl TerminalColor {
    /// Parses a color name string into a TerminalColor.
    pub fn from_name(name: &str) -> Option<Self> {
        match name.to_lowercase().as_str() {
            "black" => Some(Self::Black),
            "red" => Some(Self::Red),
            "green" => Some(Self::Green),
            "yellow" => Some(Self::Yellow),
            "blue" => Some(Self::Blue),
            "magenta" => Some(Self::Magenta),
            "cyan" => Some(Self::Cyan),
            "white" => Some(Self::White),
            "darkgray" | "dark_gray" => Some(Self::DarkGray),
            "lightred" | "light_red" => Some(Self::LightRed),
            "lightgreen" | "light_green" => Some(Self::LightGreen),
            "lightyellow" | "light_yellow" => Some(Self::LightYellow),
            "lightblue" | "light_blue" => Some(Self::LightBlue),
            "lightmagenta" | "light_magenta" => Some(Self::LightMagenta),
            "lightcyan" | "light_cyan" => Some(Self::LightCyan),
            _ => None,
        }
    }
}

/// Visual style for a cell.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct CellStyle {
    pub bold: bool,
    pub underline: bool,
    pub fg_color: Option<TerminalColor>,
    pub bg_color: Option<TerminalColor>,
}

impl Default for CellStyle {
    fn default() -> Self {
        Self { bold: false, underline: false, fg_color: None, bg_color: None }
    }
}

/// Cell formatting options.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct CellFormat {
    /// Number format
    pub number_format: NumberFormat,
    /// Cell visual style
    pub style: CellStyle,
}

impl Default for CellFormat {
    fn default() -> Self {
        Self {
            number_format: NumberFormat::General,
            style: CellStyle::default(),
        }
    }
}

/// Formats a cell value according to the given number format.
pub fn format_cell_value(value: &str, format: &CellFormat) -> String {
    match &format.number_format {
        NumberFormat::General => value.to_string(),
        NumberFormat::Number { decimals, thousands_sep } => {
            if let Ok(n) = value.parse::<f64>() {
                let formatted = format!("{:.prec$}", n, prec = *decimals as usize);
                if *thousands_sep {
                    add_thousands_separator(&formatted)
                } else {
                    formatted
                }
            } else {
                value.to_string()
            }
        }
        NumberFormat::Currency { symbol, decimals } => {
            if let Ok(n) = value.parse::<f64>() {
                let formatted = format!("{:.prec$}", n, prec = *decimals as usize);
                format!("{}{}", symbol, add_thousands_separator(&formatted))
            } else {
                value.to_string()
            }
        }
        NumberFormat::Percentage { decimals } => {
            if let Ok(n) = value.parse::<f64>() {
                format!("{:.prec$}%", n * 100.0, prec = *decimals as usize)
            } else {
                value.to_string()
            }
        }
    }
}

/// Token-aware replacement of an unquoted sheet name in a formula string.
/// Replaces `OldName!` and `'OldName'!` with `NewName!` (quoting if the new
/// name has spaces or special chars). Case-insensitive comparison.
pub fn rewrite_sheet_refs(formula: &str, old: &str, new: &str) -> String {
    if !formula.starts_with('=') {
        return formula.to_string();
    }
    let needs_quotes = new.chars().any(|c| !c.is_ascii_alphanumeric() && c != '_');
    let quoted_new = if needs_quotes {
        format!("'{}'!", new.replace('\'', "''"))
    } else {
        format!("{}!", new)
    };
    // Walk char-by-char, replacing matches at boundaries. A "match" is an
    // identifier-like token immediately followed by `!` whose name equals
    // `old` (case-insensitively). Skip the inside of `"..."` literals so we
    // don't rewrite sheet names that appear inside cell-value strings.
    let chars: Vec<char> = formula.chars().collect();
    let mut out = String::with_capacity(formula.len());
    let mut i = 0;
    let mut in_string = false;
    while i < chars.len() {
        let c = chars[i];
        if in_string {
            // Inside "..."; just copy. Handle doubled "" as an escaped quote.
            out.push(c);
            if c == '"' {
                if i + 1 < chars.len() && chars[i + 1] == '"' {
                    out.push('"');
                    i += 2;
                    continue;
                }
                in_string = false;
            }
            i += 1;
            continue;
        }
        if c == '"' {
            in_string = true;
            out.push(c);
            i += 1;
            continue;
        }
        // Try to match `'<name>'!` first.
        if c == '\'' {
            let end = chars[i + 1..]
                .iter()
                .position(|&x| x == '\'')
                .map(|p| i + 1 + p);
            if let Some(close) = end {
                if close + 1 < chars.len() && chars[close + 1] == '!' {
                    let name: String = chars[i + 1..close].iter().collect();
                    if name.eq_ignore_ascii_case(old) {
                        out.push_str(&quoted_new);
                        i = close + 2;
                        continue;
                    }
                }
            }
        }
        // Try to match a bare identifier followed by `!`.
        if c.is_ascii_alphabetic() || c == '_' {
            let mut j = i;
            while j < chars.len() && (chars[j].is_ascii_alphanumeric() || chars[j] == '_') {
                j += 1;
            }
            if j < chars.len() && chars[j] == '!' {
                let name: String = chars[i..j].iter().collect();
                if name.eq_ignore_ascii_case(old) {
                    out.push_str(&quoted_new);
                    i = j + 1;
                    continue;
                }
            }
        }
        out.push(c);
        i += 1;
    }
    out
}

/// Variant of `rewrite_sheet_refs` for named-range *values* (not formulas).
/// Named-range values look like `Sheet1!A1:B10` (no leading `=`), so we
/// prepend `=` to satisfy the formula-shape check, then strip it.
pub fn rewrite_sheet_refs_for_name_value(value: &str, old: &str, new: &str) -> String {
    let s = format!("={}", value);
    let rewritten = rewrite_sheet_refs(&s, old, new);
    rewritten.strip_prefix('=').map(|x| x.to_string()).unwrap_or(rewritten)
}

/// Merge `over` on top of `base` — boolean flags are OR'd, colors override.
fn layer_style(base: CellStyle, over: &CellStyle) -> CellStyle {
    CellStyle {
        bold: base.bold || over.bold,
        underline: base.underline || over.underline,
        fg_color: over.fg_color.clone().or(base.fg_color),
        bg_color: over.bg_color.clone().or(base.bg_color),
    }
}

fn add_thousands_separator(s: &str) -> String {
    let parts: Vec<&str> = s.splitn(2, '.').collect();
    let int_part = parts[0];
    let negative = int_part.starts_with('-');
    let digits = if negative { &int_part[1..] } else { int_part };

    let mut result = String::new();
    for (i, c) in digits.chars().rev().enumerate() {
        if i > 0 && i % 3 == 0 {
            result.push(',');
        }
        result.push(c);
    }
    let int_formatted: String = result.chars().rev().collect();
    let prefix = if negative { "-" } else { "" };

    if parts.len() > 1 {
        format!("{}{}.{}", prefix, int_formatted, parts[1])
    } else {
        format!("{}{}", prefix, int_formatted)
    }
}

/// A single spreadsheet cell. If `formula` is set, `value` is the evaluated result.
///
/// ```
/// use tshts::domain::CellData;
///
/// let cell = CellData {
///     value: "84".to_string(),
///     formula: Some("=A1*2".to_string()),
///     ..Default::default()
/// };
/// ```
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct CellData {
    /// The display value of the cell (either user input or formula result)
    pub value: String,
    /// Optional formula that generates the value (starts with '=')
    pub formula: Option<String>,
    /// Optional cell format
    pub format: Option<CellFormat>,
    /// Optional cell comment
    pub comment: Option<String>,
    /// When set, this cell is a SPILL ghost — its value derives from the
    /// anchor cell's array result. Edits land at the anchor; clearing the
    /// anchor sweeps these up. `None` for normal cells and the anchor itself.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub spill_anchor: Option<(usize, usize)>,
}

impl Default for CellData {
    fn default() -> Self {
        Self {
            value: String::new(),
            formula: None,
            format: None,
            comment: None,
            spill_anchor: None,
        }
    }
}

impl CellData {
    /// True if this cell is a spill ghost (derived from another cell's
    /// array formula). Ghosts are read-only and cleared when the anchor
    /// changes.
    #[allow(dead_code)]
    pub fn is_spill_ghost(&self) -> bool {
        self.spill_anchor.is_some()
    }
}

/// A grid of cells stored sparsely by `(row, col)` with dependency tracking.
///
/// ```
/// use tshts::domain::{Spreadsheet, CellData};
///
/// let mut sheet = Spreadsheet::default();
/// sheet.set_cell(0, 0, CellData { value: "Hello".to_string(), ..Default::default() });
/// ```
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
        use super::parser::{Parser, Value, ExpressionEvaluator, FunctionRegistry};
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
        let evaluator = if self.named_ranges.is_empty() {
            ExpressionEvaluator::new(self, &registry)
        } else {
            ExpressionEvaluator::with_names(self, &registry, &self.named_ranges)
        };
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
                if let Some(existing) = self.cells.get(&(r, c)) {
                    if existing.spill_anchor != Some((row, col))
                        && (!existing.value.is_empty() || existing.formula.is_some())
                    {
                        if let Some(cd) = self.cells.get_mut(&(row, col)) {
                            cd.value = "#SPILL!".to_string();
                        }
                        return;
                    }
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
        use super::services::FormulaEvaluator;
        let evaluator = if self.named_ranges.is_empty() {
            FormulaEvaluator::new(self)
        } else {
            FormulaEvaluator::with_names(self, &self.named_ranges)
        };
        let dependencies = evaluator.extract_cell_references(formula);
        let cell_pos = (row, col);

        if !dependencies.is_empty() {
            self.dependencies.insert(cell_pos, dependencies.iter().cloned().collect());
            for dep in dependencies {
                self.dependents.entry(dep).or_insert_with(HashSet::new).insert(cell_pos);
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
            if to_recalc.insert(dep) {
                if let Some(next) = self.dependents.get(&dep).cloned() {
                    for n in next {
                        queue.push_back(n);
                    }
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
            if to_recalc.insert(dep) {
                if let Some(next) = self.dependents.get(&dep).cloned() {
                    for n in next {
                        queue.push_back(n);
                    }
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

    /// Recalculates a single cell's value based on its formula.
    fn recalculate_cell(&mut self, row: usize, col: usize) {
        let cell_pos = (row, col);
        if let Some(cell) = self.cells.get(&cell_pos).cloned() {
            if let Some(ref formula) = cell.formula {
                use super::services::FormulaEvaluator;
                let evaluator = if self.named_ranges.is_empty() {
                    FormulaEvaluator::new(self)
                } else {
                    FormulaEvaluator::with_names(self, &self.named_ranges)
                };
                let new_value = evaluator.evaluate_formula(formula);
                let mut updated_cell = cell;
                updated_cell.value = new_value;
                self.set_cell_internal(row, col, updated_cell);
            }
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
    /// Calls `auto_resize_column` for each column in the spreadsheet.
    pub fn auto_resize_all_columns(&mut self) {
        for col in 0..self.cols {
            self.auto_resize_column(col);
        }
    }

    /// Inserts a row at the given index, shifting all rows at or below down by 1.
    /// Updates formula references accordingly.
    pub fn insert_row(&mut self, at: usize) {
        use super::services::FormulaEvaluator;

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
    }

    /// Deletes the row at the given index, shifting all rows below up by 1.
    /// Updates formula references accordingly.
    pub fn delete_row(&mut self, at: usize) {
        use super::services::FormulaEvaluator;

        let mut entries: Vec<((usize, usize), CellData)> = self.cells.drain().collect();
        entries.sort_by(|a, b| a.0.0.cmp(&b.0.0));

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
    }

    /// Inserts a column at the given index, shifting all columns at or to the right by 1.
    /// Updates formula references accordingly.
    pub fn insert_col(&mut self, at: usize) {
        use super::services::FormulaEvaluator;

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
    }

    /// Deletes the column at the given index, shifting all columns to the right left by 1.
    /// Updates formula references accordingly.
    pub fn delete_col(&mut self, at: usize) {
        use super::services::FormulaEvaluator;

        let mut entries: Vec<((usize, usize), CellData)> = self.cells.drain().collect();
        entries.sort_by(|a, b| a.0.1.cmp(&b.0.1));

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
    }

    /// Compute the conditional-format style for the given cell, if any rule
    /// matches. Walks rules in declaration order; later rules layer on top.
    /// `_` in the predicate is replaced with the cell's value (quoted if
    /// non-numeric) before evaluation.
    pub fn conditional_style_for(&self, row: usize, col: usize) -> Option<CellStyle> {
        use super::services::FormulaEvaluator;
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
                FormulaEvaluator::with_names(self, &self.named_ranges)
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

/// A cell address that includes the sheet name. Used as the key type for
/// the workbook-level cross-sheet dependency graph.
pub type CrossSheetKey = (String, usize, usize);

/// A workbook containing multiple spreadsheets (tabs).
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Workbook {
    /// The sheets in this workbook
    pub sheets: Vec<Spreadsheet>,
    /// Names for each sheet
    pub sheet_names: Vec<String>,
    /// Index of the currently active sheet
    pub active_sheet: usize,
    /// Named ranges: name -> cell reference string (e.g., "Revenue" -> "B2:B50")
    pub named_ranges: HashMap<String, String>,
    /// Workbook-level dep graph for CROSS-sheet references only. Same-sheet
    /// deps remain on each Spreadsheet. `dependents[P]` is the set of
    /// cells that reference `P` from a different sheet. Not serialized —
    /// rebuilt on load.
    #[serde(skip)]
    pub cross_sheet_dependents: HashMap<CrossSheetKey, HashSet<CrossSheetKey>>,
    /// Inverse of `cross_sheet_dependents`: `dependencies[X]` is the set of
    /// cells `X` references from other sheets. Used to clear stale deps when
    /// a formula changes.
    #[serde(skip)]
    pub cross_sheet_dependencies: HashMap<CrossSheetKey, HashSet<CrossSheetKey>>,
}

impl Default for Workbook {
    fn default() -> Self {
        Self {
            sheets: vec![Spreadsheet::default()],
            sheet_names: vec!["Sheet1".to_string()],
            active_sheet: 0,
            named_ranges: HashMap::new(),
            cross_sheet_dependents: HashMap::new(),
            cross_sheet_dependencies: HashMap::new(),
        }
    }
}

impl Workbook {
    /// Gets a reference to the active sheet.
    pub fn current_sheet(&self) -> &Spreadsheet {
        &self.sheets[self.active_sheet]
    }

    /// Gets a mutable reference to the active sheet.
    pub fn current_sheet_mut(&mut self) -> &mut Spreadsheet {
        &mut self.sheets[self.active_sheet]
    }

    /// Adds a new empty sheet with the given name.
    pub fn add_sheet(&mut self, name: String) {
        self.sheets.push(Spreadsheet::default());
        self.sheet_names.push(name);
    }

    /// Removes a sheet by index. Adjusts active_sheet if needed.
    /// Won't remove the last sheet.
    pub fn remove_sheet(&mut self, index: usize) -> bool {
        if self.sheets.len() <= 1 || index >= self.sheets.len() {
            return false;
        }
        let removed_name = self.sheet_names[index].clone();
        self.sheets.remove(index);
        self.sheet_names.remove(index);
        if self.active_sheet >= self.sheets.len() {
            self.active_sheet = self.sheets.len() - 1;
        } else if self.active_sheet > index {
            self.active_sheet -= 1;
        }
        // Purge any cross-sheet dep entries that touched the removed sheet.
        self.cross_sheet_dependents
            .retain(|k, _| !k.0.eq_ignore_ascii_case(&removed_name));
        for set in self.cross_sheet_dependents.values_mut() {
            set.retain(|k| !k.0.eq_ignore_ascii_case(&removed_name));
        }
        self.cross_sheet_dependencies
            .retain(|k, _| !k.0.eq_ignore_ascii_case(&removed_name));
        for set in self.cross_sheet_dependencies.values_mut() {
            set.retain(|k| !k.0.eq_ignore_ascii_case(&removed_name));
        }
        true
    }

    /// Renames the active sheet.
    /// Rename the active sheet. Returns false (no-op) if `new_name` is empty
    /// or duplicates another sheet's name (case-insensitive). On success,
    /// rewrites formulas in every sheet AND named-range values that
    /// referenced the old name.
    pub fn rename_sheet(&mut self, new_name: String) -> bool {
        let old_name = self.sheet_names[self.active_sheet].clone();
        if old_name == new_name {
            return true;
        }
        // Reject empty names (unreferenceable in formulas).
        if new_name.trim().is_empty() {
            return false;
        }
        // Reject duplicates against any OTHER sheet (case-insensitive).
        if self
            .sheet_names
            .iter()
            .enumerate()
            .any(|(i, n)| i != self.active_sheet && n.eq_ignore_ascii_case(&new_name))
        {
            return false;
        }
        self.sheet_names[self.active_sheet] = new_name.clone();
        // Rewrite formulas in every sheet.
        for sheet in &mut self.sheets {
            let updates: Vec<(usize, usize, String)> = sheet
                .cells
                .iter()
                .filter_map(|(&(r, c), cd)| {
                    cd.formula
                        .as_ref()
                        .map(|f| (r, c, rewrite_sheet_refs(f, &old_name, &new_name)))
                })
                .filter(|(r, c, new_formula)| {
                    sheet
                        .cells
                        .get(&(*r, *c))
                        .and_then(|cd| cd.formula.as_ref())
                        .map(|old| old != new_formula)
                        .unwrap_or(false)
                })
                .collect();
            for (r, c, formula) in updates {
                if let Some(cd) = sheet.cells.get_mut(&(r, c)) {
                    cd.formula = Some(formula);
                }
            }
            sheet.rebuild_dependencies();
        }
        // Update named-range values too — they often contain sheet-qualified
        // ranges (`Sheet1!A1:B10`) that need the rename.
        let updated: Vec<(String, String)> = self
            .named_ranges
            .iter()
            .map(|(k, v)| (k.clone(), rewrite_sheet_refs_for_name_value(v, &old_name, &new_name)))
            .filter(|(k, v)| self.named_ranges.get(k).map(|orig| orig != v).unwrap_or(false))
            .collect();
        for (k, v) in updated {
            self.named_ranges.insert(k.clone(), v.clone());
            for sheet in &mut self.sheets {
                sheet.named_ranges.insert(k.clone(), v.clone());
            }
        }
        // Cross-sheet dep keys reference sheet names by string; rebuild.
        self.rebuild_cross_sheet_deps();
        true
    }

    /// Define or replace a named range. Keys are normalized to uppercase so
    /// formulas can reference them case-insensitively.
    pub fn set_name(&mut self, name: &str, value: &str) {
        let key = name.to_uppercase();
        self.named_ranges.insert(key.clone(), value.to_string());
        for sheet in &mut self.sheets {
            sheet.named_ranges.insert(key.clone(), value.to_string());
            sheet.rebuild_dependencies();
        }
    }

    /// Remove a named range. Returns true if it existed.
    pub fn remove_name(&mut self, name: &str) -> bool {
        let key = name.to_uppercase();
        let existed = self.named_ranges.remove(&key).is_some();
        for sheet in &mut self.sheets {
            sheet.named_ranges.remove(&key);
            sheet.rebuild_dependencies();
        }
        existed
    }

    /// Creates a Workbook from a single Spreadsheet (for backward compatibility).
    pub fn from_spreadsheet(sheet: Spreadsheet) -> Self {
        Self {
            sheets: vec![sheet],
            sheet_names: vec!["Sheet1".to_string()],
            active_sheet: 0,
            named_ranges: HashMap::new(),
            cross_sheet_dependents: HashMap::new(),
            cross_sheet_dependencies: HashMap::new(),
        }
    }

    /// Re-register the cross-sheet dependencies for the cell at
    /// `(sheet_name, row, col)`. Called after every cell write. Removes
    /// stale entries from the old formula and inserts new ones from the
    /// current one.
    pub fn register_cross_sheet_deps(&mut self, sheet_name: &str, row: usize, col: usize) {
        use super::services::FormulaEvaluator;
        let key: CrossSheetKey = (sheet_name.to_string(), row, col);

        // Step 1: clear old reverse links.
        if let Some(old_precs) = self.cross_sheet_dependencies.remove(&key) {
            for p in old_precs {
                if let Some(set) = self.cross_sheet_dependents.get_mut(&p) {
                    set.remove(&key);
                    if set.is_empty() {
                        self.cross_sheet_dependents.remove(&p);
                    }
                }
            }
        }

        // Step 2: pull the current formula.
        let sheet_idx = match self.sheet_names.iter().position(|n| n == sheet_name) {
            Some(i) => i,
            None => return,
        };
        let formula = match self.sheets[sheet_idx]
            .cells
            .get(&(row, col))
            .and_then(|cd| cd.formula.clone())
        {
            Some(f) => f,
            None => return,
        };

        // Step 3: extract qualified refs from the formula. Use a snapshot
        // so the evaluator can borrow the workbook immutably.
        let qualified_refs: Vec<(Option<String>, usize, usize)> = {
            let names = self.named_ranges.clone();
            let evaluator = FormulaEvaluator::with_workbook(
                self,
                &self.sheets[sheet_idx],
                &names,
            );
            evaluator.extract_qualified_refs(&formula)
        };

        // Step 4: register the cross-sheet ones (skip refs back to the same
        // sheet — those already live in the per-sheet dep graph).
        for (ref_sheet, ref_row, ref_col) in qualified_refs {
            // Skip if no explicit sheet or if it points to the same sheet.
            let resolved_sheet = match ref_sheet {
                Some(s) if !s.eq_ignore_ascii_case(sheet_name) => s,
                _ => continue,
            };
            // Normalize to the canonical sheet-name casing in `sheet_names`.
            let canon = self
                .sheet_names
                .iter()
                .find(|n| n.eq_ignore_ascii_case(&resolved_sheet))
                .cloned()
                .unwrap_or(resolved_sheet);
            let prec: CrossSheetKey = (canon, ref_row, ref_col);
            self.cross_sheet_dependencies
                .entry(key.clone())
                .or_insert_with(HashSet::new)
                .insert(prec.clone());
            self.cross_sheet_dependents
                .entry(prec)
                .or_insert_with(HashSet::new)
                .insert(key.clone());
        }
    }

    /// Recalculate every cell that depends on `(sheet_name, row, col)` via
    /// the cross-sheet graph. Walks transitively (BFS) so chains like
    /// Sheet1!A1 → Sheet2!A1 → Sheet3!A1 all update.
    pub fn propagate_cross_sheet_changes(
        &mut self,
        sheet_name: &str,
        row: usize,
        col: usize,
    ) {
        use super::services::FormulaEvaluator;
        let mut queue: std::collections::VecDeque<CrossSheetKey> =
            std::collections::VecDeque::new();
        queue.push_back((sheet_name.to_string(), row, col));
        let mut visited: HashSet<CrossSheetKey> = HashSet::new();

        while let Some(key) = queue.pop_front() {
            if !visited.insert(key.clone()) {
                continue;
            }
            let deps: Vec<CrossSheetKey> = self
                .cross_sheet_dependents
                .get(&key)
                .cloned()
                .unwrap_or_default()
                .into_iter()
                .collect();
            if deps.is_empty() {
                continue;
            }
            // Snapshot the workbook so the evaluator can borrow it while
            // we mutate `self.sheets`. Cloning is heavy but only happens
            // when something cross-sheet actually fires; for the common
            // case of no cross-sheet deps, this is never reached.
            let snapshot = self.clone();
            let names = snapshot.named_ranges.clone();
            for dep in deps {
                let (dep_sheet, dep_row, dep_col) = dep.clone();
                let Some(dep_idx) = snapshot
                    .sheet_names
                    .iter()
                    .position(|n| n.eq_ignore_ascii_case(&dep_sheet))
                else {
                    continue;
                };
                let snap_sheet = &snapshot.sheets[dep_idx];
                let Some(cd) = snap_sheet.cells.get(&(dep_row, dep_col)) else {
                    continue;
                };
                let Some(formula) = cd.formula.clone() else { continue };
                let evaluator =
                    FormulaEvaluator::with_workbook(&snapshot, snap_sheet, &names);
                let new_value = evaluator.evaluate_formula(&formula);
                // Write to the real workbook.
                let dep_real_idx = self
                    .sheet_names
                    .iter()
                    .position(|n| n.eq_ignore_ascii_case(&dep_sheet));
                if let Some(idx) = dep_real_idx {
                    if let Some(real_cd) = self.sheets[idx].cells.get_mut(&(dep_row, dep_col))
                    {
                        if real_cd.value != new_value {
                            real_cd.value = new_value;
                            queue.push_back(dep);
                        }
                    }
                }
            }
        }
    }

    /// Check whether adding a new formula at `(sheet_name, row, col)` with
    /// the given precedents would create a cross-sheet cycle. Walks the
    /// existing cross-sheet graph from each precedent; if any path reaches
    /// `(sheet_name, row, col)`, we'd loop. The same-sheet check still runs
    /// separately via `FormulaEvaluator::would_create_circular_reference`.
    pub fn would_create_cross_sheet_cycle(
        &self,
        sheet_name: &str,
        row: usize,
        col: usize,
        candidate_precedents: &[(Option<String>, usize, usize)],
    ) -> bool {
        let target: CrossSheetKey = (sheet_name.to_string(), row, col);
        let mut stack: Vec<CrossSheetKey> = Vec::new();
        for (prec_sheet, prec_row, prec_col) in candidate_precedents {
            // Only consider cross-sheet precedents (same-sheet cycles are
            // caught by the existing AST walker).
            let Some(ps) = prec_sheet else { continue };
            if ps.eq_ignore_ascii_case(sheet_name) {
                continue;
            }
            let canon = self
                .sheet_names
                .iter()
                .find(|n| n.eq_ignore_ascii_case(ps))
                .cloned()
                .unwrap_or_else(|| ps.clone());
            stack.push((canon, *prec_row, *prec_col));
        }
        let mut visited: HashSet<CrossSheetKey> = HashSet::new();
        while let Some(node) = stack.pop() {
            if node == target {
                return true;
            }
            if !visited.insert(node.clone()) {
                continue;
            }
            // Also walk down: the cells that *node* in turn depends on.
            if let Some(deps) = self.cross_sheet_dependencies.get(&node) {
                for d in deps {
                    stack.push(d.clone());
                }
            }
        }
        false
    }

    /// Rebuild the cross-sheet dep graph from scratch by scanning every
    /// formula in every sheet. Called after load (since the graph isn't
    /// serialized) and as a fallback when state drifts.
    pub fn rebuild_cross_sheet_deps(&mut self) {
        self.cross_sheet_dependents.clear();
        self.cross_sheet_dependencies.clear();
        let cells: Vec<(String, usize, usize)> = self
            .sheet_names
            .iter()
            .enumerate()
            .flat_map(|(idx, name)| {
                self.sheets[idx]
                    .cells
                    .iter()
                    .filter(|(_, cd)| cd.formula.is_some())
                    .map(move |(&(r, c), _)| (name.clone(), r, c))
            })
            .collect();
        for (sheet, r, c) in cells {
            self.register_cross_sheet_deps(&sheet, r, c);
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

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_conditional_format_fires_on_truthy_predicate() {
        let mut sheet = Spreadsheet::default();
        sheet.set_cell(0, 0, CellData {
            value: "150".to_string(), formula: None, format: None, comment: None,
        spill_anchor: None,
        });
        sheet.set_cell(1, 0, CellData {
            value: "50".to_string(), formula: None, format: None, comment: None,
        spill_anchor: None,
        });
        sheet.conditional_formats.push(ConditionalFormat {
            column: 0,
            predicate: "_ > 100".to_string(),
            style: CellStyle {
                bold: true,
                underline: false,
                fg_color: Some(TerminalColor::Red),
                bg_color: None,
            },
        });
        let s0 = sheet.conditional_style_for(0, 0);
        assert!(s0.is_some());
        assert!(s0.as_ref().unwrap().bold);
        assert_eq!(s0.unwrap().fg_color, Some(TerminalColor::Red));
        // Row 1 doesn't satisfy the predicate.
        assert!(sheet.conditional_style_for(1, 0).is_none());
    }

    #[test]
    fn test_thousands_separator_edge_cases() {
        let fmt = CellFormat {
            number_format: NumberFormat::Number { decimals: 2, thousands_sep: true },
            style: CellStyle::default(),
        };
        assert_eq!(format_cell_value("1234567.89", &fmt), "1,234,567.89");
        assert_eq!(format_cell_value("-1234.5", &fmt), "-1,234.50");
        assert_eq!(format_cell_value("999.99", &fmt), "999.99");
        assert_eq!(format_cell_value("0", &fmt), "0.00");
        assert_eq!(format_cell_value("-0.5", &fmt), "-0.50");

        // Whole-million boundary
        let fmt0 = CellFormat {
            number_format: NumberFormat::Number { decimals: 0, thousands_sep: true },
            style: CellStyle::default(),
        };
        assert_eq!(format_cell_value("1000000", &fmt0), "1,000,000");
        assert_eq!(format_cell_value("-1000000", &fmt0), "-1,000,000");
    }

    #[test]
    fn test_cell_data_default() {
        let cell = CellData::default();
        assert!(cell.value.is_empty());
        assert!(cell.formula.is_none());
    }

    #[test]
    fn test_cell_data_creation() {
        let cell = CellData {
            value: "42".to_string(),
            formula: Some("=6*7".to_string()),
            format: None,
            comment: None,
        spill_anchor: None,
        };
        assert_eq!(cell.value, "42");
        assert_eq!(cell.formula.unwrap(), "=6*7");
    }

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

    // === Number Formatting Tests ===

    #[test]
    fn test_format_cell_value_general() {
        let fmt = CellFormat { number_format: NumberFormat::General, ..CellFormat::default() };
        assert_eq!(super::format_cell_value("42.5", &fmt), "42.5");
        assert_eq!(super::format_cell_value("hello", &fmt), "hello");
    }

    #[test]
    fn test_format_cell_value_number() {
        let fmt = CellFormat { number_format: NumberFormat::Number { decimals: 2, thousands_sep: false }, ..CellFormat::default() };
        assert_eq!(super::format_cell_value("42", &fmt), "42.00");
        assert_eq!(super::format_cell_value("3.14159", &fmt), "3.14");
        assert_eq!(super::format_cell_value("hello", &fmt), "hello"); // non-numeric passthrough
    }

    #[test]
    fn test_format_cell_value_number_thousands() {
        let fmt = CellFormat { number_format: NumberFormat::Number { decimals: 2, thousands_sep: true }, ..CellFormat::default() };
        assert_eq!(super::format_cell_value("1234567.89", &fmt), "1,234,567.89");
        assert_eq!(super::format_cell_value("42", &fmt), "42.00");
    }

    #[test]
    fn test_format_cell_value_currency() {
        let fmt = CellFormat { number_format: NumberFormat::Currency { symbol: "$".to_string(), decimals: 2 }, ..CellFormat::default() };
        assert_eq!(super::format_cell_value("1234.5", &fmt), "$1,234.50");
        assert_eq!(super::format_cell_value("42", &fmt), "$42.00");
        assert_eq!(super::format_cell_value("hello", &fmt), "hello");
    }

    #[test]
    fn test_format_cell_value_percentage() {
        let fmt = CellFormat { number_format: NumberFormat::Percentage { decimals: 1 }, ..CellFormat::default() };
        assert_eq!(super::format_cell_value("0.75", &fmt), "75.0%");
        assert_eq!(super::format_cell_value("1", &fmt), "100.0%");
        assert_eq!(super::format_cell_value("0.123", &fmt), "12.3%");
    }

    #[test]
    fn test_thousands_separator() {
        assert_eq!(super::add_thousands_separator("1234567"), "1,234,567");
        assert_eq!(super::add_thousands_separator("123"), "123");
        assert_eq!(super::add_thousands_separator("1234.56"), "1,234.56");
        assert_eq!(super::add_thousands_separator("-1234567"), "-1,234,567");
    }

    // === Insert/Delete Row/Col Tests ===

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

    #[test]
    fn test_cell_data_format_serialization() {
        let cell = CellData {
            value: "100".to_string(),
            formula: None,
            format: Some(CellFormat {
                number_format: NumberFormat::Percentage { decimals: 1 },
                ..CellFormat::default()
            }),
            comment: None,
        spill_anchor: None,
        };
        let json = serde_json::to_string(&cell).unwrap();
        let deserialized: CellData = serde_json::from_str(&json).unwrap();
        assert_eq!(deserialized.value, "100");
        assert!(deserialized.format.is_some());
        assert!(matches!(deserialized.format.unwrap().number_format, NumberFormat::Percentage { decimals: 1 }));
    }
}