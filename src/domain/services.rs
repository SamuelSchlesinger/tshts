//! Formula evaluation services for the terminal spreadsheet.
//!
//! This module provides the core formula evaluation engine that can
//! parse and execute spreadsheet formulas with cell references,
//! arithmetic operations, and built-in functions.

use super::models::Spreadsheet;
use super::parser::{Parser, ExpressionEvaluator, FunctionRegistry, Expr, Value};
use std::collections::HashSet;
use std::fs::File;

/// A formula evaluation engine that processes spreadsheet expressions.
///
/// The evaluator uses a modern recursive descent parser with formal BNF grammar.
/// All logical operations are implemented as functions for consistency and extensibility.
///
/// Supported features:
/// - Arithmetic operations: +, -, *, /, **, ^, %
/// - Comparison operators: <, >, <=, >=, =, <>
/// - Comprehensive functions: SUM, AVERAGE, MIN, MAX, IF, AND, OR, NOT, ABS, SQRT, ROUND
/// - Cell references: A1, B2, AA123, etc.
/// - Cell ranges: A1:C3, B1:B10, etc.
/// - Circular reference detection with AST analysis
/// - Extensible function registry for custom functions
///
/// # Examples
///
/// ```
/// use tshts::domain::{Spreadsheet, FormulaEvaluator};
///
/// let sheet = Spreadsheet::default();
/// let evaluator = FormulaEvaluator::new(&sheet);
/// 
/// // Arithmetic operations
/// assert_eq!(evaluator.evaluate_formula("=2+3*4"), "14");
/// 
/// // Logical functions (not operators)
/// assert_eq!(evaluator.evaluate_formula("=AND(1>0, 2<5)"), "1");
/// assert_eq!(evaluator.evaluate_formula("=OR(1>2, 3<5)"), "1");
/// assert_eq!(evaluator.evaluate_formula("=NOT(0)"), "1");
/// 
/// // Built-in functions with ranges
/// assert_eq!(evaluator.evaluate_formula("=SUM(A1:A3)"), "0");
/// ```
pub struct FormulaEvaluator<'a> {
    /// Reference to the spreadsheet for cell lookups
    spreadsheet: &'a Spreadsheet,
}

impl<'a> FormulaEvaluator<'a> {
    /// Creates a new formula evaluator for the given spreadsheet.
    ///
    /// # Arguments
    ///
    /// * `spreadsheet` - Reference to the spreadsheet for cell lookups
    ///
    /// # Examples
    ///
    /// ```
    /// use tshts::domain::{Spreadsheet, FormulaEvaluator};
    ///
    /// let sheet = Spreadsheet::default();
    /// let evaluator = FormulaEvaluator::new(&sheet);
    /// ```
    pub fn new(spreadsheet: &'a Spreadsheet) -> Self {
        Self { spreadsheet }
    }

    /// Evaluates a formula and returns the result as a string.
    ///
    /// Uses a recursive descent parser with formal BNF grammar to parse and evaluate
    /// expressions. All logical operations (AND, OR, NOT) are implemented as functions.
    /// Formulas must start with '=' to be recognized as formulas.
    /// Non-formula strings are returned unchanged.
    ///
    /// # Arguments
    ///
    /// * `formula` - Formula string to evaluate (e.g., "=A1+B1", "=AND(A1>0,B1<10)")
    ///
    /// # Returns
    ///
    /// String representation of the evaluation result, or "#ERROR" if evaluation fails
    ///
    /// # Examples
    ///
    /// ```
    /// use tshts::domain::{Spreadsheet, FormulaEvaluator};
    ///
    /// let sheet = Spreadsheet::default();
    /// let evaluator = FormulaEvaluator::new(&sheet);
    ///
    /// assert_eq!(evaluator.evaluate_formula("=2+3"), "5");
    /// assert_eq!(evaluator.evaluate_formula("=AND(1,1)"), "1");
    /// assert_eq!(evaluator.evaluate_formula("hello"), "hello");
    /// ```
    pub fn evaluate_formula(&self, formula: &str) -> String {
        if formula.starts_with('=') {
            let expr = &formula[1..];
            
            match self.parse_and_evaluate(expr) {
                Ok(result) => result.to_string(),
                Err(_) => "#ERROR".to_string(),
            }
        } else {
            formula.to_string()
        }
    }

    /// Checks if a formula would create a circular reference using AST analysis.
    ///
    /// A circular reference occurs when a cell's formula directly or indirectly
    /// references itself, which would cause infinite recursion during evaluation.
    /// This method parses the formula into an AST and analyzes all cell references.
    ///
    /// # Arguments
    ///
    /// * `formula` - Formula to check for circular references
    /// * `current_cell` - Coordinates of the cell that would contain this formula
    ///
    /// # Returns
    ///
    /// `true` if the formula would create a circular reference, `false` otherwise
    ///
    /// # Examples
    ///
    /// ```
    /// use tshts::domain::{Spreadsheet, FormulaEvaluator};
    ///
    /// let sheet = Spreadsheet::default();
    /// let evaluator = FormulaEvaluator::new(&sheet);
    ///
    /// // This would be circular: A1 referring to itself
    /// assert!(evaluator.would_create_circular_reference("=A1+1", (0, 0)));
    /// // This is fine: A1 referring to B1
    /// assert!(!evaluator.would_create_circular_reference("=AND(B1>0,C1<10)", (0, 0)));
    /// ```
    pub fn would_create_circular_reference(&self, formula: &str, current_cell: (usize, usize)) -> bool {
        if !formula.starts_with('=') {
            return false;
        }
        
        let expr = &formula[1..];
        match Parser::new(expr) {
            Ok(mut parser) => {
                match parser.parse() {
                    Ok(ast) => self.check_circular_reference_in_ast(&ast, current_cell, &mut HashSet::new()),
                    Err(_) => false, // If we can't parse it, assume it's not circular
                }
            }
            Err(_) => false,
        }
    }

    /// Parses and evaluates an expression using the new parser.
    fn parse_and_evaluate(&self, expr: &str) -> Result<Value, String> {
        let mut parser = Parser::new(expr)?;
        let ast = parser.parse()?;
        
        let function_registry = FunctionRegistry::new();
        let evaluator = ExpressionEvaluator::new(self.spreadsheet, &function_registry);
        evaluator.evaluate(&ast)
    }
    
    /// Checks for circular references in a formula string, reusing the visited set.
    fn check_circular_in_formula(&self, formula: &str, target_cell: (usize, usize), visited: &mut HashSet<(usize, usize)>) -> bool {
        if !formula.starts_with('=') {
            return false;
        }
        let expr = &formula[1..];
        match Parser::new(expr) {
            Ok(mut parser) => match parser.parse() {
                Ok(ast) => self.check_circular_reference_in_ast(&ast, target_cell, visited),
                Err(_) => false,
            },
            Err(_) => false,
        }
    }

    /// Checks for circular references in an AST.
    fn check_circular_reference_in_ast(&self, expr: &Expr, target_cell: (usize, usize), visited: &mut HashSet<(usize, usize)>) -> bool {
        match expr {
            Expr::String(_) => false,
            Expr::CellRef(cell_ref) => {
                if let Some((row, col)) = Spreadsheet::parse_cell_reference(cell_ref) {
                    if (row, col) == target_cell {
                        return true;
                    }

                    if visited.contains(&(row, col)) {
                        return false;
                    }

                    visited.insert((row, col));

                    let cell = self.spreadsheet.get_cell(row, col);
                    if let Some(ref cell_formula) = cell.formula {
                        if self.check_circular_in_formula(cell_formula, target_cell, visited) {
                            return true;
                        }
                    }
                }
                false
            }
            Expr::Range(start_cell, end_cell) => {
                // Check both start and end cells of the range
                if let (Some((start_row, start_col)), Some((end_row, end_col))) = 
                    (Spreadsheet::parse_cell_reference(start_cell), Spreadsheet::parse_cell_reference(end_cell)) {
                    for row in start_row..=end_row {
                        for col in start_col..=end_col {
                            if (row, col) == target_cell {
                                return true;
                            }
                        }
                    }
                }
                false
            }
            Expr::Binary { left, right, .. } => {
                self.check_circular_reference_in_ast(left, target_cell, visited) ||
                self.check_circular_reference_in_ast(right, target_cell, visited)
            }
            Expr::Unary { operand, .. } => {
                self.check_circular_reference_in_ast(operand, target_cell, visited)
            }
            Expr::FunctionCall { args, .. } => {
                args.iter().any(|arg| self.check_circular_reference_in_ast(arg, target_cell, visited))
            }
            Expr::Number(_) => false,
        }
    }

    /// Extracts all cell references from a formula string.
    ///
    /// Parses the formula and analyzes its AST to find all cell references.
    /// This is used for dependency tracking and automatic recalculation.
    ///
    /// # Arguments
    ///
    /// * `formula` - Formula string to analyze (should start with '=')
    ///
    /// # Returns
    ///
    /// Vector of (row, col) tuples representing the referenced cells
    pub fn extract_cell_references(&self, formula: &str) -> Vec<(usize, usize)> {
        if !formula.starts_with('=') {
            return Vec::new();
        }
        
        let expr = &formula[1..];
        match Parser::new(expr) {
            Ok(mut parser) => {
                match parser.parse() {
                    Ok(ast) => self.extract_cell_references_from_ast(&ast),
                    Err(_) => Vec::new(),
                }
            }
            Err(_) => Vec::new(),
        }
    }

    /// Extracts all cell references from an AST.
    ///
    /// This is a utility method for analyzing formula dependencies.
    ///
    /// # Arguments
    ///
    /// * `expr` - Expression AST to analyze
    ///
    /// # Returns
    ///
    /// Vector of (row, col) tuples representing the referenced cells
    fn extract_cell_references_from_ast(&self, expr: &Expr) -> Vec<(usize, usize)> {
        let mut references = Vec::new();
        
        match expr {
            Expr::String(_) => {},
            Expr::CellRef(cell_ref) => {
                if let Some((row, col)) = Spreadsheet::parse_cell_reference(cell_ref) {
                    references.push((row, col));
                }
            }
            Expr::Range(start_cell, end_cell) => {
                if let (Some((start_row, start_col)), Some((end_row, end_col))) = 
                    (Spreadsheet::parse_cell_reference(start_cell), Spreadsheet::parse_cell_reference(end_cell)) {
                    for row in start_row..=end_row {
                        for col in start_col..=end_col {
                            references.push((row, col));
                        }
                    }
                }
            }
            Expr::Binary { left, right, .. } => {
                references.extend(self.extract_cell_references_from_ast(left));
                references.extend(self.extract_cell_references_from_ast(right));
            }
            Expr::Unary { operand, .. } => {
                references.extend(self.extract_cell_references_from_ast(operand));
            }
            Expr::FunctionCall { args, .. } => {
                for arg in args {
                    references.extend(self.extract_cell_references_from_ast(arg));
                }
            }
            Expr::Number(_) => {}
        }
        
        references
    }

    /// Adjusts formula with relative references when copying to a new position.
    ///
    /// This method takes a formula and adjusts all cell references to maintain
    /// their relative positions when the formula is moved to a new location.
    ///
    /// # Arguments
    ///
    /// * `formula` - Original formula string (should start with '=')
    /// * `row_offset` - Row offset (positive = down, negative = up)
    /// * `col_offset` - Column offset (positive = right, negative = left)
    ///
    /// # Returns
    ///
    /// Adjusted formula string with updated cell references
    ///
    /// # Examples
    ///
    /// ```
    /// use tshts::domain::{Spreadsheet, FormulaEvaluator};
    ///
    /// let sheet = Spreadsheet::default();
    /// let evaluator = FormulaEvaluator::new(&sheet);
    /// 
    /// // Moving formula =SUM(B4:B6) one column to the right
    /// let adjusted = evaluator.adjust_formula_references("=SUM(B4:B6)", 0, 1);
    /// assert_eq!(adjusted, "=SUM(C4:C6)");
    /// 
    /// // Moving formula =A1+B1 one row down
    /// let adjusted = evaluator.adjust_formula_references("=A1+B1", 1, 0);
    /// assert_eq!(adjusted, "=A2+B2");
    /// ```
    pub fn adjust_formula_references(&self, formula: &str, row_offset: i32, col_offset: i32) -> String {
        if !formula.starts_with('=') {
            return formula.to_string();
        }
        
        let expr = &formula[1..];
        match Parser::new(expr) {
            Ok(mut parser) => {
                match parser.parse() {
                    Ok(ast) => {
                        let adjusted_ast = self.adjust_ast_references(&ast, row_offset, col_offset);
                        format!("={}", self.ast_to_string(&adjusted_ast))
                    }
                    Err(_) => formula.to_string(),
                }
            }
            Err(_) => formula.to_string(),
        }
    }

    /// Adjusts cell references in an AST with the given offsets.
    fn adjust_ast_references(&self, expr: &Expr, row_offset: i32, col_offset: i32) -> Expr {
        match expr {
            Expr::String(s) => Expr::String(s.clone()),
            Expr::CellRef(cell_ref) => {
                if let Some((row, col)) = Spreadsheet::parse_cell_reference(cell_ref) {
                    let new_row = (row as i32 + row_offset).max(0) as usize;
                    let new_col = (col as i32 + col_offset).max(0) as usize;
                    let new_ref = format!("{}{}", Spreadsheet::column_label(new_col), new_row + 1);
                    Expr::CellRef(new_ref)
                } else {
                    expr.clone()
                }
            }
            Expr::Range(start_ref, end_ref) => {
                let new_start = if let Some((row, col)) = Spreadsheet::parse_cell_reference(start_ref) {
                    let new_row = (row as i32 + row_offset).max(0) as usize;
                    let new_col = (col as i32 + col_offset).max(0) as usize;
                    format!("{}{}", Spreadsheet::column_label(new_col), new_row + 1)
                } else {
                    start_ref.clone()
                };
                let new_end = if let Some((row, col)) = Spreadsheet::parse_cell_reference(end_ref) {
                    let new_row = (row as i32 + row_offset).max(0) as usize;
                    let new_col = (col as i32 + col_offset).max(0) as usize;
                    format!("{}{}", Spreadsheet::column_label(new_col), new_row + 1)
                } else {
                    end_ref.clone()
                };
                Expr::Range(new_start, new_end)
            }
            Expr::Binary { left, operator, right } => {
                Expr::Binary {
                    left: Box::new(self.adjust_ast_references(left, row_offset, col_offset)),
                    operator: operator.clone(),
                    right: Box::new(self.adjust_ast_references(right, row_offset, col_offset)),
                }
            }
            Expr::Unary { operator, operand } => {
                Expr::Unary {
                    operator: operator.clone(),
                    operand: Box::new(self.adjust_ast_references(operand, row_offset, col_offset)),
                }
            }
            Expr::FunctionCall { name, args } => {
                let adjusted_args: Vec<Expr> = args.iter()
                    .map(|arg| self.adjust_ast_references(arg, row_offset, col_offset))
                    .collect();
                Expr::FunctionCall {
                    name: name.clone(),
                    args: adjusted_args,
                }
            }
            Expr::Number(n) => Expr::Number(*n),
        }
    }

    /// Converts an AST back to a formula string.
    fn ast_to_string(&self, expr: &Expr) -> String {
        match expr {
            Expr::String(s) => format!("\"{}\"", s.replace("\"", "\"\"")),
            Expr::CellRef(cell_ref) => cell_ref.clone(),
            Expr::Range(start_ref, end_ref) => format!("{}:{}", start_ref, end_ref),
            Expr::Binary { left, operator, right } => {
                format!("{}{}{}",
                    self.ast_to_string(left),
                    self.binary_op_to_string(operator),
                    self.ast_to_string(right))
            }
            Expr::Unary { operator, operand } => {
                format!("{}{}", self.unary_op_to_string(operator), self.ast_to_string(operand))
            }
            Expr::FunctionCall { name, args } => {
                let arg_strs: Vec<String> = args.iter()
                    .map(|arg| self.ast_to_string(arg))
                    .collect();
                format!("{}({})", name, arg_strs.join(","))
            }
            Expr::Number(n) => n.to_string(),
        }
    }

    /// Converts a binary operator enum back to a string.
    fn binary_op_to_string(&self, op: &super::parser::BinaryOp) -> &'static str {
        use super::parser::BinaryOp;
        match op {
            BinaryOp::Add => "+",
            BinaryOp::Subtract => "-",
            BinaryOp::Multiply => "*",
            BinaryOp::Divide => "/",
            BinaryOp::Power => "**",
            BinaryOp::Modulo => "%",
            BinaryOp::Equal => "=",
            BinaryOp::NotEqual => "<>",
            BinaryOp::Less => "<",
            BinaryOp::LessEqual => "<=",
            BinaryOp::Greater => ">",
            BinaryOp::GreaterEqual => ">=",
            BinaryOp::Concatenate => "&",
        }
    }

    /// Converts a unary operator enum back to a string.
    fn unary_op_to_string(&self, op: &super::parser::UnaryOp) -> &'static str {
        use super::parser::UnaryOp;
        match op {
            UnaryOp::Minus => "-",
            UnaryOp::Plus => "+",
        }
    }

}

// Known sequences for autofill pattern recognition
const DAYS_SHORT: &[&str] = &["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"];
const DAYS_FULL: &[&str] = &["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"];
const MONTHS_SHORT: &[&str] = &["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
const MONTHS_FULL: &[&str] = &["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
const QUARTERS: &[&str] = &["Q1", "Q2", "Q3", "Q4"];

/// Detected autofill pattern for smart sequence continuation.
#[derive(Debug, Clone, PartialEq)]
pub enum AutofillPattern {
    /// Arithmetic sequence: start + step * index
    Arithmetic { start: f64, step: f64 },
    /// Text prefix with numeric suffix following arithmetic pattern
    PrefixedNumber { prefix: String, suffix: String, start: f64, step: f64 },
    /// Known sequence (days, months, quarters) with wrap-around
    KnownSequence { sequence: Vec<String>, start_index: usize },
    /// Simple copy of single value
    Copy { value: String },
}

impl AutofillPattern {
    /// Detect pattern from non-empty cell values.
    ///
    /// Analyzes the given values and returns the detected pattern type.
    /// Pattern detection priority: Arithmetic > KnownSequence > PrefixedNumber > Copy
    ///
    /// Known sequences are checked before prefixed numbers so that "Q1, Q2"
    /// is detected as quarters rather than prefix "Q" with numbers 1, 2.
    pub fn detect(values: &[String]) -> Self {
        if values.is_empty() {
            return AutofillPattern::Copy { value: String::new() };
        }

        if values.len() == 1 {
            return AutofillPattern::Copy { value: values[0].clone() };
        }

        // Try arithmetic pattern first (all numeric values)
        if let Some((start, step)) = Self::try_parse_arithmetic(values) {
            return AutofillPattern::Arithmetic { start, step };
        }

        // Try known sequence pattern (days, months, quarters) BEFORE prefixed numbers
        // This ensures Q1, Q2 is detected as quarters, not "Q" + 1, 2
        if let Some((sequence, start_index)) = Self::try_match_known_sequence(values) {
            return AutofillPattern::KnownSequence { sequence, start_index };
        }

        // Try prefixed number pattern (e.g., "Item1", "Item2")
        if let Some((prefix, suffix, start, step)) = Self::try_parse_prefixed_number(values) {
            return AutofillPattern::PrefixedNumber { prefix, suffix, start, step };
        }

        // Fallback to copy
        AutofillPattern::Copy { value: values[0].clone() }
    }

    /// Generate value at given index (0-based from start of pattern).
    pub fn generate(&self, index: usize) -> String {
        match self {
            AutofillPattern::Arithmetic { start, step } => {
                let value = start + step * (index as f64);
                Self::format_number(value)
            }
            AutofillPattern::PrefixedNumber { prefix, suffix, start, step } => {
                let num = start + step * (index as f64);
                format!("{}{}{}", prefix, Self::format_number(num), suffix)
            }
            AutofillPattern::KnownSequence { sequence, start_index } => {
                let idx = (start_index + index) % sequence.len();
                sequence[idx].clone()
            }
            AutofillPattern::Copy { value } => value.clone(),
        }
    }

    /// Return description for status message.
    pub fn description(&self) -> String {
        match self {
            AutofillPattern::Arithmetic { step, .. } => {
                if *step >= 0.0 {
                    format!("arithmetic sequence (+{})", Self::format_number(*step))
                } else {
                    format!("arithmetic sequence ({})", Self::format_number(*step))
                }
            }
            AutofillPattern::PrefixedNumber { prefix, step, .. } => {
                format!("\"{}...\" sequence (+{})", prefix, Self::format_number(*step))
            }
            AutofillPattern::KnownSequence { sequence, .. } => {
                if sequence == &DAYS_SHORT.iter().map(|s| s.to_string()).collect::<Vec<_>>()
                    || sequence == &DAYS_FULL.iter().map(|s| s.to_string()).collect::<Vec<_>>() {
                    "days sequence".to_string()
                } else if sequence == &MONTHS_SHORT.iter().map(|s| s.to_string()).collect::<Vec<_>>()
                    || sequence == &MONTHS_FULL.iter().map(|s| s.to_string()).collect::<Vec<_>>() {
                    "months sequence".to_string()
                } else if sequence == &QUARTERS.iter().map(|s| s.to_string()).collect::<Vec<_>>() {
                    "quarters sequence".to_string()
                } else {
                    "known sequence".to_string()
                }
            }
            AutofillPattern::Copy { .. } => "copy".to_string(),
        }
    }

    /// Try to parse values as an arithmetic sequence.
    /// Returns (start, step) if successful.
    fn try_parse_arithmetic(values: &[String]) -> Option<(f64, f64)> {
        if values.len() < 2 {
            return None;
        }

        // Try to parse all values as numbers
        let nums: Option<Vec<f64>> = values.iter()
            .map(|s| s.trim().parse::<f64>().ok())
            .collect();

        let nums = nums?;

        // Check if differences are constant (within epsilon)
        let step = nums[1] - nums[0];
        const EPSILON: f64 = 1e-9;

        for i in 2..nums.len() {
            let diff = nums[i] - nums[i - 1];
            if (diff - step).abs() > EPSILON {
                return None;
            }
        }

        Some((nums[0], step))
    }

    /// Try to parse values as prefixed numbers (e.g., "Item1", "Item2").
    /// Returns (prefix, suffix, start, step) if successful.
    fn try_parse_prefixed_number(values: &[String]) -> Option<(String, String, f64, f64)> {
        if values.len() < 2 {
            return None;
        }

        // Extract prefix, number, and suffix from each value
        let mut parsed: Vec<(String, f64, String)> = Vec::new();

        for val in values {
            if let Some((prefix, num, suffix)) = Self::split_prefixed_number(val) {
                parsed.push((prefix, num, suffix));
            } else {
                return None;
            }
        }

        // Check that all prefixes and suffixes are the same
        let first_prefix = &parsed[0].0;
        let first_suffix = &parsed[0].2;

        for (prefix, _, suffix) in &parsed {
            if prefix != first_prefix || suffix != first_suffix {
                return None;
            }
        }

        // Extract numbers and check for arithmetic pattern
        let nums: Vec<f64> = parsed.iter().map(|(_, n, _)| *n).collect();
        let step = nums[1] - nums[0];
        const EPSILON: f64 = 1e-9;

        for i in 2..nums.len() {
            let diff = nums[i] - nums[i - 1];
            if (diff - step).abs() > EPSILON {
                return None;
            }
        }

        Some((first_prefix.clone(), first_suffix.clone(), nums[0], step))
    }

    /// Split a string into (prefix, number, suffix).
    /// E.g., "Item10" -> ("Item", 10.0, ""), "Row_5_data" -> ("Row_", 5.0, "_data")
    fn split_prefixed_number(s: &str) -> Option<(String, f64, String)> {
        // Find the first digit
        let first_digit = s.chars().position(|c| c.is_ascii_digit())?;

        // Find where the number ends
        let after_number = s[first_digit..].chars()
            .position(|c| !c.is_ascii_digit() && c != '.' && c != '-')
            .map(|i| first_digit + i)
            .unwrap_or(s.len());

        let prefix = &s[..first_digit];
        let num_str = &s[first_digit..after_number];
        let suffix = &s[after_number..];

        let num = num_str.parse::<f64>().ok()?;

        Some((prefix.to_string(), num, suffix.to_string()))
    }

    /// Try to match values against known sequences (days, months, quarters).
    /// Returns (sequence, start_index) if successful.
    fn try_match_known_sequence(values: &[String]) -> Option<(Vec<String>, usize)> {
        let all_sequences: &[&[&str]] = &[
            DAYS_SHORT, DAYS_FULL, MONTHS_SHORT, MONTHS_FULL, QUARTERS
        ];

        for seq in all_sequences {
            if let Some(start_idx) = Self::match_sequence(values, seq) {
                let owned_seq: Vec<String> = seq.iter().map(|s| s.to_string()).collect();
                return Some((owned_seq, start_idx));
            }
        }

        None
    }

    /// Check if values match a sequence starting at some index.
    /// Returns the starting index if matched.
    fn match_sequence(values: &[String], sequence: &[&str]) -> Option<usize> {
        if values.is_empty() || values.len() > sequence.len() {
            return None;
        }

        // Find where the first value matches in the sequence (case-insensitive)
        let first_lower = values[0].to_lowercase();
        let start_idx = sequence.iter()
            .position(|s| s.to_lowercase() == first_lower)?;

        // Check if all subsequent values match the sequence
        for (i, val) in values.iter().enumerate() {
            let seq_idx = (start_idx + i) % sequence.len();
            if val.to_lowercase() != sequence[seq_idx].to_lowercase() {
                return None;
            }
        }

        Some(start_idx)
    }

    /// Format a number smartly: show as integer if whole, otherwise as decimal.
    fn format_number(n: f64) -> String {
        if n.fract().abs() < 1e-9 {
            if n.abs() < (i64::MAX as f64) {
                format!("{}", n as i64)
            } else {
                format!("{:.0}", n)
            }
        } else {
            // Remove trailing zeros
            let s = format!("{}", n);
            s.trim_end_matches('0').trim_end_matches('.').to_string()
        }
    }
}

/// CSV export service for converting spreadsheets to CSV format.
///
/// Provides functionality to export spreadsheet data to CSV files with
/// configurable options for data inclusion and formatting.
pub struct CsvExporter;

impl CsvExporter {
    /// Exports a spreadsheet to a CSV file.
    ///
    /// Writes all non-empty cells from the spreadsheet to a CSV file.
    /// Only exports the rectangular region containing data (from A1 to the
    /// bottom-right cell with content).
    ///
    /// # Arguments
    ///
    /// * `spreadsheet` - Reference to the spreadsheet to export
    /// * `filename` - Path where the CSV file should be saved
    ///
    /// # Returns
    ///
    /// Result containing the filename on success, or error message on failure
    ///
    /// # Examples
    ///
    /// ```
    /// use tshts::domain::{Spreadsheet, CsvExporter};
    ///
    /// let sheet = Spreadsheet::default();
    /// match CsvExporter::export_to_csv(&sheet, "data.csv") {
    ///     Ok(filename) => println!("Exported to {}", filename),
    ///     Err(error) => println!("Export failed: {}", error),
    /// }
    /// ```
    pub fn export_to_csv(spreadsheet: &Spreadsheet, filename: &str) -> Result<String, String> {
        // Find the bounds of actual data
        let (max_row, max_col) = Self::find_data_bounds(spreadsheet);
        
        if max_row == 0 && max_col == 0 && spreadsheet.get_cell(0, 0).value.is_empty() {
            return Err("No data to export".to_string());
        }
        
        let file = File::create(filename).map_err(|e| format!("Failed to create file: {}", e))?;
        let mut writer = csv::Writer::from_writer(file);
        
        // Export data row by row
        for row in 0..=max_row {
            let mut record = Vec::new();
            for col in 0..=max_col {
                let cell = spreadsheet.get_cell(row, col);
                record.push(cell.value.clone());
            }
            writer.write_record(&record).map_err(|e| format!("Failed to write row: {}", e))?;
        }
        
        writer.flush().map_err(|e| format!("Failed to flush CSV writer: {}", e))?;
        Ok(filename.to_string())
    }
    
    /// Imports data from a CSV file into a spreadsheet.
    ///
    /// Reads CSV data and populates a new spreadsheet with the values.
    /// Each CSV row becomes a spreadsheet row, and each CSV column becomes a spreadsheet column.
    /// Empty cells in the CSV are preserved as empty cells in the spreadsheet.
    ///
    /// # Arguments
    ///
    /// * `filename` - Path to the CSV file to import
    ///
    /// # Returns
    ///
    /// Result containing the populated spreadsheet on success, or error message on failure
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use tshts::domain::CsvExporter;
    ///
    /// match CsvExporter::import_from_csv("data.csv") {
    ///     Ok(spreadsheet) => println!("Imported {} rows", spreadsheet.rows),
    ///     Err(error) => println!("Import failed: {}", error),
    /// }
    /// ```
    pub fn import_from_csv(filename: &str) -> Result<Spreadsheet, String> {
        let file = File::open(filename).map_err(|e| format!("Failed to open file: {}", e))?;
        let mut reader = csv::ReaderBuilder::new()
            .has_headers(false) // Don't treat first row as headers
            .from_reader(file);
        
        let mut spreadsheet = Spreadsheet::default();
        let mut max_row = 0;
        let mut max_col = 0;
        
        for (row_index, result) in reader.records().enumerate() {
            let record = result.map_err(|e| format!("Failed to read CSV row {}: {}", row_index + 1, e))?;
            
            for (col_index, field) in record.iter().enumerate() {
                if !field.is_empty() {
                    let cell_data = super::models::CellData {
                        value: field.to_string(),
                        formula: None,
                    };
                    spreadsheet.set_cell(row_index, col_index, cell_data);
                }
                max_col = max_col.max(col_index);
            }
            max_row = max_row.max(row_index);
        }
        
        // Update spreadsheet dimensions based on imported data
        if max_row > 0 || max_col > 0 {
            spreadsheet.rows = spreadsheet.rows.max(max_row + 10); // Add some buffer
            spreadsheet.cols = spreadsheet.cols.max(max_col + 5);   // Add some buffer
        }
        
        // Rebuild dependencies in case any imported cells contain formulas
        spreadsheet.rebuild_dependencies();
        
        Ok(spreadsheet)
    }

    /// Finds the bounds of the data in the spreadsheet.
    ///
    /// Returns the maximum row and column indices that contain non-empty data.
    /// This is used to determine the rectangular region to export.
    ///
    /// # Arguments
    ///
    /// * `spreadsheet` - Reference to the spreadsheet to analyze
    ///
    /// # Returns
    ///
    /// Tuple containing (max_row, max_col) with zero-based indices
    fn find_data_bounds(spreadsheet: &Spreadsheet) -> (usize, usize) {
        let mut max_row = 0;
        let mut max_col = 0;
        
        for ((row, col), cell) in &spreadsheet.cells {
            if !cell.value.is_empty() {
                max_row = max_row.max(*row);
                max_col = max_col.max(*col);
            }
        }
        
        (max_row, max_col)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::domain::CellData;

    fn create_test_spreadsheet() -> Spreadsheet {
        let mut sheet = Spreadsheet::default();
        
        // Set up some test data
        sheet.set_cell(0, 0, CellData { value: "10".to_string(), formula: None });
        sheet.set_cell(0, 1, CellData { value: "20".to_string(), formula: None });
        sheet.set_cell(0, 2, CellData { value: "30".to_string(), formula: None });
        sheet.set_cell(1, 0, CellData { value: "5".to_string(), formula: None });
        sheet.set_cell(1, 1, CellData { value: "15".to_string(), formula: None });
        sheet.set_cell(1, 2, CellData { value: "25".to_string(), formula: None });
        
        sheet
    }

    #[test]
    fn test_non_formula_passthrough() {
        let sheet = create_test_spreadsheet();
        let evaluator = FormulaEvaluator::new(&sheet);
        
        assert_eq!(evaluator.evaluate_formula("hello"), "hello");
        assert_eq!(evaluator.evaluate_formula("123"), "123");
        assert_eq!(evaluator.evaluate_formula(""), "");
    }

    #[test]
    fn test_simple_arithmetic() {
        let sheet = create_test_spreadsheet();
        let evaluator = FormulaEvaluator::new(&sheet);
        
        assert_eq!(evaluator.evaluate_formula("=2+3"), "5");
        assert_eq!(evaluator.evaluate_formula("=10-3"), "7");
        assert_eq!(evaluator.evaluate_formula("=4*5"), "20");
        assert_eq!(evaluator.evaluate_formula("=15/3"), "5");
        assert_eq!(evaluator.evaluate_formula("=2**3"), "8");
        assert_eq!(evaluator.evaluate_formula("=3^2"), "9");
        assert_eq!(evaluator.evaluate_formula("=10%3"), "1");
    }

    #[test]
    fn test_comparison_operators() {
        let sheet = create_test_spreadsheet();
        let evaluator = FormulaEvaluator::new(&sheet);
        
        assert_eq!(evaluator.evaluate_formula("=5<10"), "1");
        assert_eq!(evaluator.evaluate_formula("=10<5"), "0");
        assert_eq!(evaluator.evaluate_formula("=10>5"), "1");
        assert_eq!(evaluator.evaluate_formula("=5>10"), "0");
        assert_eq!(evaluator.evaluate_formula("=5<=5"), "1");
        assert_eq!(evaluator.evaluate_formula("=5<=4"), "0");
        assert_eq!(evaluator.evaluate_formula("=5>=5"), "1");
        assert_eq!(evaluator.evaluate_formula("=4>=5"), "0");
        assert_eq!(evaluator.evaluate_formula("=5<>5"), "0");
        assert_eq!(evaluator.evaluate_formula("=5<>4"), "1");
    }

    #[test]
    fn test_cell_references() {
        let sheet = create_test_spreadsheet();
        let evaluator = FormulaEvaluator::new(&sheet);
        
        assert_eq!(evaluator.evaluate_formula("=A1"), "10");
        assert_eq!(evaluator.evaluate_formula("=B1"), "20");
        assert_eq!(evaluator.evaluate_formula("=C1"), "30");
        assert_eq!(evaluator.evaluate_formula("=A2"), "5");
    }

    #[test]
    fn test_cell_arithmetic() {
        let sheet = create_test_spreadsheet();
        let evaluator = FormulaEvaluator::new(&sheet);
        
        assert_eq!(evaluator.evaluate_formula("=A1+B1"), "30"); // 10 + 20
        assert_eq!(evaluator.evaluate_formula("=C1-A1"), "20"); // 30 - 10
        assert_eq!(evaluator.evaluate_formula("=A1*A2"), "50"); // 10 * 5
        assert_eq!(evaluator.evaluate_formula("=B1/A2"), "4"); // 20 / 5
    }

    #[test]
    fn test_sum_function() {
        let sheet = create_test_spreadsheet();
        let evaluator = FormulaEvaluator::new(&sheet);
        
        assert_eq!(evaluator.evaluate_formula("=SUM(A1,B1,C1)"), "60"); // 10+20+30
        assert_eq!(evaluator.evaluate_formula("=SUM(A1:C1)"), "60"); // Range sum
        assert_eq!(evaluator.evaluate_formula("=SUM(A1:A2)"), "15"); // 10+5
        assert_eq!(evaluator.evaluate_formula("=SUM(5,10,15)"), "30"); // Literal values
    }

    #[test]
    fn test_average_function() {
        let sheet = create_test_spreadsheet();
        let evaluator = FormulaEvaluator::new(&sheet);
        
        assert_eq!(evaluator.evaluate_formula("=AVERAGE(A1,B1,C1)"), "20"); // (10+20+30)/3
        assert_eq!(evaluator.evaluate_formula("=AVERAGE(A1:C1)"), "20"); // Range average
        assert_eq!(evaluator.evaluate_formula("=AVERAGE(10,20)"), "15");
    }

    #[test]
    fn test_min_max_functions() {
        let sheet = create_test_spreadsheet();
        let evaluator = FormulaEvaluator::new(&sheet);
        
        assert_eq!(evaluator.evaluate_formula("=MIN(A1,B1,C1)"), "10");
        assert_eq!(evaluator.evaluate_formula("=MAX(A1,B1,C1)"), "30");
        assert_eq!(evaluator.evaluate_formula("=MIN(A1:C1)"), "10");
        assert_eq!(evaluator.evaluate_formula("=MAX(A1:C1)"), "30");
    }

    #[test]
    fn test_if_function() {
        let sheet = create_test_spreadsheet();
        let evaluator = FormulaEvaluator::new(&sheet);
        
        assert_eq!(evaluator.evaluate_formula("=IF(1,100,200)"), "100"); // True condition
        assert_eq!(evaluator.evaluate_formula("=IF(0,100,200)"), "200"); // False condition
        // Note: Complex comparisons in IF functions need a more sophisticated parser
        // For now, test with simple values
        assert_eq!(evaluator.evaluate_formula("=IF(1,1,0)"), "1"); // Simple true
        assert_eq!(evaluator.evaluate_formula("=IF(0,1,0)"), "0"); // Simple false
    }

    #[test]
    fn test_logical_functions() {
        let sheet = create_test_spreadsheet();
        let evaluator = FormulaEvaluator::new(&sheet);
        
        // Test basic functions first
        assert_eq!(evaluator.evaluate_formula("=SUM(1,2)"), "3"); // Simple sum
        
        // Test logical functions (these are built-in functions in the registry)
        assert_eq!(evaluator.evaluate_formula("=AND(1,1)"), "1"); // Both true  
        assert_eq!(evaluator.evaluate_formula("=AND(1,0)"), "0"); // One false
        assert_eq!(evaluator.evaluate_formula("=OR(0,1)"), "1"); // One true
        assert_eq!(evaluator.evaluate_formula("=OR(0,0)"), "0"); // Both false
        assert_eq!(evaluator.evaluate_formula("=NOT(0)"), "1"); // Not false
        assert_eq!(evaluator.evaluate_formula("=NOT(1)"), "0"); // Not true
        
        // All logical operations are now functions
        // (No need for separate binary operator tests since they're all functions now)
    }

    #[test]
    fn test_range_parsing() {
        let sheet = create_test_spreadsheet();
        let evaluator = FormulaEvaluator::new(&sheet);
        
        // Test different range formats
        assert_eq!(evaluator.evaluate_formula("=SUM(A1:C1)"), "60"); // Row range
        assert_eq!(evaluator.evaluate_formula("=SUM(A1:A2)"), "15"); // Column range
        assert_eq!(evaluator.evaluate_formula("=SUM(A1:B2)"), "50"); // Rectangle range
    }

    #[test]
    fn test_error_cases() {
        let sheet = create_test_spreadsheet();
        let evaluator = FormulaEvaluator::new(&sheet);
        
        assert_eq!(evaluator.evaluate_formula("=1/0"), "#ERROR"); // Division by zero
        assert_eq!(evaluator.evaluate_formula("=10%0"), "#ERROR"); // Modulo by zero
        assert_eq!(evaluator.evaluate_formula("=INVALID()"), "#ERROR"); // Unknown function
        assert_eq!(evaluator.evaluate_formula("=AVERAGE()"), "#ERROR"); // No args for average
    }

    #[test]
    fn test_circular_reference_detection() {
        let mut sheet = Spreadsheet::default();
        // Set up a cell that would reference itself
        sheet.set_cell(0, 0, CellData {
            value: "10".to_string(),
            formula: Some("=B1+1".to_string()),
        });
        
        // Set up indirect circular reference chain
        sheet.set_cell(0, 1, CellData {
            value: "20".to_string(),
            formula: Some("=C1+1".to_string()),
        });
        
        let evaluator = FormulaEvaluator::new(&sheet);
        
        // Direct self-reference
        assert!(evaluator.would_create_circular_reference("=A1+1", (0, 0)));
        
        // Non-circular reference
        assert!(!evaluator.would_create_circular_reference("=B1+1", (0, 0)));
        
        // This would create A1->C1->A1 if we set C1 to reference A1
        assert!(evaluator.would_create_circular_reference("=A1+1", (0, 2)));
    }

    #[test]
    fn test_extract_cell_references_from_ast() {
        let sheet = create_test_spreadsheet();
        let evaluator = FormulaEvaluator::new(&sheet);
        
        // Test cell reference extraction from AST
        let mut parser = Parser::new("A1 + B2 * C3").unwrap();
        let ast = parser.parse().unwrap();
        let refs = evaluator.extract_cell_references_from_ast(&ast);
        assert_eq!(refs.len(), 3);
        assert!(refs.contains(&(0, 0))); // A1
        assert!(refs.contains(&(1, 1))); // B2
        assert!(refs.contains(&(2, 2))); // C3
        
        // Test range extraction
        let mut parser = Parser::new("SUM(A1:A3)").unwrap();
        let ast = parser.parse().unwrap();
        let refs = evaluator.extract_cell_references_from_ast(&ast);
        assert_eq!(refs.len(), 3); // Should find A1, A2, A3 from the range
        assert!(refs.contains(&(0, 0))); // A1
        assert!(refs.contains(&(1, 0))); // A2
        assert!(refs.contains(&(2, 0))); // A3
    }

    #[test]
    fn test_case_insensitive_functions() {
        let sheet = create_test_spreadsheet();
        let evaluator = FormulaEvaluator::new(&sheet);
        
        assert_eq!(evaluator.evaluate_formula("=sum(A1,B1)"), "30");
        assert_eq!(evaluator.evaluate_formula("=SUM(A1,B1)"), "30");
        assert_eq!(evaluator.evaluate_formula("=Sum(A1,B1)"), "30");
        assert_eq!(evaluator.evaluate_formula("=average(A1,B1)"), "15");
        assert_eq!(evaluator.evaluate_formula("=AVERAGE(A1,B1)"), "15");
    }

    #[test]
    fn test_whitespace_handling() {
        let sheet = create_test_spreadsheet();
        let evaluator = FormulaEvaluator::new(&sheet);
        
        assert_eq!(evaluator.evaluate_formula("= 2 + 3 "), "5");
        assert_eq!(evaluator.evaluate_formula("=SUM( A1 , B1 )"), "30");
        assert_eq!(evaluator.evaluate_formula("= A1 * 2 "), "20");
    }

    #[test]
    fn test_complex_expressions() {
        let sheet = create_test_spreadsheet();
        let evaluator = FormulaEvaluator::new(&sheet);
        
        // Test power operations (simple arithmetic)
        assert_eq!(evaluator.evaluate_formula("=2**3+1"), "9"); // 8+1
        assert_eq!(evaluator.evaluate_formula("=3*4+2"), "14"); // 12+2
        
        // Test functions work correctly
        assert_eq!(evaluator.evaluate_formula("=SUM(A1,B1)"), "30"); // 10+20
        assert_eq!(evaluator.evaluate_formula("=SUM(A1:B1)"), "30"); // 10+20
    }

    #[test]
    fn test_string_literals() {
        let sheet = create_test_spreadsheet();
        let evaluator = FormulaEvaluator::new(&sheet);
        
        assert_eq!(evaluator.evaluate_formula("=\"Hello World\""), "Hello World");
        assert_eq!(evaluator.evaluate_formula("=\"\""), "");
        assert_eq!(evaluator.evaluate_formula("=\"Test\""), "Test");
    }

    #[test]
    fn test_string_concatenation() {
        let sheet = create_test_spreadsheet();
        let evaluator = FormulaEvaluator::new(&sheet);
        
        assert_eq!(evaluator.evaluate_formula("=\"Hello\" & \" \" & \"World\""), "Hello World");
        assert_eq!(evaluator.evaluate_formula("=\"Number: \" & 42"), "Number: 42");
        assert_eq!(evaluator.evaluate_formula("=\"Result: \" & (2 + 3)"), "Result: 5");
    }

    #[test]
    fn test_string_functions() {
        let sheet = create_test_spreadsheet();
        let evaluator = FormulaEvaluator::new(&sheet);
        
        // Test LEN function
        assert_eq!(evaluator.evaluate_formula("=LEN(\"Hello\")"), "5");
        assert_eq!(evaluator.evaluate_formula("=LEN(\"\")"), "0");
        
        // Test UPPER/LOWER functions
        assert_eq!(evaluator.evaluate_formula("=UPPER(\"hello\")"), "HELLO");
        assert_eq!(evaluator.evaluate_formula("=LOWER(\"WORLD\")"), "world");
        
        // Test TRIM function
        assert_eq!(evaluator.evaluate_formula("=TRIM(\"  spaces  \")"), "spaces");
        
        // Test LEFT/RIGHT functions
        assert_eq!(evaluator.evaluate_formula("=LEFT(\"Hello World\", 5)"), "Hello");
        assert_eq!(evaluator.evaluate_formula("=RIGHT(\"Hello World\", 5)"), "World");
        
        // Test MID function (0-based indexing)
        assert_eq!(evaluator.evaluate_formula("=MID(\"Hello World\", 6, 5)"), "World");
        
        // Test FIND function (0-based indexing)
        assert_eq!(evaluator.evaluate_formula("=FIND(\"lo\", \"Hello\")"), "3");
        assert_eq!(evaluator.evaluate_formula("=FIND(\"World\", \"Hello World\")"), "6");
        
        // Test CONCAT function
        assert_eq!(evaluator.evaluate_formula("=CONCAT(\"A\", \"B\", \"C\")"), "ABC");
        assert_eq!(evaluator.evaluate_formula("=CONCAT(\"Number: \", 123)"), "Number: 123");
    }

    #[test]
    fn test_string_cell_references() {
        let mut sheet = Spreadsheet::default();
        sheet.set_cell(0, 0, CellData { value: "Hello".to_string(), formula: None });
        sheet.set_cell(0, 1, CellData { value: "World".to_string(), formula: None });
        sheet.set_cell(0, 2, CellData { value: "123".to_string(), formula: None });
        
        let evaluator = FormulaEvaluator::new(&sheet);
        
        // Test string cell concatenation
        assert_eq!(evaluator.evaluate_formula("=A1 & \" \" & B1"), "Hello World");
        
        // Test string functions with cell references
        assert_eq!(evaluator.evaluate_formula("=LEN(A1)"), "5");
        assert_eq!(evaluator.evaluate_formula("=UPPER(A1)"), "HELLO");
        
        // Test numeric conversion from string cells
        assert_eq!(evaluator.evaluate_formula("=C1 + 456"), "579");
    }

    #[test]
    fn test_string_equality() {
        let sheet = create_test_spreadsheet();
        let evaluator = FormulaEvaluator::new(&sheet);
        
        // String equality
        assert_eq!(evaluator.evaluate_formula("=\"Hello\" = \"Hello\""), "1");
        assert_eq!(evaluator.evaluate_formula("=\"Hello\" = \"World\""), "0");
        
        // String inequality
        assert_eq!(evaluator.evaluate_formula("=\"Hello\" <> \"World\""), "1");
        assert_eq!(evaluator.evaluate_formula("=\"Hello\" <> \"Hello\""), "0");
    }

    #[test]
    fn test_if_with_strings() {
        let mut sheet = Spreadsheet::default();
        sheet.set_cell(0, 0, CellData { value: "Hello".to_string(), formula: None });
        
        let evaluator = FormulaEvaluator::new(&sheet);
        
        // Test IF with string conditions and results
        assert_eq!(evaluator.evaluate_formula("=IF(A1=\"Hello\", \"Found\", \"Not Found\")"), "Found");
        assert_eq!(evaluator.evaluate_formula("=IF(LEN(A1)>3, \"Long\", \"Short\")"), "Long");
        assert_eq!(evaluator.evaluate_formula("=IF(A1=\"World\", \"Found\", \"Not Found\")"), "Not Found");
    }

    #[test]
    fn test_string_function_errors() {
        let sheet = create_test_spreadsheet();
        let evaluator = FormulaEvaluator::new(&sheet);
        
        // Test FIND with no match
        assert_eq!(evaluator.evaluate_formula("=FIND(\"xyz\", \"Hello\")"), "#ERROR");
        
        // Test functions with wrong argument count
        assert_eq!(evaluator.evaluate_formula("=LEN()"), "#ERROR");
        assert_eq!(evaluator.evaluate_formula("=LEN(\"a\", \"b\")"), "#ERROR");
    }

    #[test]
    fn test_get_function_basic() {
        let sheet = create_test_spreadsheet();
        let evaluator = FormulaEvaluator::new(&sheet);
        
        // Test GET function with wrong argument count
        assert_eq!(evaluator.evaluate_formula("=GET()"), "#ERROR");
        assert_eq!(evaluator.evaluate_formula("=GET(\"url1\", \"url2\")"), "#ERROR");
        
        // Note: We can't easily test actual HTTP requests in unit tests
        // since they depend on external services. In a real application,
        // you might want to use dependency injection or mock HTTP clients
        // for testing. For now, we just test the error cases.
    }

    #[test]
    fn test_get_function_invalid_url() {
        let sheet = create_test_spreadsheet();
        let evaluator = FormulaEvaluator::new(&sheet);
        
        // Test with invalid URL - this should return an error
        let result = evaluator.evaluate_formula("=GET(\"not-a-valid-url\")");
        assert_eq!(result, "#ERROR");
        
        // Test with empty string URL
        let result = evaluator.evaluate_formula("=GET(\"\")");
        assert_eq!(result, "#ERROR");
    }

    #[test]
    fn test_get_function_real_http_requests() {
        let sheet = create_test_spreadsheet();
        let evaluator = FormulaEvaluator::new(&sheet);
        
        // Test fetching crypto price from real API
        let result = evaluator.evaluate_formula("=GET(\"https://cryptoprices.cc/ADA\")");
        // The result should be a valid response (not #ERROR) and contain price data
        assert_ne!(result, "#ERROR");
        assert!(!result.is_empty());
        
        // Test another crypto ticker
        let result = evaluator.evaluate_formula("=GET(\"https://cryptoprices.cc/BTC\")");
        assert_ne!(result, "#ERROR");
        assert!(!result.is_empty());
    }

    #[test]
    fn test_nested_get_with_string_functions() {
        let sheet = create_test_spreadsheet();
        let evaluator = FormulaEvaluator::new(&sheet);
        
        // Test LEN with GET - get length of response
        let result = evaluator.evaluate_formula("=LEN(GET(\"https://cryptoprices.cc/ADA\"))");
        assert_ne!(result, "#ERROR");
        // Should be a positive number (length of response)
        if let Ok(len) = result.parse::<f64>() {
            assert!(len > 0.0);
        } else {
            panic!("Expected numeric result for LEN(GET(...)), got: {}", result);
        }
        
        // Test UPPER with GET - convert response to uppercase
        let result = evaluator.evaluate_formula("=UPPER(GET(\"https://cryptoprices.cc/ADA\"))");
        assert_ne!(result, "#ERROR");
        assert!(!result.is_empty());
        
        // Test TRIM with GET - trim whitespace from response
        let result = evaluator.evaluate_formula("=TRIM(GET(\"https://cryptoprices.cc/ADA\"))");
        assert_ne!(result, "#ERROR");
        assert!(!result.is_empty());
    }

    #[test]
    fn test_nested_concat_with_get() {
        let sheet = create_test_spreadsheet();
        let evaluator = FormulaEvaluator::new(&sheet);
        
        // Test CONCAT with GET results
        let result = evaluator.evaluate_formula("=CONCAT(\"ADA Price: \", GET(\"https://cryptoprices.cc/ADA\"))");
        assert_ne!(result, "#ERROR");
        assert!(result.starts_with("ADA Price: "));
        
        // Test concatenating multiple GET requests
        let result = evaluator.evaluate_formula("=CONCAT(\"ADA: \", GET(\"https://cryptoprices.cc/ADA\"), \" | BTC: \", GET(\"https://cryptoprices.cc/BTC\"))");
        assert_ne!(result, "#ERROR");
        assert!(result.contains("ADA: "));
        assert!(result.contains(" | BTC: "));
    }

    #[test]
    fn test_complex_nested_expressions_with_get() {
        let sheet = create_test_spreadsheet();
        let evaluator = FormulaEvaluator::new(&sheet);
        
        // Test deeply nested: UPPER(CONCAT("Price: ", TRIM(GET(url))))
        let result = evaluator.evaluate_formula("=UPPER(CONCAT(\"Ada price: \", TRIM(GET(\"https://cryptoprices.cc/ADA\"))))");
        assert_ne!(result, "#ERROR");
        assert!(result.starts_with("ADA PRICE: "));
        
        // Test conditional with GET: IF(LEN(GET(url)) > 0, "Got data", "No data")
        let result = evaluator.evaluate_formula("=IF(LEN(GET(\"https://cryptoprices.cc/ADA\"))>0, \"Got data\", \"No data\")");
        assert_ne!(result, "#ERROR");
        assert_eq!(result, "Got data");
        
        // Test FIND within GET results
        let result = evaluator.evaluate_formula("=FIND(\".\", GET(\"https://cryptoprices.cc/ADA\"))");
        // Should find a decimal point in the price response (most crypto prices have decimals)
        // If no decimal found, it will return #ERROR, but most crypto prices should have decimals
        if result != "#ERROR" {
            if let Ok(pos) = result.parse::<f64>() {
                assert!(pos >= 0.0);
            }
        }
    }

    #[test]
    fn test_get_with_cell_references() {
        let mut sheet = Spreadsheet::default();
        // Set up a cell with a URL
        sheet.set_cell(0, 0, CellData { 
            value: "https://cryptoprices.cc/ADA".to_string(), 
            formula: None 
        });
        
        let evaluator = FormulaEvaluator::new(&sheet);
        
        // Test GET with cell reference
        let result = evaluator.evaluate_formula("=GET(A1)");
        assert_ne!(result, "#ERROR");
        assert!(!result.is_empty());
        
        // Test nested function with cell reference: LEN(GET(A1))
        let result = evaluator.evaluate_formula("=LEN(GET(A1))");
        assert_ne!(result, "#ERROR");
        if let Ok(len) = result.parse::<f64>() {
            assert!(len > 0.0);
        }
        
        // Test CONCAT with cell reference and GET
        let result = evaluator.evaluate_formula("=CONCAT(\"Price from \", A1, \": \", GET(A1))");
        assert_ne!(result, "#ERROR");
        assert!(result.contains("Price from https://cryptoprices.cc/ADA: "));
    }

    #[test]
    fn test_multiple_nested_function_levels() {
        let sheet = create_test_spreadsheet();
        let evaluator = FormulaEvaluator::new(&sheet);
        
        // Test 4-level nesting: LEFT(UPPER(TRIM(GET(url))), 10)
        let result = evaluator.evaluate_formula("=LEFT(UPPER(TRIM(GET(\"https://cryptoprices.cc/ADA\"))), 10)");
        assert_ne!(result, "#ERROR");
        assert!(!result.is_empty());
        assert!(result.len() <= 10);
        
        // Test 5-level nesting with conditional: IF(LEN(TRIM(GET(url))) > 5, LEFT(UPPER(GET(url)), 20), "Short")
        let result = evaluator.evaluate_formula("=IF(LEN(TRIM(GET(\"https://cryptoprices.cc/ADA\")))>5, LEFT(UPPER(GET(\"https://cryptoprices.cc/ADA\")), 20), \"Short\")");
        assert_ne!(result, "#ERROR");
        // Should not be "Short" since crypto price responses are typically longer than 5 characters
        assert_ne!(result, "Short");
    }

    #[test]
    fn test_error_propagation_in_nested_functions() {
        let sheet = create_test_spreadsheet();
        let evaluator = FormulaEvaluator::new(&sheet);
        
        // Test that errors in inner functions propagate outward
        let result = evaluator.evaluate_formula("=LEN(GET(\"invalid-url\"))");
        assert_eq!(result, "#ERROR");
        
        let result = evaluator.evaluate_formula("=CONCAT(\"Price: \", GET(\"invalid-url\"))");
        assert_eq!(result, "#ERROR");
        
        let result = evaluator.evaluate_formula("=IF(LEN(GET(\"invalid-url\"))>0, \"Good\", \"Bad\")");
        assert_eq!(result, "#ERROR");
        
        // Test FIND with invalid search in valid GET
        let result = evaluator.evaluate_formula("=FIND(\"xyz123notfound\", GET(\"https://cryptoprices.cc/ADA\"))");
        // This should return #ERROR because the search string likely won't be found
        assert_eq!(result, "#ERROR");
    }

    #[test]
    fn test_large_numbers() {
        let mut sheet = Spreadsheet::default();
        sheet.set_cell(0, 0, CellData { value: "1000000".to_string(), formula: None });
        sheet.set_cell(0, 1, CellData { value: "2000000".to_string(), formula: None });
        
        let evaluator = FormulaEvaluator::new(&sheet);
        assert_eq!(evaluator.evaluate_formula("=A1+B1"), "3000000");
        assert_eq!(evaluator.evaluate_formula("=A1*B1"), "2000000000000");
    }

    #[test]
    fn test_negative_numbers() {
        let mut sheet = Spreadsheet::default();
        sheet.set_cell(0, 0, CellData { value: "-10".to_string(), formula: None });
        sheet.set_cell(0, 1, CellData { value: "5".to_string(), formula: None });
        
        let evaluator = FormulaEvaluator::new(&sheet);
        assert_eq!(evaluator.evaluate_formula("=A1+B1"), "-5");
        assert_eq!(evaluator.evaluate_formula("=A1*B1"), "-50");
        assert_eq!(evaluator.evaluate_formula("=-5+10"), "5");
    }

    #[test]
    fn test_decimal_precision() {
        let sheet = create_test_spreadsheet();
        let evaluator = FormulaEvaluator::new(&sheet);
        
        assert_eq!(evaluator.evaluate_formula("=1/3"), "0.3333333333333333");
        assert_eq!(evaluator.evaluate_formula("=22/7"), "3.142857142857143");
    }

    #[test]
    fn test_csv_export_basic() {
        use tempfile::NamedTempFile;
        
        let mut sheet = Spreadsheet::default();
        sheet.set_cell(0, 0, CellData { value: "Name".to_string(), formula: None });
        sheet.set_cell(0, 1, CellData { value: "Age".to_string(), formula: None });
        sheet.set_cell(1, 0, CellData { value: "Alice".to_string(), formula: None });
        sheet.set_cell(1, 1, CellData { value: "30".to_string(), formula: None });
        sheet.set_cell(2, 0, CellData { value: "Bob".to_string(), formula: None });
        sheet.set_cell(2, 1, CellData { value: "25".to_string(), formula: None });
        
        let temp_file = NamedTempFile::new().expect("Failed to create temp file");
        let file_path = temp_file.path().to_str().unwrap();
        
        let result = CsvExporter::export_to_csv(&sheet, file_path);
        assert!(result.is_ok());
        assert_eq!(result.unwrap(), file_path);
        
        // Read back the CSV and verify content
        let content = std::fs::read_to_string(file_path).expect("Failed to read CSV file");
        let lines: Vec<&str> = content.lines().collect();
        assert_eq!(lines.len(), 3);
        assert_eq!(lines[0], "Name,Age");
        assert_eq!(lines[1], "Alice,30");
        assert_eq!(lines[2], "Bob,25");
    }

    #[test]
    fn test_csv_export_empty_sheet() {
        use tempfile::NamedTempFile;
        
        let sheet = Spreadsheet::default();
        let temp_file = NamedTempFile::new().expect("Failed to create temp file");
        let file_path = temp_file.path().to_str().unwrap();
        
        let result = CsvExporter::export_to_csv(&sheet, file_path);
        assert!(result.is_err());
        assert_eq!(result.unwrap_err(), "No data to export");
    }

    #[test]
    fn test_csv_export_sparse_data() {
        use tempfile::NamedTempFile;
        
        let mut sheet = Spreadsheet::default();
        sheet.set_cell(0, 0, CellData { value: "A1".to_string(), formula: None });
        sheet.set_cell(2, 3, CellData { value: "D3".to_string(), formula: None });
        sheet.set_cell(1, 1, CellData { value: "B2".to_string(), formula: None });
        
        let temp_file = NamedTempFile::new().expect("Failed to create temp file");
        let file_path = temp_file.path().to_str().unwrap();
        
        let result = CsvExporter::export_to_csv(&sheet, file_path);
        assert!(result.is_ok());
        
        // Read back the CSV and verify content
        let content = std::fs::read_to_string(file_path).expect("Failed to read CSV file");
        let lines: Vec<&str> = content.lines().collect();
        assert_eq!(lines.len(), 3); // 0 to 2 (max_row)
        assert_eq!(lines[0], "A1,,,"); // A1 with empty cells up to column 3
        assert_eq!(lines[1], ",B2,,"); // Empty, B2, empty, empty
        assert_eq!(lines[2], ",,,D3"); // Empty cells then D3
    }

    #[test]
    fn test_csv_export_with_formulas() {
        use tempfile::NamedTempFile;
        
        let mut sheet = Spreadsheet::default();
        sheet.set_cell(0, 0, CellData { value: "10".to_string(), formula: None });
        sheet.set_cell(0, 1, CellData { value: "20".to_string(), formula: None });
        sheet.set_cell(0, 2, CellData { 
            value: "30".to_string(), 
            formula: Some("=A1+B1".to_string()) 
        });
        
        let temp_file = NamedTempFile::new().expect("Failed to create temp file");
        let file_path = temp_file.path().to_str().unwrap();
        
        let result = CsvExporter::export_to_csv(&sheet, file_path);
        assert!(result.is_ok());
        
        // Read back the CSV - should contain values, not formulas
        let content = std::fs::read_to_string(file_path).expect("Failed to read CSV file");
        let lines: Vec<&str> = content.lines().collect();
        assert_eq!(lines.len(), 1);
        assert_eq!(lines[0], "10,20,30"); // Values, not formulas
    }

    #[test]
    fn test_csv_export_special_characters() {
        use tempfile::NamedTempFile;
        
        let mut sheet = Spreadsheet::default();
        sheet.set_cell(0, 0, CellData { value: "Hello, World!".to_string(), formula: None });
        sheet.set_cell(0, 1, CellData { value: "\"Quoted\"".to_string(), formula: None });
        sheet.set_cell(0, 2, CellData { value: "Line\nBreak".to_string(), formula: None });
        
        let temp_file = NamedTempFile::new().expect("Failed to create temp file");
        let file_path = temp_file.path().to_str().unwrap();
        
        let result = CsvExporter::export_to_csv(&sheet, file_path);
        assert!(result.is_ok());
        
        // The CSV library should handle proper escaping
        let content = std::fs::read_to_string(file_path).expect("Failed to read CSV file");
        assert!(content.contains("Hello, World!"));
        assert!(content.contains("\"Quoted\""));
        assert!(content.contains("Line\nBreak"));
    }

    #[test]
    fn test_find_data_bounds() {
        let mut sheet = Spreadsheet::default();
        
        // Test empty sheet
        let (max_row, max_col) = CsvExporter::find_data_bounds(&sheet);
        assert_eq!((max_row, max_col), (0, 0));
        
        // Add some data
        sheet.set_cell(5, 3, CellData { value: "data".to_string(), formula: None });
        sheet.set_cell(2, 7, CellData { value: "more".to_string(), formula: None });
        sheet.set_cell(0, 0, CellData { value: "start".to_string(), formula: None });
        
        let (max_row, max_col) = CsvExporter::find_data_bounds(&sheet);
        assert_eq!((max_row, max_col), (5, 7));
    }

    #[test]
    fn test_csv_import_basic() {
        use tempfile::NamedTempFile;
        use std::io::Write;
        
        let mut temp_file = NamedTempFile::new().expect("Failed to create temp file");
        writeln!(temp_file, "Name,Age,City").expect("Failed to write to temp file");
        writeln!(temp_file, "Alice,30,New York").expect("Failed to write to temp file");
        writeln!(temp_file, "Bob,25,Los Angeles").expect("Failed to write to temp file");
        writeln!(temp_file, "Charlie,35,Chicago").expect("Failed to write to temp file");
        
        let file_path = temp_file.path().to_str().unwrap();
        let result = CsvExporter::import_from_csv(file_path);
        assert!(result.is_ok());
        
        let sheet = result.unwrap();
        
        // Check headers
        assert_eq!(sheet.get_cell(0, 0).value, "Name");
        assert_eq!(sheet.get_cell(0, 1).value, "Age");
        assert_eq!(sheet.get_cell(0, 2).value, "City");
        
        // Check data rows
        assert_eq!(sheet.get_cell(1, 0).value, "Alice");
        assert_eq!(sheet.get_cell(1, 1).value, "30");
        assert_eq!(sheet.get_cell(1, 2).value, "New York");
        
        assert_eq!(sheet.get_cell(2, 0).value, "Bob");
        assert_eq!(sheet.get_cell(2, 1).value, "25");
        assert_eq!(sheet.get_cell(2, 2).value, "Los Angeles");
        
        assert_eq!(sheet.get_cell(3, 0).value, "Charlie");
        assert_eq!(sheet.get_cell(3, 1).value, "35");
        assert_eq!(sheet.get_cell(3, 2).value, "Chicago");
        
        // Check that dimensions were updated appropriately
        assert!(sheet.rows >= 4);
        assert!(sheet.cols >= 3);
    }

    #[test]
    fn test_csv_import_empty_cells() {
        use tempfile::NamedTempFile;
        use std::io::Write;
        
        let mut temp_file = NamedTempFile::new().expect("Failed to create temp file");
        writeln!(temp_file, "A,,C").expect("Failed to write to temp file");
        writeln!(temp_file, ",B,").expect("Failed to write to temp file");
        writeln!(temp_file, ",,").expect("Failed to write to temp file");
        writeln!(temp_file, "D,E,F").expect("Failed to write to temp file");
        
        let file_path = temp_file.path().to_str().unwrap();
        let result = CsvExporter::import_from_csv(file_path);
        assert!(result.is_ok());
        
        let sheet = result.unwrap();
        
        // Check non-empty cells
        assert_eq!(sheet.get_cell(0, 0).value, "A");
        assert_eq!(sheet.get_cell(0, 2).value, "C");
        assert_eq!(sheet.get_cell(1, 1).value, "B");
        assert_eq!(sheet.get_cell(3, 0).value, "D");
        assert_eq!(sheet.get_cell(3, 1).value, "E");
        assert_eq!(sheet.get_cell(3, 2).value, "F");
        
        // Check empty cells remain empty
        assert!(sheet.get_cell(0, 1).value.is_empty());
        assert!(sheet.get_cell(1, 0).value.is_empty());
        assert!(sheet.get_cell(1, 2).value.is_empty());
        assert!(sheet.get_cell(2, 0).value.is_empty());
        assert!(sheet.get_cell(2, 1).value.is_empty());
        assert!(sheet.get_cell(2, 2).value.is_empty());
    }

    #[test]
    fn test_csv_import_special_characters() {
        use tempfile::NamedTempFile;
        use std::io::Write;
        
        let mut temp_file = NamedTempFile::new().expect("Failed to create temp file");
        writeln!(temp_file, r#""Hello, World!","Quote ""Test""","Line
Break""#).expect("Failed to write to temp file");
        writeln!(temp_file, "Héllo Wörld,🌍,Тест").expect("Failed to write to temp file");
        
        let file_path = temp_file.path().to_str().unwrap();
        let result = CsvExporter::import_from_csv(file_path);
        assert!(result.is_ok());
        
        let sheet = result.unwrap();
        
        // Check that special characters are preserved
        assert_eq!(sheet.get_cell(0, 0).value, "Hello, World!");
        assert_eq!(sheet.get_cell(0, 1).value, "Quote \"Test\"");
        assert_eq!(sheet.get_cell(0, 2).value, "Line\nBreak");
        assert_eq!(sheet.get_cell(1, 0).value, "Héllo Wörld");
        assert_eq!(sheet.get_cell(1, 1).value, "🌍");
        assert_eq!(sheet.get_cell(1, 2).value, "Тест");
    }

    #[test]
    fn test_csv_import_numbers() {
        use tempfile::NamedTempFile;
        use std::io::Write;
        
        let mut temp_file = NamedTempFile::new().expect("Failed to create temp file");
        writeln!(temp_file, "Integer,Decimal,Negative,Scientific").expect("Failed to write to temp file");
        writeln!(temp_file, "42,3.14159,-273.15,6.022e23").expect("Failed to write to temp file");
        writeln!(temp_file, "0,0.0,-0,1.0e-10").expect("Failed to write to temp file");
        
        let file_path = temp_file.path().to_str().unwrap();
        let result = CsvExporter::import_from_csv(file_path);
        assert!(result.is_ok());
        
        let sheet = result.unwrap();
        
        // Check that numbers are imported as strings (no automatic conversion)
        assert_eq!(sheet.get_cell(1, 0).value, "42");
        assert_eq!(sheet.get_cell(1, 1).value, "3.14159");
        assert_eq!(sheet.get_cell(1, 2).value, "-273.15");
        assert_eq!(sheet.get_cell(1, 3).value, "6.022e23");
        assert_eq!(sheet.get_cell(2, 0).value, "0");
        assert_eq!(sheet.get_cell(2, 1).value, "0.0");
        assert_eq!(sheet.get_cell(2, 2).value, "-0");
        assert_eq!(sheet.get_cell(2, 3).value, "1.0e-10");
        
        // Verify that none of these have formulas
        assert!(sheet.get_cell(1, 0).formula.is_none());
        assert!(sheet.get_cell(1, 1).formula.is_none());
        assert!(sheet.get_cell(1, 2).formula.is_none());
        assert!(sheet.get_cell(1, 3).formula.is_none());
    }

    #[test]
    fn test_csv_import_nonexistent_file() {
        let result = CsvExporter::import_from_csv("/nonexistent/file.csv");
        assert!(result.is_err());
        assert!(result.unwrap_err().contains("Failed to open file"));
    }

    #[test]
    fn test_csv_import_invalid_csv() {
        use tempfile::NamedTempFile;
        use std::io::Write;
        
        let mut temp_file = NamedTempFile::new().expect("Failed to create temp file");
        // Write invalid CSV (unmatched quote)
        writeln!(temp_file, r#"Valid,"Unmatched quote"#).expect("Failed to write to temp file");
        writeln!(temp_file, "Another line").expect("Failed to write to temp file");
        
        let file_path = temp_file.path().to_str().unwrap();
        let result = CsvExporter::import_from_csv(file_path);
        // This might succeed or fail depending on the CSV parser's tolerance
        // The main thing is that it doesn't panic
        match result {
            Ok(_) => {
                // CSV parser was tolerant of the malformed input
            }
            Err(err) => {
                // CSV parser rejected the malformed input
                assert!(err.contains("Failed to read CSV row"));
            }
        }
    }

    #[test]
    fn test_csv_import_empty_file() {
        use tempfile::NamedTempFile;
        
        let temp_file = NamedTempFile::new().expect("Failed to create temp file");
        let file_path = temp_file.path().to_str().unwrap();
        
        let result = CsvExporter::import_from_csv(file_path);
        assert!(result.is_ok());
        
        let sheet = result.unwrap();
        
        // Empty file should result in empty spreadsheet
        assert_eq!(sheet.rows, 100); // Default dimensions
        assert_eq!(sheet.cols, 26);
        assert!(sheet.cells.is_empty());
    }

    #[test]
    fn test_csv_roundtrip() {
        use tempfile::NamedTempFile;
        
        // Create original spreadsheet
        let mut original = Spreadsheet::default();
        original.set_cell(0, 0, CellData { value: "Name".to_string(), formula: None });
        original.set_cell(0, 1, CellData { value: "Score".to_string(), formula: None });
        original.set_cell(1, 0, CellData { value: "Alice".to_string(), formula: None });
        original.set_cell(1, 1, CellData { value: "95".to_string(), formula: None });
        original.set_cell(2, 0, CellData { value: "Bob".to_string(), formula: None });
        original.set_cell(2, 1, CellData { value: "87".to_string(), formula: None });
        
        // Export to CSV
        let temp_file1 = NamedTempFile::new().expect("Failed to create temp file");
        let export_path = temp_file1.path().to_str().unwrap();
        let export_result = CsvExporter::export_to_csv(&original, export_path);
        assert!(export_result.is_ok());
        
        // Import back from CSV
        let import_result = CsvExporter::import_from_csv(export_path);
        assert!(import_result.is_ok());
        
        let imported = import_result.unwrap();
        
        // Verify data integrity
        assert_eq!(imported.get_cell(0, 0).value, "Name");
        assert_eq!(imported.get_cell(0, 1).value, "Score");
        assert_eq!(imported.get_cell(1, 0).value, "Alice");
        assert_eq!(imported.get_cell(1, 1).value, "95");
        assert_eq!(imported.get_cell(2, 0).value, "Bob");
        assert_eq!(imported.get_cell(2, 1).value, "87");
        
        // All imported cells should have no formulas
        assert!(imported.get_cell(0, 0).formula.is_none());
        assert!(imported.get_cell(0, 1).formula.is_none());
        assert!(imported.get_cell(1, 0).formula.is_none());
        assert!(imported.get_cell(1, 1).formula.is_none());
        assert!(imported.get_cell(2, 0).formula.is_none());
        assert!(imported.get_cell(2, 1).formula.is_none());
    }

    #[test]
    fn test_adjust_formula_references_horizontal() {
        let sheet = Spreadsheet::default();
        let evaluator = FormulaEvaluator::new(&sheet);
        
        // Test moving formula one column to the right
        let adjusted = evaluator.adjust_formula_references("=A1+B1", 0, 1);
        assert_eq!(adjusted, "=B1+C1");
        
        // Test moving formula two columns to the right
        let adjusted = evaluator.adjust_formula_references("=SUM(A1:B3)", 0, 2);
        assert_eq!(adjusted, "=SUM(C1:D3)");
        
        // Test moving formula with multiple references
        let adjusted = evaluator.adjust_formula_references("=A1*B2+C3", 0, 1);
        assert_eq!(adjusted, "=B1*C2+D3");
    }

    #[test]
    fn test_adjust_formula_references_vertical() {
        let sheet = Spreadsheet::default();
        let evaluator = FormulaEvaluator::new(&sheet);
        
        // Test moving formula one row down
        let adjusted = evaluator.adjust_formula_references("=A1+A2", 1, 0);
        assert_eq!(adjusted, "=A2+A3");
        
        // Test moving formula two rows down with range
        let adjusted = evaluator.adjust_formula_references("=SUM(A1:C1)", 2, 0);
        assert_eq!(adjusted, "=SUM(A3:C3)");
        
        // Test moving formula with mixed references
        let adjusted = evaluator.adjust_formula_references("=A1*B2+SUM(C3:D4)", 1, 0);
        assert_eq!(adjusted, "=A2*B3+SUM(C4:D5)");
    }

    #[test]
    fn test_adjust_formula_references_diagonal() {
        let sheet = Spreadsheet::default();
        let evaluator = FormulaEvaluator::new(&sheet);
        
        // Test moving formula diagonally (down and right)
        let adjusted = evaluator.adjust_formula_references("=A1+B2", 1, 1);
        assert_eq!(adjusted, "=B2+C3");
        
        // Test range adjustment diagonally
        let adjusted = evaluator.adjust_formula_references("=SUM(A1:B2)", 2, 3);
        assert_eq!(adjusted, "=SUM(D3:E4)");
    }

    #[test]
    fn test_adjust_formula_references_complex() {
        let sheet = Spreadsheet::default();
        let evaluator = FormulaEvaluator::new(&sheet);
        
        // Test complex formula with multiple functions and references
        let adjusted = evaluator.adjust_formula_references("=IF(A1>B1,SUM(C1:C3),AVERAGE(D1:E2))", 1, 1);
        assert_eq!(adjusted, "=IF(B2>C2,SUM(D2:D4),AVERAGE(E2:F3))");
        
        // Test formula with string literals (should not be affected)
        let adjusted = evaluator.adjust_formula_references("=CONCAT(A1,\"test\",B1)", 0, 1);
        assert_eq!(adjusted, "=CONCAT(B1,\"test\",C1)");
    }

    #[test]
    fn test_adjust_formula_references_edge_cases() {
        let sheet = Spreadsheet::default();
        let evaluator = FormulaEvaluator::new(&sheet);

        // Test non-formula (should return unchanged)
        let adjusted = evaluator.adjust_formula_references("Hello World", 1, 1);
        assert_eq!(adjusted, "Hello World");

        // Test formula with no cell references
        let adjusted = evaluator.adjust_formula_references("=5+10", 1, 1);
        assert_eq!(adjusted, "=5+10");

        // Test negative offsets (moving up/left) - should not go below A1
        let adjusted = evaluator.adjust_formula_references("=A1+B1", -1, -1);
        assert_eq!(adjusted, "=A1+A1");
    }

    // ==================== AutofillPattern Tests ====================

    #[test]
    fn test_autofill_pattern_arithmetic_positive_step() {
        let values = vec!["1".to_string(), "2".to_string(), "3".to_string()];
        let pattern = AutofillPattern::detect(&values);

        assert!(matches!(pattern, AutofillPattern::Arithmetic { start: 1.0, step: 1.0 }));
        assert_eq!(pattern.generate(0), "1");
        assert_eq!(pattern.generate(1), "2");
        assert_eq!(pattern.generate(2), "3");
        assert_eq!(pattern.generate(3), "4");
        assert_eq!(pattern.generate(4), "5");
    }

    #[test]
    fn test_autofill_pattern_arithmetic_larger_step() {
        let values = vec!["10".to_string(), "20".to_string(), "30".to_string()];
        let pattern = AutofillPattern::detect(&values);

        assert!(matches!(pattern, AutofillPattern::Arithmetic { start: 10.0, step: 10.0 }));
        assert_eq!(pattern.generate(3), "40");
        assert_eq!(pattern.generate(4), "50");
    }

    #[test]
    fn test_autofill_pattern_arithmetic_negative_step() {
        let values = vec!["10".to_string(), "5".to_string(), "0".to_string()];
        let pattern = AutofillPattern::detect(&values);

        assert!(matches!(pattern, AutofillPattern::Arithmetic { start: 10.0, step: -5.0 }));
        assert_eq!(pattern.generate(3), "-5");
        assert_eq!(pattern.generate(4), "-10");
    }

    #[test]
    fn test_autofill_pattern_arithmetic_decimal() {
        let values = vec!["0.5".to_string(), "1.0".to_string(), "1.5".to_string()];
        let pattern = AutofillPattern::detect(&values);

        assert!(matches!(pattern, AutofillPattern::Arithmetic { .. }));
        assert_eq!(pattern.generate(3), "2");
        assert_eq!(pattern.generate(4), "2.5");
    }

    #[test]
    fn test_autofill_pattern_prefixed_number_simple() {
        let values = vec!["Item1".to_string(), "Item2".to_string()];
        let pattern = AutofillPattern::detect(&values);

        assert!(matches!(pattern, AutofillPattern::PrefixedNumber { .. }));
        assert_eq!(pattern.generate(0), "Item1");
        assert_eq!(pattern.generate(1), "Item2");
        assert_eq!(pattern.generate(2), "Item3");
        assert_eq!(pattern.generate(3), "Item4");
    }

    #[test]
    fn test_autofill_pattern_prefixed_number_with_gap() {
        let values = vec!["Test10".to_string(), "Test20".to_string()];
        let pattern = AutofillPattern::detect(&values);

        assert!(matches!(pattern, AutofillPattern::PrefixedNumber { .. }));
        assert_eq!(pattern.generate(2), "Test30");
        assert_eq!(pattern.generate(3), "Test40");
    }

    #[test]
    fn test_autofill_pattern_prefixed_number_with_suffix() {
        let values = vec!["Row_1_data".to_string(), "Row_2_data".to_string()];
        let pattern = AutofillPattern::detect(&values);

        assert!(matches!(pattern, AutofillPattern::PrefixedNumber { .. }));
        assert_eq!(pattern.generate(2), "Row_3_data");
        assert_eq!(pattern.generate(3), "Row_4_data");
    }

    #[test]
    fn test_autofill_pattern_days_short() {
        let values = vec!["Mon".to_string(), "Tue".to_string()];
        let pattern = AutofillPattern::detect(&values);

        assert!(matches!(pattern, AutofillPattern::KnownSequence { .. }));
        assert_eq!(pattern.generate(2), "Wed");
        assert_eq!(pattern.generate(3), "Thu");
        assert_eq!(pattern.generate(4), "Fri");
        assert_eq!(pattern.generate(5), "Sat");
        assert_eq!(pattern.generate(6), "Sun");
        // Wraps around
        assert_eq!(pattern.generate(7), "Mon");
    }

    #[test]
    fn test_autofill_pattern_days_full() {
        let values = vec!["Monday".to_string(), "Tuesday".to_string()];
        let pattern = AutofillPattern::detect(&values);

        assert!(matches!(pattern, AutofillPattern::KnownSequence { .. }));
        assert_eq!(pattern.generate(2), "Wednesday");
        assert_eq!(pattern.generate(3), "Thursday");
    }

    #[test]
    fn test_autofill_pattern_days_case_insensitive() {
        let values = vec!["MON".to_string(), "TUE".to_string()];
        let pattern = AutofillPattern::detect(&values);

        // Should still detect as days sequence
        assert!(matches!(pattern, AutofillPattern::KnownSequence { .. }));
    }

    #[test]
    fn test_autofill_pattern_months_short() {
        let values = vec!["Jan".to_string(), "Feb".to_string(), "Mar".to_string()];
        let pattern = AutofillPattern::detect(&values);

        assert!(matches!(pattern, AutofillPattern::KnownSequence { .. }));
        assert_eq!(pattern.generate(3), "Apr");
        assert_eq!(pattern.generate(4), "May");
        // Wraps around
        assert_eq!(pattern.generate(11), "Dec");
        assert_eq!(pattern.generate(12), "Jan");
    }

    #[test]
    fn test_autofill_pattern_months_full() {
        let values = vec!["January".to_string(), "February".to_string()];
        let pattern = AutofillPattern::detect(&values);

        assert!(matches!(pattern, AutofillPattern::KnownSequence { .. }));
        assert_eq!(pattern.generate(2), "March");
        assert_eq!(pattern.generate(3), "April");
    }

    #[test]
    fn test_autofill_pattern_quarters() {
        let values = vec!["Q1".to_string(), "Q2".to_string()];
        let pattern = AutofillPattern::detect(&values);

        assert!(matches!(pattern, AutofillPattern::KnownSequence { .. }));
        assert_eq!(pattern.generate(2), "Q3");
        assert_eq!(pattern.generate(3), "Q4");
        // Wraps around
        assert_eq!(pattern.generate(4), "Q1");
    }

    #[test]
    fn test_autofill_pattern_single_value_copy() {
        let values = vec!["Hello".to_string()];
        let pattern = AutofillPattern::detect(&values);

        assert!(matches!(pattern, AutofillPattern::Copy { .. }));
        assert_eq!(pattern.generate(0), "Hello");
        assert_eq!(pattern.generate(1), "Hello");
        assert_eq!(pattern.generate(100), "Hello");
    }

    #[test]
    fn test_autofill_pattern_empty_values() {
        let values: Vec<String> = vec![];
        let pattern = AutofillPattern::detect(&values);

        assert!(matches!(pattern, AutofillPattern::Copy { .. }));
    }

    #[test]
    fn test_autofill_pattern_mixed_types_fallback() {
        // Mixed types should fall back to copy
        let values = vec!["1".to_string(), "hello".to_string(), "3".to_string()];
        let pattern = AutofillPattern::detect(&values);

        assert!(matches!(pattern, AutofillPattern::Copy { .. }));
        assert_eq!(pattern.generate(0), "1");
    }

    #[test]
    fn test_autofill_pattern_non_arithmetic_numbers() {
        // Numbers that don't form an arithmetic sequence
        let values = vec!["1".to_string(), "2".to_string(), "4".to_string()];
        let pattern = AutofillPattern::detect(&values);

        // Should fall back to copy since 1, 2, 4 is not arithmetic
        assert!(matches!(pattern, AutofillPattern::Copy { .. }));
    }

    #[test]
    fn test_autofill_pattern_description() {
        let arith = AutofillPattern::Arithmetic { start: 1.0, step: 2.0 };
        assert_eq!(arith.description(), "arithmetic sequence (+2)");

        let arith_neg = AutofillPattern::Arithmetic { start: 10.0, step: -5.0 };
        assert_eq!(arith_neg.description(), "arithmetic sequence (-5)");

        let prefixed = AutofillPattern::PrefixedNumber {
            prefix: "Item".to_string(),
            suffix: "".to_string(),
            start: 1.0,
            step: 1.0
        };
        assert_eq!(prefixed.description(), "\"Item...\" sequence (+1)");

        let copy = AutofillPattern::Copy { value: "test".to_string() };
        assert_eq!(copy.description(), "copy");
    }

    #[test]
    fn test_autofill_pattern_format_number() {
        // Whole numbers should not have decimal point
        assert_eq!(AutofillPattern::format_number(5.0), "5");
        assert_eq!(AutofillPattern::format_number(-10.0), "-10");
        assert_eq!(AutofillPattern::format_number(0.0), "0");

        // Decimals should be preserved (trailing zeros removed)
        assert_eq!(AutofillPattern::format_number(5.5), "5.5");
        assert_eq!(AutofillPattern::format_number(3.14159), "3.14159");
    }

    #[test]
    fn test_autofill_pattern_starting_mid_sequence() {
        // Start from Wednesday
        let values = vec!["Wed".to_string(), "Thu".to_string()];
        let pattern = AutofillPattern::detect(&values);

        assert!(matches!(pattern, AutofillPattern::KnownSequence { start_index: 2, .. }));
        assert_eq!(pattern.generate(0), "Wed");
        assert_eq!(pattern.generate(1), "Thu");
        assert_eq!(pattern.generate(2), "Fri");
        assert_eq!(pattern.generate(3), "Sat");
        assert_eq!(pattern.generate(4), "Sun");
        assert_eq!(pattern.generate(5), "Mon");
    }

    #[test]
    fn test_format_number_large_values() {
        // Values within i64 range should format as integers
        assert_eq!(AutofillPattern::format_number(1000.0), "1000");
        assert_eq!(AutofillPattern::format_number(-1000.0), "-1000");

        // Values beyond i64 range should not panic and should format correctly
        let result = AutofillPattern::format_number(1e19);
        assert!(!result.is_empty());
        // Should not produce incorrect i64-saturated value
        assert!(result.starts_with("1000000000000000000"));

        let result = AutofillPattern::format_number(1e20);
        assert!(!result.is_empty());
        assert!(result.starts_with("1000000000000000000"));
    }
}