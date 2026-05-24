//! Formula evaluator. Public API is [`FormulaEvaluator`]; the impl is split
//! across files for readability:
//!
//! - `mod.rs` — struct, constructors, top-level `evaluate_formula` /
//!   `would_create_circular_reference` entry points, and the parse-and-eval
//!   inner helper.
//! - `circular.rs` — circular-reference detection (AST walker).
//! - `refs.rs` — cell-reference extraction (drives the dep graph).
//! - `adjust.rs` — same-sheet ref adjustment for paste/autofill/row+col
//!   structural mutations, plus AST↔string serialization.
//! - `cross_sheet.rs` — sheet-qualified ref adjustment for cross-sheet
//!   structural mutations.

#![allow(unused_imports)]
use super::*;
use crate::domain::models::Spreadsheet;
use crate::domain::parser::ErrorKind;

mod circular;
mod refs;
mod adjust;
mod cross_sheet;

#[cfg(test)]
mod tests;

/// Wrap a sheet name in `'...'` quotes if the bare form can't be lexed as
/// an identifier. Names need quoting when they're empty, contain anything
/// outside `[A-Za-z0-9_]`, or start with a digit (the lexer would otherwise
/// tokenize the leading digit as a number — e.g. emitting `1Q!A5` breaks
/// the lexer at the `1`). Apostrophes inside the name are escaped by
/// doubling, matching the lexer's quoted-sheet syntax.
pub(super) fn format_sheet_name(name: &str) -> String {
    let starts_with_digit = name.chars().next().is_some_and(|c| c.is_ascii_digit());
    let has_non_ident = !name.chars().all(|c| c.is_ascii_alphanumeric() || c == '_');
    let needs_quotes = name.is_empty() || starts_with_digit || has_non_ident;
    if needs_quotes {
        format!("'{}'", name.replace('\'', "''"))
    } else {
        name.to_string()
    }
}

pub struct FormulaEvaluator<'a> {
    pub(super) spreadsheet: &'a Spreadsheet,
    /// Optional named-ranges context. Resolution of bare identifiers in
    /// formulas (e.g. `=Revenue + 10`) uses this map; absent → unknown.
    pub(super) names: Option<&'a HashMap<String, String>>,
    /// Optional workbook for cross-sheet refs (`Sheet2!A1`). When absent,
    /// sheet-qualified refs fail with `#REF!`.
    pub(super) workbook: Option<&'a Workbook>,
}

impl<'a> FormulaEvaluator<'a> {
    /// Single-sheet evaluator. Cross-sheet refs return `#REF!` because no
    /// workbook context is supplied; bare identifiers don't resolve because
    /// no named-ranges map is supplied. Use the builder methods
    /// `with_names` / `with_workbook` to enrich the context.
    pub fn new(spreadsheet: &'a Spreadsheet) -> Self {
        Self { spreadsheet, names: None, workbook: None }
    }

    /// Attach a named-ranges map to the evaluator. Bare identifiers in
    /// formulas (e.g. `=Revenue + 10`) resolve through it.
    pub fn with_names(mut self, names: &'a HashMap<String, String>) -> Self {
        self.names = Some(names);
        self
    }

    /// Attach a workbook handle to the evaluator. Sheet-qualified refs
    /// (`Sheet2!A1`) and 3-D ranges (`Sheet1:Sheet3!A1`) resolve through it.
    pub fn with_workbook(mut self, workbook: &'a Workbook) -> Self {
        self.workbook = Some(workbook);
        self
    }

    /// Convenience for the common "I have a workbook plus its named ranges"
    /// case — equivalent to `new(sheet).with_names(...).with_workbook(...)`.
    pub fn for_workbook(
        workbook: &'a Workbook,
        spreadsheet: &'a Spreadsheet,
        names: &'a HashMap<String, String>,
    ) -> Self {
        FormulaEvaluator::new(spreadsheet)
            .with_names(names)
            .with_workbook(workbook)
    }

    /// Returns the evaluated result, "#ERROR" on failure, or `formula`
    /// unchanged if it doesn't start with `=`.
    ///
    /// ```
    /// use tshts::domain::{Spreadsheet, FormulaEvaluator};
    ///
    /// let sheet = Spreadsheet::default();
    /// let evaluator = FormulaEvaluator::new(&sheet);
    /// assert_eq!(evaluator.evaluate_formula("=2+3"), "5");
    /// assert_eq!(evaluator.evaluate_formula("hello"), "hello");
    /// ```
    pub fn evaluate_formula(&self, formula: &str) -> String {
        if let Some(expr) = formula.strip_prefix('=') {
            match self.parse_and_evaluate(expr) {
                Ok(result) => result.to_string(),
                // Unhandled-error fallback: Excel's default is #VALUE! for
                // generic argument issues, but we keep "#ERROR" for the
                // truly-untyped failure mode (parse errors, unknown
                // functions). Heuristic: if the error message implies a
                // recognized error class, surface that; else "#ERROR".
                Err(msg) => classify_err(&msg).to_string(),
            }
        } else {
            formula.to_string()
        }
    }

    /// Walks the AST checking whether `formula` directly or transitively
    /// references `current_cell`.
    ///
    /// ```
    /// use tshts::domain::{Spreadsheet, FormulaEvaluator};
    ///
    /// let sheet = Spreadsheet::default();
    /// let evaluator = FormulaEvaluator::new(&sheet);
    /// assert!(evaluator.would_create_circular_reference("=A1+1", (0, 0)));
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

    /// Parses and evaluates an expression using the new parser. Uses a
    /// thread-local LRU cache so repeated formulas (common during paste/
    /// autofill) don't re-tokenize.
    fn parse_and_evaluate(&self, expr: &str) -> Result<Value, String> {
        let ast = with_parse_cache(expr, |s| {
            let mut parser = Parser::new(s)?;
            parser.parse()
        })?;
        // Reuse the thread-local shared built-in registry instead of
        // rebuilding a ~140-entry HashMap on every formula evaluation.
        let function_registry = FunctionRegistry::shared_builtin();
        let evaluator = ExpressionEvaluator::new(
            self.spreadsheet,
            &function_registry,
            self.names,
            self.workbook,
        );
        evaluator.evaluate(&ast)
    }
}
