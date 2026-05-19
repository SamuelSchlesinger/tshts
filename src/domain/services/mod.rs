//! Formula evaluation services for the terminal spreadsheet.
//!
//! This module provides the core formula evaluation engine that can
//! parse and execute spreadsheet formulas with cell references,
//! arithmetic operations, and built-in functions.

#[allow(unused_imports)]
use super::models::{Spreadsheet, Workbook};
#[allow(unused_imports)]
use super::parser::{Parser, ExpressionEvaluator, FunctionRegistry, Expr, Value};
use std::cell::RefCell;
use std::collections::{HashMap, HashSet};
use std::fs::File;

/// Map an unhandled `Err(String)` to the closest Excel error code based on
/// keywords in the message. Falls back to `#ERROR` for genuinely
/// unclassifiable failures (parser errors, unknown functions).
fn classify_err(msg: &str) -> &'static str {
    let m = msg.to_lowercase();
    if m.contains("not found") || m.contains("no match") {
        "#N/A"
    } else if m.contains("division by zero") || m.contains("modulo by zero") {
        "#DIV/0!"
    } else if m.contains("out of range")
        || m.contains("invalid cell reference")
        || m.contains("unknown sheet")
    {
        "#REF!"
    } else if m.contains("unknown function") || m.contains("unknown name") {
        "#NAME?"
    } else if m.contains("requires") || m.contains("expected") || m.contains("bad ") {
        "#VALUE!"
    } else if m.contains("out-of-domain") || m.contains("nan") {
        "#NUM!"
    } else {
        "#ERROR"
    }
}

const PARSE_CACHE_CAP: usize = 256;

thread_local! {
    /// Bounded per-thread cache of recently-parsed formulas. Insertion order
    /// is preserved; oldest entry evicted when full. Small enough to be cheap
    /// and to fit common workload patterns (autofill into a column).
    static PARSE_CACHE: RefCell<Vec<(String, Expr)>> = RefCell::new(Vec::with_capacity(PARSE_CACHE_CAP));
}

/// Run `parse` on `expr`, returning a cached AST if one exists. The cache is
/// keyed by the exact formula string. Lookups are O(N) over a small cap so
/// the linear scan is faster than re-hashing for the typical N.
fn with_parse_cache<F>(expr: &str, parse: F) -> Result<Expr, String>
where
    F: FnOnce(&str) -> Result<Expr, String>,
{
    PARSE_CACHE.with(|cache| {
        let cache_ref = cache.borrow();
        for (k, v) in cache_ref.iter() {
            if k == expr {
                return Ok(v.clone());
            }
        }
        drop(cache_ref);
        let ast = parse(expr)?;
        let mut cache_mut = cache.borrow_mut();
        if cache_mut.len() >= PARSE_CACHE_CAP {
            cache_mut.remove(0);
        }
        cache_mut.push((expr.to_string(), ast.clone()));
        Ok(ast)
    })
}

/// Evaluates spreadsheet formulas. Logical ops (AND/OR/NOT) are functions, not
/// operators, so they participate in the standard precedence chain.
///
/// ```
/// use tshts::domain::{Spreadsheet, FormulaEvaluator};
///
/// let sheet = Spreadsheet::default();
/// let evaluator = FormulaEvaluator::new(&sheet);
/// assert_eq!(evaluator.evaluate_formula("=2+3*4"), "14");
/// assert_eq!(evaluator.evaluate_formula("=AND(1>0, 2<5)"), "1");
/// ```

// Submodules.
mod evaluator;
mod autofill_pattern;
mod csv;

pub use evaluator::FormulaEvaluator;
pub use autofill_pattern::AutofillPattern;
pub use csv::CsvExporter;

#[cfg(test)]
mod tests {
}
