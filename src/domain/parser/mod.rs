//! Expression parser for spreadsheet formulas.
//!
//! This module implements a recursive descent parser for spreadsheet expressions
//! based on a formal BNF grammar. It provides a clean separation between binary
//! operators and function calls, with all logical operations implemented as functions
//! for consistency and extensibility.
//!
//! # BNF Grammar
//!
//! The parser implements the following BNF grammar for expressions:
//!
//! ```bnf
//! Expression     ::= Equality
//! Equality       ::= Comparison ( ( "<>" | "=" ) Comparison )*
//! Comparison     ::= Concatenation ( ( "<" | "<=" | ">" | ">=" ) Concatenation )*
//! Concatenation  ::= Addition ( "&" Addition )*
//! Addition       ::= Multiplication ( ( "+" | "-" ) Multiplication )*
//! Multiplication ::= Power ( ( "*" | "/" | "%" ) Power )*
//! Power          ::= Unary ( ( "**" | "^" ) Unary )*
//! Unary          ::= ( "+" | "-" )? Primary
//! Primary        ::= Number | CellRef | Range | FunctionCall | "(" Expression ")"
//! FunctionCall   ::= Identifier "(" ArgumentList? ")"
//! ArgumentList   ::= Expression ( "," Expression )*
//! Range          ::= CellRef ":" CellRef
//! CellRef        ::= [A-Z]+ [0-9]+
//! Number         ::= [0-9]+ ( "." [0-9]+ )?
//! Identifier     ::= [A-Z][A-Z0-9_]*
//! ```
//!
//! This grammar ensures proper operator precedence and associativity:
//! - Comparison operators (<, >, <=, >=, <>, =) have lowest precedence
//! - Arithmetic operators (+, -, *, /, %)
//! - Power operators (**, ^) have highest precedence among binary operators
//! - Unary operators (+, -) have higher precedence than binary
//! - Parentheses override precedence
//! - Function calls and primary expressions have highest precedence
//! - Logical operations (AND, OR, NOT) are implemented as functions

use std::collections::HashMap;
use super::models::Spreadsheet;

/// Represents a token in the expression.
#[derive(Debug, Clone, PartialEq)]
pub enum Token {
    // Literals
    Number(f64),
    String(String),
    CellRef(String),
    Identifier(String),
    /// An Excel-style error literal lexed from the source (`#REF!`, `#N/A`,
    /// `#DIV/0!`, …). The parser turns this into `Expr::ErrorLit` so error
    /// propagation works through the AST in the same way evaluated errors do.
    ErrorLit(ErrorKind),
    
    // Operators
    Plus,
    Minus,
    Multiply,
    Divide,
    Modulo,
    Power,
    PowerAlt,     // ^ alternative to **
    
    // Comparison operators
    Less,
    LessEqual,
    Greater,
    GreaterEqual,
    NotEqual,
    Equal,
    
    
    // Delimiters
    LeftParen,
    RightParen,
    Comma,
    Colon,
    /// `!` — separates a sheet name from a cell/range ref: `Sheet2!A1`.
    Bang,
    /// `{` — opens an array literal: `{1,2;3,4}`.
    LeftBrace,
    /// `}` — closes an array literal.
    RightBrace,
    /// `;` — separates rows in an array literal.
    Semicolon,
    
    // String concatenation
    Ampersand,
    
    // End of input
    Eof,
}

/// A spreadsheet value. `List` is a 1-D vector (e.g. flattened row);
/// `Array` carries explicit (rows × cols) shape so range-aware functions
/// like `VLOOKUP`/`HLOOKUP`/`INDEX` can index by row+col instead of guessing
/// the width. `Error` lets formulas carry typed Excel-style errors that
/// `IFERROR`/`ISERROR` can trap and arithmetic propagates.
#[derive(Debug, Clone, PartialEq)]
pub enum Value {
    Number(f64),
    String(String),
    Bool(bool),
    List(Vec<Value>),
    Array {
        rows: usize,
        cols: usize,
        data: Vec<Value>,
    },
    Error(ErrorKind),
}

/// Excel-style error codes. Display strings match Excel exactly.
#[allow(dead_code)] // Several variants are reserved for upcoming Phase items.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ErrorKind {
    /// `#DIV/0!` — division/modulo by zero.
    Div0,
    /// `#REF!` — invalid or deleted reference.
    Ref,
    /// `#VALUE!` — wrong type for an operation.
    Value,
    /// `#NAME?` — unknown function or named range.
    Name,
    /// `#NUM!` — out-of-domain numeric result.
    Num,
    /// `#N/A` — value not available (lookup miss).
    NA,
    /// `#NULL!` — intersection of non-overlapping ranges (rare).
    Null,
    /// `#SPILL!` — dynamic array can't expand into occupied cells.
    Spill,
}

impl ErrorKind {
    pub fn as_str(&self) -> &'static str {
        match self {
            ErrorKind::Div0 => "#DIV/0!",
            ErrorKind::Ref => "#REF!",
            ErrorKind::Value => "#VALUE!",
            ErrorKind::Name => "#NAME?",
            ErrorKind::Num => "#NUM!",
            ErrorKind::NA => "#N/A",
            ErrorKind::Null => "#NULL!",
            ErrorKind::Spill => "#SPILL!",
        }
    }
}

impl Value {
    /// Converts value to string representation.
    /// For lists/arrays, returns the first element's string (matches Excel's
    /// "implicit intersection" — a multi-cell value in a scalar context
    /// degrades to the top-left).
    pub fn to_string(&self) -> String {
        match self {
            Value::Number(n) => n.to_string(),
            Value::String(s) => s.clone(),
            Value::Bool(b) => if *b { "TRUE".to_string() } else { "FALSE".to_string() },
            Value::List(items) => items
                .first()
                .map(|v| v.to_string())
                .unwrap_or_default(),
            Value::Array { data, .. } => data
                .first()
                .map(|v| v.to_string())
                .unwrap_or_default(),
            Value::Error(e) => e.as_str().to_string(),
        }
    }

    pub fn to_number(&self) -> f64 {
        match self {
            Value::Number(n) => *n,
            Value::String(s) => s.parse::<f64>().unwrap_or(0.0),
            Value::Bool(b) => if *b { 1.0 } else { 0.0 },
            Value::List(items) => items.first().map(|v| v.to_number()).unwrap_or(0.0),
            Value::Array { data, .. } => data.first().map(|v| v.to_number()).unwrap_or(0.0),
            Value::Error(_) => 0.0,
        }
    }

    pub fn is_truthy(&self) -> bool {
        match self {
            Value::Number(n) => *n != 0.0,
            Value::String(s) => !s.is_empty(),
            Value::Bool(b) => *b,
            Value::List(items) => items.first().map(|v| v.is_truthy()).unwrap_or(false),
            Value::Array { data, .. } => data.first().map(|v| v.is_truthy()).unwrap_or(false),
            Value::Error(_) => false,
        }
    }

    /// True if this value is an error (or a list/array containing one).
    pub fn is_error(&self) -> bool {
        match self {
            Value::Error(_) => true,
            Value::List(items) => items.iter().any(|v| v.is_error()),
            Value::Array { data, .. } => data.iter().any(|v| v.is_error()),
            _ => false,
        }
    }

    /// Return the first contained error, if any.
    pub fn first_error(&self) -> Option<ErrorKind> {
        match self {
            Value::Error(e) => Some(*e),
            Value::List(items) | Value::Array { data: items, .. } => {
                items.iter().find_map(|v| v.first_error())
            }
            _ => None,
        }
    }

    /// Flatten arrays and lists into a vector of scalar values.
    pub fn flatten(&self) -> Vec<Value> {
        match self {
            Value::List(items) => items.iter().flat_map(|v| v.flatten()).collect(),
            Value::Array { data, .. } => data.iter().flat_map(|v| v.flatten()).collect(),
            _ => vec![self.clone()],
        }
    }
}

/// Convenience: flatten an iterator of values into a single vec.
pub fn flatten_args(args: &[Value]) -> Vec<Value> {
    args.iter().flat_map(|v| v.flatten()).collect()
}

/// Returns (rows, cols, data) for a value. Scalars are 1×1. `List` is 1×N
/// (or N×1 — we treat it as a single row by convention).
pub fn shape_of(v: &Value) -> (usize, usize, Vec<Value>) {
    match v {
        Value::Array { rows, cols, data } => (*rows, *cols, data.clone()),
        Value::List(items) => (1, items.len(), items.clone()),
        scalar => (1, 1, vec![scalar.clone()]),
    }
}

/// Broadcast a binary scalar op across array operands. Scalar × scalar is a
/// scalar result. Array × scalar broadcasts the scalar across the array.
/// Array × array element-wise (must match shape).
pub fn broadcast_binary<F>(left: &Value, right: &Value, op: F) -> Result<Value, String>
where
    F: Fn(&Value, &Value) -> Result<Value, String>,
{
    use Value::*;
    match (left, right) {
        (Array { rows: lr, cols: lc, data: ld },
         Array { rows: rr, cols: rc, data: rd }) => {
            if lr != rr || lc != rc {
                return Err(format!(
                    "Array shape mismatch: {}x{} vs {}x{}",
                    lr, lc, rr, rc
                ));
            }
            let data: Result<Vec<_>, _> =
                ld.iter().zip(rd.iter()).map(|(a, b)| op(a, b)).collect();
            Ok(Array { rows: *lr, cols: *lc, data: data? })
        }
        (Array { rows, cols, data }, scalar) => {
            let data: Result<Vec<_>, _> =
                data.iter().map(|a| op(a, scalar)).collect();
            Ok(Array { rows: *rows, cols: *cols, data: data? })
        }
        (scalar, Array { rows, cols, data }) => {
            let data: Result<Vec<_>, _> =
                data.iter().map(|b| op(scalar, b)).collect();
            Ok(Array { rows: *rows, cols: *cols, data: data? })
        }
        (List(la), List(lb)) => {
            if la.len() != lb.len() {
                return Err(format!("List length mismatch: {} vs {}", la.len(), lb.len()));
            }
            let data: Result<Vec<_>, _> =
                la.iter().zip(lb.iter()).map(|(a, b)| op(a, b)).collect();
            Ok(List(data?))
        }
        (List(items), scalar) => {
            let data: Result<Vec<_>, _> =
                items.iter().map(|a| op(a, scalar)).collect();
            Ok(List(data?))
        }
        (scalar, List(items)) => {
            let data: Result<Vec<_>, _> =
                items.iter().map(|b| op(scalar, b)).collect();
            Ok(List(data?))
        }
        // Scalar × scalar
        _ => op(left, right),
    }
}

/// Represents an Abstract Syntax Tree node for expressions.
#[derive(Debug, Clone, PartialEq)]
pub enum Expr {
    Number(f64),
    String(String),
    CellRef(String),
    Range(String, String),
    /// Named reference (e.g. `Revenue` resolves to whatever range/cell the
    /// workbook's `named_ranges` says). Evaluated by looking up the name in
    /// the evaluator's context and recursively evaluating the resolved expr.
    NamedRef(String),
    /// `LET(name1, value1, name2, value2, ..., body)` — local bindings.
    Let {
        bindings: Vec<(String, Box<Expr>)>,
        body: Box<Expr>,
    },
    /// `LAMBDA(param1, param2, ..., body)` — a first-class function. Calling
    /// it (`=LAMBDA(x, x*2)(5)`) evaluates `body` with params bound.
    Lambda {
        params: Vec<String>,
        body: Box<Expr>,
    },
    /// `{1,2;3,4}` array literal. Inner Vec is one row.
    ArrayLiteral {
        rows: Vec<Vec<Expr>>,
    },

    Binary {
        left: Box<Expr>,
        operator: BinaryOp,
        right: Box<Expr>,
    },
    Unary {
        operator: UnaryOp,
        operand: Box<Expr>,
    },
    FunctionCall {
        name: String,
        args: Vec<Expr>,
    },
    /// Excel-style error literal carried through the AST. Evaluates to
    /// `Value::Error(kind)`, and the existing first-error propagation in
    /// Binary/Unary/FunctionCall cascades it correctly. Emitted by the lexer
    /// for source-level `#REF!` etc., and by `adjust_*` when a relative
    /// shift takes a cell reference past the origin or onto a deleted row.
    ErrorLit(ErrorKind),
}

/// Binary operators with their precedence and evaluation behavior.
#[derive(Debug, Clone, PartialEq)]
pub enum BinaryOp {
    
    // Equality
    Equal,
    NotEqual,
    
    // Comparison
    Less,
    LessEqual,
    Greater,
    GreaterEqual,
    
    // Arithmetic
    Add,
    Subtract,
    Multiply,
    Divide,
    Modulo,
    
    // Power (highest precedence among binary)
    Power,
    
    // String operations
    Concatenate, // & operator for string concatenation
}

/// Unary operators.
#[derive(Debug, Clone, PartialEq)]
pub enum UnaryOp {
    Plus,
    Minus,
}

// Submodules — separated by structural concern.
mod lexer;
mod registry;
mod registry_fns; // category-split builtin function files
mod parser_impl;
mod evaluator;

pub use lexer::Lexer;
pub use registry::FunctionRegistry;
pub use parser_impl::Parser;
pub use evaluator::ExpressionEvaluator;
pub use evaluator::formula_purity;


/// Function signature for built-in and user-defined functions.
pub type FunctionImpl = fn(&[Value]) -> Result<Value, String>;

/// How a function interacts with shared state and parallel evaluation.
///
/// The classification follows Excel's terminology and OpenFormula
/// §3.5, refined into categories that have distinct scheduling
/// implications under parallel calc:
///
/// - **Pure**: idempotent, side-effect-free. Safe to dispatch on any
///   worker; safe to memoize.
/// - **VolatileClock**: reads the wall clock (NOW, TODAY). The recalc
///   executor snapshots a single timestamp at pass start into a
///   `RecalcContext`; all volatile-clock cells read the snapshot, so
///   they remain parallel-safe and return consistent values within a
///   pass.
/// - **VolatileRandom**: advances a random-number generator (RAND,
///   RANDBETWEEN, RANDARRAY). tshts seeds a thread-local PRNG per
///   worker; cells get unrelated random values across runs but
///   correct within-pass independence.
/// - **VolatileStructural**: changes the dep graph based on cell
///   *values*, not just formulas (INDIRECT, OFFSET when args are
///   dynamic, CELL("address" / "format"), FORMULATEXT, INFO). The
///   scheduler dispatches these cells serially because the parallel
///   path assumes a static graph for the duration of a level.
/// - **SideEffecting**: external I/O (GET). Rate-limited and cached
///   to avoid hammering external services from N workers in parallel.
///
/// Combining: when a formula references multiple functions, its
/// effective purity is the **maximum** of their purities under the
/// total order `Pure < VolatileClock < VolatileStructural <
/// VolatileRandom < SideEffecting`. The ordering reflects how strongly
/// each category constrains parallel dispatch (Random shares mutable
/// state ordering across cells more loosely than Structural shares
/// graph identity, and SideEffecting is the strongest).
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub enum FunctionPurity {
    Pure,
    VolatileClock,
    VolatileStructural,
    VolatileRandom,
    SideEffecting,
}

impl FunctionPurity {
    /// True if a formula with this purity can be safely dispatched on
    /// a rayon worker as part of a level batch. Pure / VolatileClock /
    /// VolatileRandom qualify; VolatileStructural and SideEffecting
    /// run sequentially.
    #[allow(dead_code)]
    pub fn is_parallel_safe(self) -> bool {
        matches!(
            self,
            FunctionPurity::Pure
                | FunctionPurity::VolatileClock
                | FunctionPurity::VolatileRandom
        )
    }

    /// True if this purity requires the recalc engine to re-evaluate the
    /// cell every pass even when its precedents are unchanged (e.g. NOW
    /// changes between calls without any cell having been mutated).
    /// Used by `Workbook::recalc_via_graph` to seed volatile cells into
    /// the dirty set automatically at pass start.
    #[allow(dead_code)]
    pub fn is_volatile(self) -> bool {
        self != FunctionPurity::Pure
    }

    /// The join (least upper bound) on the purity lattice. Used by the
    /// AST walker to combine the purities of a formula's subexpressions.
    pub fn join(self, other: FunctionPurity) -> FunctionPurity {
        self.max(other)
    }
}

/// Returns names of every registered built-in function. Used for autocomplete
/// and the help library.
pub fn builtin_function_names() -> Vec<&'static str> {
    // Hand-maintained — registry doesn't preserve order or expose iteration.
    // Update when adding functions in `register_builtin_functions`.
    vec![
        "SUM", "AVERAGE", "MIN", "MAX", "IF", "AND", "OR", "NOT",
        "ABS", "SQRT", "ROUND", "CONCAT", "LEN", "LEFT", "RIGHT", "MID", "FIND",
        "UPPER", "LOWER", "TRIM", "GET",
        "CEILING", "FLOOR", "INT", "MOD", "LOG", "LN", "EXP", "PI", "RAND",
        "RANDBETWEEN", "SIGN", "POWER",
        "SUBSTITUTE", "REPLACE", "REPT", "EXACT", "PROPER", "CLEAN",
        "CHAR", "CODE", "TEXT", "VALUE", "NUMBERVALUE",
        "IFERROR", "IFNA", "ISERROR", "ISERR", "ISNA", "NA", "ERROR.TYPE",
        "ISBLANK", "ISNUMBER", "ISTEXT", "TYPE", "COUNT", "COUNTA",
        "SPARKLINE",
        "SUMIF", "COUNTIF", "AVERAGEIF",
        "VLOOKUP", "HLOOKUP", "XLOOKUP", "INDEX", "MATCH",
        "TRUE", "FALSE",
        "TODAY", "NOW", "DATE", "YEAR", "MONTH", "DAY",
        "TIME", "HOUR", "MINUTE", "SECOND",
        "DATEDIF", "WEEKDAY", "EDATE", "EOMONTH", "DAYS",
        "IFS", "SWITCH", "XOR",
        "MEDIAN", "STDEV.S", "STDEV", "STDEV.P", "VAR.S", "VAR.P",
        "LARGE", "SMALL", "RANK.EQ", "PERCENTILE.INC", "CORREL",
        "PMT", "FV", "PV", "NPV",
        "TRUNC", "ATAN", "ATAN2", "ASIN", "ACOS", "SINH", "COSH", "TANH",
        "SIN", "COS", "TAN", "DEGREES", "RADIANS", "FACT", "COMBIN",
        "GCD", "LCM", "ROUNDUP", "ROUNDDOWN", "MROUND", "EVEN", "ODD",
        "TEXTJOIN", "SEARCH", "TEXTBEFORE", "TEXTAFTER",
        "REGEXMATCH", "REGEXEXTRACT", "REGEXREPLACE",
        "NETWORKDAYS", "WORKDAY", "DATEVALUE", "TIMEVALUE", "YEARFRAC",
        "UNICHAR", "UNICODE", "DOLLAR", "FIXED",
        "ARRAYTOTEXT", "FREQUENCY",
        "MAP", "REDUCE", "BYROW", "BYCOL", "SCAN", "MAKEARRAY",
        "SUMPRODUCT", "TRANSPOSE", "SEQUENCE", "FILTER", "SORT", "UNIQUE",
        "INDIRECT", "OFFSET",
        "LET", "LAMBDA",
    ]
}

/// Tests a value against an Excel-style criteria string.
/// Supported forms: "5" (equality), ">5", "<5", ">=5", "<=5", "<>5",
/// "*wildcard*" (glob; `?` matches one char, `*` matches any).
pub fn criteria_matches(value: &Value, criteria: &str) -> bool {
    let c = criteria.trim();
    let (op, rest) = if let Some(r) = c.strip_prefix(">=") {
        (">=", r)
    } else if let Some(r) = c.strip_prefix("<=") {
        ("<=", r)
    } else if let Some(r) = c.strip_prefix("<>") {
        ("<>", r)
    } else if let Some(r) = c.strip_prefix('>') {
        (">", r)
    } else if let Some(r) = c.strip_prefix('<') {
        ("<", r)
    } else if let Some(r) = c.strip_prefix('=') {
        ("=", r)
    } else {
        ("=", c)
    };
    let rest = rest.trim();
    // Numeric comparison only when both sides parse as numbers AND the value
    // is either a Number or a String whose contents parse cleanly. Bool/Error
    // values fall through to string compare (Excel: `COUNTIF(["#REF!"], ">3")`
    // does not match).
    let value_num: Option<f64> = match value {
        Value::Number(n) => Some(*n),
        Value::String(s) => s.trim().parse::<f64>().ok(),
        Value::Bool(b) => Some(if *b { 1.0 } else { 0.0 }),
        _ => None,
    };
    if let (Ok(a), Some(b)) = (rest.parse::<f64>(), value_num) {
        match op {
            ">"  => return b >  a,
            "<"  => return b <  a,
            ">=" => return b >= a,
            "<=" => return b <= a,
            "<>" => return (b - a).abs() > 1e-9,
            "="  => return (b - a).abs() < 1e-9,
            _ => {}
        }
    }
    // String comparison with optional wildcards
    let s = value.to_string();
    match op {
        "=" => glob_match(&s, rest),
        "<>" => !glob_match(&s, rest),
        ">" => s.as_str() > rest,
        "<" => s.as_str() < rest,
        ">=" => s.as_str() >= rest,
        "<=" => s.as_str() <= rest,
        _ => s == rest,
    }
}

/// Glob match supporting `*` (any) and `?` (one). Case-sensitive.
pub(super) fn glob_match(s: &str, pattern: &str) -> bool {
    if !pattern.contains('*') && !pattern.contains('?') {
        return s == pattern;
    }
    fn rec(s: &[char], p: &[char]) -> bool {
        match (s.split_first(), p.split_first()) {
            (None, None) => true,
            (_, Some((&'*', rest))) => rec(s, rest) || s.split_first().is_some_and(|(_, sr)| rec(sr, p)),
            (Some((_, sr)), Some((&'?', rest))) => rec(sr, rest),
            (Some((sc, sr)), Some((pc, rest))) if sc == pc => rec(sr, rest),
            _ => false,
        }
    }
    let sc: Vec<char> = s.chars().collect();
    let pc: Vec<char> = pattern.chars().collect();
    rec(&sc, &pc)
}

// Date serial helpers — Excel-style epoch 1899-12-30 = 0, modulo the famous
// 1900-leap-year bug which we elide here (we treat days correctly).
pub(super) fn date_to_serial(year: i32, month: u32, day: u32) -> f64 {
    // Excel-style rollover: normalize month into [1..=12] (carrying year),
    // then compute the serial for day=1 of that month and add (day - 1) so
    // out-of-range days roll into adjacent months. This makes
    // DATE(2023,2,29) → 2023-03-01 and DATE(2023,3,0) → 2023-02-28.
    let mut y_norm = year as i64;
    let m_signed = month as i64;
    let total_months = y_norm * 12 + (m_signed - 1);
    let m_norm = total_months.rem_euclid(12) + 1;
    y_norm = total_months.div_euclid(12);

    let y_for_algo = if m_norm <= 2 { y_norm - 1 } else { y_norm };
    let m_for_algo = if m_norm <= 2 { m_norm + 12 } else { m_norm };
    let era = if y_for_algo >= 0 { y_for_algo } else { y_for_algo - 399 } / 400;
    let yoe = y_for_algo - era * 400;
    let doy = (153 * (m_for_algo - 3) + 2) / 5; // day-of-year for day=1
    let doe = yoe * 365 + yoe / 4 - yoe / 100 + doy;
    let days_since_1970 = era * 146097 + doe - 719468;
    // Add (day - 1) as a signed offset so day=0 rolls back.
    let serial = (days_since_1970 + 25569) + (day as i64 - 1);
    serial as f64
}

/// Public alias for use by the XLSX importer, which needs to render Excel
/// date serials as ISO strings instead of bare floats.
pub fn serial_to_date_pub(serial: f64) -> (i32, u32, u32) {
    serial_to_date(serial)
}

pub(super) fn serial_to_date(serial: f64) -> (i32, u32, u32) {
    let z = (serial as i64) - 25569 + 719468;
    let era = if z >= 0 { z } else { z - 146096 } / 146097;
    let doe = (z - era * 146097) as u64;
    let yoe = (doe - doe / 1460 + doe / 36524 - doe / 146096) / 365;
    let y = yoe as i64 + era * 400;
    let doy = doe - (365 * yoe + yoe / 4 - yoe / 100);
    let mp = (5 * doy + 2) / 153;
    let d = (doy - (153 * mp + 2) / 5 + 1) as u32;
    let m = if mp < 10 { mp + 3 } else { mp - 9 } as u32;
    let year = if m <= 2 { y + 1 } else { y } as i32;
    (year, m, d)
}

/// Parse `YYYY-MM-DD` into (year, month, day).
pub(super) fn parse_iso_date(s: &str) -> Result<(i32, u32, u32), ()> {
    let parts: Vec<&str> = s.split('-').collect();
    if parts.len() != 3 {
        return Err(());
    }
    let y: i32 = parts[0].parse().map_err(|_| ())?;
    let m: u32 = parts[1].parse().map_err(|_| ())?;
    let d: u32 = parts[2].parse().map_err(|_| ())?;
    Ok((y, m, d))
}

/// Add thousands-separator commas to an integer-string.
pub(super) fn add_commas(int_str: &str) -> String {
    let mut out = String::with_capacity(int_str.len() + int_str.len() / 3);
    for (i, c) in int_str.chars().rev().enumerate() {
        if i > 0 && i % 3 == 0 {
            out.push(',');
        }
        out.push(c);
    }
    out.chars().rev().collect()
}

pub(super) fn days_in_month(year: i32, month: u32) -> u32 {
    match month {
        1 | 3 | 5 | 7 | 8 | 10 | 12 => 31,
        4 | 6 | 9 | 11 => 30,
        2 => {
            // Gregorian leap rule.
            let leap = (year % 4 == 0 && year % 100 != 0) || year % 400 == 0;
            if leap { 29 } else { 28 }
        }
        _ => 30,
    }
}

pub(super) fn today_serial() -> f64 {
    use std::time::{SystemTime, UNIX_EPOCH};
    let secs = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .map(|d| d.as_secs())
        .unwrap_or(0);
    let days = secs / 86400;
    25569.0 + days as f64
}

pub(super) fn now_serial() -> f64 {
    use std::time::{SystemTime, UNIX_EPOCH};
    let secs = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .map(|d| d.as_secs_f64())
        .unwrap_or(0.0);
    25569.0 + secs / 86400.0
}

/// Tolerance-aware float equality.
/// Returns true when the two numbers are within a relative epsilon scaled
/// to their magnitude, with an absolute floor at `f64::EPSILON` so genuinely-
/// near-zero values (e.g. `0.1 + 0.2 - 0.3`) compare equal.
/// `=0.1+0.2 = 0.3` evaluates true. `=1e-15 = 2e-15` evaluates false —
/// the previous absolute-1e-12 floor lumped far-apart small numbers
/// together, breaking VLOOKUP / MATCH / numeric COUNTIF equality on
/// scientific-data sheets.
pub(crate) fn numbers_equal(l: f64, r: f64) -> bool {
    if l == r {
        return true;
    }
    if !l.is_finite() || !r.is_finite() {
        return false;
    }
    let diff = (l - r).abs();
    let scale = l.abs().max(r.abs());
    diff < scale * 1e-12 || diff < f64::EPSILON
}

/// Expression evaluator that walks the AST and computes results.

#[cfg(test)]
mod tests {
}
