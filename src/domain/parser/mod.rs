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
//! Comparison     ::= Addition ( ( "<" | "<=" | ">" | ">=" ) Addition )*
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
            Value::Bool(b) => if *b { "1".to_string() } else { "0".to_string() },
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

/// Lexical analyzer for tokenizing expressions.

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


/// Function signature for built-in and user-defined functions.
pub type FunctionImpl = fn(&[Value]) -> Result<Value, String>;

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
    // Numeric comparison if both parse as numbers
    if let (Ok(a), Ok(b)) = (rest.parse::<f64>(), Ok::<f64, ()>(value.to_number())) {
        match op {
            ">" => return value.to_number() > a,
            "<" => return value.to_number() < a,
            ">=" => return value.to_number() >= a,
            "<=" => return value.to_number() <= a,
            "<>" => return (b - a).abs() > 1e-9 || !matches!(value, Value::Number(_) | Value::String(_) if {
                // hack: keep numeric check fast
                true
            }),
            "=" => {
                if let Value::Number(_) = value { return (b - a).abs() < 1e-9; }
                // fall through to string match if value isn't numeric
            }
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
    let m_norm = total_months.rem_euclid(12) as i64 + 1;
    y_norm = total_months.div_euclid(12);

    let y_for_algo = if m_norm <= 2 { y_norm - 1 } else { y_norm };
    let m_for_algo = if m_norm <= 2 { m_norm + 12 } else { m_norm };
    let era = if y_for_algo >= 0 { y_for_algo } else { y_for_algo - 399 } / 400;
    let yoe = (y_for_algo - era * 400) as i64;
    let doy = (153 * (m_for_algo - 3) + 2) / 5; // day-of-year for day=1
    let doe = yoe * 365 + yoe / 4 - yoe / 100 + doy;
    let days_since_1970 = era * 146097 + doe - 719468;
    // Add (day - 1) as a signed offset so day=0 rolls back.
    let serial = (days_since_1970 + 25569) as i64 + (day as i64 - 1);
    serial as f64
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

/// Registry for spreadsheet functions.

/// Recursive descent parser for spreadsheet expressions.

/// Tolerance-aware float equality.
/// Returns true if numbers are within absolute or relative epsilon.
/// `=0.1+0.2 = 0.3` should evaluate true; very small numbers near zero should
/// also compare equal.
fn numbers_equal(l: f64, r: f64) -> bool {
    if l == r {
        return true;
    }
    let diff = (l - r).abs();
    if diff < 1e-12 {
        return true;
    }
    let scale = l.abs().max(r.abs());
    diff < scale * 1e-9
}

/// Expression evaluator that walks the AST and computes results.

#[cfg(test)]
mod tests {
    use super::*;
    use crate::domain::CellData;

    fn create_test_spreadsheet() -> Spreadsheet {
        let mut sheet = Spreadsheet::default();
        sheet.set_cell(0, 0, CellData { value: "10".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        sheet.set_cell(0, 1, CellData { value: "20".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        sheet.set_cell(0, 2, CellData { value: "30".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        sheet.set_cell(1, 0, CellData { value: "5".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        sheet.set_cell(1, 1, CellData { value: "15".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        sheet.set_cell(1, 2, CellData { value: "25".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        sheet
    }

    #[test]
    fn test_lexer_numbers() {
        let mut lexer = Lexer::new("42 3.14 0.5");
        
        assert_eq!(lexer.next_token().unwrap(), Token::Number(42.0));
        assert_eq!(lexer.next_token().unwrap(), Token::Number(3.14));
        assert_eq!(lexer.next_token().unwrap(), Token::Number(0.5));
        assert_eq!(lexer.next_token().unwrap(), Token::Eof);
    }

    #[test]
    fn test_lexer_operators() {
        let mut lexer = Lexer::new("+ - * / % ** ^ < <= > >= <> =");
        
        assert_eq!(lexer.next_token().unwrap(), Token::Plus);
        assert_eq!(lexer.next_token().unwrap(), Token::Minus);
        assert_eq!(lexer.next_token().unwrap(), Token::Multiply);
        assert_eq!(lexer.next_token().unwrap(), Token::Divide);
        assert_eq!(lexer.next_token().unwrap(), Token::Modulo);
        assert_eq!(lexer.next_token().unwrap(), Token::Power);
        assert_eq!(lexer.next_token().unwrap(), Token::PowerAlt);
        assert_eq!(lexer.next_token().unwrap(), Token::Less);
        assert_eq!(lexer.next_token().unwrap(), Token::LessEqual);
        assert_eq!(lexer.next_token().unwrap(), Token::Greater);
        assert_eq!(lexer.next_token().unwrap(), Token::GreaterEqual);
        assert_eq!(lexer.next_token().unwrap(), Token::NotEqual);
        assert_eq!(lexer.next_token().unwrap(), Token::Equal);
        assert_eq!(lexer.next_token().unwrap(), Token::Eof);
    }

    #[test]
    fn test_lexer_identifiers_and_keywords() {
        let mut lexer = Lexer::new("SUM AVERAGE AND OR NOT A1 B2 AA123");
        
        assert_eq!(lexer.next_token().unwrap(), Token::Identifier("SUM".to_string()));
        assert_eq!(lexer.next_token().unwrap(), Token::Identifier("AVERAGE".to_string()));
        assert_eq!(lexer.next_token().unwrap(), Token::Identifier("AND".to_string()));
        assert_eq!(lexer.next_token().unwrap(), Token::Identifier("OR".to_string()));
        assert_eq!(lexer.next_token().unwrap(), Token::Identifier("NOT".to_string()));
        assert_eq!(lexer.next_token().unwrap(), Token::CellRef("A1".to_string()));
        assert_eq!(lexer.next_token().unwrap(), Token::CellRef("B2".to_string()));
        assert_eq!(lexer.next_token().unwrap(), Token::CellRef("AA123".to_string()));
        assert_eq!(lexer.next_token().unwrap(), Token::Eof);
    }

    #[test]
    fn test_lexer_delimiters() {
        let mut lexer = Lexer::new("( ) , :");
        
        assert_eq!(lexer.next_token().unwrap(), Token::LeftParen);
        assert_eq!(lexer.next_token().unwrap(), Token::RightParen);
        assert_eq!(lexer.next_token().unwrap(), Token::Comma);
        assert_eq!(lexer.next_token().unwrap(), Token::Colon);
        assert_eq!(lexer.next_token().unwrap(), Token::Eof);
    }

    #[test]
    fn test_parser_numbers() {
        let mut parser = Parser::new("42").unwrap();
        let expr = parser.parse().unwrap();
        assert_eq!(expr, Expr::Number(42.0));
        
        let mut parser = Parser::new("3.14").unwrap();
        let expr = parser.parse().unwrap();
        assert_eq!(expr, Expr::Number(3.14));
    }

    #[test]
    fn test_parser_cell_references() {
        let mut parser = Parser::new("A1").unwrap();
        let expr = parser.parse().unwrap();
        assert_eq!(expr, Expr::CellRef("A1".to_string()));
        
        let mut parser = Parser::new("B2").unwrap();
        let expr = parser.parse().unwrap();
        assert_eq!(expr, Expr::CellRef("B2".to_string()));
    }

    #[test]
    fn test_parser_ranges() {
        let mut parser = Parser::new("A1:C3").unwrap();
        let expr = parser.parse().unwrap();
        assert_eq!(expr, Expr::Range("A1".to_string(), "C3".to_string()));
    }

    #[test]
    fn test_parser_binary_operations() {
        let mut parser = Parser::new("2 + 3").unwrap();
        let expr = parser.parse().unwrap();
        match expr {
            Expr::Binary { left, operator, right } => {
                assert!(matches!(left.as_ref(), &Expr::Number(2.0)));
                assert_eq!(operator, BinaryOp::Add);
                assert!(matches!(right.as_ref(), &Expr::Number(3.0)));
            }
            _ => panic!("Expected binary expression"),
        }
        
        let mut parser = Parser::new("A1 * B1").unwrap();
        let expr = parser.parse().unwrap();
        match expr {
            Expr::Binary { left, operator, right } => {
                assert!(matches!(left.as_ref(), &Expr::CellRef(ref s) if s == "A1"));
                assert_eq!(operator, BinaryOp::Multiply);
                assert!(matches!(right.as_ref(), &Expr::CellRef(ref s) if s == "B1"));
            }
            _ => panic!("Expected binary expression"),
        }
    }

    #[test]
    fn test_parser_operator_precedence() {
        // Test that 2 + 3 * 4 is parsed as 2 + (3 * 4)
        let mut parser = Parser::new("2 + 3 * 4").unwrap();
        let expr = parser.parse().unwrap();
        match expr {
            Expr::Binary { left, operator: BinaryOp::Add, right } => {
                assert!(matches!(left.as_ref(), &Expr::Number(2.0)));
                match right.as_ref() {
                    Expr::Binary { left: mult_left, operator: BinaryOp::Multiply, right: mult_right } => {
                        assert!(matches!(mult_left.as_ref(), &Expr::Number(3.0)));
                        assert!(matches!(mult_right.as_ref(), &Expr::Number(4.0)));
                    }
                    _ => panic!("Expected multiplication as right operand"),
                }
            }
            _ => panic!("Expected addition at top level"),
        }
    }

    #[test]
    fn test_parser_power_right_associative() {
        // Test that 2 ** 3 ** 2 is parsed as 2 ** (3 ** 2)
        let mut parser = Parser::new("2 ** 3 ** 2").unwrap();
        let expr = parser.parse().unwrap();
        match expr {
            Expr::Binary { left, operator: BinaryOp::Power, right } => {
                assert!(matches!(left.as_ref(), &Expr::Number(2.0)));
                match right.as_ref() {
                    Expr::Binary { left: pow_left, operator: BinaryOp::Power, right: pow_right } => {
                        assert!(matches!(pow_left.as_ref(), &Expr::Number(3.0)));
                        assert!(matches!(pow_right.as_ref(), &Expr::Number(2.0)));
                    }
                    _ => panic!("Expected power as right operand"),
                }
            }
            _ => panic!("Expected power at top level"),
        }
    }

    #[test]
    fn test_parser_unary_operations() {
        let mut parser = Parser::new("-5").unwrap();
        let expr = parser.parse().unwrap();
        match expr {
            Expr::Unary { operator, operand } => {
                assert_eq!(operator, UnaryOp::Minus);
                assert!(matches!(operand.as_ref(), &Expr::Number(5.0)));
            }
            _ => panic!("Expected unary expression"),
        }
        
        // NOT is now a function, not a unary operator
        let mut parser = Parser::new("NOT(A1)").unwrap();
        let expr = parser.parse().unwrap();
        match expr {
            Expr::FunctionCall { name, args } => {
                assert_eq!(name, "NOT");
                assert_eq!(args.len(), 1);
                assert!(matches!(args[0], Expr::CellRef(ref s) if s == "A1"));
            }
            _ => panic!("Expected function call expression"),
        }
    }

    #[test]
    fn test_parser_parentheses() {
        let mut parser = Parser::new("(2 + 3) * 4").unwrap();
        let expr = parser.parse().unwrap();
        match expr {
            Expr::Binary { left, operator: BinaryOp::Multiply, right } => {
                match left.as_ref() {
                    Expr::Binary { left: add_left, operator: BinaryOp::Add, right: add_right } => {
                        assert!(matches!(add_left.as_ref(), &Expr::Number(2.0)));
                        assert!(matches!(add_right.as_ref(), &Expr::Number(3.0)));
                    }
                    _ => panic!("Expected addition in parentheses"),
                }
                assert!(matches!(right.as_ref(), &Expr::Number(4.0)));
            }
            _ => panic!("Expected multiplication at top level"),
        }
    }

    #[test]
    fn test_parser_function_calls() {
        let mut parser = Parser::new("SUM(A1, B1, C1)").unwrap();
        let expr = parser.parse().unwrap();
        match expr {
            Expr::FunctionCall { name, args } => {
                assert_eq!(name, "SUM");
                assert_eq!(args.len(), 3);
                assert_eq!(args[0], Expr::CellRef("A1".to_string()));
                assert_eq!(args[1], Expr::CellRef("B1".to_string()));
                assert_eq!(args[2], Expr::CellRef("C1".to_string()));
            }
            _ => panic!("Expected function call"),
        }
        
        let mut parser = Parser::new("SUM(A1:C1)").unwrap();
        let expr = parser.parse().unwrap();
        match expr {
            Expr::FunctionCall { name, args } => {
                assert_eq!(name, "SUM");
                assert_eq!(args.len(), 1);
                assert_eq!(args[0], Expr::Range("A1".to_string(), "C1".to_string()));
            }
            _ => panic!("Expected function call"),
        }
    }

    #[test]
    fn test_parser_comparison_operations() {
        let mut parser = Parser::new("A1 > B1").unwrap();
        let expr = parser.parse().unwrap();
        match expr {
            Expr::Binary { left, operator, right } => {
                assert!(matches!(left.as_ref(), &Expr::CellRef(ref s) if s == "A1"));
                assert_eq!(operator, BinaryOp::Greater);
                assert!(matches!(right.as_ref(), &Expr::CellRef(ref s) if s == "B1"));
            }
            _ => panic!("Expected binary expression"),
        }
        
        let mut parser = Parser::new("5 <= 10").unwrap();
        let expr = parser.parse().unwrap();
        match expr {
            Expr::Binary { left, operator, right } => {
                assert!(matches!(left.as_ref(), &Expr::Number(5.0)));
                assert_eq!(operator, BinaryOp::LessEqual);
                assert!(matches!(right.as_ref(), &Expr::Number(10.0)));
            }
            _ => panic!("Expected binary expression"),
        }
    }

    #[test]
    fn test_parser_logical_operations() {
        // Logical operations are now functions, test AND function call
        let mut parser = Parser::new("AND(A1 > 5, B1 < 10)").unwrap();
        let expr = parser.parse().unwrap();
        match expr {
            Expr::FunctionCall { name, args } => {
                assert_eq!(name, "AND");
                assert_eq!(args.len(), 2);
                
                // First argument should be A1 > 5
                match &args[0] {
                    Expr::Binary { left: comp_left, operator: BinaryOp::Greater, right: comp_right } => {
                        assert!(matches!(comp_left.as_ref(), &Expr::CellRef(ref s) if s == "A1"));
                        assert!(matches!(comp_right.as_ref(), &Expr::Number(5.0)));
                    }
                    _ => panic!("Expected comparison in first argument"),
                }
                
                // Second argument should be B1 < 10
                match &args[1] {
                    Expr::Binary { left: comp_left, operator: BinaryOp::Less, right: comp_right } => {
                        assert!(matches!(comp_left.as_ref(), &Expr::CellRef(ref s) if s == "B1"));
                        assert!(matches!(comp_right.as_ref(), &Expr::Number(10.0)));
                    }
                    _ => panic!("Expected comparison in second argument"),
                }
            }
            _ => panic!("Expected function call"),
        }
    }

    #[test]
    fn test_expression_evaluator_numbers() {
        let sheet = create_test_spreadsheet();
        let registry = FunctionRegistry::new();
        let evaluator = ExpressionEvaluator::new(&sheet, &registry);
        
        let expr = Expr::Number(42.5);
        match evaluator.evaluate(&expr).unwrap() {
            Value::Number(n) => assert_eq!(n, 42.5),
            _ => panic!("Expected number"),
        }
        
        let expr = Expr::String("Hello".to_string());
        match evaluator.evaluate(&expr).unwrap() {
            Value::String(s) => assert_eq!(s, "Hello"),
            _ => panic!("Expected string"),
        }
    }

    #[test]
    fn test_expression_evaluator_cell_refs() {
        let sheet = create_test_spreadsheet();
        let registry = FunctionRegistry::new();
        let evaluator = ExpressionEvaluator::new(&sheet, &registry);
        
        let expr = Expr::CellRef("A1".to_string());
        match evaluator.evaluate(&expr).unwrap() {
            Value::Number(n) => assert_eq!(n, 10.0),
            _ => panic!("Expected number"),
        }
        
        let expr = Expr::CellRef("B1".to_string());
        match evaluator.evaluate(&expr).unwrap() {
            Value::Number(n) => assert_eq!(n, 20.0),
            _ => panic!("Expected number"),
        }
    }

    #[test]
    fn test_expression_evaluator_string_cells() {
        let mut sheet = Spreadsheet::default();
        sheet.set_cell(0, 0, CellData { value: "Hello".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        sheet.set_cell(0, 1, CellData { value: "World".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        sheet.set_cell(0, 2, CellData { value: "123".to_string(), formula: None, format: None, comment: None, spill_anchor: None }); // Number as string
        
        let registry = FunctionRegistry::new();
        let evaluator = ExpressionEvaluator::new(&sheet, &registry);
        
        // String cell reference
        let expr = Expr::CellRef("A1".to_string());
        match evaluator.evaluate(&expr).unwrap() {
            Value::String(s) => assert_eq!(s, "Hello"),
            _ => panic!("Expected string"),
        }
        
        // Numeric string cell reference  
        let expr = Expr::CellRef("C1".to_string());
        match evaluator.evaluate(&expr).unwrap() {
            Value::Number(n) => assert_eq!(n, 123.0),
            _ => panic!("Expected number"),
        }
    }

    #[test]
    fn test_expression_evaluator_binary_ops() {
        let sheet = create_test_spreadsheet();
        let registry = FunctionRegistry::new();
        let evaluator = ExpressionEvaluator::new(&sheet, &registry);
        
        let expr = Expr::Binary {
            left: Box::new(Expr::Number(10.0)),
            operator: BinaryOp::Add,
            right: Box::new(Expr::Number(5.0)),
        };
        match evaluator.evaluate(&expr).unwrap() {
            Value::Number(n) => assert_eq!(n, 15.0),
            _ => panic!("Expected number"),
        }
        
        let expr = Expr::Binary {
            left: Box::new(Expr::CellRef("A1".to_string())),
            operator: BinaryOp::Multiply,
            right: Box::new(Expr::CellRef("B1".to_string())),
        };
        match evaluator.evaluate(&expr).unwrap() {
            Value::Number(n) => assert_eq!(n, 200.0), // 10 * 20
            _ => panic!("Expected number"),
        }
    }

    #[test]
    fn test_expression_evaluator_unary_ops() {
        let sheet = create_test_spreadsheet();
        let registry = FunctionRegistry::new();
        let evaluator = ExpressionEvaluator::new(&sheet, &registry);
        
        let expr = Expr::Unary {
            operator: UnaryOp::Minus,
            operand: Box::new(Expr::Number(5.0)),
        };
        match evaluator.evaluate(&expr).unwrap() {
            Value::Number(n) => assert_eq!(n, -5.0),
            _ => panic!("Expected number"),
        }
        
        // NOT is now a function, not a unary operator
        let expr = Expr::FunctionCall {
            name: "NOT".to_string(),
            args: vec![Expr::Number(0.0)],
        };
        match evaluator.evaluate(&expr).unwrap() {
            Value::Number(n) => assert_eq!(n, 1.0),
            _ => panic!("Expected number"),
        }
    }

    #[test]
    fn test_expression_evaluator_functions() {
        let sheet = create_test_spreadsheet();
        let registry = FunctionRegistry::new();
        let evaluator = ExpressionEvaluator::new(&sheet, &registry);
        
        let expr = Expr::FunctionCall {
            name: "SUM".to_string(),
            args: vec![
                Expr::CellRef("A1".to_string()),
                Expr::CellRef("B1".to_string()),
            ],
        };
        match evaluator.evaluate(&expr).unwrap() {
            Value::Number(n) => assert_eq!(n, 30.0),
            _ => panic!("Expected number"),
        }
        
        let expr = Expr::FunctionCall {
            name: "IF".to_string(),
            args: vec![
                Expr::Number(1.0),
                Expr::Number(100.0),
                Expr::Number(200.0),
            ],
        };
        match evaluator.evaluate(&expr).unwrap() {
            Value::Number(n) => assert_eq!(n, 100.0),
            _ => panic!("Expected number"),
        }
    }

    #[test]
    fn test_expression_evaluator_string_functions() {
        let sheet = create_test_spreadsheet();
        let registry = FunctionRegistry::new();
        let evaluator = ExpressionEvaluator::new(&sheet, &registry);
        
        // Test CONCAT function
        let expr = Expr::FunctionCall {
            name: "CONCAT".to_string(),
            args: vec![
                Expr::String("Hello".to_string()),
                Expr::String(" ".to_string()),
                Expr::String("World".to_string()),
            ],
        };
        match evaluator.evaluate(&expr).unwrap() {
            Value::String(s) => assert_eq!(s, "Hello World"),
            _ => panic!("Expected string"),
        }
        
        // Test LEN function
        let expr = Expr::FunctionCall {
            name: "LEN".to_string(),
            args: vec![Expr::String("Hello".to_string())],
        };
        match evaluator.evaluate(&expr).unwrap() {
            Value::Number(n) => assert_eq!(n, 5.0),
            _ => panic!("Expected number"),
        }
        
        // Test UPPER function
        let expr = Expr::FunctionCall {
            name: "UPPER".to_string(),
            args: vec![Expr::String("hello".to_string())],
        };
        match evaluator.evaluate(&expr).unwrap() {
            Value::String(s) => assert_eq!(s, "HELLO"),
            _ => panic!("Expected string"),
        }
        
        // Test LEFT function
        let expr = Expr::FunctionCall {
            name: "LEFT".to_string(),
            args: vec![
                Expr::String("Hello World".to_string()),
                Expr::Number(5.0),
            ],
        };
        match evaluator.evaluate(&expr).unwrap() {
            Value::String(s) => assert_eq!(s, "Hello"),
            _ => panic!("Expected string"),
        }
        
        // Test FIND function
        let expr = Expr::FunctionCall {
            name: "FIND".to_string(),
            args: vec![
                Expr::String("lo".to_string()),
                Expr::String("Hello".to_string()),
            ],
        };
        match evaluator.evaluate(&expr).unwrap() {
            Value::Number(n) => assert_eq!(n, 4.0), // 1-based - "lo" starts at position 4
            _ => panic!("Expected number"),
        }
    }

    #[test]
    fn test_expression_evaluator_string_concatenation() {
        let sheet = create_test_spreadsheet();
        let registry = FunctionRegistry::new();
        let evaluator = ExpressionEvaluator::new(&sheet, &registry);
        
        let expr = Expr::Binary {
            left: Box::new(Expr::String("Hello".to_string())),
            operator: BinaryOp::Concatenate,
            right: Box::new(Expr::String(" World".to_string())),
        };
        match evaluator.evaluate(&expr).unwrap() {
            Value::String(s) => assert_eq!(s, "Hello World"),
            _ => panic!("Expected string"),
        }
        
        // Test mixed concatenation
        let expr = Expr::Binary {
            left: Box::new(Expr::String("Number: ".to_string())),
            operator: BinaryOp::Concatenate,
            right: Box::new(Expr::Number(42.0)),
        };
        match evaluator.evaluate(&expr).unwrap() {
            Value::String(s) => assert_eq!(s, "Number: 42"),
            _ => panic!("Expected string"),
        }
    }

    #[test]
    fn test_expression_evaluator_string_equality() {
        let sheet = create_test_spreadsheet();
        let registry = FunctionRegistry::new();
        let evaluator = ExpressionEvaluator::new(&sheet, &registry);
        
        // String equality
        let expr = Expr::Binary {
            left: Box::new(Expr::String("Hello".to_string())),
            operator: BinaryOp::Equal,
            right: Box::new(Expr::String("Hello".to_string())),
        };
        match evaluator.evaluate(&expr).unwrap() {
            Value::Number(n) => assert_eq!(n, 1.0),
            _ => panic!("Expected number"),
        }
        
        // String inequality
        let expr = Expr::Binary {
            left: Box::new(Expr::String("Hello".to_string())),
            operator: BinaryOp::NotEqual,
            right: Box::new(Expr::String("World".to_string())),
        };
        match evaluator.evaluate(&expr).unwrap() {
            Value::Number(n) => assert_eq!(n, 1.0),
            _ => panic!("Expected number"),
        }
    }

    #[test]
    fn test_expression_evaluator_ranges() {
        let sheet = create_test_spreadsheet();
        let registry = FunctionRegistry::new();
        let evaluator = ExpressionEvaluator::new(&sheet, &registry);
        
        let expr = Expr::FunctionCall {
            name: "SUM".to_string(),
            args: vec![Expr::Range("A1".to_string(), "B1".to_string())],
        };
        match evaluator.evaluate(&expr).unwrap() {
            Value::Number(n) => assert_eq!(n, 30.0),
            _ => panic!("Expected number"),
        }
        
        let expr = Expr::FunctionCall {
            name: "AVERAGE".to_string(),
            args: vec![Expr::Range("A1".to_string(), "C1".to_string())],
        };
        match evaluator.evaluate(&expr).unwrap() {
            Value::Number(n) => assert_eq!(n, 20.0),
            _ => panic!("Expected number"),
        }
    }

    #[test]
    fn test_function_registry() {
        let mut registry = FunctionRegistry::new();
        
        // Test that built-in functions are registered
        assert!(registry.get_function("SUM").is_some());
        assert!(registry.get_function("AVERAGE").is_some());
        assert!(registry.get_function("MIN").is_some());
        assert!(registry.get_function("MAX").is_some());
        assert!(registry.get_function("IF").is_some());
        
        // Test case insensitivity
        assert!(registry.get_function("sum").is_some());
        assert!(registry.get_function("Sum").is_some());
        
        // Test unknown function
        assert!(registry.get_function("UNKNOWN").is_none());
        
        // Test registering custom function
        registry.register_function("DOUBLE", |args| {
            if args.len() == 1 {
                Ok(Value::Number(args[0].to_number() * 2.0))
            } else {
                Err("DOUBLE requires exactly 1 argument".to_string())
            }
        });
        
        assert!(registry.get_function("DOUBLE").is_some());
        let double_func = registry.get_function("DOUBLE").unwrap();
        match double_func(&[Value::Number(5.0)]).unwrap() {
            Value::Number(n) => assert_eq!(n, 10.0),
            _ => panic!("Expected number"),
        }
    }

    #[test]
    fn test_complex_expression_parsing_and_evaluation() {
        let sheet = create_test_spreadsheet();
        let registry = FunctionRegistry::new();
        let evaluator = ExpressionEvaluator::new(&sheet, &registry);
        
        // Test complex expression: IF(SUM(A1:B1) > 25, MAX(A1:C1), MIN(A1:C1))
        let mut parser = Parser::new("IF(SUM(A1:B1) > 25, MAX(A1:C1), MIN(A1:C1))").unwrap();
        let ast = parser.parse().unwrap();
        let result = evaluator.evaluate(&ast).unwrap();
        
        // SUM(A1:B1) = 10 + 20 = 30, which is > 25, so we take MAX(A1:C1) = 30
        match result {
            Value::Number(n) => assert_eq!(n, 30.0),
            _ => panic!("Expected number"),
        }
        
        // Test arithmetic with functions: SUM(A1:B1) + 5
        let mut parser = Parser::new("SUM(A1:B1) + 5").unwrap();
        let ast = parser.parse().unwrap();
        let result = evaluator.evaluate(&ast).unwrap();
        match result {
            Value::Number(n) => assert_eq!(n, 35.0),
            _ => panic!("Expected number"),
        }
        
        // Test power operations: 2 ** 3 + 1
        let mut parser = Parser::new("2 ** 3 + 1").unwrap();
        let ast = parser.parse().unwrap();
        let result = evaluator.evaluate(&ast).unwrap();
        match result {
            Value::Number(n) => assert_eq!(n, 9.0),
            _ => panic!("Expected number"),
        }
    }

    #[test]
    fn test_error_handling() {
        let sheet = create_test_spreadsheet();
        let registry = FunctionRegistry::new();
        let evaluator = ExpressionEvaluator::new(&sheet, &registry);
        
        // Division by zero now yields a typed Value::Error(Div0).
        let expr = Expr::Binary {
            left: Box::new(Expr::Number(10.0)),
            operator: BinaryOp::Divide,
            right: Box::new(Expr::Number(0.0)),
        };
        let v = evaluator.evaluate(&expr).unwrap();
        assert_eq!(v, Value::Error(ErrorKind::Div0));
        
        // Test unknown function
        let expr = Expr::FunctionCall {
            name: "UNKNOWN".to_string(),
            args: vec![Expr::Number(5.0)],
        };
        assert!(evaluator.evaluate(&expr).is_err());
        
        // Test invalid cell reference
        let expr = Expr::CellRef("INVALID".to_string());
        assert!(evaluator.evaluate(&expr).is_err());
    }

    #[test]
    fn test_lexer_strings() {
        let mut lexer = Lexer::new("\"Hello World\"");
        assert_eq!(lexer.next_token().unwrap(), Token::String("Hello World".to_string()));
        assert_eq!(lexer.next_token().unwrap(), Token::Eof);
        
        let mut lexer = Lexer::new("\"\"");
        assert_eq!(lexer.next_token().unwrap(), Token::String("".to_string()));
        assert_eq!(lexer.next_token().unwrap(), Token::Eof);
        
        let mut lexer = Lexer::new("\"Quote\"\"Test\"");
        assert_eq!(lexer.next_token().unwrap(), Token::String("Quote\"Test".to_string()));
        assert_eq!(lexer.next_token().unwrap(), Token::Eof);
    }

    #[test]
    fn test_lexer_ampersand() {
        let mut lexer = Lexer::new("&");
        assert_eq!(lexer.next_token().unwrap(), Token::Ampersand);
        assert_eq!(lexer.next_token().unwrap(), Token::Eof);
    }

    #[test]
    fn test_lexer_error_handling() {
        let mut lexer = Lexer::new("@#$");
        assert!(lexer.next_token().is_err());
    }

    #[test]
    fn test_parser_string_literals() {
        let mut parser = Parser::new("\"Hello\"").unwrap();
        let expr = parser.parse().unwrap();
        assert_eq!(expr, Expr::String("Hello".to_string()));
        
        let mut parser = Parser::new("\"\"").unwrap();
        let expr = parser.parse().unwrap();
        assert_eq!(expr, Expr::String("".to_string()));
    }

    #[test]
    fn test_parser_string_concatenation() {
        let mut parser = Parser::new("\"Hello\" & \"World\"").unwrap();
        let expr = parser.parse().unwrap();
        match expr {
            Expr::Binary { left, operator, right } => {
                assert!(matches!(left.as_ref(), &Expr::String(ref s) if s == "Hello"));
                assert_eq!(operator, BinaryOp::Concatenate);
                assert!(matches!(right.as_ref(), &Expr::String(ref s) if s == "World"));
            }
            _ => panic!("Expected binary expression"),
        }
    }

    #[test]
    fn test_parser_mixed_concatenation() {
        let mut parser = Parser::new("\"Number: \" & 42").unwrap();
        let expr = parser.parse().unwrap();
        match expr {
            Expr::Binary { left, operator, right } => {
                assert!(matches!(left.as_ref(), &Expr::String(ref s) if s == "Number: "));
                assert_eq!(operator, BinaryOp::Concatenate);
                assert!(matches!(right.as_ref(), &Expr::Number(42.0)));
            }
            _ => panic!("Expected binary expression"),
        }
    }

    #[test]
    fn test_parser_error_handling() {
        // Test unexpected token
        let result = Parser::new("2 +");
        assert!(result.is_ok()); // Parser creation should succeed
        let mut parser = result.unwrap();
        assert!(parser.parse().is_err()); // But parsing should fail
        
        // Test mismatched parentheses
        let mut parser = Parser::new("(2 + 3").unwrap();
        assert!(parser.parse().is_err());
        
        // Test invalid function call
        let mut parser = Parser::new("SUM(").unwrap();
        assert!(parser.parse().is_err());
        
        // Test unterminated string
        let result = Parser::new("\"Hello");
        assert!(result.is_err());
    }

    #[test]
    fn test_parser_complex_nested_functions() {
        // Test parsing of deeply nested function calls
        let mut parser = Parser::new("UPPER(CONCAT(\"Price: \", TRIM(GET(\"https://example.com\"))))").unwrap();
        let expr = parser.parse().unwrap();
        
        // Verify the AST structure for nested functions
        match expr {
            Expr::FunctionCall { name, args } => {
                assert_eq!(name, "UPPER");
                assert_eq!(args.len(), 1);
                
                // Check that the inner argument is a CONCAT function
                match &args[0] {
                    Expr::FunctionCall { name: inner_name, args: inner_args } => {
                        assert_eq!(inner_name, "CONCAT");
                        assert_eq!(inner_args.len(), 2);
                        
                        // First arg should be string literal "Price: "
                        assert!(matches!(inner_args[0], Expr::String(ref s) if s == "Price: "));
                        
                        // Second arg should be TRIM function
                        match &inner_args[1] {
                            Expr::FunctionCall { name: trim_name, args: trim_args } => {
                                assert_eq!(trim_name, "TRIM");
                                assert_eq!(trim_args.len(), 1);
                                
                                // TRIM's arg should be GET function
                                match &trim_args[0] {
                                    Expr::FunctionCall { name: get_name, args: get_args } => {
                                        assert_eq!(get_name, "GET");
                                        assert_eq!(get_args.len(), 1);
                                        assert!(matches!(get_args[0], Expr::String(ref s) if s == "https://example.com"));
                                    }
                                    _ => panic!("Expected GET function call"),
                                }
                            }
                            _ => panic!("Expected TRIM function call"),
                        }
                    }
                    _ => panic!("Expected CONCAT function call"),
                }
            }
            _ => panic!("Expected function call at top level"),
        }
    }

    #[test]
    fn test_parser_multiple_nested_levels() {
        // Test 5-level nesting: IF(LEN(TRIM(GET(url))) > 5, LEFT(UPPER(GET(url)), 20), "Short")
        let mut parser = Parser::new("IF(LEN(TRIM(GET(\"https://api.com\")))>5, LEFT(UPPER(GET(\"https://api.com\")), 20), \"Short\")").unwrap();
        let expr = parser.parse().unwrap();
        
        // Verify this parses correctly as an IF function
        match expr {
            Expr::FunctionCall { name, args } => {
                assert_eq!(name, "IF");
                assert_eq!(args.len(), 3);
                
                // First argument should be a comparison (LEN(...) > 5)
                match &args[0] {
                    Expr::Binary { left, operator, right } => {
                        assert_eq!(*operator, BinaryOp::Greater);
                        assert!(matches!(right.as_ref(), &Expr::Number(5.0)));
                        
                        // Left side should be LEN(TRIM(GET(...)))
                        match left.as_ref() {
                            Expr::FunctionCall { name: len_name, args: len_args } => {
                                assert_eq!(len_name, "LEN");
                                assert_eq!(len_args.len(), 1);
                                
                                // LEN's arg should be TRIM(GET(...))
                                match &len_args[0] {
                                    Expr::FunctionCall { name: trim_name, args: trim_args } => {
                                        assert_eq!(trim_name, "TRIM");
                                        assert_eq!(trim_args.len(), 1);
                                        
                                        // TRIM's arg should be GET(...)
                                        match &trim_args[0] {
                                            Expr::FunctionCall { name: get_name, .. } => {
                                                assert_eq!(get_name, "GET");
                                            }
                                            _ => panic!("Expected GET function"),
                                        }
                                    }
                                    _ => panic!("Expected TRIM function"),
                                }
                            }
                            _ => panic!("Expected LEN function"),
                        }
                    }
                    _ => panic!("Expected comparison expression"),
                }
                
                // Second argument should be LEFT(UPPER(GET(...)), 20)
                match &args[1] {
                    Expr::FunctionCall { name: left_name, args: left_args } => {
                        assert_eq!(left_name, "LEFT");
                        assert_eq!(left_args.len(), 2);
                        
                        // First arg should be UPPER(GET(...))
                        match &left_args[0] {
                            Expr::FunctionCall { name: upper_name, args: upper_args } => {
                                assert_eq!(upper_name, "UPPER");
                                assert_eq!(upper_args.len(), 1);
                                
                                // UPPER's arg should be GET(...)
                                match &upper_args[0] {
                                    Expr::FunctionCall { name: get_name, .. } => {
                                        assert_eq!(get_name, "GET");
                                    }
                                    _ => panic!("Expected GET function"),
                                }
                            }
                            _ => panic!("Expected UPPER function"),
                        }
                        
                        // Second arg should be number 20
                        assert!(matches!(left_args[1], Expr::Number(20.0)));
                    }
                    _ => panic!("Expected LEFT function"),
                }
                
                // Third argument should be string "Short"
                assert!(matches!(args[2], Expr::String(ref s) if s == "Short"));
            }
            _ => panic!("Expected IF function call"),
        }
    }

    #[test]
    fn test_parser_mixed_operators_and_functions() {
        // Test expression mixing operators and functions: GET(A1) & " - " & UPPER(GET(B1))
        let mut parser = Parser::new("GET(A1) & \" - \" & UPPER(GET(B1))").unwrap();
        let expr = parser.parse().unwrap();
        
        // Should parse as nested concatenation operations
        match expr {
            Expr::Binary { left, operator: op1, right } => {
                assert_eq!(op1, BinaryOp::Concatenate);
                
                // Left side should be another concatenation: GET(A1) & " - "
                match left.as_ref() {
                    Expr::Binary { left: inner_left, operator: op2, right: inner_right } => {
                        assert_eq!(*op2, BinaryOp::Concatenate);
                        
                        // Inner left should be GET(A1)
                        match inner_left.as_ref() {
                            Expr::FunctionCall { name, args } => {
                                assert_eq!(name, "GET");
                                assert_eq!(args.len(), 1);
                                assert!(matches!(args[0], Expr::CellRef(ref s) if s == "A1"));
                            }
                            _ => panic!("Expected GET function call"),
                        }
                        
                        // Inner right should be " - "
                        assert!(matches!(inner_right.as_ref(), &Expr::String(ref s) if s == " - "));
                    }
                    _ => panic!("Expected concatenation expression"),
                }
                
                // Right side should be UPPER(GET(B1))
                match right.as_ref() {
                    Expr::FunctionCall { name: upper_name, args: upper_args } => {
                        assert_eq!(upper_name, "UPPER");
                        assert_eq!(upper_args.len(), 1);
                        
                        match &upper_args[0] {
                            Expr::FunctionCall { name: get_name, args: get_args } => {
                                assert_eq!(get_name, "GET");
                                assert_eq!(get_args.len(), 1);
                                assert!(matches!(get_args[0], Expr::CellRef(ref s) if s == "B1"));
                            }
                            _ => panic!("Expected GET function call"),
                        }
                    }
                    _ => panic!("Expected UPPER function call"),
                }
            }
            _ => panic!("Expected concatenation expression at top level"),
        }
    }

    #[test]
    fn test_sparkline_basic() {
        let mut sheet = Spreadsheet::default();
        sheet.set_cell(0, 0, CellData { value: "1".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        sheet.set_cell(0, 1, CellData { value: "5".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        sheet.set_cell(0, 2, CellData { value: "10".to_string(), formula: None, format: None, comment: None, spill_anchor: None });

        let registry = FunctionRegistry::new();
        let evaluator = ExpressionEvaluator::new(&sheet, &registry);
        let expr = Parser::new("SPARKLINE(A1:C1)").unwrap().parse().unwrap();
        let result = evaluator.evaluate(&expr);
        assert!(result.is_ok());
        match result.unwrap() {
            Value::String(val) => assert_eq!(val.chars().count(), 3),
            _ => panic!("Expected string from SPARKLINE"),
        }
    }

    #[test]
    fn test_sparkline_all_equal() {
        let mut sheet = Spreadsheet::default();
        sheet.set_cell(0, 0, CellData { value: "5".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        sheet.set_cell(0, 1, CellData { value: "5".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        sheet.set_cell(0, 2, CellData { value: "5".to_string(), formula: None, format: None, comment: None, spill_anchor: None });

        let registry = FunctionRegistry::new();
        let evaluator = ExpressionEvaluator::new(&sheet, &registry);
        let expr = Parser::new("SPARKLINE(A1:C1)").unwrap().parse().unwrap();
        let result = evaluator.evaluate(&expr);
        assert!(result.is_ok());
        match result.unwrap() {
            Value::String(val) => {
                assert_eq!(val.chars().count(), 3);
                let chars: Vec<char> = val.chars().collect();
                assert_eq!(chars[0], chars[1]);
                assert_eq!(chars[1], chars[2]);
            }
            _ => panic!("Expected string from SPARKLINE"),
        }
    }

    #[test]
    fn test_sparkline_single_value() {
        let mut sheet = Spreadsheet::default();
        sheet.set_cell(0, 0, CellData { value: "7".to_string(), formula: None, format: None, comment: None, spill_anchor: None });

        let registry = FunctionRegistry::new();
        let evaluator = ExpressionEvaluator::new(&sheet, &registry);
        let expr = Parser::new("SPARKLINE(A1)").unwrap().parse().unwrap();
        let result = evaluator.evaluate(&expr);
        assert!(result.is_ok());
        match result.unwrap() {
            Value::String(val) => assert_eq!(val.chars().count(), 1),
            _ => panic!("Expected string from SPARKLINE"),
        }
    }
}