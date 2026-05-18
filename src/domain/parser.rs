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
pub struct Lexer {
    input: Vec<char>,
    position: usize,
    current_char: Option<char>,
}

impl Lexer {
    /// Creates a new lexer for the given input string.
    pub fn new(input: &str) -> Self {
        let chars: Vec<char> = input.chars().collect();
        let current_char = chars.get(0).copied();
        
        Self {
            input: chars,
            position: 0,
            current_char,
        }
    }
    
    /// Advances to the next character in the input.
    fn advance(&mut self) {
        self.position += 1;
        self.current_char = self.input.get(self.position).copied();
    }
    
    
    /// Skips whitespace characters.
    fn skip_whitespace(&mut self) {
        while let Some(ch) = self.current_char {
            if ch.is_whitespace() {
                self.advance();
            } else {
                break;
            }
        }
    }
    
    /// Reads a number token (integer or decimal).
    fn read_number(&mut self) -> Result<f64, String> {
        let mut number_str = String::new();
        
        // Read integer part
        while let Some(ch) = self.current_char {
            if ch.is_ascii_digit() {
                number_str.push(ch);
                self.advance();
            } else {
                break;
            }
        }
        
        // Read decimal part if present
        if self.current_char == Some('.') {
            number_str.push('.');
            self.advance();
            
            while let Some(ch) = self.current_char {
                if ch.is_ascii_digit() {
                    number_str.push(ch);
                    self.advance();
                } else {
                    break;
                }
            }
        }
        
        number_str.parse::<f64>()
            .map_err(|_| format!("Invalid number: {}", number_str))
    }
    
    /// Reads an identifier or cell reference. Cell refs may include `$`.
    /// Dotted function names (`STDEV.S`) and structured table refs
    /// (`Table1[Col1]`) are folded into a single identifier token.
    fn read_identifier(&mut self) -> String {
        let mut identifier = String::new();
        while let Some(ch) = self.current_char {
            if ch.is_ascii_alphanumeric() || ch == '_' || ch == '$' {
                identifier.push(if ch == '$' { '$' } else { ch.to_ascii_uppercase() });
                self.advance();
            } else if ch == '.' {
                let next = self.input.get(self.position + 1).copied();
                if next.map(|c| c.is_ascii_alphabetic()).unwrap_or(false) {
                    identifier.push('.');
                    self.advance();
                } else {
                    break;
                }
            } else if ch == '[' {
                // Structured table ref: read up to matching ']' inclusive.
                identifier.push('[');
                self.advance();
                while let Some(c) = self.current_char {
                    if c == ']' {
                        identifier.push(']');
                        self.advance();
                        break;
                    } else {
                        identifier.push(c.to_ascii_uppercase());
                        self.advance();
                    }
                }
            } else {
                break;
            }
        }
        identifier
    }
    
    /// Reads a string literal (between quotes).
    fn read_string(&mut self) -> Result<String, String> {
        let mut string_value = String::new();
        
        while let Some(ch) = self.current_char {
            if ch == '"' {
                // Check for escaped quote
                self.advance();
                if self.current_char == Some('"') {
                    string_value.push('"');
                    self.advance();
                } else {
                    // End of string
                    return Ok(string_value);
                }
            } else {
                string_value.push(ch);
                self.advance();
            }
        }
        
        Err("Unterminated string literal".to_string())
    }
    
    /// Determines if an identifier is a cell reference or function name.
    /// Cell references accept optional `$` markers: `$?[A-Z]+\$?[0-9]+`.
    /// Anything containing `.` or `[` is an identifier (function or table ref).
    fn classify_identifier(&self, identifier: &str) -> Token {
        if identifier.contains('.') || identifier.contains('[') {
            return Token::Identifier(identifier.to_string());
        }
        let mut chars = identifier.chars().peekable();
        // Optional leading $
        if chars.peek() == Some(&'$') {
            chars.next();
        }
        let mut has_letters = false;
        while let Some(&c) = chars.peek() {
            if c.is_ascii_alphabetic() {
                has_letters = true;
                chars.next();
            } else {
                break;
            }
        }
        // Optional $ between letters and digits
        if chars.peek() == Some(&'$') {
            chars.next();
        }
        let mut has_digits = false;
        while let Some(&c) = chars.peek() {
            if c.is_ascii_digit() {
                has_digits = true;
                chars.next();
            } else {
                break;
            }
        }
        if has_letters && has_digits && chars.peek().is_none() {
            Token::CellRef(identifier.to_string())
        } else {
            // If it contains `$`, it can't be a function name either; bubble up
            // as identifier and the parser will reject it.
            Token::Identifier(identifier.to_string())
        }
    }
    
    /// Gets the next token from the input.
    pub fn next_token(&mut self) -> Result<Token, String> {
        self.skip_whitespace();
        
        match self.current_char {
            None => Ok(Token::Eof),
            
            Some(ch) => match ch {
                // Numbers
                '0'..='9' => {
                    let number = self.read_number()?;
                    Ok(Token::Number(number))
                }
                
                // Identifiers and cell references (incl. absolute refs starting with `$`)
                'A'..='Z' | 'a'..='z' | '$' => {
                    let identifier = self.read_identifier();
                    Ok(self.classify_identifier(&identifier))
                }
                
                // Operators and delimiters
                '+' => {
                    self.advance();
                    Ok(Token::Plus)
                }
                
                '-' => {
                    self.advance();
                    Ok(Token::Minus)
                }
                
                '*' => {
                    self.advance();
                    if self.current_char == Some('*') {
                        self.advance();
                        Ok(Token::Power)
                    } else {
                        Ok(Token::Multiply)
                    }
                }
                
                '/' => {
                    self.advance();
                    Ok(Token::Divide)
                }
                
                '%' => {
                    self.advance();
                    Ok(Token::Modulo)
                }
                
                '^' => {
                    self.advance();
                    Ok(Token::PowerAlt)
                }
                
                '<' => {
                    self.advance();
                    match self.current_char {
                        Some('=') => {
                            self.advance();
                            Ok(Token::LessEqual)
                        }
                        Some('>') => {
                            self.advance();
                            Ok(Token::NotEqual)
                        }
                        _ => Ok(Token::Less),
                    }
                }
                
                '>' => {
                    self.advance();
                    if self.current_char == Some('=') {
                        self.advance();
                        Ok(Token::GreaterEqual)
                    } else {
                        Ok(Token::Greater)
                    }
                }
                
                '=' => {
                    self.advance();
                    Ok(Token::Equal)
                }
                
                '(' => {
                    self.advance();
                    Ok(Token::LeftParen)
                }
                
                ')' => {
                    self.advance();
                    Ok(Token::RightParen)
                }
                
                ',' => {
                    self.advance();
                    Ok(Token::Comma)
                }
                
                ':' => {
                    self.advance();
                    Ok(Token::Colon)
                }
                
                '&' => {
                    self.advance();
                    Ok(Token::Ampersand)
                }

                '!' => {
                    self.advance();
                    Ok(Token::Bang)
                }

                '{' => {
                    self.advance();
                    Ok(Token::LeftBrace)
                }

                '}' => {
                    self.advance();
                    Ok(Token::RightBrace)
                }

                ';' => {
                    self.advance();
                    Ok(Token::Semicolon)
                }


                '\'' => {
                    // Quoted sheet name: 'Some Sheet' (only used before `!`).
                    self.advance();
                    let mut name = String::new();
                    while let Some(c) = self.current_char {
                        if c == '\'' {
                            self.advance();
                            // Doubled '' inside is an escape for a literal '.
                            if self.current_char == Some('\'') {
                                name.push('\'');
                                self.advance();
                            } else {
                                break;
                            }
                        } else {
                            name.push(c);
                            self.advance();
                        }
                    }
                    Ok(Token::Identifier(format!("'{}'", name)))
                }

                '"' => {
                    self.advance();
                    let string_value = self.read_string()?;
                    Ok(Token::String(string_value))
                }

                _ => Err(format!("Unexpected character: '{}'", ch)),
            }
        }
    }
}

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
fn glob_match(s: &str, pattern: &str) -> bool {
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
fn date_to_serial(year: i32, month: u32, day: u32) -> f64 {
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

fn serial_to_date(serial: f64) -> (i32, u32, u32) {
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
fn parse_iso_date(s: &str) -> Result<(i32, u32, u32), ()> {
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
fn add_commas(int_str: &str) -> String {
    let mut out = String::with_capacity(int_str.len() + int_str.len() / 3);
    for (i, c) in int_str.chars().rev().enumerate() {
        if i > 0 && i % 3 == 0 {
            out.push(',');
        }
        out.push(c);
    }
    out.chars().rev().collect()
}

fn days_in_month(year: i32, month: u32) -> u32 {
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

fn today_serial() -> f64 {
    use std::time::{SystemTime, UNIX_EPOCH};
    let secs = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .map(|d| d.as_secs())
        .unwrap_or(0);
    let days = secs / 86400;
    25569.0 + days as f64
}

fn now_serial() -> f64 {
    use std::time::{SystemTime, UNIX_EPOCH};
    let secs = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .map(|d| d.as_secs_f64())
        .unwrap_or(0.0);
    25569.0 + secs / 86400.0
}

/// Registry for spreadsheet functions.
pub struct FunctionRegistry {
    functions: HashMap<String, FunctionImpl>,
}

impl FunctionRegistry {
    /// Creates a new function registry with built-in functions.
    pub fn new() -> Self {
        let mut registry = Self {
            functions: HashMap::new(),
        };
        
        // Register built-in functions
        registry.register_builtin_functions();
        registry
    }
    
    /// Registers a new function in the registry.
    pub fn register_function(&mut self, name: &str, func: FunctionImpl) {
        self.functions.insert(name.to_uppercase(), func);
    }
    
    /// Gets a function by name.
    pub fn get_function(&self, name: &str) -> Option<&FunctionImpl> {
        self.functions.get(&name.to_uppercase())
    }
    
    /// Registers all built-in spreadsheet functions.
    fn register_builtin_functions(&mut self) {
        // Numeric functions
        self.register_function("SUM", |args| {
            let flat = flatten_args(args);
            // Excel: SUM propagates the first error it encounters.
            if let Some(e) = flat.iter().find_map(|v| v.first_error()) {
                return Ok(Value::Error(e));
            }
            let sum: f64 = flat.iter().map(|v| v.to_number()).sum();
            Ok(Value::Number(sum))
        });

        self.register_function("AVERAGE", |args| {
            let flat = flatten_args(args);
            if flat.is_empty() {
                Err("AVERAGE requires at least one argument".to_string())
            } else {
                if let Some(e) = flat.iter().find_map(|v| v.first_error()) {
                    return Ok(Value::Error(e));
                }
                // Only count cells that parse as numbers (Excel semantics).
                let mut sum = 0.0;
                let mut count = 0usize;
                for v in &flat {
                    match v {
                        Value::Number(n) => { sum += *n; count += 1; }
                        Value::String(s) if !s.is_empty() => {
                            if let Ok(n) = s.parse::<f64>() { sum += n; count += 1; }
                        }
                        Value::Bool(b) => { sum += if *b { 1.0 } else { 0.0 }; count += 1; }
                        _ => {}
                    }
                }
                if count == 0 {
                    Err("AVERAGE: no numeric values".to_string())
                } else {
                    Ok(Value::Number(sum / count as f64))
                }
            }
        });

        self.register_function("MIN", |args| {
            let flat = flatten_args(args);
            if let Some(e) = flat.iter().find_map(|v| v.first_error()) {
                return Ok(Value::Error(e));
            }
            flat.iter().map(|v| v.to_number()).fold(None, |acc: Option<f64>, x| {
                Some(acc.map_or(x, |a| a.min(x)))
            }).map(Value::Number).ok_or_else(|| "MIN requires at least one argument".to_string())
        });

        self.register_function("MAX", |args| {
            let flat = flatten_args(args);
            if let Some(e) = flat.iter().find_map(|v| v.first_error()) {
                return Ok(Value::Error(e));
            }
            flat.iter().map(|v| v.to_number()).fold(None, |acc: Option<f64>, x| {
                Some(acc.map_or(x, |a| a.max(x)))
            }).map(Value::Number).ok_or_else(|| "MAX requires at least one argument".to_string())
        });

        self.register_function("IF", |args| {
            if args.len() != 3 {
                Err("IF requires exactly 3 arguments".to_string())
            } else {
                Ok(if args[0].is_truthy() { args[1].clone() } else { args[2].clone() })
            }
        });

        self.register_function("AND", |args| {
            let flat = flatten_args(args);
            let result = flat.iter().all(|v| v.is_truthy());
            Ok(Value::Bool(result))
        });

        self.register_function("OR", |args| {
            let flat = flatten_args(args);
            let result = flat.iter().any(|v| v.is_truthy());
            Ok(Value::Bool(result))
        });

        self.register_function("NOT", |args| {
            if args.len() != 1 {
                Err("NOT requires exactly 1 argument".to_string())
            } else {
                let result = !args[0].is_truthy();
                Ok(Value::Number(if result { 1.0 } else { 0.0 }))
            }
        });
        
        self.register_function("ABS", |args| {
            if args.len() != 1 {
                Err("ABS requires exactly 1 argument".to_string())
            } else {
                Ok(Value::Number(args[0].to_number().abs()))
            }
        });
        
        self.register_function("SQRT", |args| {
            if args.len() != 1 {
                Err("SQRT requires exactly 1 argument".to_string())
            } else {
                let num = args[0].to_number();
                if num < 0.0 {
                    Ok(Value::Error(ErrorKind::Num))
                } else {
                    Ok(Value::Number(num.sqrt()))
                }
            }
        });
        
        self.register_function("ROUND", |args| {
            match args.len() {
                1 => Ok(Value::Number(args[0].to_number().round())),
                2 => {
                    let num = args[0].to_number();
                    let places = args[1].to_number() as i32;
                    let multiplier = 10f64.powi(places);
                    Ok(Value::Number((num * multiplier).round() / multiplier))
                }
                _ => Err("ROUND requires 1 or 2 arguments".to_string()),
            }
        });
        
        // String functions
        self.register_function("CONCAT", |args| {
            let flat = flatten_args(args);
            let result = flat.iter().map(|v| v.to_string()).collect::<String>();
            Ok(Value::String(result))
        });
        
        self.register_function("LEN", |args| {
            if args.len() != 1 {
                Err("LEN requires exactly 1 argument".to_string())
            } else {
                let len = args[0].to_string().chars().count() as f64;
                Ok(Value::Number(len))
            }
        });
        
        self.register_function("LEFT", |args| {
            if args.len() != 2 {
                Err("LEFT requires exactly 2 arguments".to_string())
            } else {
                let text = args[0].to_string();
                let num_chars = args[1].to_number() as usize;
                let result = text.chars().take(num_chars).collect::<String>();
                Ok(Value::String(result))
            }
        });
        
        self.register_function("RIGHT", |args| {
            if args.len() != 2 {
                Err("RIGHT requires exactly 2 arguments".to_string())
            } else {
                let text = args[0].to_string();
                let num_chars = args[1].to_number() as usize;
                let chars: Vec<char> = text.chars().collect();
                let start = chars.len().saturating_sub(num_chars);
                let result = chars[start..].iter().collect::<String>();
                Ok(Value::String(result))
            }
        });
        
        self.register_function("MID", |args| {
            if args.len() != 3 {
                Err("MID requires exactly 3 arguments".to_string())
            } else {
                let text = args[0].to_string();
                // 1-based start position (Excel convention). start < 1
                // or length < 0 → #VALUE! (Excel behavior).
                let start_one = args[1].to_number() as i64;
                let length_raw = args[2].to_number();
                if start_one < 1 || length_raw < 0.0 {
                    return Ok(Value::Error(ErrorKind::Value));
                }
                let length = length_raw as usize;
                let chars: Vec<char> = text.chars().collect();
                let start = (start_one as usize) - 1;
                let end = (start + length).min(chars.len());
                let result = if start < chars.len() {
                    chars[start..end].iter().collect::<String>()
                } else {
                    String::new()
                };
                Ok(Value::String(result))
            }
        });

        self.register_function("FIND", |args| {
            if args.len() < 2 || args.len() > 3 {
                Err("FIND requires 2 or 3 arguments".to_string())
            } else {
                let search_text = args[0].to_string();
                let within_text = args[1].to_string();
                // 1-based start position (Excel convention).
                let start_one = if args.len() == 3 {
                    args[2].to_number() as i64
                } else {
                    1
                };
                let start_pos = if start_one < 1 { 0 } else { (start_one as usize) - 1 };

                let within_chars: Vec<char> = within_text.chars().collect();
                if start_pos > within_chars.len() {
                    return Err("Start position is beyond text length".to_string());
                }

                let search_in = within_chars[start_pos..].iter().collect::<String>();
                match search_in.find(&search_text) {
                    Some(byte_pos) => {
                        // Convert byte offset back to char offset within search_in.
                        let char_offset = search_in[..byte_pos].chars().count();
                        Ok(Value::Number((start_pos + char_offset + 1) as f64))
                    }
                    None => Err("Search text not found".to_string()),
                }
            }
        });
        
        self.register_function("UPPER", |args| {
            if args.len() != 1 {
                Err("UPPER requires exactly 1 argument".to_string())
            } else {
                Ok(Value::String(args[0].to_string().to_uppercase()))
            }
        });
        
        self.register_function("LOWER", |args| {
            if args.len() != 1 {
                Err("LOWER requires exactly 1 argument".to_string())
            } else {
                Ok(Value::String(args[0].to_string().to_lowercase()))
            }
        });
        
        self.register_function("TRIM", |args| {
            if args.len() != 1 {
                Err("TRIM requires exactly 1 argument".to_string())
            } else {
                Ok(Value::String(args[0].to_string().trim().to_string()))
            }
        });
        
        self.register_function("GET", |args| {
            if args.len() != 1 {
                return Err("GET requires exactly 1 argument (URL)".to_string());
            }
            let url = args[0].to_string();
            if url.is_empty() {
                return Err("GET: empty URL".to_string());
            }
            use crate::infrastructure::fetcher::{fetch, FetchResult};
            match fetch(&url) {
                FetchResult::Value(body) => Ok(Value::String(body)),
                FetchResult::Loading => Ok(Value::String("Loading…".to_string())),
                FetchResult::Error(msg) => Err(msg),
            }
        });

        // --- Math functions ---

        self.register_function("CEILING", |args| {
            if args.len() != 1 {
                Err("CEILING requires exactly 1 argument".to_string())
            } else {
                Ok(Value::Number(args[0].to_number().ceil()))
            }
        });

        self.register_function("FLOOR", |args| {
            if args.len() != 1 {
                Err("FLOOR requires exactly 1 argument".to_string())
            } else {
                Ok(Value::Number(args[0].to_number().floor()))
            }
        });

        self.register_function("INT", |args| {
            if args.len() != 1 {
                Err("INT requires exactly 1 argument".to_string())
            } else {
                // Excel `INT` is floor, not truncate (matters for negatives).
                Ok(Value::Number(args[0].to_number().floor()))
            }
        });

        self.register_function("MOD", |args| {
            if args.len() != 2 {
                Err("MOD requires exactly 2 arguments".to_string())
            } else {
                let dividend = args[0].to_number();
                let divisor = args[1].to_number();
                if divisor == 0.0 {
                    return Ok(Value::Error(ErrorKind::Div0));
                }
                // Excel MOD = dividend - divisor*INT(dividend/divisor),
                // where INT is floor — so the result takes the divisor's sign.
                let result = dividend - divisor * (dividend / divisor).floor();
                Ok(Value::Number(result))
            }
        });

        self.register_function("LOG", |args| {
            match args.len() {
                1 => {
                    let n = args[0].to_number();
                    if n <= 0.0 { Ok(Value::Error(ErrorKind::Num)) }
                    else { Ok(Value::Number(n.log10())) }
                }
                2 => {
                    let n = args[0].to_number();
                    let base = args[1].to_number();
                    if n <= 0.0 || base <= 0.0 || base == 1.0 {
                        Ok(Value::Error(ErrorKind::Num))
                    } else {
                        Ok(Value::Number(n.log(base)))
                    }
                }
                _ => Err("LOG requires 1 or 2 arguments".to_string()),
            }
        });

        self.register_function("LN", |args| {
            if args.len() != 1 {
                Err("LN requires exactly 1 argument".to_string())
            } else {
                let n = args[0].to_number();
                if n <= 0.0 {
                    Ok(Value::Error(ErrorKind::Num))
                } else {
                    Ok(Value::Number(n.ln()))
                }
            }
        });

        self.register_function("EXP", |args| {
            if args.len() != 1 {
                Err("EXP requires exactly 1 argument".to_string())
            } else {
                Ok(Value::Number(args[0].to_number().exp()))
            }
        });

        self.register_function("PI", |args| {
            if !args.is_empty() {
                Err("PI takes no arguments".to_string())
            } else {
                Ok(Value::Number(std::f64::consts::PI))
            }
        });

        self.register_function("RAND", |_args| {
            use std::time::SystemTime;
            let nanos = SystemTime::now()
                .duration_since(SystemTime::UNIX_EPOCH)
                .unwrap_or_default()
                .subsec_nanos();
            Ok(Value::Number(nanos as f64 / 1_000_000_000.0))
        });

        self.register_function("RANDBETWEEN", |args| {
            if args.len() != 2 {
                Err("RANDBETWEEN requires exactly 2 arguments".to_string())
            } else {
                use std::time::SystemTime;
                let low = args[0].to_number();
                let high = args[1].to_number();
                let seed = SystemTime::now()
                    .duration_since(SystemTime::UNIX_EPOCH)
                    .unwrap_or_default()
                    .subsec_nanos();
                let result = (low + (seed as f64 / u32::MAX as f64) * (high - low + 1.0)).floor();
                Ok(Value::Number(result))
            }
        });

        self.register_function("SIGN", |args| {
            if args.len() != 1 {
                Err("SIGN requires exactly 1 argument".to_string())
            } else {
                let n = args[0].to_number();
                let result = if n > 0.0 {
                    1.0
                } else if n < 0.0 {
                    -1.0
                } else {
                    0.0
                };
                Ok(Value::Number(result))
            }
        });

        self.register_function("POWER", |args| {
            if args.len() != 2 {
                Err("POWER requires exactly 2 arguments".to_string())
            } else {
                Ok(Value::Number(args[0].to_number().powf(args[1].to_number())))
            }
        });

        // --- String functions ---

        self.register_function("SUBSTITUTE", |args| {
            if args.len() != 3 {
                Err("SUBSTITUTE requires exactly 3 arguments".to_string())
            } else {
                let text = args[0].to_string();
                let old = args[1].to_string();
                let new = args[2].to_string();
                // Excel: empty `old` is a no-op (otherwise `replace("","x")`
                // inserts between every char).
                if old.is_empty() {
                    Ok(Value::String(text))
                } else {
                    Ok(Value::String(text.replace(&old, &new)))
                }
            }
        });

        self.register_function("REPLACE", |args| {
            if args.len() != 4 {
                Err("REPLACE requires exactly 4 arguments".to_string())
            } else {
                let text = args[0].to_string();
                let start = args[1].to_number() as usize; // 1-based
                let num_chars = args[2].to_number() as usize;
                let new_text = args[3].to_string();
                let chars: Vec<char> = text.chars().collect();
                let start_idx = if start > 0 { start - 1 } else { 0 };
                let end_idx = (start_idx + num_chars).min(chars.len());
                let mut result = chars[..start_idx].iter().collect::<String>();
                result.push_str(&new_text);
                if end_idx < chars.len() {
                    result.extend(chars[end_idx..].iter());
                }
                Ok(Value::String(result))
            }
        });

        self.register_function("REPT", |args| {
            if args.len() != 2 {
                Err("REPT requires exactly 2 arguments".to_string())
            } else {
                let text = args[0].to_string();
                let count_raw = args[1].to_number();
                if count_raw < 0.0 {
                    return Ok(Value::Error(ErrorKind::Value));
                }
                Ok(Value::String(text.repeat(count_raw as usize)))
            }
        });

        self.register_function("EXACT", |args| {
            if args.len() != 2 {
                Err("EXACT requires exactly 2 arguments".to_string())
            } else {
                let a = args[0].to_string();
                let b = args[1].to_string();
                Ok(Value::Number(if a == b { 1.0 } else { 0.0 }))
            }
        });

        self.register_function("PROPER", |args| {
            if args.len() != 1 {
                Err("PROPER requires exactly 1 argument".to_string())
            } else {
                let text = args[0].to_string();
                let mut result = String::new();
                let mut capitalize_next = true;
                for ch in text.chars() {
                    if ch.is_whitespace() || ch == '-' || ch == '_' {
                        result.push(ch);
                        capitalize_next = true;
                    } else if capitalize_next {
                        for upper in ch.to_uppercase() {
                            result.push(upper);
                        }
                        capitalize_next = false;
                    } else {
                        for lower in ch.to_lowercase() {
                            result.push(lower);
                        }
                    }
                }
                Ok(Value::String(result))
            }
        });

        self.register_function("CLEAN", |args| {
            if args.len() != 1 {
                Err("CLEAN requires exactly 1 argument".to_string())
            } else {
                let text = args[0].to_string();
                let cleaned: String = text
                    .chars()
                    .filter(|c| c.is_ascii_graphic() || *c == ' ')
                    .collect();
                Ok(Value::String(cleaned))
            }
        });

        self.register_function("CHAR", |args| {
            if args.len() != 1 {
                Err("CHAR requires exactly 1 argument".to_string())
            } else {
                let n = args[0].to_number() as u32;
                match char::from_u32(n) {
                    Some(c) => Ok(Value::String(String::from(c))),
                    None => Err(format!("CHAR: {} is not a valid character code", n)),
                }
            }
        });

        self.register_function("CODE", |args| {
            if args.len() != 1 {
                Err("CODE requires exactly 1 argument".to_string())
            } else {
                let text = args[0].to_string();
                if let Some(ch) = text.chars().next() {
                    Ok(Value::Number(ch as u32 as f64))
                } else {
                    Err("CODE requires a non-empty string".to_string())
                }
            }
        });

        self.register_function("TEXT", |args| {
            if args.len() < 1 || args.len() > 2 {
                Err("TEXT requires 1 or 2 arguments".to_string())
            } else {
                Ok(Value::String(args[0].to_string()))
            }
        });

        self.register_function("VALUE", |args| {
            if args.len() != 1 {
                Err("VALUE requires exactly 1 argument".to_string())
            } else {
                let text = args[0].to_string();
                match text.parse::<f64>() {
                    Ok(n) => Ok(Value::Number(n)),
                    Err(_) => Err(format!("VALUE: cannot convert '{}' to number", text)),
                }
            }
        });

        self.register_function("NUMBERVALUE", |args| {
            if args.len() != 1 {
                Err("NUMBERVALUE requires exactly 1 argument".to_string())
            } else {
                let text = args[0].to_string();
                match text.parse::<f64>() {
                    Ok(n) => Ok(Value::Number(n)),
                    Err(_) => Err(format!("NUMBERVALUE: cannot convert '{}' to number", text)),
                }
            }
        });

        // --- Info functions ---

        // --- Error trapping & inspection ---
        self.register_function("IFERROR", |args| {
            if args.len() != 2 {
                return Err("IFERROR requires 2 arguments".to_string());
            }
            if args[0].is_error() {
                Ok(args[1].clone())
            } else {
                Ok(args[0].clone())
            }
        });

        self.register_function("IFNA", |args| {
            if args.len() != 2 {
                return Err("IFNA requires 2 arguments".to_string());
            }
            if matches!(args[0].first_error(), Some(ErrorKind::NA)) {
                Ok(args[1].clone())
            } else {
                Ok(args[0].clone())
            }
        });

        self.register_function("ISERROR", |args| {
            if args.len() != 1 {
                return Err("ISERROR requires 1 argument".to_string());
            }
            Ok(Value::Bool(args[0].is_error()))
        });

        self.register_function("ISERR", |args| {
            // ISERR: error EXCEPT #N/A.
            if args.len() != 1 {
                return Err("ISERR requires 1 argument".to_string());
            }
            let result = match args[0].first_error() {
                Some(ErrorKind::NA) | None => false,
                _ => true,
            };
            Ok(Value::Bool(result))
        });

        self.register_function("ISNA", |args| {
            if args.len() != 1 {
                return Err("ISNA requires 1 argument".to_string());
            }
            Ok(Value::Bool(matches!(args[0].first_error(), Some(ErrorKind::NA))))
        });

        self.register_function("NA", |args| {
            if !args.is_empty() {
                return Err("NA takes no arguments".to_string());
            }
            Ok(Value::Error(ErrorKind::NA))
        });

        self.register_function("ERROR.TYPE", |args| {
            if args.len() != 1 {
                return Err("ERROR.TYPE requires 1 argument".to_string());
            }
            // Excel codes: 1=#NULL!, 2=#DIV/0!, 3=#VALUE!, 4=#REF!,
            // 5=#NAME?, 6=#NUM!, 7=#N/A.
            let code = match args[0].first_error() {
                Some(ErrorKind::Null) => 1.0,
                Some(ErrorKind::Div0) => 2.0,
                Some(ErrorKind::Value) => 3.0,
                Some(ErrorKind::Ref) => 4.0,
                Some(ErrorKind::Name) => 5.0,
                Some(ErrorKind::Num) => 6.0,
                Some(ErrorKind::NA) => 7.0,
                Some(ErrorKind::Spill) => 14.0,
                None => return Ok(Value::Error(ErrorKind::NA)),
            };
            Ok(Value::Number(code))
        });

        self.register_function("ISBLANK", |args| {
            if args.len() != 1 {
                Err("ISBLANK requires exactly 1 argument".to_string())
            } else {
                let is_blank = match &args[0] {
                    Value::String(s) => s.is_empty(),
                    Value::List(l) => l.is_empty(),
                    _ => false,
                };
                Ok(Value::Number(if is_blank { 1.0 } else { 0.0 }))
            }
        });

        self.register_function("ISNUMBER", |args| {
            if args.len() != 1 {
                Err("ISNUMBER requires exactly 1 argument".to_string())
            } else {
                let is_num = matches!(&args[0], Value::Number(_));
                Ok(Value::Number(if is_num { 1.0 } else { 0.0 }))
            }
        });

        self.register_function("ISTEXT", |args| {
            if args.len() != 1 {
                Err("ISTEXT requires exactly 1 argument".to_string())
            } else {
                let is_text = matches!(&args[0], Value::String(_));
                Ok(Value::Number(if is_text { 1.0 } else { 0.0 }))
            }
        });

        self.register_function("TYPE", |args| {
            if args.len() != 1 {
                Err("TYPE requires exactly 1 argument".to_string())
            } else {
                let type_num = match &args[0] {
                    Value::Number(_) => 1.0,
                    Value::String(_) => 2.0,
                    Value::Bool(_) => 4.0,
                    Value::Error(_) => 16.0,
                    Value::List(_) | Value::Array { .. } => 64.0,
                };
                Ok(Value::Number(type_num))
            }
        });

        // --- Stats functions ---

        self.register_function("COUNT", |args| {
            let flat = flatten_args(args);
            let count = flat.iter().filter(|v| matches!(v, Value::Number(_))).count();
            Ok(Value::Number(count as f64))
        });

        self.register_function("COUNTA", |args| {
            let flat = flatten_args(args);
            let count = flat.iter().filter(|v| match v {
                Value::Number(_) => true,
                Value::String(s) => !s.is_empty(),
                Value::Bool(_) => true,
                // Errors and nested aggregates are not "values" for COUNTA.
                Value::Error(_) | Value::List(_) | Value::Array { .. } => false,
            }).count();
            Ok(Value::Number(count as f64))
        });

        // --- Visualization functions ---

        self.register_function("SPARKLINE", |args| {
            let flat = flatten_args(args);
            if flat.is_empty() {
                return Err("SPARKLINE requires at least one argument".to_string());
            }
            let values: Vec<f64> = flat.iter().map(|v| v.to_number()).collect();
            let min = values.iter().cloned().fold(f64::INFINITY, f64::min);
            let max = values.iter().cloned().fold(f64::NEG_INFINITY, f64::max);
            let blocks = [' ', '\u{2581}', '\u{2582}', '\u{2583}', '\u{2584}', '\u{2585}', '\u{2586}', '\u{2587}', '\u{2588}'];
            let range = max - min;
            let sparkline: String = values.iter().map(|&v| {
                if range == 0.0 {
                    blocks[4]
                } else {
                    let idx = ((v - min) / range * 8.0).round() as usize;
                    blocks[idx.min(8)]
                }
            }).collect();
            Ok(Value::String(sparkline))
        });

        // --- Lookup & conditional aggregates ---

        // SUMIF(range, criteria) — sums values in `range` matching `criteria`.
        // Criteria can be a number ("5"), a string ("apple"), or a comparison
        // ">5", "<=10", "<>foo", "*wild*".
        self.register_function("SUMIF", |args| {
            if args.len() != 2 && args.len() != 3 {
                return Err("SUMIF requires 2 or 3 arguments".to_string());
            }
            let range = args[0].flatten();
            let criteria = args[1].to_string();
            let sum_range = if args.len() == 3 { args[2].flatten() } else { range.clone() };
            let mut sum = 0.0;
            for (i, v) in range.iter().enumerate() {
                if criteria_matches(v, &criteria) {
                    if let Some(target) = sum_range.get(i) {
                        sum += target.to_number();
                    }
                }
            }
            Ok(Value::Number(sum))
        });

        self.register_function("COUNTIF", |args| {
            if args.len() != 2 {
                return Err("COUNTIF requires 2 arguments".to_string());
            }
            let range = args[0].flatten();
            let criteria = args[1].to_string();
            let count = range.iter().filter(|v| criteria_matches(v, &criteria)).count();
            Ok(Value::Number(count as f64))
        });

        self.register_function("AVERAGEIF", |args| {
            if args.len() != 2 && args.len() != 3 {
                return Err("AVERAGEIF requires 2 or 3 arguments".to_string());
            }
            let range = args[0].flatten();
            let criteria = args[1].to_string();
            let sum_range = if args.len() == 3 { args[2].flatten() } else { range.clone() };
            let mut sum = 0.0;
            let mut count = 0usize;
            for (i, v) in range.iter().enumerate() {
                if criteria_matches(v, &criteria) {
                    if let Some(target) = sum_range.get(i) {
                        sum += target.to_number();
                        count += 1;
                    }
                }
            }
            if count == 0 {
                Err("AVERAGEIF: no matching values".to_string())
            } else {
                Ok(Value::Number(sum / count as f64))
            }
        });

        // VLOOKUP(value, range, col_index, [exact])
        // VLOOKUP(lookup, range, col_index, [exact])
        // Range is a 2-D block; we walk col 1 for the key, return col_index.
        self.register_function("VLOOKUP", |args| {
            if args.len() < 3 || args.len() > 4 {
                return Err("VLOOKUP requires 3 or 4 arguments".to_string());
            }
            let lookup = args[0].to_string();
            let col_index = args[2].to_number() as usize;
            let exact = args.get(3).map(|v| !v.is_truthy()).unwrap_or(false);
            let (rows, cols, data) = shape_of(&args[1]);
            if col_index == 0 || col_index > cols {
                return Err("VLOOKUP: col_index out of range".to_string());
            }
            // Approximate-match mode requires sorted keys (Excel semantics):
            // we stop at the first key strictly greater than the target.
            // If keys are unsorted, the result is undefined per Excel docs.
            let target_num = lookup.parse::<f64>().ok();
            let mut last_match: Option<usize> = None;
            for r in 0..rows {
                let key = &data[r * cols];
                if exact {
                    if key.to_string() == lookup {
                        return Ok(data[r * cols + col_index - 1].clone());
                    }
                } else if let Some(t) = target_num {
                    let k = key.to_number();
                    if k > t {
                        // Sorted-ascending assumption: nothing further can match.
                        break;
                    }
                    last_match = Some(r);
                } else if key.to_string() <= lookup {
                    // Non-numeric approximate: string compare.
                    last_match = Some(r);
                }
            }
            if let Some(r) = last_match {
                return Ok(data[r * cols + col_index - 1].clone());
            }
            Err("VLOOKUP: value not found".to_string())
        });

        // HLOOKUP(lookup, range, row_index, [exact]) — horizontal twin.
        self.register_function("HLOOKUP", |args| {
            if args.len() < 3 || args.len() > 4 {
                return Err("HLOOKUP requires 3 or 4 arguments".to_string());
            }
            let lookup = args[0].to_string();
            let row_index = args[2].to_number() as usize;
            let exact = args.get(3).map(|v| !v.is_truthy()).unwrap_or(false);
            let (rows, cols, data) = shape_of(&args[1]);
            if row_index == 0 || row_index > rows {
                return Err("HLOOKUP: row_index out of range".to_string());
            }
            let mut last_match: Option<usize> = None;
            for c in 0..cols {
                let key = &data[c]; // top row at index 0..cols
                let matches = if exact {
                    key.to_string() == lookup
                } else {
                    key.to_number() <= lookup.parse::<f64>().unwrap_or(0.0)
                };
                if matches {
                    if exact {
                        return Ok(data[(row_index - 1) * cols + c].clone());
                    }
                    last_match = Some(c);
                }
            }
            if let Some(c) = last_match {
                return Ok(data[(row_index - 1) * cols + c].clone());
            }
            Err("HLOOKUP: value not found".to_string())
        });

        // XLOOKUP(lookup, lookup_range, return_range, [if_not_found])
        // Modern Excel: lookup_range and return_range are independent ranges
        // of the same length. No col_index needed.
        // XLOOKUP(lookup, lookup_array, return_array, [if_not_found],
        //         [match_mode], [search_mode])
        //   match_mode:  0 = exact (default), -1 = exact or next-smaller,
        //                1 = exact or next-larger, 2 = wildcard.
        //   search_mode: 1 = first-to-last (default), -1 = last-to-first,
        //                2 = binary asc, -2 = binary desc.
        self.register_function("XLOOKUP", |args| {
            if args.len() < 3 || args.len() > 6 {
                return Err("XLOOKUP requires 3-6 arguments".to_string());
            }
            let lookup = args[0].to_string();
            let keys = args[1].flatten();
            let values = args[2].flatten();
            if keys.len() != values.len() {
                return Err("XLOOKUP: lookup and return ranges must match in length".to_string());
            }
            let match_mode = args
                .get(4)
                .map(|v| v.to_number() as i64)
                .unwrap_or(0);
            let search_mode = args
                .get(5)
                .map(|v| v.to_number() as i64)
                .unwrap_or(1);

            // Build an index order based on search_mode. Binary search modes
            // assume the array is already sorted.
            let mut indices: Vec<usize> = (0..keys.len()).collect();
            if search_mode == -1 {
                indices.reverse();
            }
            // (Binary modes 2/-2: we fall through to linear; the spec lets us
            // exploit sortedness but the linear walk still finds the answer.)

            let needle_num = lookup.parse::<f64>().ok();
            let mut exact_hit: Option<usize> = None;
            let mut next_smaller: Option<usize> = None; // largest <= target
            let mut next_larger: Option<usize> = None;  // smallest >= target

            for i in &indices {
                let k = &keys[*i];
                let matched = match match_mode {
                    2 => glob_match(&k.to_string(), &lookup),
                    _ => k.to_string() == lookup,
                };
                if matched {
                    exact_hit = Some(*i);
                    break;
                }
                if let (Some(t), Some(n)) = (needle_num, Some(k.to_number())) {
                    match match_mode {
                        -1 if n <= t => {
                            if next_smaller
                                .map(|si| keys[si].to_number())
                                .map(|sv| n > sv)
                                .unwrap_or(true)
                            {
                                next_smaller = Some(*i);
                            }
                        }
                        1 if n >= t => {
                            if next_larger
                                .map(|li| keys[li].to_number())
                                .map(|lv| n < lv)
                                .unwrap_or(true)
                            {
                                next_larger = Some(*i);
                            }
                        }
                        _ => {}
                    }
                }
            }
            if let Some(i) = exact_hit {
                return Ok(values[i].clone());
            }
            if match_mode == -1 {
                if let Some(i) = next_smaller {
                    return Ok(values[i].clone());
                }
            }
            if match_mode == 1 {
                if let Some(i) = next_larger {
                    return Ok(values[i].clone());
                }
            }
            args.get(3)
                .cloned()
                .ok_or_else(|| "XLOOKUP: value not found".to_string())
        });

        // INDEX(range, row, [col]) — 1-based row/col into the range.
        self.register_function("INDEX", |args| {
            if args.len() < 2 || args.len() > 3 {
                return Err("INDEX requires 2 or 3 arguments".to_string());
            }
            let row = args[1].to_number() as usize;
            let col = args.get(2).map(|v| v.to_number() as usize).unwrap_or(1);
            if row == 0 || col == 0 {
                return Err("INDEX: row/col are 1-based".to_string());
            }
            let (rows, cols, data) = shape_of(&args[0]);
            if row > rows || col > cols {
                return Err("INDEX: out of range".to_string());
            }
            Ok(data[(row - 1) * cols + (col - 1)].clone())
        });

        // MATCH(value, range, [type]) — returns 1-based position of value in
        // range. type: 1 (exact or largest <=), 0 (exact), -1 (smallest >=).
        // We implement type=0 (exact) and type=1 (approx, default).
        // SUMPRODUCT(arr1, arr2, ...) — multiply arrays element-wise, then sum.
        // All arrays must share shape.
        self.register_function("SUMPRODUCT", |args| {
            if args.is_empty() {
                return Err("SUMPRODUCT requires at least 1 argument".to_string());
            }
            let mut acc: Vec<f64> = args[0]
                .flatten()
                .iter()
                .map(|v| v.to_number())
                .collect();
            for arg in &args[1..] {
                let next: Vec<f64> = arg.flatten().iter().map(|v| v.to_number()).collect();
                if next.len() != acc.len() {
                    return Err("SUMPRODUCT: array shape mismatch".to_string());
                }
                for (i, n) in next.iter().enumerate() {
                    acc[i] *= n;
                }
            }
            Ok(Value::Number(acc.iter().sum()))
        });

        // TRANSPOSE — swap rows and cols of a 2-D range.
        self.register_function("TRANSPOSE", |args| {
            if args.len() != 1 {
                return Err("TRANSPOSE requires 1 argument".to_string());
            }
            let (rows, cols, data) = shape_of(&args[0]);
            let mut out = Vec::with_capacity(rows * cols);
            for c in 0..cols {
                for r in 0..rows {
                    out.push(data[r * cols + c].clone());
                }
            }
            Ok(Value::Array { rows: cols, cols: rows, data: out })
        });

        // SEQUENCE(rows, [cols], [start], [step]) — generate a numeric sequence.
        self.register_function("SEQUENCE", |args| {
            if args.is_empty() || args.len() > 4 {
                return Err("SEQUENCE requires 1-4 arguments".to_string());
            }
            let rows = args[0].to_number() as usize;
            let cols = args.get(1).map(|v| v.to_number() as usize).unwrap_or(1);
            let start = args.get(2).map(|v| v.to_number()).unwrap_or(1.0);
            let step = args.get(3).map(|v| v.to_number()).unwrap_or(1.0);
            let mut data = Vec::with_capacity(rows * cols);
            for i in 0..(rows * cols) {
                data.push(Value::Number(start + step * i as f64));
            }
            Ok(Value::Array { rows, cols, data })
        });

        // FILTER(range, predicate_array) — keep rows where the predicate is truthy.
        // Predicate must be a 1-D mask matching the range's row count.
        self.register_function("FILTER", |args| {
            if args.len() < 2 || args.len() > 3 {
                return Err("FILTER requires 2 or 3 arguments".to_string());
            }
            let (rows, cols, data) = shape_of(&args[0]);
            let mask = args[1].flatten();
            if mask.len() != rows && mask.len() != rows * cols {
                return Err("FILTER: predicate length must match range rows".to_string());
            }
            let mut out_rows: Vec<Value> = Vec::new();
            let mut kept = 0;
            for r in 0..rows {
                let keep = mask.get(r).map(|v| v.is_truthy()).unwrap_or(false);
                if keep {
                    for c in 0..cols {
                        out_rows.push(data[r * cols + c].clone());
                    }
                    kept += 1;
                }
            }
            if kept == 0 {
                if let Some(fallback) = args.get(2) {
                    return Ok(fallback.clone());
                }
                return Err("FILTER: no matches".to_string());
            }
            Ok(Value::Array {
                rows: kept,
                cols,
                data: out_rows,
            })
        });

        // SORT(range, [sort_index], [order])
        // order: 1 (ascending, default), -1 (descending). sort_index is 1-based.
        self.register_function("SORT", |args| {
            if args.is_empty() || args.len() > 3 {
                return Err("SORT requires 1-3 arguments".to_string());
            }
            let (rows, cols, data) = shape_of(&args[0]);
            let sort_col = args.get(1).map(|v| v.to_number() as usize).unwrap_or(1);
            let order = args.get(2).map(|v| v.to_number() as i64).unwrap_or(1);
            if sort_col == 0 || sort_col > cols {
                return Err("SORT: sort_index out of range".to_string());
            }
            let mut row_indices: Vec<usize> = (0..rows).collect();
            row_indices.sort_by(|a, b| {
                let av = &data[*a * cols + sort_col - 1];
                let bv = &data[*b * cols + sort_col - 1];
                let cmp = match (av.to_number(), bv.to_number()) {
                    // Try numeric first, fall back to string.
                    (an, bn) if !an.is_nan() && !bn.is_nan() && (av.to_string().parse::<f64>().is_ok() || bv.to_string().parse::<f64>().is_ok()) => {
                        an.partial_cmp(&bn).unwrap_or(std::cmp::Ordering::Equal)
                    }
                    _ => av.to_string().cmp(&bv.to_string()),
                };
                if order < 0 { cmp.reverse() } else { cmp }
            });
            let mut out = Vec::with_capacity(rows * cols);
            for r in &row_indices {
                for c in 0..cols {
                    out.push(data[r * cols + c].clone());
                }
            }
            Ok(Value::Array { rows, cols, data: out })
        });

        // UNIQUE(range) — drop duplicate rows (string-equality on the full row).
        self.register_function("UNIQUE", |args| {
            if args.len() != 1 {
                return Err("UNIQUE requires 1 argument".to_string());
            }
            let (rows, cols, data) = shape_of(&args[0]);
            let mut seen: std::collections::HashSet<String> = std::collections::HashSet::new();
            let mut out = Vec::new();
            let mut kept = 0;
            for r in 0..rows {
                let row_key: String = (0..cols)
                    .map(|c| data[r * cols + c].to_string())
                    .collect::<Vec<_>>()
                    .join("\x1f");
                if seen.insert(row_key) {
                    for c in 0..cols {
                        out.push(data[r * cols + c].clone());
                    }
                    kept += 1;
                }
            }
            Ok(Value::Array { rows: kept, cols, data: out })
        });

        self.register_function("MATCH", |args| {
            if args.len() < 2 || args.len() > 3 {
                return Err("MATCH requires 2 or 3 arguments".to_string());
            }
            let needle = args[0].to_string();
            let range = args[1].flatten();
            let match_type = args.get(2).map(|v| v.to_number() as i64).unwrap_or(1);
            if match_type == 0 {
                // Excel MATCH exact mode supports `*` / `?` wildcards.
                let has_wild = needle.contains('*') || needle.contains('?');
                for (i, v) in range.iter().enumerate() {
                    let s = v.to_string();
                    let matched = if has_wild {
                        glob_match(&s, &needle)
                    } else {
                        s == needle
                    };
                    if matched {
                        return Ok(Value::Number((i + 1) as f64));
                    }
                }
                Err("MATCH: not found".to_string())
            } else {
                let target = needle.parse::<f64>().unwrap_or(0.0);
                let mut last_idx: Option<usize> = None;
                for (i, v) in range.iter().enumerate() {
                    let n = v.to_number();
                    if (match_type > 0 && n <= target) || (match_type < 0 && n >= target) {
                        last_idx = Some(i);
                    }
                }
                last_idx
                    .map(|i| Value::Number((i + 1) as f64))
                    .ok_or_else(|| "MATCH: not found".to_string())
            }
        });

        // --- Booleans ---
        self.register_function("TRUE", |args| {
            if !args.is_empty() {
                return Err("TRUE takes no arguments".to_string());
            }
            Ok(Value::Bool(true))
        });
        self.register_function("FALSE", |args| {
            if !args.is_empty() {
                return Err("FALSE takes no arguments".to_string());
            }
            Ok(Value::Bool(false))
        });

        // --- Date/time ---
        // Stored as Excel-style serial days since 1899-12-30 epoch.
        // TODAY() returns days; NOW() returns days + fractional time-of-day.
        self.register_function("TODAY", |args| {
            if !args.is_empty() {
                return Err("TODAY takes no arguments".to_string());
            }
            Ok(Value::Number(today_serial()))
        });
        self.register_function("NOW", |args| {
            if !args.is_empty() {
                return Err("NOW takes no arguments".to_string());
            }
            Ok(Value::Number(now_serial()))
        });
        self.register_function("DATE", |args| {
            if args.len() != 3 {
                return Err("DATE requires 3 arguments (year, month, day)".to_string());
            }
            let y = args[0].to_number() as i32;
            let m = args[1].to_number() as u32;
            let d = args[2].to_number() as u32;
            Ok(Value::Number(date_to_serial(y, m, d)))
        });
        self.register_function("YEAR", |args| {
            if args.len() != 1 {
                return Err("YEAR requires 1 argument".to_string());
            }
            let (y, _, _) = serial_to_date(args[0].to_number());
            Ok(Value::Number(y as f64))
        });
        self.register_function("MONTH", |args| {
            if args.len() != 1 {
                return Err("MONTH requires 1 argument".to_string());
            }
            let (_, m, _) = serial_to_date(args[0].to_number());
            Ok(Value::Number(m as f64))
        });
        self.register_function("DAY", |args| {
            if args.len() != 1 {
                return Err("DAY requires 1 argument".to_string());
            }
            let (_, _, d) = serial_to_date(args[0].to_number());
            Ok(Value::Number(d as f64))
        });

        // TIME(h, m, s) — fractional day.
        self.register_function("TIME", |args| {
            if args.len() != 3 {
                return Err("TIME requires 3 arguments".to_string());
            }
            let h = args[0].to_number();
            let m = args[1].to_number();
            let s = args[2].to_number();
            Ok(Value::Number((h * 3600.0 + m * 60.0 + s) / 86400.0))
        });
        self.register_function("HOUR", |args| {
            let frac = args[0].to_number().fract();
            let secs = (frac * 86400.0).round() as i64;
            Ok(Value::Number(((secs / 3600) % 24) as f64))
        });
        self.register_function("MINUTE", |args| {
            let frac = args[0].to_number().fract();
            let secs = (frac * 86400.0).round() as i64;
            Ok(Value::Number(((secs / 60) % 60) as f64))
        });
        self.register_function("SECOND", |args| {
            let frac = args[0].to_number().fract();
            let secs = (frac * 86400.0).round() as i64;
            Ok(Value::Number((secs % 60) as f64))
        });

        // DATEDIF(start, end, unit) — units: "D", "M", "Y", "MD", "YM", "YD".
        self.register_function("DATEDIF", |args| {
            if args.len() != 3 {
                return Err("DATEDIF requires 3 arguments".to_string());
            }
            let start = args[0].to_number().floor();
            let end = args[1].to_number().floor();
            if end < start {
                return Ok(Value::Error(ErrorKind::Num));
            }
            let unit = args[2].to_string().to_uppercase();
            let (sy, sm, sd) = serial_to_date(start);
            let (ey, em, ed) = serial_to_date(end);
            let years = ey - sy - if (em, ed) < (sm, sd) { 1 } else { 0 };
            let months = {
                let mut m = (ey - sy) * 12 + (em as i32 - sm as i32);
                if (ed as i32) < (sd as i32) {
                    m -= 1;
                }
                m
            };
            match unit.as_str() {
                "D" => Ok(Value::Number(end - start)),
                "M" => Ok(Value::Number(months as f64)),
                "Y" => Ok(Value::Number(years as f64)),
                "MD" => {
                    // Days component, ignoring months and years. If end-day
                    // ≥ start-day, it's a simple subtraction. Otherwise we
                    // borrow from the previous calendar month, using its
                    // actual day count (not a flat 30).
                    if ed >= sd {
                        Ok(Value::Number((ed - sd) as f64))
                    } else {
                        let prev_month = if em == 1 { 12 } else { em - 1 };
                        let prev_year = if em == 1 { ey - 1 } else { ey };
                        let borrowed = days_in_month(prev_year, prev_month);
                        Ok(Value::Number(
                            (borrowed as i64 - sd as i64 + ed as i64) as f64,
                        ))
                    }
                }
                "YM" => Ok(Value::Number((months % 12 + 12) as f64 % 12.0)),
                "YD" => {
                    // Days as if the start were in the same year as `end`.
                    // If end already passes through start's (month, day) in
                    // its year, use that. Otherwise use the previous year.
                    let candidate_year = if (em, ed) >= (sm, sd) { ey } else { ey - 1 };
                    let s2 = date_to_serial(candidate_year, sm, sd);
                    Ok(Value::Number(end - s2))
                }
                _ => Err(format!("DATEDIF: unknown unit '{}'", unit)),
            }
        });

        // WEEKDAY(serial, [type]). type 1 (default): Sun=1..Sat=7.
        // 2: Mon=1..Sun=7. 3: Mon=0..Sun=6.
        self.register_function("WEEKDAY", |args| {
            if args.is_empty() {
                return Err("WEEKDAY requires 1 or 2 arguments".to_string());
            }
            let serial = args[0].to_number().floor() as i64;
            let ty = args.get(1).map(|v| v.to_number() as i64).unwrap_or(1);
            // 1899-12-30 (serial 0) was a Saturday → (0 + 6) % 7 = 6 (Sat in 0-based Mon=0).
            let mon_based = ((serial + 5).rem_euclid(7)) as i64; // Mon=0..Sun=6
            let v = match ty {
                1 => ((mon_based + 1) % 7) + 1, // Sun=1..Sat=7
                2 => mon_based + 1,             // Mon=1..Sun=7
                3 => mon_based,                 // Mon=0..Sun=6
                _ => return Err(format!("WEEKDAY: bad type {}", ty)),
            };
            Ok(Value::Number(v as f64))
        });

        // EDATE(start, months) — date shifted by `months`. Clamps day if needed.
        self.register_function("EDATE", |args| {
            if args.len() != 2 {
                return Err("EDATE requires 2 arguments".to_string());
            }
            let (y, m, d) = serial_to_date(args[0].to_number());
            let total = (y as i64) * 12 + (m as i64 - 1) + args[1].to_number() as i64;
            let new_y = total.div_euclid(12) as i32;
            let new_m = (total.rem_euclid(12) + 1) as u32;
            let last = days_in_month(new_y, new_m);
            let new_d = d.min(last);
            Ok(Value::Number(date_to_serial(new_y, new_m, new_d)))
        });

        // EOMONTH(start, months) — last day of EDATE result month.
        self.register_function("EOMONTH", |args| {
            if args.len() != 2 {
                return Err("EOMONTH requires 2 arguments".to_string());
            }
            let (y, m, _) = serial_to_date(args[0].to_number());
            let total = (y as i64) * 12 + (m as i64 - 1) + args[1].to_number() as i64;
            let new_y = total.div_euclid(12) as i32;
            let new_m = (total.rem_euclid(12) + 1) as u32;
            let last = days_in_month(new_y, new_m);
            Ok(Value::Number(date_to_serial(new_y, new_m, last)))
        });

        // DAYS(end, start) — simple days between dates.
        self.register_function("DAYS", |args| {
            if args.len() != 2 {
                return Err("DAYS requires 2 arguments".to_string());
            }
            Ok(Value::Number(args[0].to_number().floor() - args[1].to_number().floor()))
        });

        // --- Modern logic operators ---

        self.register_function("IFS", |args| {
            if args.len() < 2 || args.len() % 2 != 0 {
                return Err("IFS requires pairs (cond, value), at least one pair".to_string());
            }
            let mut i = 0;
            while i < args.len() {
                if args[i].is_truthy() {
                    return Ok(args[i + 1].clone());
                }
                i += 2;
            }
            Ok(Value::Error(ErrorKind::NA))
        });

        self.register_function("SWITCH", |args| {
            if args.len() < 3 {
                return Err("SWITCH requires expr + at least one match pair".to_string());
            }
            let expr_s = args[0].to_string();
            let mut i = 1;
            // Pairs (case, result); optional trailing default has odd count.
            while i + 1 < args.len() {
                if args[i].to_string() == expr_s {
                    return Ok(args[i + 1].clone());
                }
                i += 2;
            }
            if i < args.len() {
                Ok(args[i].clone())
            } else {
                Ok(Value::Error(ErrorKind::NA))
            }
        });

        self.register_function("XOR", |args| {
            let flat = flatten_args(args);
            let count_true = flat.iter().filter(|v| v.is_truthy()).count();
            Ok(Value::Bool(count_true % 2 == 1))
        });

        // --- Statistical functions ---
        // Each takes a single flattened range of numbers (Excel-compatible
        // behavior: STDEV.S uses N-1, STDEV.P uses N).
        self.register_function("MEDIAN", |args| {
            let mut nums: Vec<f64> = flatten_args(args)
                .iter()
                .filter_map(|v| match v {
                    Value::Number(n) => Some(*n),
                    Value::String(s) => s.parse::<f64>().ok(),
                    _ => None,
                })
                .filter(|n| n.is_finite())
                .collect();
            if nums.is_empty() {
                return Err("MEDIAN: no numeric values".to_string());
            }
            nums.sort_by(|a, b| a.partial_cmp(b).unwrap_or(std::cmp::Ordering::Equal));
            let n = nums.len();
            let m = if n % 2 == 0 {
                (nums[n / 2 - 1] + nums[n / 2]) / 2.0
            } else {
                nums[n / 2]
            };
            Ok(Value::Number(m))
        });

        fn collect_numbers(args: &[Value]) -> Vec<f64> {
            flatten_args(args)
                .iter()
                .filter_map(|v| match v {
                    Value::Number(n) => Some(*n),
                    Value::String(s) => s.parse::<f64>().ok(),
                    _ => None,
                })
                .filter(|n| n.is_finite())
                .collect()
        }

        self.register_function("STDEV.S", |args| {
            let nums = collect_numbers(args);
            if nums.len() < 2 {
                return Err("STDEV.S requires at least 2 values".to_string());
            }
            let mean: f64 = nums.iter().sum::<f64>() / nums.len() as f64;
            let var: f64 = nums.iter().map(|n| (n - mean).powi(2)).sum::<f64>()
                / (nums.len() - 1) as f64;
            Ok(Value::Number(var.sqrt()))
        });
        self.register_function("STDEV", |args| {
            // Excel legacy alias for STDEV.S
            let nums = collect_numbers(args);
            if nums.len() < 2 {
                return Err("STDEV requires at least 2 values".to_string());
            }
            let mean: f64 = nums.iter().sum::<f64>() / nums.len() as f64;
            let var: f64 = nums.iter().map(|n| (n - mean).powi(2)).sum::<f64>()
                / (nums.len() - 1) as f64;
            Ok(Value::Number(var.sqrt()))
        });
        self.register_function("STDEV.P", |args| {
            let nums = collect_numbers(args);
            if nums.is_empty() {
                return Err("STDEV.P: empty".to_string());
            }
            let mean: f64 = nums.iter().sum::<f64>() / nums.len() as f64;
            let var: f64 = nums.iter().map(|n| (n - mean).powi(2)).sum::<f64>()
                / nums.len() as f64;
            Ok(Value::Number(var.sqrt()))
        });
        self.register_function("VAR.S", |args| {
            let nums = collect_numbers(args);
            if nums.len() < 2 {
                return Err("VAR.S requires at least 2 values".to_string());
            }
            let mean: f64 = nums.iter().sum::<f64>() / nums.len() as f64;
            let var: f64 = nums.iter().map(|n| (n - mean).powi(2)).sum::<f64>()
                / (nums.len() - 1) as f64;
            Ok(Value::Number(var))
        });
        self.register_function("VAR.P", |args| {
            let nums = collect_numbers(args);
            if nums.is_empty() {
                return Err("VAR.P: empty".to_string());
            }
            let mean: f64 = nums.iter().sum::<f64>() / nums.len() as f64;
            let var: f64 = nums.iter().map(|n| (n - mean).powi(2)).sum::<f64>()
                / nums.len() as f64;
            Ok(Value::Number(var))
        });
        self.register_function("LARGE", |args| {
            if args.len() != 2 {
                return Err("LARGE requires 2 arguments".to_string());
            }
            let mut nums = collect_numbers(&args[..1]);
            let k = args[1].to_number() as usize;
            if k == 0 || k > nums.len() {
                return Ok(Value::Error(ErrorKind::Num));
            }
            nums.sort_by(|a, b| b.partial_cmp(a).unwrap_or(std::cmp::Ordering::Equal));
            Ok(Value::Number(nums[k - 1]))
        });
        self.register_function("SMALL", |args| {
            if args.len() != 2 {
                return Err("SMALL requires 2 arguments".to_string());
            }
            let mut nums = collect_numbers(&args[..1]);
            let k = args[1].to_number() as usize;
            if k == 0 || k > nums.len() {
                return Ok(Value::Error(ErrorKind::Num));
            }
            nums.sort_by(|a, b| a.partial_cmp(b).unwrap_or(std::cmp::Ordering::Equal));
            Ok(Value::Number(nums[k - 1]))
        });

        // RANK.EQ(value, range, [order]) — ties get the same rank.
        self.register_function("RANK.EQ", |args| {
            if args.len() < 2 || args.len() > 3 {
                return Err("RANK.EQ requires 2 or 3 arguments".to_string());
            }
            let v = args[0].to_number();
            let nums = collect_numbers(&args[1..2]);
            let ascending = args.get(2).map(|x| x.to_number() != 0.0).unwrap_or(false);
            let rank = if ascending {
                nums.iter().filter(|n| **n < v).count() + 1
            } else {
                nums.iter().filter(|n| **n > v).count() + 1
            };
            Ok(Value::Number(rank as f64))
        });

        // PERCENTILE.INC(range, p) — linear-interp percentile, 0 ≤ p ≤ 1.
        self.register_function("PERCENTILE.INC", |args| {
            if args.len() != 2 {
                return Err("PERCENTILE.INC requires 2 arguments".to_string());
            }
            let mut nums = collect_numbers(&args[..1]);
            let p = args[1].to_number();
            if nums.is_empty() || !(0.0..=1.0).contains(&p) {
                return Ok(Value::Error(ErrorKind::Num));
            }
            nums.sort_by(|a, b| a.partial_cmp(b).unwrap_or(std::cmp::Ordering::Equal));
            let n = nums.len();
            let h = p * (n - 1) as f64;
            let lo = h.floor() as usize;
            let hi = h.ceil() as usize;
            let frac = h - lo as f64;
            let result = nums[lo] + frac * (nums[hi] - nums[lo]);
            Ok(Value::Number(result))
        });

        // CORREL(arr1, arr2) — Pearson correlation.
        self.register_function("CORREL", |args| {
            if args.len() != 2 {
                return Err("CORREL requires 2 arguments".to_string());
            }
            let x = collect_numbers(&args[..1]);
            let y = collect_numbers(&args[1..2]);
            if x.len() != y.len() || x.len() < 2 {
                return Err("CORREL: arrays must match length, min 2".to_string());
            }
            let mx: f64 = x.iter().sum::<f64>() / x.len() as f64;
            let my: f64 = y.iter().sum::<f64>() / y.len() as f64;
            let num: f64 = x.iter().zip(y.iter())
                .map(|(a, b)| (a - mx) * (b - my))
                .sum();
            let den_x: f64 = x.iter().map(|a| (a - mx).powi(2)).sum::<f64>().sqrt();
            let den_y: f64 = y.iter().map(|b| (b - my).powi(2)).sum::<f64>().sqrt();
            if den_x == 0.0 || den_y == 0.0 {
                return Ok(Value::Error(ErrorKind::Div0));
            }
            Ok(Value::Number(num / (den_x * den_y)))
        });

        // --- Financial functions ---
        // PMT(rate, nper, pv, [fv], [type]) — periodic payment for a loan.
        self.register_function("PMT", |args| {
            if args.len() < 3 || args.len() > 5 {
                return Err("PMT requires 3-5 arguments".to_string());
            }
            let rate = args[0].to_number();
            let nper = args[1].to_number();
            let pv = args[2].to_number();
            let fv = args.get(3).map(|v| v.to_number()).unwrap_or(0.0);
            let type_ = args.get(4).map(|v| v.to_number()).unwrap_or(0.0);
            let pmt = if rate == 0.0 {
                -(pv + fv) / nper
            } else {
                let pow = (1.0 + rate).powf(nper);
                -(pv * pow + fv) * rate / ((1.0 + rate * type_) * (pow - 1.0))
            };
            Ok(Value::Number(pmt))
        });

        // FV(rate, nper, pmt, [pv], [type]) — future value.
        self.register_function("FV", |args| {
            if args.len() < 3 || args.len() > 5 {
                return Err("FV requires 3-5 arguments".to_string());
            }
            let rate = args[0].to_number();
            let nper = args[1].to_number();
            let pmt = args[2].to_number();
            let pv = args.get(3).map(|v| v.to_number()).unwrap_or(0.0);
            let type_ = args.get(4).map(|v| v.to_number()).unwrap_or(0.0);
            let fv = if rate == 0.0 {
                -(pv + pmt * nper)
            } else {
                let pow = (1.0 + rate).powf(nper);
                -(pv * pow + pmt * (1.0 + rate * type_) * (pow - 1.0) / rate)
            };
            Ok(Value::Number(fv))
        });

        // PV(rate, nper, pmt, [fv], [type]) — present value.
        self.register_function("PV", |args| {
            if args.len() < 3 || args.len() > 5 {
                return Err("PV requires 3-5 arguments".to_string());
            }
            let rate = args[0].to_number();
            let nper = args[1].to_number();
            let pmt = args[2].to_number();
            let fv = args.get(3).map(|v| v.to_number()).unwrap_or(0.0);
            let type_ = args.get(4).map(|v| v.to_number()).unwrap_or(0.0);
            let pv = if rate == 0.0 {
                -(fv + pmt * nper)
            } else {
                let pow = (1.0 + rate).powf(nper);
                -(fv + pmt * (1.0 + rate * type_) * (pow - 1.0) / rate) / pow
            };
            Ok(Value::Number(pv))
        });

        // NPV(rate, val1, val2, ...) — net present value of a cashflow series
        // starting at period 1.
        self.register_function("NPV", |args| {
            if args.len() < 2 {
                return Err("NPV requires rate + at least one value".to_string());
            }
            let rate = args[0].to_number();
            let flat = flatten_args(&args[1..]);
            let mut acc = 0.0;
            for (i, v) in flat.iter().enumerate() {
                acc += v.to_number() / (1.0 + rate).powi(i as i32 + 1);
            }
            Ok(Value::Number(acc))
        });

        // --- More math ---
        self.register_function("TRUNC", |args| {
            if args.is_empty() || args.len() > 2 {
                return Err("TRUNC requires 1 or 2 arguments".to_string());
            }
            let n = args[0].to_number();
            let digits = args.get(1).map(|v| v.to_number() as i32).unwrap_or(0);
            let scale = 10f64.powi(digits);
            Ok(Value::Number((n * scale).trunc() / scale))
        });
        self.register_function("ATAN2", |args| {
            if args.len() != 2 {
                return Err("ATAN2 requires 2 arguments".to_string());
            }
            // Excel signature: ATAN2(x, y) (yes, x first).
            Ok(Value::Number(args[1].to_number().atan2(args[0].to_number())))
        });
        self.register_function("ATAN", |args| Ok(Value::Number(args[0].to_number().atan())));
        self.register_function("ASIN", |args| Ok(Value::Number(args[0].to_number().asin())));
        self.register_function("ACOS", |args| Ok(Value::Number(args[0].to_number().acos())));
        self.register_function("SINH", |args| Ok(Value::Number(args[0].to_number().sinh())));
        self.register_function("COSH", |args| Ok(Value::Number(args[0].to_number().cosh())));
        self.register_function("TANH", |args| Ok(Value::Number(args[0].to_number().tanh())));
        self.register_function("SIN", |args| Ok(Value::Number(args[0].to_number().sin())));
        self.register_function("COS", |args| Ok(Value::Number(args[0].to_number().cos())));
        self.register_function("TAN", |args| Ok(Value::Number(args[0].to_number().tan())));
        self.register_function("DEGREES", |args| {
            Ok(Value::Number(args[0].to_number().to_degrees()))
        });
        self.register_function("RADIANS", |args| {
            Ok(Value::Number(args[0].to_number().to_radians()))
        });
        self.register_function("FACT", |args| {
            let n = args[0].to_number() as u64;
            let mut r: f64 = 1.0;
            for i in 2..=n {
                r *= i as f64;
            }
            Ok(Value::Number(r))
        });
        self.register_function("COMBIN", |args| {
            if args.len() != 2 {
                return Err("COMBIN requires 2 arguments".to_string());
            }
            let n = args[0].to_number() as i64;
            let k = args[1].to_number() as i64;
            if k < 0 || n < 0 || k > n {
                return Ok(Value::Error(ErrorKind::Num));
            }
            let k = k.min(n - k);
            let mut r: f64 = 1.0;
            for i in 0..k {
                r *= (n - i) as f64;
                r /= (i + 1) as f64;
            }
            Ok(Value::Number(r))
        });
        self.register_function("GCD", |args| {
            let nums: Vec<i64> = flatten_args(args)
                .iter()
                .map(|v| v.to_number() as i64)
                .collect();
            fn gcd(a: i64, b: i64) -> i64 {
                if b == 0 { a.abs() } else { gcd(b, a % b) }
            }
            let g = nums.iter().fold(0, |a, &b| gcd(a, b));
            Ok(Value::Number(g as f64))
        });
        self.register_function("LCM", |args| {
            let nums: Vec<i64> = flatten_args(args)
                .iter()
                .map(|v| v.to_number() as i64)
                .collect();
            fn gcd(a: i64, b: i64) -> i64 {
                if b == 0 { a.abs() } else { gcd(b, a % b) }
            }
            let l = nums.iter().fold(1, |a, &b| if b == 0 { 0 } else { (a * b).abs() / gcd(a, b) });
            Ok(Value::Number(l as f64))
        });
        self.register_function("ROUNDUP", |args| {
            if args.len() != 2 {
                return Err("ROUNDUP requires 2 arguments".to_string());
            }
            let n = args[0].to_number();
            let digits = args[1].to_number() as i32;
            let scale = 10f64.powi(digits);
            let v = (n * scale).abs().ceil() * n.signum() / scale;
            Ok(Value::Number(v))
        });
        self.register_function("ROUNDDOWN", |args| {
            if args.len() != 2 {
                return Err("ROUNDDOWN requires 2 arguments".to_string());
            }
            let n = args[0].to_number();
            let digits = args[1].to_number() as i32;
            let scale = 10f64.powi(digits);
            let v = (n * scale).abs().floor() * n.signum() / scale;
            Ok(Value::Number(v))
        });
        self.register_function("MROUND", |args| {
            if args.len() != 2 {
                return Err("MROUND requires 2 arguments".to_string());
            }
            let n = args[0].to_number();
            let m = args[1].to_number();
            if m == 0.0 {
                return Ok(Value::Number(0.0));
            }
            Ok(Value::Number((n / m).round() * m))
        });
        self.register_function("EVEN", |args| {
            let n = args[0].to_number();
            let v = if n >= 0.0 {
                (n / 2.0).ceil() * 2.0
            } else {
                (n / 2.0).floor() * 2.0
            };
            Ok(Value::Number(v))
        });
        self.register_function("ODD", |args| {
            let n = args[0].to_number();
            let v = if n >= 0.0 {
                let c = ((n + 1.0) / 2.0).ceil() * 2.0 - 1.0;
                if c == -1.0 { 1.0 } else { c }
            } else {
                ((n - 1.0) / 2.0).floor() * 2.0 + 1.0
            };
            Ok(Value::Number(v))
        });

        // --- More text functions ---

        // TEXTJOIN(delim, ignore_empty, ...) — concatenate with delimiter.
        self.register_function("TEXTJOIN", |args| {
            if args.len() < 3 {
                return Err("TEXTJOIN requires at least 3 arguments".to_string());
            }
            let delim = args[0].to_string();
            let ignore_empty = args[1].is_truthy();
            let parts: Vec<String> = flatten_args(&args[2..])
                .iter()
                .filter_map(|v| {
                    let s = v.to_string();
                    if ignore_empty && s.is_empty() { None } else { Some(s) }
                })
                .collect();
            Ok(Value::String(parts.join(&delim)))
        });

        // SEARCH(needle, hay, [start]) — case-insensitive 1-based position.
        self.register_function("SEARCH", |args| {
            if args.len() < 2 || args.len() > 3 {
                return Err("SEARCH requires 2 or 3 arguments".to_string());
            }
            let needle = args[0].to_string().to_lowercase();
            let hay = args[1].to_string();
            let hay_lc = hay.to_lowercase();
            let start = args.get(2).map(|v| v.to_number() as usize).unwrap_or(1).max(1) - 1;
            let chars: Vec<char> = hay_lc.chars().collect();
            if start > chars.len() {
                return Ok(Value::Error(ErrorKind::Value));
            }
            let tail: String = chars[start..].iter().collect();
            match tail.find(&needle) {
                Some(byte_pos) => {
                    let char_offset = tail[..byte_pos].chars().count();
                    Ok(Value::Number((start + char_offset + 1) as f64))
                }
                None => Ok(Value::Error(ErrorKind::Value)),
            }
        });

        // TEXTBEFORE(text, delim, [instance]) / TEXTAFTER (modern Excel).
        self.register_function("TEXTBEFORE", |args| {
            if args.len() < 2 || args.len() > 3 {
                return Err("TEXTBEFORE requires 2 or 3 arguments".to_string());
            }
            let text = args[0].to_string();
            let delim = args[1].to_string();
            let n = args.get(2).map(|v| v.to_number() as usize).unwrap_or(1);
            let mut start = 0usize;
            for _ in 0..n - 1 {
                if let Some(idx) = text[start..].find(&delim) {
                    start += idx + delim.len();
                } else {
                    return Ok(Value::Error(ErrorKind::NA));
                }
            }
            if let Some(idx) = text[start..].find(&delim) {
                Ok(Value::String(text[..start + idx].to_string()))
            } else {
                Ok(Value::Error(ErrorKind::NA))
            }
        });
        self.register_function("TEXTAFTER", |args| {
            if args.len() < 2 || args.len() > 3 {
                return Err("TEXTAFTER requires 2 or 3 arguments".to_string());
            }
            let text = args[0].to_string();
            let delim = args[1].to_string();
            let n = args.get(2).map(|v| v.to_number() as usize).unwrap_or(1);
            let mut start = 0usize;
            for _ in 0..n {
                if let Some(idx) = text[start..].find(&delim) {
                    start += idx + delim.len();
                } else {
                    return Ok(Value::Error(ErrorKind::NA));
                }
            }
            Ok(Value::String(text[start..].to_string()))
        });

        // REGEXMATCH(text, pattern) / REGEXEXTRACT / REGEXREPLACE.
        self.register_function("REGEXMATCH", |args| {
            if args.len() != 2 {
                return Err("REGEXMATCH requires 2 arguments".to_string());
            }
            let re = regex::Regex::new(&args[1].to_string())
                .map_err(|e| format!("REGEXMATCH: bad pattern: {}", e))?;
            Ok(Value::Bool(re.is_match(&args[0].to_string())))
        });
        self.register_function("REGEXEXTRACT", |args| {
            if args.len() != 2 {
                return Err("REGEXEXTRACT requires 2 arguments".to_string());
            }
            let re = regex::Regex::new(&args[1].to_string())
                .map_err(|e| format!("REGEXEXTRACT: bad pattern: {}", e))?;
            let text = args[0].to_string();
            if let Some(caps) = re.captures(&text) {
                // Prefer first capture group; otherwise the whole match.
                let s = caps.get(1).or_else(|| caps.get(0)).map(|m| m.as_str()).unwrap_or("");
                Ok(Value::String(s.to_string()))
            } else {
                Ok(Value::Error(ErrorKind::NA))
            }
        });
        self.register_function("REGEXREPLACE", |args| {
            if args.len() != 3 {
                return Err("REGEXREPLACE requires 3 arguments".to_string());
            }
            let re = regex::Regex::new(&args[1].to_string())
                .map_err(|e| format!("REGEXREPLACE: bad pattern: {}", e))?;
            let replacement = args[2].to_string();
            Ok(Value::String(re.replace_all(&args[0].to_string(), replacement.as_str()).into_owned()))
        });

        // --- Workday math ---
        // NETWORKDAYS(start, end, [holidays]) — business days between dates,
        // inclusive, Mon-Fri.
        self.register_function("NETWORKDAYS", |args| {
            if args.len() < 2 || args.len() > 3 {
                return Err("NETWORKDAYS requires 2 or 3 arguments".to_string());
            }
            let start = args[0].to_number().floor() as i64;
            let end = args[1].to_number().floor() as i64;
            let (lo, hi, sign) = if start <= end { (start, end, 1) } else { (end, start, -1) };
            let holidays: std::collections::HashSet<i64> = args
                .get(2)
                .map(|v| {
                    v.flatten()
                        .iter()
                        .map(|x| x.to_number().floor() as i64)
                        .collect()
                })
                .unwrap_or_default();
            let mut count = 0i64;
            for d in lo..=hi {
                // 1899-12-30 (serial 0) was Saturday → Mon-based day = (d+5) mod 7
                let dow = (d + 5).rem_euclid(7);
                if dow < 5 && !holidays.contains(&d) {
                    count += 1;
                }
            }
            Ok(Value::Number((count * sign) as f64))
        });

        // WORKDAY(start, days, [holidays]) — add business days to a start date.
        self.register_function("WORKDAY", |args| {
            if args.len() < 2 || args.len() > 3 {
                return Err("WORKDAY requires 2 or 3 arguments".to_string());
            }
            let start = args[0].to_number().floor() as i64;
            let days = args[1].to_number() as i64;
            let holidays: std::collections::HashSet<i64> = args
                .get(2)
                .map(|v| {
                    v.flatten()
                        .iter()
                        .map(|x| x.to_number().floor() as i64)
                        .collect()
                })
                .unwrap_or_default();
            let step: i64 = if days >= 0 { 1 } else { -1 };
            let mut remaining = days.abs();
            let mut current = start;
            while remaining > 0 {
                current += step;
                let dow = (current + 5).rem_euclid(7);
                if dow < 5 && !holidays.contains(&current) {
                    remaining -= 1;
                }
            }
            Ok(Value::Number(current as f64))
        });

        // DATEVALUE — parse a date string into a serial. Accepts ISO `YYYY-MM-DD`
        // and Excel-ish `M/D/YYYY` for now.
        self.register_function("DATEVALUE", |args| {
            if args.len() != 1 {
                return Err("DATEVALUE requires 1 argument".to_string());
            }
            let s = args[0].to_string();
            // ISO
            if let Ok(parts) = parse_iso_date(&s) {
                return Ok(Value::Number(date_to_serial(parts.0, parts.1, parts.2)));
            }
            // M/D/YYYY (US-ish)
            let parts: Vec<&str> = s.split('/').collect();
            if parts.len() == 3 {
                if let (Ok(m), Ok(d), Ok(y)) = (
                    parts[0].parse::<u32>(),
                    parts[1].parse::<u32>(),
                    parts[2].parse::<i32>(),
                ) {
                    let year = if y < 100 { 2000 + y } else { y };
                    return Ok(Value::Number(date_to_serial(year, m, d)));
                }
            }
            Ok(Value::Error(ErrorKind::Value))
        });

        // TIMEVALUE — parse `HH:MM[:SS]` to a fraction of a day.
        self.register_function("TIMEVALUE", |args| {
            if args.len() != 1 {
                return Err("TIMEVALUE requires 1 argument".to_string());
            }
            let s = args[0].to_string();
            let parts: Vec<&str> = s.split(':').collect();
            if parts.len() < 2 || parts.len() > 3 {
                return Ok(Value::Error(ErrorKind::Value));
            }
            let h: f64 = parts[0].parse().unwrap_or(-1.0);
            let m: f64 = parts[1].parse().unwrap_or(-1.0);
            let sec: f64 = parts.get(2).map(|x| x.parse().unwrap_or(0.0)).unwrap_or(0.0);
            if h < 0.0 || m < 0.0 || sec < 0.0 {
                return Ok(Value::Error(ErrorKind::Value));
            }
            Ok(Value::Number((h * 3600.0 + m * 60.0 + sec) / 86400.0))
        });

        // --- Text helpers ---
        self.register_function("UNICHAR", |args| {
            let n = args[0].to_number() as u32;
            match char::from_u32(n) {
                Some(c) => Ok(Value::String(c.to_string())),
                None => Ok(Value::Error(ErrorKind::Value)),
            }
        });
        self.register_function("UNICODE", |args| {
            let s = args[0].to_string();
            match s.chars().next() {
                Some(c) => Ok(Value::Number(c as u32 as f64)),
                None => Ok(Value::Error(ErrorKind::Value)),
            }
        });
        self.register_function("DOLLAR", |args| {
            if args.is_empty() || args.len() > 2 {
                return Err("DOLLAR requires 1 or 2 arguments".to_string());
            }
            let n = args[0].to_number();
            let decimals = args.get(1).map(|v| v.to_number() as i32).unwrap_or(2);
            let scale = 10f64.powi(decimals);
            let rounded = (n * scale).round() / scale;
            let sign = if rounded < 0.0 { "-" } else { "" };
            let abs = rounded.abs();
            let mut s = format!("${:.*}", decimals.max(0) as usize, abs);
            // Add thousands separators.
            if let Some(dot) = s.find('.') {
                let (int_part, dec_part) = s.split_at(dot);
                let int_with_seps = add_commas(int_part.trim_start_matches('$'));
                s = format!("${}{}", int_with_seps, dec_part);
            } else {
                let int_with_seps = add_commas(&s[1..]);
                s = format!("${}", int_with_seps);
            }
            Ok(Value::String(format!("{}{}", sign, s)))
        });
        self.register_function("FIXED", |args| {
            if args.is_empty() || args.len() > 3 {
                return Err("FIXED requires 1-3 arguments".to_string());
            }
            let n = args[0].to_number();
            let decimals = args.get(1).map(|v| v.to_number() as i32).unwrap_or(2);
            let no_commas = args.get(2).map(|v| v.is_truthy()).unwrap_or(false);
            let mut s = format!("{:.*}", decimals.max(0) as usize, n);
            if !no_commas {
                if let Some(dot) = s.find('.') {
                    let (int_part, dec_part) = s.split_at(dot);
                    let sign = int_part.starts_with('-');
                    let int_clean = int_part.trim_start_matches('-');
                    let with_commas = add_commas(int_clean);
                    s = format!("{}{}{}", if sign { "-" } else { "" }, with_commas, dec_part);
                } else {
                    let sign = s.starts_with('-');
                    let int_clean = s.trim_start_matches('-').to_string();
                    let with_commas = add_commas(&int_clean);
                    s = format!("{}{}", if sign { "-" } else { "" }, with_commas);
                }
            }
            Ok(Value::String(s))
        });

        // ARRAYTOTEXT(range, [format]) — string representation of array.
        // format: 0 (concise, default) joins by ", "; 1 (strict) wraps in {}.
        self.register_function("ARRAYTOTEXT", |args| {
            if args.is_empty() || args.len() > 2 {
                return Err("ARRAYTOTEXT requires 1 or 2 arguments".to_string());
            }
            let strict = args.get(1).map(|v| v.to_number() as i32 == 1).unwrap_or(false);
            let (rows, cols, data) = shape_of(&args[0]);
            let mut rows_str = Vec::new();
            for r in 0..rows {
                let mut cells = Vec::new();
                for c in 0..cols {
                    let v = &data[r * cols + c];
                    if strict {
                        match v {
                            Value::String(s) => cells.push(format!("\"{}\"", s)),
                            _ => cells.push(v.to_string()),
                        }
                    } else {
                        cells.push(v.to_string());
                    }
                }
                rows_str.push(cells.join(if strict { "," } else { ", " }));
            }
            let joined = rows_str.join(if strict { ";" } else { ", " });
            Ok(Value::String(if strict { format!("{{{}}}", joined) } else { joined }))
        });

        // FREQUENCY(data, bins) — count of `data` values ≤ each bin.
        self.register_function("FREQUENCY", |args| {
            if args.len() != 2 {
                return Err("FREQUENCY requires 2 arguments".to_string());
            }
            let data: Vec<f64> = args[0].flatten().iter().map(|v| v.to_number()).collect();
            let bins: Vec<f64> = args[1].flatten().iter().map(|v| v.to_number()).collect();
            let mut counts: Vec<usize> = vec![0; bins.len() + 1];
            for v in &data {
                let mut placed = false;
                for (i, &b) in bins.iter().enumerate() {
                    if *v <= b {
                        counts[i] += 1;
                        placed = true;
                        break;
                    }
                }
                if !placed {
                    counts[bins.len()] += 1;
                }
            }
            let result: Vec<Value> = counts.iter().map(|&c| Value::Number(c as f64)).collect();
            Ok(Value::Array {
                rows: result.len(),
                cols: 1,
                data: result,
            })
        });

        // --- LAMBDA helpers ---
        // MAP / REDUCE / BYROW / BYCOL / SCAN need the evaluator's lambda
        // invocation machinery, which we can't reach from a plain FunctionImpl
        // (no eval context). They're handled as special forms in the
        // FunctionCall dispatch path of `evaluate`.

        // YEARFRAC(start, end, [basis]) — fractional year between dates.
        // basis: 0 (default, US 30/360), 1 (actual/actual), 2 (actual/360),
        // 3 (actual/365), 4 (European 30/360). We implement the common ones.
        self.register_function("YEARFRAC", |args| {
            if args.len() < 2 || args.len() > 3 {
                return Err("YEARFRAC requires 2 or 3 arguments".to_string());
            }
            let start = args[0].to_number().floor() as i64;
            let end = args[1].to_number().floor() as i64;
            let basis = args.get(2).map(|v| v.to_number() as i32).unwrap_or(0);
            let days = (end - start).abs();
            let denom = match basis {
                1 => 365.25,
                2 => 360.0,
                3 => 365.0,
                _ => 360.0, // 30/360-ish; we approximate
            };
            Ok(Value::Number(days as f64 / denom))
        });
    }
}

impl Default for FunctionRegistry {
    fn default() -> Self {
        Self::new()
    }
}

/// Recursive descent parser for spreadsheet expressions.
pub struct Parser {
    lexer: Lexer,
    current_token: Token,
}

impl Parser {
    /// Creates a new parser for the given expression.
    pub fn new(input: &str) -> Result<Self, String> {
        let mut lexer = Lexer::new(input);
        let current_token = lexer.next_token()?;
        
        Ok(Self {
            lexer,
            current_token,
        })
    }
    
    
    /// Advances to the next token.
    fn advance(&mut self) -> Result<(), String> {
        self.current_token = self.lexer.next_token()?;
        Ok(())
    }
    
    /// Checks if the current token matches the expected token and advances.
    fn expect(&mut self, expected: Token) -> Result<(), String> {
        if std::mem::discriminant(&self.current_token) == std::mem::discriminant(&expected) {
            self.advance()
        } else {
            Err(format!("Expected {:?}, found {:?}", expected, self.current_token))
        }
    }
    
    /// Parses the top-level expression.
    pub fn parse(&mut self) -> Result<Expr, String> {
        let expr = self.parse_equality()?;
        
        if self.current_token != Token::Eof {
            return Err(format!("Unexpected token at end: {:?}", self.current_token));
        }
        
        Ok(expr)
    }
    
    
    /// Parses equality expressions.
    fn parse_equality(&mut self) -> Result<Expr, String> {
        let mut left = self.parse_comparison()?;
        
        while matches!(self.current_token, Token::Equal | Token::NotEqual) {
            let op = match self.current_token {
                Token::Equal => BinaryOp::Equal,
                Token::NotEqual => BinaryOp::NotEqual,
                _ => unreachable!(),
            };
            self.advance()?;
            let right = self.parse_comparison()?;
            left = Expr::Binary {
                left: Box::new(left),
                operator: op,
                right: Box::new(right),
            };
        }
        
        Ok(left)
    }
    
    /// Parses comparison expressions.
    fn parse_comparison(&mut self) -> Result<Expr, String> {
        let mut left = self.parse_addition()?;
        
        while matches!(self.current_token, Token::Less | Token::LessEqual | Token::Greater | Token::GreaterEqual) {
            let op = match self.current_token {
                Token::Less => BinaryOp::Less,
                Token::LessEqual => BinaryOp::LessEqual,
                Token::Greater => BinaryOp::Greater,
                Token::GreaterEqual => BinaryOp::GreaterEqual,
                _ => unreachable!(),
            };
            self.advance()?;
            let right = self.parse_addition()?;
            left = Expr::Binary {
                left: Box::new(left),
                operator: op,
                right: Box::new(right),
            };
        }
        
        Ok(left)
    }
    
    /// Parses addition and subtraction expressions.
    fn parse_addition(&mut self) -> Result<Expr, String> {
        let mut left = self.parse_concatenation()?;
        
        while matches!(self.current_token, Token::Plus | Token::Minus) {
            let op = match self.current_token {
                Token::Plus => BinaryOp::Add,
                Token::Minus => BinaryOp::Subtract,
                _ => unreachable!(),
            };
            self.advance()?;
            let right = self.parse_concatenation()?;
            left = Expr::Binary {
                left: Box::new(left),
                operator: op,
                right: Box::new(right),
            };
        }
        
        Ok(left)
    }
    
    /// Parses string concatenation expressions.
    fn parse_concatenation(&mut self) -> Result<Expr, String> {
        let mut left = self.parse_multiplication()?;
        
        while matches!(self.current_token, Token::Ampersand) {
            self.advance()?;
            let right = self.parse_multiplication()?;
            left = Expr::Binary {
                left: Box::new(left),
                operator: BinaryOp::Concatenate,
                right: Box::new(right),
            };
        }
        
        Ok(left)
    }
    
    /// Parses multiplication, division, and modulo expressions.
    fn parse_multiplication(&mut self) -> Result<Expr, String> {
        let mut left = self.parse_power()?;
        
        while matches!(self.current_token, Token::Multiply | Token::Divide | Token::Modulo) {
            let op = match self.current_token {
                Token::Multiply => BinaryOp::Multiply,
                Token::Divide => BinaryOp::Divide,
                Token::Modulo => BinaryOp::Modulo,
                _ => unreachable!(),
            };
            self.advance()?;
            let right = self.parse_power()?;
            left = Expr::Binary {
                left: Box::new(left),
                operator: op,
                right: Box::new(right),
            };
        }
        
        Ok(left)
    }
    
    /// Parses power expressions (right-associative).
    fn parse_power(&mut self) -> Result<Expr, String> {
        let left = self.parse_unary()?;
        
        if matches!(self.current_token, Token::Power | Token::PowerAlt) {
            self.advance()?;
            let right = self.parse_power()?; // Right-associative
            Ok(Expr::Binary {
                left: Box::new(left),
                operator: BinaryOp::Power,
                right: Box::new(right),
            })
        } else {
            Ok(left)
        }
    }
    
    /// Parses unary expressions.
    fn parse_unary(&mut self) -> Result<Expr, String> {
        match self.current_token {
            Token::Plus => {
                self.advance()?;
                let operand = self.parse_unary()?;
                Ok(Expr::Unary {
                    operator: UnaryOp::Plus,
                    operand: Box::new(operand),
                })
            }
            Token::Minus => {
                self.advance()?;
                let operand = self.parse_unary()?;
                Ok(Expr::Unary {
                    operator: UnaryOp::Minus,
                    operand: Box::new(operand),
                })
            }
            _ => self.parse_primary(),
        }
    }
    
    /// Builds an `Expr::Let` from a parsed argument list.
    /// LET(name1, value1, name2, value2, ..., body) — at least 3 args; odd count.
    fn build_let(mut args: Vec<Expr>) -> Result<Expr, String> {
        if args.len() < 3 || args.len() % 2 == 0 {
            return Err("LET requires an odd number of arguments ≥ 3".to_string());
        }
        let body = args.pop().unwrap();
        let mut bindings = Vec::new();
        let mut iter = args.into_iter();
        while let Some(name_expr) = iter.next() {
            let name = match name_expr {
                Expr::NamedRef(n) => n,
                other => {
                    return Err(format!(
                        "LET: expected a name (got {:?})",
                        other
                    ))
                }
            };
            let value = iter.next().unwrap();
            bindings.push((name, Box::new(value)));
        }
        Ok(Expr::Let {
            bindings,
            body: Box::new(body),
        })
    }

    /// Builds an `Expr::Lambda` from a parsed argument list.
    /// LAMBDA(param1, param2, ..., body) — last arg is the body; preceding
    /// args must be bare NamedRefs (parameter names).
    fn build_lambda(mut args: Vec<Expr>) -> Result<Expr, String> {
        if args.is_empty() {
            return Err("LAMBDA requires at least a body".to_string());
        }
        let body = args.pop().unwrap();
        let mut params = Vec::new();
        for a in args {
            match a {
                Expr::NamedRef(n) => params.push(n),
                other => {
                    return Err(format!(
                        "LAMBDA: parameter must be a bare name (got {:?})",
                        other
                    ))
                }
            }
        }
        Ok(Expr::Lambda {
            params,
            body: Box::new(body),
        })
    }

    /// Called after we've consumed `<sheet_name> !`. Reads the cell or range
    /// that follows and produces a sheet-qualified Expr. Also supports 3-D
    /// refs of the form `Sheet1:Sheet3!A1`. The `sheet` token is the prefix
    /// before the `!` and may be a bare name or a quoted `'Some Sheet'`.
    fn continue_sheet_qualified_ref(&mut self, sheet: &str) -> Result<Expr, String> {
        // 3-D range over sheets: the sheet itself may be a range `S1:S3`.
        // In that case `sheet` is `"S1"`, and we have already consumed S1.
        // The colon would have come from a *previous* parse position — but
        // our caller already consumed it as part of the prefix? Actually we
        // detect 3-D form before the bang: if the user typed `Sheet1:Sheet3!A1`,
        // the lexer produced CellRef("Sheet1"), Colon, CellRef("Sheet3"),
        // Bang, CellRef("A1"). That won't reach this helper because the
        // colon-after-cell branch above would have folded it into a Range.
        // We catch that case before emitting Range by peeking for Bang.
        // (See cellref branch.) The simple case here is sheet ! ref.
        match &self.current_token {
            Token::CellRef(cell) => {
                let cell = cell.clone();
                self.advance()?;
                if self.current_token == Token::Colon {
                    self.advance()?;
                    if let Token::CellRef(end_cell) = &self.current_token {
                        let end_cell = end_cell.clone();
                        self.advance()?;
                        Ok(Expr::Range(
                            format!("{}!{}", sheet, cell),
                            format!("{}!{}", sheet, end_cell),
                        ))
                    } else {
                        Err("Expected cell reference after ':'".to_string())
                    }
                } else {
                    Ok(Expr::CellRef(format!("{}!{}", sheet, cell)))
                }
            }
            other => Err(format!(
                "Expected cell reference after '!', found {:?}",
                other
            )),
        }
    }

    /// Parses primary expressions (highest precedence).
    fn parse_primary(&mut self) -> Result<Expr, String> {
        match &self.current_token {
            Token::Number(value) => {
                let value = *value;
                self.advance()?;
                Ok(Expr::Number(value))
            }
            
            Token::String(text) => {
                let text = text.clone();
                self.advance()?;
                Ok(Expr::String(text))
            }
            
            Token::CellRef(cell) => {
                let cell = cell.clone();
                self.advance()?;
                // CellRef ! CellRef is a sheet-qualified ref where the
                // identifier-shape sheet name happened to be lex-classified
                // as a CellRef (e.g. `Sheet2`). Fold it into a prefixed ref.
                if self.current_token == Token::Bang {
                    self.advance()?;
                    return self.continue_sheet_qualified_ref(&cell);
                }
                // Range or 3-D range across sheets.
                if self.current_token == Token::Colon {
                    self.advance()?;
                    if let Token::CellRef(end_cell) = &self.current_token {
                        let end_cell = end_cell.clone();
                        self.advance()?;
                        // 3-D form: Sheet1:Sheet3!A1 — `cell` and `end_cell`
                        // are actually sheet names. Followed by `!` + a real
                        // cell ref. We emit a `Range` whose endpoints carry
                        // a special `__sheets__:S1..S3` marker for the
                        // evaluator to detect and span sheets.
                        if self.current_token == Token::Bang {
                            self.advance()?;
                            if let Token::CellRef(real_cell) = &self.current_token {
                                let real_cell = real_cell.clone();
                                self.advance()?;
                                // Encode as a Range whose endpoints share a
                                // `<sheet1>..<sheet2>!<cell>` form. Both
                                // endpoints carry the sheet range so the
                                // evaluator can split them out.
                                let marker = format!("{}..{}!{}", cell, end_cell, real_cell);
                                return Ok(Expr::Range(marker.clone(), marker));
                            }
                            return Err("Expected cell reference after '!' in 3-D ref".to_string());
                        }
                        Ok(Expr::Range(cell, end_cell))
                    } else {
                        Err("Expected cell reference after ':'".to_string())
                    }
                } else {
                    Ok(Expr::CellRef(cell))
                }
            }

            Token::Identifier(name) => {
                let name = name.clone();
                self.advance()?;
                // Sheet-qualified ref: Identifier ! ...
                if self.current_token == Token::Bang {
                    self.advance()?;
                    return self.continue_sheet_qualified_ref(&name);
                }
                // Quoted 3-D ref: 'Sheet1':'Sheet3'!A1 or 'Sheet1':Sheet3!A1
                if self.current_token == Token::Colon {
                    let lookahead = self.current_token.clone();
                    let _ = lookahead;
                    // Peek ahead to see if we have a sheet name + Bang.
                    self.advance()?; // consume ':'
                    let second_name: Option<String> = match &self.current_token {
                        Token::Identifier(n) => Some(n.clone()),
                        Token::CellRef(n) => Some(n.clone()),
                        _ => None,
                    };
                    if let Some(s2) = second_name {
                        self.advance()?;
                        if self.current_token == Token::Bang {
                            self.advance()?;
                            if let Token::CellRef(cell) = &self.current_token {
                                let cell = cell.clone();
                                self.advance()?;
                                // Strip surrounding quotes on each name for
                                // the marker. The evaluator will compare
                                // case-insensitively against sheet_names.
                                let strip = |s: &str| -> String {
                                    if s.starts_with('\'') && s.ends_with('\'') && s.len() >= 2 {
                                        s[1..s.len() - 1].replace("''", "'")
                                    } else {
                                        s.to_string()
                                    }
                                };
                                let s1c = strip(&name);
                                let s2c = strip(&s2);
                                let marker = format!("{}..{}!{}", s1c, s2c, cell);
                                return Ok(Expr::Range(marker.clone(), marker));
                            }
                            return Err("Expected cell reference after '!' in 3-D ref".to_string());
                        }
                        return Err("Expected '!' after 3-D sheet range".to_string());
                    }
                    return Err("Expected sheet name after ':'".to_string());
                }
                if self.current_token == Token::LeftParen {
                    self.advance()?;
                    let args = self.parse_argument_list()?;
                    self.expect(Token::RightParen)?;
                    // Detect LET / LAMBDA at parse time so the evaluator
                    // gets a structured node instead of an opaque FunctionCall.
                    let upper = name.to_uppercase();
                    if upper == "LET" {
                        return Self::build_let(args);
                    }
                    if upper == "LAMBDA" {
                        return Self::build_lambda(args);
                    }
                    Ok(Expr::FunctionCall { name, args })
                } else {
                    Ok(Expr::NamedRef(name))
                }
            }
            
            Token::LeftParen => {
                self.advance()?;
                let expr = self.parse_equality()?;
                self.expect(Token::RightParen)?;
                Ok(expr)
            }

            Token::LeftBrace => {
                self.advance()?;
                let mut rows: Vec<Vec<Expr>> = Vec::new();
                let mut current_row: Vec<Expr> = Vec::new();
                if self.current_token != Token::RightBrace {
                    current_row.push(self.parse_equality()?);
                    loop {
                        match self.current_token {
                            Token::Comma => {
                                self.advance()?;
                                current_row.push(self.parse_equality()?);
                            }
                            Token::Semicolon => {
                                self.advance()?;
                                rows.push(std::mem::take(&mut current_row));
                                current_row.push(self.parse_equality()?);
                            }
                            _ => break,
                        }
                    }
                }
                if !current_row.is_empty() {
                    rows.push(current_row);
                }
                self.expect(Token::RightBrace)?;
                // Validate rectangular shape.
                if let Some(first) = rows.first() {
                    let cols = first.len();
                    if rows.iter().any(|r| r.len() != cols) {
                        return Err("Array literal: ragged rows".to_string());
                    }
                }
                Ok(Expr::ArrayLiteral { rows })
            }
            
            _ => Err(format!("Unexpected token: {:?}", self.current_token)),
        }
    }
    
    /// Parses function argument lists.
    fn parse_argument_list(&mut self) -> Result<Vec<Expr>, String> {
        let mut args = Vec::new();
        
        // Empty argument list
        if self.current_token == Token::RightParen {
            return Ok(args);
        }
        
        // Parse first argument
        args.push(self.parse_equality()?);
        
        // Parse remaining arguments
        while self.current_token == Token::Comma {
            self.advance()?;
            args.push(self.parse_equality()?);
        }
        
        Ok(args)
    }
}

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
pub struct ExpressionEvaluator<'a> {
    spreadsheet: &'a Spreadsheet,
    function_registry: &'a FunctionRegistry,
    named_ranges: Option<&'a HashMap<String, String>>,
    /// Workbook for cross-sheet refs. When absent, sheet-qualified refs
    /// return `#REF!`.
    workbook: Option<&'a super::models::Workbook>,
    /// Stack of local-binding scopes for LET / LAMBDA. Innermost is last.
    /// Uses interior mutability so `evaluate(&self)` can push/pop.
    local_scope: std::cell::RefCell<Vec<HashMap<String, Value>>>,
}

impl<'a> ExpressionEvaluator<'a> {
    pub fn new(spreadsheet: &'a Spreadsheet, function_registry: &'a FunctionRegistry) -> Self {
        Self {
            spreadsheet,
            function_registry,
            named_ranges: None,
            workbook: None,
            local_scope: std::cell::RefCell::new(Vec::new()),
        }
    }

    pub fn with_names(
        spreadsheet: &'a Spreadsheet,
        function_registry: &'a FunctionRegistry,
        names: &'a HashMap<String, String>,
    ) -> Self {
        Self {
            spreadsheet,
            function_registry,
            named_ranges: Some(names),
            workbook: None,
            local_scope: std::cell::RefCell::new(Vec::new()),
        }
    }

    pub fn with_workbook(
        workbook: &'a super::models::Workbook,
        spreadsheet: &'a Spreadsheet,
        function_registry: &'a FunctionRegistry,
        names: &'a HashMap<String, String>,
    ) -> Self {
        Self {
            spreadsheet,
            function_registry,
            named_ranges: Some(names),
            workbook: Some(workbook),
            local_scope: std::cell::RefCell::new(Vec::new()),
        }
    }

    /// Look up `name` in the local scope (innermost first).
    fn lookup_local(&self, name: &str) -> Option<Value> {
        let scope = self.local_scope.borrow();
        for frame in scope.iter().rev() {
            if let Some(v) = frame.get(name).or_else(|| frame.get(&name.to_uppercase())) {
                return Some(v.clone());
            }
        }
        None
    }

    /// Resolve a possibly sheet-qualified cell ref to the addressed sheet.
    /// Returns the "current" sheet for unqualified refs. Currently used only
    /// for cross-sheet INDIRECT/OFFSET work; kept here for symmetry.
    #[allow(dead_code)]
    fn sheet_for_ref(&self, cell_ref: &str) -> &Spreadsheet {
        if let (Some(wb), Some((Some(sheet_name), _, _, _, _))) = (
            self.workbook,
            super::models::Spreadsheet::parse_qualified_reference(cell_ref),
        ) {
            if let Some(idx) = wb.sheet_names.iter().position(|n| n == &sheet_name) {
                return &wb.sheets[idx];
            }
        }
        self.spreadsheet
    }

    /// Look up a name and parse its value into an `Expr`. The value may be
    /// a single cell ref or a range like `A1:B10`.
    fn resolve_name(&self, name: &str) -> Result<Expr, String> {
        let names = self
            .named_ranges
            .ok_or_else(|| format!("Unknown identifier: {}", name))?;
        let value = names
            .get(&name.to_uppercase())
            .or_else(|| names.get(name))
            .ok_or_else(|| format!("Unknown name: {}", name))?;
        let mut parser = Parser::new(value)?;
        parser.parse()
    }
    
    /// Evaluates an expression AST to a value result.
    pub fn evaluate(&self, expr: &Expr) -> Result<Value, String> {
        match expr {
            Expr::Number(value) => Ok(Value::Number(*value)),

            Expr::String(text) => Ok(Value::String(text.clone())),

            Expr::NamedRef(name) => {
                // Local scope (LET / LAMBDA params) takes priority over
                // workbook-level named ranges.
                if let Some(v) = self.lookup_local(name) {
                    return Ok(v);
                }
                let resolved = self.resolve_name(name)?;
                self.evaluate(&resolved)
            }

            Expr::Let { bindings, body } => {
                let mut frame: HashMap<String, Value> = HashMap::new();
                for (name, value_expr) in bindings {
                    // Bindings see earlier siblings — push the frame
                    // incrementally so later RHS expressions can refer to
                    // earlier ones (Excel LET semantics).
                    self.local_scope.borrow_mut().push(frame.clone());
                    let v = self.evaluate(value_expr);
                    self.local_scope.borrow_mut().pop();
                    let v = v?;
                    frame.insert(name.to_uppercase(), v);
                }
                self.local_scope.borrow_mut().push(frame);
                let result = self.evaluate(body);
                self.local_scope.borrow_mut().pop();
                result
            }

            // A bare LAMBDA outside an invocation evaluates to itself as a
            // single-cell value. The string form is its formula source so
            // round-trip through named ranges works.
            Expr::Lambda { .. } => Ok(Value::String(format!("[LAMBDA]"))),

            Expr::ArrayLiteral { rows } => {
                let r = rows.len();
                let c = rows.first().map(|r| r.len()).unwrap_or(0);
                if r == 0 || c == 0 {
                    return Ok(Value::Array { rows: 0, cols: 0, data: Vec::new() });
                }
                let mut data = Vec::with_capacity(r * c);
                for row in rows {
                    for cell in row {
                        data.push(self.evaluate(cell)?);
                    }
                }
                Ok(Value::Array { rows: r, cols: c, data })
            }

            Expr::CellRef(cell_ref) => {
                let parsed = super::models::Spreadsheet::parse_qualified_reference(cell_ref)
                    .ok_or_else(|| format!("Invalid cell reference: {}", cell_ref))?;
                let (sheet_opt, row, col, _, _) = parsed;
                let sheet = if let Some(name) = sheet_opt {
                    let wb = self.workbook.ok_or_else(|| {
                        format!("Cross-sheet ref {} but no workbook context", cell_ref)
                    })?;
                    // Sheet names are case-insensitive (Excel convention),
                    // and the lexer uppercases unquoted identifiers.
                    let idx = wb
                        .sheet_names
                        .iter()
                        .position(|n| n.eq_ignore_ascii_case(&name))
                        .ok_or_else(|| format!("Unknown sheet: {}", name))?;
                    &wb.sheets[idx]
                } else {
                    self.spreadsheet
                };
                let cell = sheet.get_cell(row, col);
                if let Ok(num) = cell.value.parse::<f64>() {
                    Ok(Value::Number(num))
                } else {
                    Ok(Value::String(cell.value))
                }
            }

            Expr::Range(start_cell, end_cell) => {
                // Expand the range to a Value::Array so it can participate in
                // arithmetic broadcasting (e.g. `=A1:A10 * 2`). Aggregate
                // functions flatten it transparently.
                let arr = self.evaluate_function_args(&[Expr::Range(
                    start_cell.clone(),
                    end_cell.clone(),
                )])?;
                Ok(arr.into_iter().next().unwrap_or(Value::List(Vec::new())))
            }
            
            Expr::Binary { left, operator, right } => {
                let left_val = self.evaluate(left)?;
                let right_val = self.evaluate(right)?;

                // Error propagation: any operand carrying an error returns
                // that error directly. This matches Excel semantics.
                if let Some(e) = left_val.first_error() {
                    return Ok(Value::Error(e));
                }
                if let Some(e) = right_val.first_error() {
                    return Ok(Value::Error(e));
                }
                match operator {
                    BinaryOp::Add => broadcast_binary(&left_val, &right_val, |a, b| {
                        Ok(Value::Number(a.to_number() + b.to_number()))
                    }),
                    BinaryOp::Subtract => broadcast_binary(&left_val, &right_val, |a, b| {
                        Ok(Value::Number(a.to_number() - b.to_number()))
                    }),
                    BinaryOp::Multiply => broadcast_binary(&left_val, &right_val, |a, b| {
                        Ok(Value::Number(a.to_number() * b.to_number()))
                    }),
                    BinaryOp::Divide => broadcast_binary(&left_val, &right_val, |a, b| {
                        let r = b.to_number();
                        if r == 0.0 {
                            Ok(Value::Error(ErrorKind::Div0))
                        } else {
                            Ok(Value::Number(a.to_number() / r))
                        }
                    }),
                    BinaryOp::Modulo => broadcast_binary(&left_val, &right_val, |a, b| {
                        let r = b.to_number();
                        if r == 0.0 {
                            Ok(Value::Error(ErrorKind::Div0))
                        } else {
                            Ok(Value::Number(a.to_number() % r))
                        }
                    }),
                    BinaryOp::Power => broadcast_binary(&left_val, &right_val, |a, b| {
                        Ok(Value::Number(a.to_number().powf(b.to_number())))
                    }),
                    BinaryOp::Concatenate => {
                        let left_str = left_val.to_string();
                        let right_str = right_val.to_string();
                        Ok(Value::String(format!("{}{}", left_str, right_str)))
                    }
                    BinaryOp::Less => {
                        let left_num = left_val.to_number();
                        let right_num = right_val.to_number();
                        Ok(Value::Number(if left_num < right_num { 1.0 } else { 0.0 }))
                    }
                    BinaryOp::LessEqual => {
                        let left_num = left_val.to_number();
                        let right_num = right_val.to_number();
                        Ok(Value::Number(if left_num <= right_num { 1.0 } else { 0.0 }))
                    }
                    BinaryOp::Greater => {
                        let left_num = left_val.to_number();
                        let right_num = right_val.to_number();
                        Ok(Value::Number(if left_num > right_num { 1.0 } else { 0.0 }))
                    }
                    BinaryOp::GreaterEqual => {
                        let left_num = left_val.to_number();
                        let right_num = right_val.to_number();
                        Ok(Value::Number(if left_num >= right_num { 1.0 } else { 0.0 }))
                    }
                    BinaryOp::Equal => {
                        let result = match (&left_val, &right_val) {
                            (Value::Number(l), Value::Number(r)) => numbers_equal(*l, *r),
                            (Value::String(l), Value::String(r)) => l == r,
                            (Value::Bool(l), Value::Bool(r)) => l == r,
                            _ => left_val.to_string() == right_val.to_string(),
                        };
                        Ok(Value::Number(if result { 1.0 } else { 0.0 }))
                    }
                    BinaryOp::NotEqual => {
                        let result = match (&left_val, &right_val) {
                            (Value::Number(l), Value::Number(r)) => !numbers_equal(*l, *r),
                            (Value::String(l), Value::String(r)) => l != r,
                            (Value::Bool(l), Value::Bool(r)) => l != r,
                            _ => left_val.to_string() != right_val.to_string(),
                        };
                        Ok(Value::Number(if result { 1.0 } else { 0.0 }))
                    }
                }
            }
            
            Expr::Unary { operator, operand } => {
                let operand_val = self.evaluate(operand)?;
                
                match operator {
                    UnaryOp::Plus => Ok(Value::Number(operand_val.to_number())),
                    UnaryOp::Minus => Ok(Value::Number(-operand_val.to_number())),
                }
            }
            
            Expr::FunctionCall { name, args } => {
                let upper = name.to_uppercase();
                if upper == "INDIRECT" {
                    return self.eval_indirect(args);
                }
                if upper == "OFFSET" {
                    return self.eval_offset(args);
                }
                // Higher-order lambda helpers: last arg must be a LAMBDA.
                if matches!(upper.as_str(), "MAP" | "REDUCE" | "BYROW" | "BYCOL" | "SCAN" | "MAKEARRAY") {
                    return self.eval_lambda_helper(&upper, args);
                }
                // User-defined LAMBDA stored as a named range:
                // `:name DOUBLE LAMBDA(x, x*2)` then `=DOUBLE(7)` looks up
                // DOUBLE, parses it as a lambda, and invokes with the args.
                if let Some(names) = self.named_ranges {
                    if let Some(value) = names
                        .get(&upper)
                        .or_else(|| names.get(name))
                    {
                        if let Ok(mut p) = Parser::new(value) {
                            if let Ok(Expr::Lambda { params, body }) = p.parse() {
                                if params.len() != args.len() {
                                    return Err(format!(
                                        "{}: expected {} arg(s), got {}",
                                        name,
                                        params.len(),
                                        args.len()
                                    ));
                                }
                                let mut frame: HashMap<String, Value> = HashMap::new();
                                for (p_name, a) in params.iter().zip(args.iter()) {
                                    let v = self.evaluate(a)?;
                                    frame.insert(p_name.to_uppercase(), v);
                                }
                                self.local_scope.borrow_mut().push(frame);
                                let result = self.evaluate(&body);
                                self.local_scope.borrow_mut().pop();
                                return result;
                            }
                        }
                    }
                }
                let func = self.function_registry.get_function(name)
                    .ok_or_else(|| format!("Unknown function: {}", name))?;
                let arg_values = self.evaluate_function_args(args)?;
                func(&arg_values)
            }
        }
    }

    /// Dispatch table for higher-order lambda helpers: MAP, REDUCE, BYROW,
    /// BYCOL, SCAN, MAKEARRAY. Each expects a LAMBDA as the last argument.
    fn eval_lambda_helper(&self, name: &str, args: &[Expr]) -> Result<Value, String> {
        let (params, body) = match args.last() {
            Some(Expr::Lambda { params, body }) => (params.clone(), body.clone()),
            _ => return Err(format!("{}: last argument must be a LAMBDA", name)),
        };
        let invoke = |inputs: Vec<Value>| -> Result<Value, String> {
            if inputs.len() != params.len() {
                return Err(format!("{}: lambda arity mismatch", name));
            }
            let mut frame: HashMap<String, Value> = HashMap::new();
            for (p, v) in params.iter().zip(inputs.iter()) {
                frame.insert(p.to_uppercase(), v.clone());
            }
            self.local_scope.borrow_mut().push(frame);
            let result = self.evaluate(&body);
            self.local_scope.borrow_mut().pop();
            result
        };
        match name {
            "MAP" => {
                // MAP(array1, [array2, ...], lambda) — element-wise.
                let arrays: Vec<_> = args[..args.len() - 1]
                    .iter()
                    .map(|a| self.evaluate(a))
                    .collect::<Result<Vec<_>, _>>()?;
                if arrays.is_empty() {
                    return Err("MAP requires at least one array".to_string());
                }
                let first_shape = shape_of(&arrays[0]);
                let rows = first_shape.0;
                let cols = first_shape.1;
                let mut out = Vec::with_capacity(rows * cols);
                for i in 0..(rows * cols) {
                    let inputs: Vec<Value> = arrays
                        .iter()
                        .map(|a| {
                            let (_, _, d) = shape_of(a);
                            d.get(i).cloned().unwrap_or(Value::String(String::new()))
                        })
                        .collect();
                    out.push(invoke(inputs)?);
                }
                Ok(Value::Array { rows, cols, data: out })
            }
            "REDUCE" => {
                // REDUCE(initial, array, lambda(acc, val))
                if args.len() != 3 {
                    return Err("REDUCE requires 3 arguments".to_string());
                }
                let mut acc = self.evaluate(&args[0])?;
                let arr = self.evaluate(&args[1])?;
                for v in arr.flatten() {
                    acc = invoke(vec![acc, v])?;
                }
                Ok(acc)
            }
            "BYROW" | "BYCOL" => {
                if args.len() != 2 {
                    return Err(format!("{} requires 2 arguments", name));
                }
                let arr = self.evaluate(&args[0])?;
                let (rows, cols, data) = shape_of(&arr);
                if name == "BYROW" {
                    let mut out = Vec::with_capacity(rows);
                    for r in 0..rows {
                        let row: Vec<Value> = (0..cols)
                            .map(|c| data[r * cols + c].clone())
                            .collect();
                        let v = invoke(vec![Value::Array {
                            rows: 1,
                            cols,
                            data: row,
                        }])?;
                        out.push(v);
                    }
                    Ok(Value::Array { rows, cols: 1, data: out })
                } else {
                    let mut out = Vec::with_capacity(cols);
                    for c in 0..cols {
                        let col: Vec<Value> = (0..rows)
                            .map(|r| data[r * cols + c].clone())
                            .collect();
                        let v = invoke(vec![Value::Array {
                            rows,
                            cols: 1,
                            data: col,
                        }])?;
                        out.push(v);
                    }
                    Ok(Value::Array { rows: 1, cols, data: out })
                }
            }
            "SCAN" => {
                // SCAN(initial, array, lambda(acc, val))
                if args.len() != 3 {
                    return Err("SCAN requires 3 arguments".to_string());
                }
                let mut acc = self.evaluate(&args[0])?;
                let arr = self.evaluate(&args[1])?;
                let flat = arr.flatten();
                let mut out = Vec::with_capacity(flat.len());
                for v in flat {
                    acc = invoke(vec![acc.clone(), v])?;
                    out.push(acc.clone());
                }
                let cols = out.len();
                Ok(Value::Array { rows: 1, cols, data: out })
            }
            "MAKEARRAY" => {
                // MAKEARRAY(rows, cols, lambda(r, c))
                if args.len() != 3 {
                    return Err("MAKEARRAY requires 3 arguments".to_string());
                }
                let rows = self.evaluate(&args[0])?.to_number() as usize;
                let cols = self.evaluate(&args[1])?.to_number() as usize;
                let mut out = Vec::with_capacity(rows * cols);
                for r in 0..rows {
                    for c in 0..cols {
                        out.push(invoke(vec![
                            Value::Number((r + 1) as f64),
                            Value::Number((c + 1) as f64),
                        ])?);
                    }
                }
                Ok(Value::Array { rows, cols, data: out })
            }
            _ => unreachable!(),
        }
    }

    /// INDIRECT(ref_text) → value at the cell whose A1-style address is `ref_text`.
    /// Errors if the string can't be parsed as a cell reference.
    fn eval_indirect(&self, args: &[Expr]) -> Result<Value, String> {
        if args.len() != 1 {
            return Err("INDIRECT requires exactly 1 argument".to_string());
        }
        let ref_text = self.evaluate(&args[0])?.to_string();
        // Accept either a single cell ref or a range. Ranges in scalar
        // context degrade to first cell.
        if let Some(colon) = ref_text.find(':') {
            let start = &ref_text[..colon];
            let end = &ref_text[colon + 1..];
            let range_expr = Expr::Range(start.to_string(), end.to_string());
            // Implicit intersection: first cell of range.
            let values = self.evaluate_function_args(&[range_expr])?;
            return Ok(values.into_iter().next().unwrap_or(Value::String(String::new())));
        }
        let (row, col) = Spreadsheet::parse_cell_reference(&ref_text)
            .ok_or_else(|| format!("INDIRECT: invalid reference: {}", ref_text))?;
        let cell = self.spreadsheet.get_cell(row, col);
        if let Ok(n) = cell.value.parse::<f64>() {
            Ok(Value::Number(n))
        } else {
            Ok(Value::String(cell.value))
        }
    }

    /// OFFSET(base, rows, cols, [height], [width]) → value or range starting
    /// at (base.row + rows, base.col + cols) of the given size. Height and
    /// width default to 1. Returns implicit-intersection value when size is 1×1.
    fn eval_offset(&self, args: &[Expr]) -> Result<Value, String> {
        if args.len() < 3 || args.len() > 5 {
            return Err("OFFSET requires 3-5 arguments".to_string());
        }
        // Base must be a CellRef (extract row/col without evaluating).
        let (base_row, base_col) = match &args[0] {
            Expr::CellRef(r) => Spreadsheet::parse_cell_reference(r)
                .ok_or_else(|| format!("OFFSET: invalid base: {}", r))?,
            Expr::NamedRef(name) => {
                let resolved = self.resolve_name(name)?;
                match resolved {
                    Expr::CellRef(r) => Spreadsheet::parse_cell_reference(&r)
                        .ok_or_else(|| format!("OFFSET: invalid base: {}", r))?,
                    Expr::Range(start, _) => Spreadsheet::parse_cell_reference(&start)
                        .ok_or_else(|| format!("OFFSET: invalid base: {}", start))?,
                    _ => return Err("OFFSET: base must be a cell ref".to_string()),
                }
            }
            _ => return Err("OFFSET: base must be a cell ref".to_string()),
        };
        let row_off = self.evaluate(&args[1])?.to_number() as i64;
        let col_off = self.evaluate(&args[2])?.to_number() as i64;
        let height = args
            .get(3)
            .map(|e| self.evaluate(e).map(|v| v.to_number() as i64))
            .transpose()?
            .unwrap_or(1)
            .max(1);
        let width = args
            .get(4)
            .map(|e| self.evaluate(e).map(|v| v.to_number() as i64))
            .transpose()?
            .unwrap_or(1)
            .max(1);
        let new_row_i = base_row as i64 + row_off;
        let new_col_i = base_col as i64 + col_off;
        // Negative offsets and rows/cols beyond the sheet are `#REF!`.
        if new_row_i < 0
            || new_col_i < 0
            || new_row_i as usize >= self.spreadsheet.rows
            || new_col_i as usize >= self.spreadsheet.cols
        {
            return Ok(Value::Error(ErrorKind::Ref));
        }
        let new_row = new_row_i as usize;
        let new_col = new_col_i as usize;
        // Also bounds-check the height/width corner.
        if new_row + height as usize > self.spreadsheet.rows
            || new_col + width as usize > self.spreadsheet.cols
        {
            return Ok(Value::Error(ErrorKind::Ref));
        }
        if height == 1 && width == 1 {
            let cell = self.spreadsheet.get_cell(new_row, new_col);
            if let Ok(n) = cell.value.parse::<f64>() {
                return Ok(Value::Number(n));
            }
            return Ok(Value::String(cell.value));
        }
        // Build a List of the resulting block.
        let mut list = Vec::with_capacity((height * width) as usize);
        for r in new_row..new_row + height as usize {
            for c in new_col..new_col + width as usize {
                let cell = self.spreadsheet.get_cell(r, c);
                if let Ok(n) = cell.value.parse::<f64>() {
                    list.push(Value::Number(n));
                } else {
                    list.push(Value::String(cell.value));
                }
            }
        }
        Ok(Value::List(list))
    }
    
    /// Evaluates function arguments. Ranges produce a single `Value::List`
    /// so range-aware functions (SUMIF, VLOOKUP, ...) can preserve structure.
    /// Scalar aggregate functions (SUM, AVG, ...) call `flatten_args`.
    fn evaluate_function_args(&self, args: &[Expr]) -> Result<Vec<Value>, String> {
        let mut values = Vec::new();
        for arg in args {
            // Resolution priority for bare names:
            // 1. Local scope (LET / LAMBDA params) — use the bound value as-is.
            // 2. Named range — substitute its parsed Expr (lets ranges expand).
            // 3. Fall through to scalar eval (will error if truly unknown).
            let effective = if let Expr::NamedRef(name) = arg {
                if self.lookup_local(name).is_some() {
                    None // handled by evaluate(target)
                } else {
                    Some(self.resolve_name(name)?)
                }
            } else {
                None
            };
            let target = effective.as_ref().unwrap_or(arg);
            match target {
                Expr::Range(start_cell, end_cell) => {
                    // 3-D range (Sheet1:Sheet3!A1) — both endpoints carry the
                    // same `<S1>..<S3>!<cell>` marker. Expand across sheets.
                    if let Some((s1, s2, cell)) =
                        super::models::Spreadsheet::parse_three_d_marker(start_cell)
                    {
                        let wb = self.workbook.ok_or_else(|| {
                            "3-D range needs workbook context".to_string()
                        })?;
                        let i1 = wb
                            .sheet_names
                            .iter()
                            .position(|n| n.eq_ignore_ascii_case(&s1))
                            .ok_or_else(|| format!("Unknown sheet: {}", s1))?;
                        let i2 = wb
                            .sheet_names
                            .iter()
                            .position(|n| n.eq_ignore_ascii_case(&s2))
                            .ok_or_else(|| format!("Unknown sheet: {}", s2))?;
                        let (lo, hi) = if i1 <= i2 { (i1, i2) } else { (i2, i1) };
                        let (row, col) =
                            super::models::Spreadsheet::parse_cell_reference(&cell)
                                .ok_or_else(|| format!("Invalid cell: {}", cell))?;
                        let mut list = Vec::new();
                        for i in lo..=hi {
                            let c = wb.sheets[i].get_cell(row, col);
                            if let Ok(n) = c.value.parse::<f64>() {
                                list.push(Value::Number(n));
                            } else {
                                list.push(Value::String(c.value));
                            }
                        }
                        values.push(Value::List(list));
                        continue;
                    }
                    // Regular range. Endpoints may be sheet-qualified;
                    // both must address the same sheet.
                    let sp = super::models::Spreadsheet::parse_qualified_reference(start_cell)
                        .ok_or_else(|| format!("Invalid cell reference: {}", start_cell))?;
                    let ep = super::models::Spreadsheet::parse_qualified_reference(end_cell)
                        .ok_or_else(|| format!("Invalid cell reference: {}", end_cell))?;
                    let sheet = if let Some(name) = &sp.0 {
                        let wb = self.workbook.ok_or_else(|| {
                            "Cross-sheet range needs workbook context".to_string()
                        })?;
                        let idx = wb
                            .sheet_names
                            .iter()
                            .position(|n| n.eq_ignore_ascii_case(name))
                            .ok_or_else(|| format!("Unknown sheet: {}", name))?;
                        &wb.sheets[idx]
                    } else {
                        self.spreadsheet
                    };
                    let start = (sp.1, sp.2);
                    let end = (ep.1, ep.2);
                    let r0 = start.0.min(end.0);
                    let r1 = start.0.max(end.0);
                    let c0 = start.1.min(end.1);
                    let c1 = start.1.max(end.1);
                    let rows = r1 - r0 + 1;
                    let cols = c1 - c0 + 1;
                    let mut data = Vec::with_capacity(rows * cols);
                    for row in r0..=r1 {
                        for col in c0..=c1 {
                            let cell = sheet.get_cell(row, col);
                            if let Ok(num) = cell.value.parse::<f64>() {
                                data.push(Value::Number(num));
                            } else {
                                data.push(Value::String(cell.value));
                            }
                        }
                    }
                    values.push(Value::Array { rows, cols, data });
                }
                _ => {
                    values.push(self.evaluate(target)?);
                }
            }
        }
        Ok(values)
    }
}

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