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
    
    // String concatenation
    Ampersand,
    
    // End of input
    Eof,
}

/// Represents a value in the spreadsheet system that can be either a number or string.
#[derive(Debug, Clone, PartialEq)]
pub enum Value {
    Number(f64),
    String(String),
}

impl Value {
    /// Converts value to string representation
    pub fn to_string(&self) -> String {
        match self {
            Value::Number(n) => n.to_string(),
            Value::String(s) => s.clone(),
        }
    }
    
    /// Attempts to convert value to number, returns 0.0 for non-numeric strings
    pub fn to_number(&self) -> f64 {
        match self {
            Value::Number(n) => *n,
            Value::String(s) => s.parse::<f64>().unwrap_or(0.0),
        }
    }
    
    /// Returns true if this value is truthy (non-zero number or non-empty string)
    pub fn is_truthy(&self) -> bool {
        match self {
            Value::Number(n) => *n != 0.0,
            Value::String(s) => !s.is_empty(),
        }
    }
}

/// Represents an Abstract Syntax Tree node for expressions.
#[derive(Debug, Clone, PartialEq)]
pub enum Expr {
    // Literals
    Number(f64),
    String(String),
    CellRef(String),
    Range(String, String), // start_cell, end_cell
    
    // Binary operations
    Binary {
        left: Box<Expr>,
        operator: BinaryOp,
        right: Box<Expr>,
    },
    
    // Unary operations
    Unary {
        operator: UnaryOp,
        operand: Box<Expr>,
    },
    
    // Function calls
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
    
    /// Reads an identifier (function name or cell reference).
    fn read_identifier(&mut self) -> String {
        let mut identifier = String::new();
        
        while let Some(ch) = self.current_char {
            if ch.is_ascii_alphanumeric() || ch == '_' {
                identifier.push(ch.to_ascii_uppercase());
                self.advance();
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
    fn classify_identifier(&self, identifier: &str) -> Token {
        // Check if it's a cell reference (letters followed by numbers)
        let mut has_letters = false;
        let mut has_numbers = false;
        let mut letters_first = true;
        
        for ch in identifier.chars() {
            if ch.is_ascii_alphabetic() {
                if has_numbers {
                    letters_first = false;
                }
                has_letters = true;
            } else if ch.is_ascii_digit() {
                has_numbers = true;
            }
        }
        
        // Valid cell reference: letters followed by numbers (A1, B2, AA123, etc.)
        if has_letters && has_numbers && letters_first {
            Token::CellRef(identifier.to_string())
        } else {
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
                
                // Identifiers and cell references
                'A'..='Z' | 'a'..='z' => {
                    let identifier = self.read_identifier();
                    
                    // All identifiers are treated as either cell references or function names
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
            let sum: f64 = args.iter().map(|v| v.to_number()).sum();
            Ok(Value::Number(sum))
        });
        
        self.register_function("AVERAGE", |args| {
            if args.is_empty() {
                Err("AVERAGE requires at least one argument".to_string())
            } else {
                let sum: f64 = args.iter().map(|v| v.to_number()).sum();
                Ok(Value::Number(sum / args.len() as f64))
            }
        });
        
        self.register_function("MIN", |args| {
            args.iter().map(|v| v.to_number()).fold(None, |acc: Option<f64>, x| {
                Some(acc.map_or(x, |a| a.min(x)))
            }).map(Value::Number).ok_or_else(|| "MIN requires at least one argument".to_string())
        });
        
        self.register_function("MAX", |args| {
            args.iter().map(|v| v.to_number()).fold(None, |acc: Option<f64>, x| {
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
            let result = args.iter().all(|v| v.is_truthy());
            Ok(Value::Number(if result { 1.0 } else { 0.0 }))
        });
        
        self.register_function("OR", |args| {
            let result = args.iter().any(|v| v.is_truthy());
            Ok(Value::Number(if result { 1.0 } else { 0.0 }))
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
                    Err("SQRT of negative number".to_string())
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
            let result = args.iter().map(|v| v.to_string()).collect::<String>();
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
                let start = args[1].to_number() as usize; // 0-based indexing
                let length = args[2].to_number() as usize;
                let chars: Vec<char> = text.chars().collect();
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
                let start_pos = if args.len() == 3 {
                    args[2].to_number() as usize // 0-based indexing
                } else {
                    0
                };
                
                let within_chars: Vec<char> = within_text.chars().collect();
                if start_pos >= within_chars.len() {
                    return Err("Start position is beyond text length".to_string());
                }
                
                let search_in = within_chars[start_pos..].iter().collect::<String>();
                match search_in.find(&search_text) {
                    Some(pos) => Ok(Value::Number((start_pos + pos) as f64)), // 0-based result
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
                Err("GET requires exactly 1 argument (URL)".to_string())
            } else {
                let url = args[0].to_string();
                match reqwest::blocking::get(&url) {
                    Ok(response) => {
                        match response.text() {
                            Ok(content) => Ok(Value::String(content)),
                            Err(e) => Err(format!("Failed to read response: {}", e)),
                        }
                    }
                    Err(e) => Err(format!("Failed to fetch URL {}: {}", url, e)),
                }
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
                Ok(Value::Number(args[0].to_number().trunc()))
            }
        });

        self.register_function("MOD", |args| {
            if args.len() != 2 {
                Err("MOD requires exactly 2 arguments".to_string())
            } else {
                let divisor = args[1].to_number();
                if divisor == 0.0 {
                    Err("MOD division by zero".to_string())
                } else {
                    Ok(Value::Number(args[0].to_number() % divisor))
                }
            }
        });

        self.register_function("LOG", |args| {
            match args.len() {
                1 => Ok(Value::Number(args[0].to_number().log10())),
                2 => {
                    let base = args[1].to_number();
                    Ok(Value::Number(args[0].to_number().log(base)))
                }
                _ => Err("LOG requires 1 or 2 arguments".to_string()),
            }
        });

        self.register_function("LN", |args| {
            if args.len() != 1 {
                Err("LN requires exactly 1 argument".to_string())
            } else {
                Ok(Value::Number(args[0].to_number().ln()))
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
                Ok(Value::String(text.replace(&old, &new)))
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
                let count = args[1].to_number() as usize;
                Ok(Value::String(text.repeat(count)))
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

        self.register_function("ISBLANK", |args| {
            if args.len() != 1 {
                Err("ISBLANK requires exactly 1 argument".to_string())
            } else {
                let is_blank = match &args[0] {
                    Value::String(s) => s.is_empty(),
                    Value::Number(_) => false,
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
                };
                Ok(Value::Number(type_num))
            }
        });

        // --- Stats functions ---

        self.register_function("COUNT", |args| {
            let count = args
                .iter()
                .filter(|v| matches!(v, Value::Number(_)))
                .count();
            Ok(Value::Number(count as f64))
        });

        self.register_function("COUNTA", |args| {
            let count = args
                .iter()
                .filter(|v| match v {
                    Value::Number(_) => true,
                    Value::String(s) => !s.is_empty(),
                })
                .count();
            Ok(Value::Number(count as f64))
        });

        // --- Visualization functions ---

        self.register_function("SPARKLINE", |args| {
            if args.is_empty() {
                return Err("SPARKLINE requires at least one argument".to_string());
            }
            let values: Vec<f64> = args.iter().map(|v| v.to_number()).collect();
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
                
                // Check if this is the start of a range
                if self.current_token == Token::Colon {
                    self.advance()?;
                    if let Token::CellRef(end_cell) = &self.current_token {
                        let end_cell = end_cell.clone();
                        self.advance()?;
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
                
                // Check if this is a function call
                if self.current_token == Token::LeftParen {
                    self.advance()?;
                    let args = self.parse_argument_list()?;
                    self.expect(Token::RightParen)?;
                    Ok(Expr::FunctionCall { name, args })
                } else {
                    Err(format!("Unknown identifier: {}", name))
                }
            }
            
            Token::LeftParen => {
                self.advance()?;
                let expr = self.parse_equality()?;
                self.expect(Token::RightParen)?;
                Ok(expr)
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

/// Expression evaluator that walks the AST and computes results.
pub struct ExpressionEvaluator<'a> {
    spreadsheet: &'a Spreadsheet,
    function_registry: &'a FunctionRegistry,
}

impl<'a> ExpressionEvaluator<'a> {
    /// Creates a new expression evaluator.
    pub fn new(spreadsheet: &'a Spreadsheet, function_registry: &'a FunctionRegistry) -> Self {
        Self {
            spreadsheet,
            function_registry,
        }
    }
    
    /// Evaluates an expression AST to a value result.
    pub fn evaluate(&self, expr: &Expr) -> Result<Value, String> {
        match expr {
            Expr::Number(value) => Ok(Value::Number(*value)),
            
            Expr::String(text) => Ok(Value::String(text.clone())),
            
            Expr::CellRef(cell_ref) => {
                let (row, col) = super::models::Spreadsheet::parse_cell_reference(cell_ref)
                    .ok_or_else(|| format!("Invalid cell reference: {}", cell_ref))?;
                let cell = self.spreadsheet.get_cell(row, col);
                // Try to parse as number first, otherwise return as string
                if let Ok(num) = cell.value.parse::<f64>() {
                    Ok(Value::Number(num))
                } else {
                    Ok(Value::String(cell.value))
                }
            }
            
            Expr::Range(start_cell, end_cell) => {
                // This shouldn't be called directly - ranges are handled by functions
                Err(format!("Range {}:{} cannot be evaluated directly", start_cell, end_cell))
            }
            
            Expr::Binary { left, operator, right } => {
                let left_val = self.evaluate(left)?;
                let right_val = self.evaluate(right)?;
                
                match operator {
                    BinaryOp::Add => {
                        let left_num = left_val.to_number();
                        let right_num = right_val.to_number();
                        Ok(Value::Number(left_num + right_num))
                    }
                    BinaryOp::Subtract => {
                        let left_num = left_val.to_number();
                        let right_num = right_val.to_number();
                        Ok(Value::Number(left_num - right_num))
                    }
                    BinaryOp::Multiply => {
                        let left_num = left_val.to_number();
                        let right_num = right_val.to_number();
                        Ok(Value::Number(left_num * right_num))
                    }
                    BinaryOp::Divide => {
                        let left_num = left_val.to_number();
                        let right_num = right_val.to_number();
                        if right_num == 0.0 {
                            Err("Division by zero".to_string())
                        } else {
                            Ok(Value::Number(left_num / right_num))
                        }
                    }
                    BinaryOp::Modulo => {
                        let left_num = left_val.to_number();
                        let right_num = right_val.to_number();
                        if right_num == 0.0 {
                            Err("Modulo by zero".to_string())
                        } else {
                            Ok(Value::Number(left_num % right_num))
                        }
                    }
                    BinaryOp::Power => {
                        let left_num = left_val.to_number();
                        let right_num = right_val.to_number();
                        Ok(Value::Number(left_num.powf(right_num)))
                    }
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
                        // Support both numeric and string equality
                        let result = match (&left_val, &right_val) {
                            (Value::Number(l), Value::Number(r)) => (l - r).abs() < f64::EPSILON,
                            (Value::String(l), Value::String(r)) => l == r,
                            _ => {
                                // Mixed types: compare as strings
                                left_val.to_string() == right_val.to_string()
                            }
                        };
                        Ok(Value::Number(if result { 1.0 } else { 0.0 }))
                    }
                    BinaryOp::NotEqual => {
                        // Support both numeric and string inequality
                        let result = match (&left_val, &right_val) {
                            (Value::Number(l), Value::Number(r)) => (l - r).abs() >= f64::EPSILON,
                            (Value::String(l), Value::String(r)) => l != r,
                            _ => {
                                // Mixed types: compare as strings
                                left_val.to_string() != right_val.to_string()
                            }
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
                let func = self.function_registry.get_function(name)
                    .ok_or_else(|| format!("Unknown function: {}", name))?;
                
                let arg_values = self.evaluate_function_args(args)?;
                func(&arg_values)
            }
        }
    }
    
    /// Evaluates function arguments, handling ranges.
    fn evaluate_function_args(&self, args: &[Expr]) -> Result<Vec<Value>, String> {
        let mut values = Vec::new();
        
        for arg in args {
            match arg {
                Expr::Range(start_cell, end_cell) => {
                    let start = super::models::Spreadsheet::parse_cell_reference(start_cell)
                        .ok_or_else(|| format!("Invalid cell reference: {}", start_cell))?;
                    let end = super::models::Spreadsheet::parse_cell_reference(end_cell)
                        .ok_or_else(|| format!("Invalid cell reference: {}", end_cell))?;
                    
                    for row in start.0..=end.0 {
                        for col in start.1..=end.1 {
                            let cell = self.spreadsheet.get_cell(row, col);
                            // Try to parse as number first, otherwise return as string
                            if let Ok(num) = cell.value.parse::<f64>() {
                                values.push(Value::Number(num));
                            } else {
                                values.push(Value::String(cell.value));
                            }
                        }
                    }
                }
                _ => {
                    values.push(self.evaluate(arg)?);
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
        sheet.set_cell(0, 0, CellData { value: "10".to_string(), formula: None, format: None, comment: None });
        sheet.set_cell(0, 1, CellData { value: "20".to_string(), formula: None, format: None, comment: None });
        sheet.set_cell(0, 2, CellData { value: "30".to_string(), formula: None, format: None, comment: None });
        sheet.set_cell(1, 0, CellData { value: "5".to_string(), formula: None, format: None, comment: None });
        sheet.set_cell(1, 1, CellData { value: "15".to_string(), formula: None, format: None, comment: None });
        sheet.set_cell(1, 2, CellData { value: "25".to_string(), formula: None, format: None, comment: None });
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
        sheet.set_cell(0, 0, CellData { value: "Hello".to_string(), formula: None, format: None, comment: None });
        sheet.set_cell(0, 1, CellData { value: "World".to_string(), formula: None, format: None, comment: None });
        sheet.set_cell(0, 2, CellData { value: "123".to_string(), formula: None, format: None, comment: None }); // Number as string
        
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
            Value::Number(n) => assert_eq!(n, 3.0), // 0-based indexing - "lo" found at position 3
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
        
        // Test division by zero
        let expr = Expr::Binary {
            left: Box::new(Expr::Number(10.0)),
            operator: BinaryOp::Divide,
            right: Box::new(Expr::Number(0.0)),
        };
        assert!(evaluator.evaluate(&expr).is_err());
        
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
        sheet.set_cell(0, 0, CellData { value: "1".to_string(), formula: None, format: None, comment: None });
        sheet.set_cell(0, 1, CellData { value: "5".to_string(), formula: None, format: None, comment: None });
        sheet.set_cell(0, 2, CellData { value: "10".to_string(), formula: None, format: None, comment: None });

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
        sheet.set_cell(0, 0, CellData { value: "5".to_string(), formula: None, format: None, comment: None });
        sheet.set_cell(0, 1, CellData { value: "5".to_string(), formula: None, format: None, comment: None });
        sheet.set_cell(0, 2, CellData { value: "5".to_string(), formula: None, format: None, comment: None });

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
        sheet.set_cell(0, 0, CellData { value: "7".to_string(), formula: None, format: None, comment: None });

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