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
    
    // End of input
    Eof,
}

/// Represents an Abstract Syntax Tree node for expressions.
#[derive(Debug, Clone, PartialEq)]
pub enum Expr {
    // Literals
    Number(f64),
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
                
                _ => Err(format!("Unexpected character: '{}'", ch)),
            }
        }
    }
}

/// Function signature for built-in and user-defined functions.
pub type FunctionImpl = fn(&[f64]) -> Result<f64, String>;

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
        self.register_function("SUM", |args| {
            Ok(args.iter().sum())
        });
        
        self.register_function("AVERAGE", |args| {
            if args.is_empty() {
                Err("AVERAGE requires at least one argument".to_string())
            } else {
                Ok(args.iter().sum::<f64>() / args.len() as f64)
            }
        });
        
        self.register_function("MIN", |args| {
            args.iter().fold(None, |acc: Option<f64>, &x| {
                Some(acc.map_or(x, |a| a.min(x)))
            }).ok_or_else(|| "MIN requires at least one argument".to_string())
        });
        
        self.register_function("MAX", |args| {
            args.iter().fold(None, |acc: Option<f64>, &x| {
                Some(acc.map_or(x, |a| a.max(x)))
            }).ok_or_else(|| "MAX requires at least one argument".to_string())
        });
        
        self.register_function("IF", |args| {
            if args.len() != 3 {
                Err("IF requires exactly 3 arguments".to_string())
            } else {
                Ok(if args[0] != 0.0 { args[1] } else { args[2] })
            }
        });
        
        self.register_function("AND", |args| {
            Ok(if args.iter().all(|&x| x != 0.0) { 1.0 } else { 0.0 })
        });
        
        self.register_function("OR", |args| {
            Ok(if args.iter().any(|&x| x != 0.0) { 1.0 } else { 0.0 })
        });
        
        self.register_function("NOT", |args| {
            if args.len() != 1 {
                Err("NOT requires exactly 1 argument".to_string())
            } else {
                Ok(if args[0] == 0.0 { 1.0 } else { 0.0 })
            }
        });
        
        self.register_function("ABS", |args| {
            if args.len() != 1 {
                Err("ABS requires exactly 1 argument".to_string())
            } else {
                Ok(args[0].abs())
            }
        });
        
        self.register_function("SQRT", |args| {
            if args.len() != 1 {
                Err("SQRT requires exactly 1 argument".to_string())
            } else if args[0] < 0.0 {
                Err("SQRT of negative number".to_string())
            } else {
                Ok(args[0].sqrt())
            }
        });
        
        self.register_function("ROUND", |args| {
            match args.len() {
                1 => Ok(args[0].round()),
                2 => {
                    let places = args[1] as i32;
                    let multiplier = 10f64.powi(places);
                    Ok((args[0] * multiplier).round() / multiplier)
                }
                _ => Err("ROUND requires 1 or 2 arguments".to_string()),
            }
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
        let mut left = self.parse_multiplication()?;
        
        while matches!(self.current_token, Token::Plus | Token::Minus) {
            let op = match self.current_token {
                Token::Plus => BinaryOp::Add,
                Token::Minus => BinaryOp::Subtract,
                _ => unreachable!(),
            };
            self.advance()?;
            let right = self.parse_multiplication()?;
            left = Expr::Binary {
                left: Box::new(left),
                operator: op,
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
    
    /// Evaluates an expression AST to a numeric result.
    pub fn evaluate(&self, expr: &Expr) -> Result<f64, String> {
        match expr {
            Expr::Number(value) => Ok(*value),
            
            Expr::CellRef(cell_ref) => {
                let (row, col) = super::models::Spreadsheet::parse_cell_reference(cell_ref)
                    .ok_or_else(|| format!("Invalid cell reference: {}", cell_ref))?;
                Ok(self.spreadsheet.get_cell_value_for_formula(row, col))
            }
            
            Expr::Range(start_cell, end_cell) => {
                // This shouldn't be called directly - ranges are handled by functions
                Err(format!("Range {}:{} cannot be evaluated directly", start_cell, end_cell))
            }
            
            Expr::Binary { left, operator, right } => {
                let left_val = self.evaluate(left)?;
                let right_val = self.evaluate(right)?;
                
                match operator {
                    BinaryOp::Add => Ok(left_val + right_val),
                    BinaryOp::Subtract => Ok(left_val - right_val),
                    BinaryOp::Multiply => Ok(left_val * right_val),
                    BinaryOp::Divide => {
                        if right_val == 0.0 {
                            Err("Division by zero".to_string())
                        } else {
                            Ok(left_val / right_val)
                        }
                    }
                    BinaryOp::Modulo => {
                        if right_val == 0.0 {
                            Err("Modulo by zero".to_string())
                        } else {
                            Ok(left_val % right_val)
                        }
                    }
                    BinaryOp::Power => Ok(left_val.powf(right_val)),
                    BinaryOp::Less => Ok(if left_val < right_val { 1.0 } else { 0.0 }),
                    BinaryOp::LessEqual => Ok(if left_val <= right_val { 1.0 } else { 0.0 }),
                    BinaryOp::Greater => Ok(if left_val > right_val { 1.0 } else { 0.0 }),
                    BinaryOp::GreaterEqual => Ok(if left_val >= right_val { 1.0 } else { 0.0 }),
                    BinaryOp::Equal => Ok(if (left_val - right_val).abs() < f64::EPSILON { 1.0 } else { 0.0 }),
                    BinaryOp::NotEqual => Ok(if (left_val - right_val).abs() >= f64::EPSILON { 1.0 } else { 0.0 }),
                }
            }
            
            Expr::Unary { operator, operand } => {
                let operand_val = self.evaluate(operand)?;
                
                match operator {
                    UnaryOp::Plus => Ok(operand_val),
                    UnaryOp::Minus => Ok(-operand_val),
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
    fn evaluate_function_args(&self, args: &[Expr]) -> Result<Vec<f64>, String> {
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
                            values.push(self.spreadsheet.get_cell_value_for_formula(row, col));
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
        sheet.set_cell(0, 0, CellData { value: "10".to_string(), formula: None });
        sheet.set_cell(0, 1, CellData { value: "20".to_string(), formula: None });
        sheet.set_cell(0, 2, CellData { value: "30".to_string(), formula: None });
        sheet.set_cell(1, 0, CellData { value: "5".to_string(), formula: None });
        sheet.set_cell(1, 1, CellData { value: "15".to_string(), formula: None });
        sheet.set_cell(1, 2, CellData { value: "25".to_string(), formula: None });
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
        assert_eq!(evaluator.evaluate(&expr).unwrap(), 42.5);
    }

    #[test]
    fn test_expression_evaluator_cell_refs() {
        let sheet = create_test_spreadsheet();
        let registry = FunctionRegistry::new();
        let evaluator = ExpressionEvaluator::new(&sheet, &registry);
        
        let expr = Expr::CellRef("A1".to_string());
        assert_eq!(evaluator.evaluate(&expr).unwrap(), 10.0);
        
        let expr = Expr::CellRef("B1".to_string());
        assert_eq!(evaluator.evaluate(&expr).unwrap(), 20.0);
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
        assert_eq!(evaluator.evaluate(&expr).unwrap(), 15.0);
        
        let expr = Expr::Binary {
            left: Box::new(Expr::CellRef("A1".to_string())),
            operator: BinaryOp::Multiply,
            right: Box::new(Expr::CellRef("B1".to_string())),
        };
        assert_eq!(evaluator.evaluate(&expr).unwrap(), 200.0); // 10 * 20
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
        assert_eq!(evaluator.evaluate(&expr).unwrap(), -5.0);
        
        // NOT is now a function, not a unary operator
        let expr = Expr::FunctionCall {
            name: "NOT".to_string(),
            args: vec![Expr::Number(0.0)],
        };
        assert_eq!(evaluator.evaluate(&expr).unwrap(), 1.0);
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
        assert_eq!(evaluator.evaluate(&expr).unwrap(), 30.0); // 10 + 20
        
        let expr = Expr::FunctionCall {
            name: "IF".to_string(),
            args: vec![
                Expr::Number(1.0),
                Expr::Number(100.0),
                Expr::Number(200.0),
            ],
        };
        assert_eq!(evaluator.evaluate(&expr).unwrap(), 100.0);
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
        assert_eq!(evaluator.evaluate(&expr).unwrap(), 30.0); // 10 + 20
        
        let expr = Expr::FunctionCall {
            name: "AVERAGE".to_string(),
            args: vec![Expr::Range("A1".to_string(), "C1".to_string())],
        };
        assert_eq!(evaluator.evaluate(&expr).unwrap(), 20.0); // (10 + 20 + 30) / 3
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
                Ok(args[0] * 2.0)
            } else {
                Err("DOUBLE requires exactly 1 argument".to_string())
            }
        });
        
        assert!(registry.get_function("DOUBLE").is_some());
        let double_func = registry.get_function("DOUBLE").unwrap();
        assert_eq!(double_func(&[5.0]).unwrap(), 10.0);
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
        assert_eq!(result, 30.0);
        
        // Test arithmetic with functions: SUM(A1:B1) + 5
        let mut parser = Parser::new("SUM(A1:B1) + 5").unwrap();
        let ast = parser.parse().unwrap();
        let result = evaluator.evaluate(&ast).unwrap();
        assert_eq!(result, 35.0); // (10 + 20) + 5
        
        // Test power operations: 2 ** 3 + 1
        let mut parser = Parser::new("2 ** 3 + 1").unwrap();
        let ast = parser.parse().unwrap();
        let result = evaluator.evaluate(&ast).unwrap();
        assert_eq!(result, 9.0); // 8 + 1
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
    fn test_lexer_error_handling() {
        let mut lexer = Lexer::new("@#$");
        assert!(lexer.next_token().is_err());
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
    }
}