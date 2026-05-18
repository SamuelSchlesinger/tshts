//! Submodule of `parser` — see parser/mod.rs.

#![allow(unused_imports)]
use super::*;

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
