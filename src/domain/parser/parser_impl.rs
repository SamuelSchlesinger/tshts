//! Submodule of `parser` — see parser/mod.rs.

#![allow(unused_imports)]
use super::*;

/// Hard cap on recursive parser descents. Prevents stack overflow on
/// adversarial input like `=(((((...)))))` thousands deep. Excel's nesting
/// limit is 64 — we go higher for tooling but still bounded.
pub(crate) const MAX_PARSE_DEPTH: u32 = 256;

pub struct Parser {
    lexer: Lexer,
    current_token: Token,
    depth: u32,
}

impl Parser {
    /// Creates a new parser for the given expression.
    pub fn new(input: &str) -> Result<Self, String> {
        let mut lexer = Lexer::new(input);
        let current_token = lexer.next_token()?;

        Ok(Self {
            lexer,
            current_token,
            depth: 0,
        })
    }

    /// Bump the recursion counter on entry to a recursive rule. Caller must
    /// pair with `ascend()` on every return path; a `?` after `descend` keeps
    /// the parser interruptible at the deepest level the cap allows.
    fn descend(&mut self) -> Result<(), String> {
        if self.depth >= MAX_PARSE_DEPTH {
            return Err("Formula nests too deeply".to_string());
        }
        self.depth += 1;
        Ok(())
    }

    fn ascend(&mut self) {
        self.depth = self.depth.saturating_sub(1);
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
        self.depth = 0;
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
        let mut left = self.parse_concatenation()?;

        while matches!(self.current_token, Token::Less | Token::LessEqual | Token::Greater | Token::GreaterEqual) {
            let op = match self.current_token {
                Token::Less => BinaryOp::Less,
                Token::LessEqual => BinaryOp::LessEqual,
                Token::Greater => BinaryOp::Greater,
                Token::GreaterEqual => BinaryOp::GreaterEqual,
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

    /// Parses string concatenation expressions. Sits between comparison and
    /// addition: Excel binds `&` looser than `+`/`-` but tighter than
    /// comparisons, so `="a" & 1+2` is `"a3"` and `="a" & 1 = "a1"` is TRUE.
    fn parse_concatenation(&mut self) -> Result<Expr, String> {
        let mut left = self.parse_addition()?;

        while matches!(self.current_token, Token::Ampersand) {
            self.advance()?;
            let right = self.parse_addition()?;
            left = Expr::Binary {
                left: Box::new(left),
                operator: BinaryOp::Concatenate,
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
    
    /// Parses power expressions (left-associative, matching Excel/Sheets).
    fn parse_power(&mut self) -> Result<Expr, String> {
        let mut left = self.parse_unary()?;

        while matches!(self.current_token, Token::Power | Token::PowerAlt) {
            self.advance()?;
            let right = self.parse_unary()?;
            left = Expr::Binary {
                left: Box::new(left),
                operator: BinaryOp::Power,
                right: Box::new(right),
            };
        }
        Ok(left)
    }
    
    /// Parses unary expressions.
    fn parse_unary(&mut self) -> Result<Expr, String> {
        self.descend()?;
        let result = match self.current_token {
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
        };
        self.ascend();
        result
    }
    
    /// Builds an `Expr::Let` from a parsed argument list.
    /// LET(name1, value1, name2, value2, ..., body) — at least 3 args; odd count.
    fn build_let(mut args: Vec<Expr>) -> Result<Expr, String> {
        if args.len() < 3 || args.len().is_multiple_of(2) {
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

            // Source-level error literal (`#REF!`, `#N/A`, etc.). Becomes
            // Value::Error at eval time; error propagation in Binary,
            // FunctionCall, etc. cascades it through containing expressions.
            Token::ErrorLit(kind) => {
                let kind = *kind;
                self.advance()?;
                Ok(Expr::ErrorLit(kind))
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


#[cfg(test)]
mod tests {
    use super::*;
    use crate::domain::{Spreadsheet, CellData, FormulaEvaluator};
    use crate::domain::parser::*;



    #[test]
    fn test_parser_numbers() {
        let mut parser = Parser::new("42").unwrap();
        let expr = parser.parse().unwrap();
        assert_eq!(expr, Expr::Number(42.0));
        
        let mut parser = Parser::new("3.14").unwrap();
        let expr = parser.parse().unwrap();
        #[allow(clippy::approx_constant)]
        let three_fourteen = 3.14;
        assert_eq!(expr, Expr::Number(three_fourteen));
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
                assert!(matches!(left.as_ref(), Expr::CellRef(s) if s == "A1"));
                assert_eq!(operator, BinaryOp::Multiply);
                assert!(matches!(right.as_ref(), Expr::CellRef(s) if s == "B1"));
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
    fn test_parser_power_left_associative() {
        // 2 ** 3 ** 2 parses as (2 ** 3) ** 2 — Excel/Sheets convention.
        let mut parser = Parser::new("2 ** 3 ** 2").unwrap();
        let expr = parser.parse().unwrap();
        match expr {
            Expr::Binary { left, operator: BinaryOp::Power, right } => {
                assert!(matches!(right.as_ref(), &Expr::Number(2.0)));
                match left.as_ref() {
                    Expr::Binary { left: pow_left, operator: BinaryOp::Power, right: pow_right } => {
                        assert!(matches!(pow_left.as_ref(), &Expr::Number(2.0)));
                        assert!(matches!(pow_right.as_ref(), &Expr::Number(3.0)));
                    }
                    _ => panic!("Expected power as left operand"),
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
                assert!(matches!(left.as_ref(), Expr::CellRef(s) if s == "A1"));
                assert_eq!(operator, BinaryOp::Greater);
                assert!(matches!(right.as_ref(), Expr::CellRef(s) if s == "B1"));
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
                        assert!(matches!(comp_left.as_ref(), Expr::CellRef(s) if s == "A1"));
                        assert!(matches!(comp_right.as_ref(), &Expr::Number(5.0)));
                    }
                    _ => panic!("Expected comparison in first argument"),
                }
                
                // Second argument should be B1 < 10
                match &args[1] {
                    Expr::Binary { left: comp_left, operator: BinaryOp::Less, right: comp_right } => {
                        assert!(matches!(comp_left.as_ref(), Expr::CellRef(s) if s == "B1"));
                        assert!(matches!(comp_right.as_ref(), &Expr::Number(10.0)));
                    }
                    _ => panic!("Expected comparison in second argument"),
                }
            }
            _ => panic!("Expected function call"),
        }
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
                assert!(matches!(left.as_ref(), Expr::String(s) if s == "Hello"));
                assert_eq!(operator, BinaryOp::Concatenate);
                assert!(matches!(right.as_ref(), Expr::String(s) if s == "World"));
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
                assert!(matches!(left.as_ref(), Expr::String(s) if s == "Number: "));
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
                        assert!(matches!(inner_right.as_ref(), Expr::String(s) if s == " - "));
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

}
