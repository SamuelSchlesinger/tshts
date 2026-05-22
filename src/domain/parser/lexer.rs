//! Submodule of `parser` — see parser/mod.rs.

#![allow(unused_imports)]
use super::*;

pub struct Lexer {
    input: Vec<char>,
    position: usize,
    current_char: Option<char>,
}

impl Lexer {
    /// Creates a new lexer for the given input string.
    pub fn new(input: &str) -> Self {
        let chars: Vec<char> = input.chars().collect();
        let current_char = chars.first().copied();
        
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
    /// Match the longest known Excel error literal starting at the current
    /// `#`. Longest-first ordering matters: `#DIV/0!` must be tried before
    /// any potential shorter prefix collisions. Returns Err if `#` is
    /// followed by something we don't recognize so callers see a clean
    /// parse error instead of mis-tokenizing.
    fn read_error_literal(&mut self) -> Result<Token, String> {
        // Build the remaining slice from the current position so we can
        // do simple prefix matching. The input is Vec<char>, so we collect
        // a temporary String — error literals are short (max 7 chars) so
        // this stays cheap.
        let rest: String = self.input[self.position..].iter().collect();
        // (literal_str, kind). Order longest-first so e.g. `#NULL!` can't
        // be mis-matched as a hypothetical prefix.
        const LITERALS: &[(&str, ErrorKind)] = &[
            ("#DIV/0!", ErrorKind::Div0),
            ("#VALUE!", ErrorKind::Value),
            ("#SPILL!", ErrorKind::Spill),
            ("#NAME?", ErrorKind::Name),
            ("#NULL!", ErrorKind::Null),
            ("#REF!", ErrorKind::Ref),
            ("#NUM!", ErrorKind::Num),
            ("#N/A", ErrorKind::NA),
        ];
        for (lit, kind) in LITERALS {
            if rest.starts_with(lit) {
                for _ in 0..lit.chars().count() {
                    self.advance();
                }
                return Ok(Token::ErrorLit(*kind));
            }
        }
        Err(format!("Unrecognized error literal after '#': {}", rest))
    }

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

                // Excel-style error literals: `#REF!`, `#N/A`, `#DIV/0!`, etc.
                // These are first-class expressions — they evaluate to
                // Value::Error and propagate through arithmetic the same way
                // a #DIV/0! produced at runtime would. The lexer matches the
                // longest known literal greedily.
                '#' => self.read_error_literal(),

                _ => Err(format!("Unexpected character: '{}'", ch)),
            }
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::domain::{Spreadsheet, CellData, FormulaEvaluator};
    use crate::domain::parser::*;

    #[test]
    fn test_lexer_error_literals() {
        // All Excel error literals tokenize cleanly. Longest-match-first
        // matters for `#DIV/0!` and `#N/A` (which contain `/`).
        let mut lexer = Lexer::new("#REF! #N/A #DIV/0! #VALUE! #NAME? #NUM! #NULL! #SPILL!");
        assert_eq!(lexer.next_token().unwrap(), Token::ErrorLit(ErrorKind::Ref));
        assert_eq!(lexer.next_token().unwrap(), Token::ErrorLit(ErrorKind::NA));
        assert_eq!(lexer.next_token().unwrap(), Token::ErrorLit(ErrorKind::Div0));
        assert_eq!(lexer.next_token().unwrap(), Token::ErrorLit(ErrorKind::Value));
        assert_eq!(lexer.next_token().unwrap(), Token::ErrorLit(ErrorKind::Name));
        assert_eq!(lexer.next_token().unwrap(), Token::ErrorLit(ErrorKind::Num));
        assert_eq!(lexer.next_token().unwrap(), Token::ErrorLit(ErrorKind::Null));
        assert_eq!(lexer.next_token().unwrap(), Token::ErrorLit(ErrorKind::Spill));
        assert_eq!(lexer.next_token().unwrap(), Token::Eof);
    }

    #[test]
    fn test_lexer_error_literal_in_expression() {
        // Errors compose: `#REF!+1` lexes as three tokens, not one.
        let mut lexer = Lexer::new("#REF!+1");
        assert_eq!(lexer.next_token().unwrap(), Token::ErrorLit(ErrorKind::Ref));
        assert_eq!(lexer.next_token().unwrap(), Token::Plus);
        assert_eq!(lexer.next_token().unwrap(), Token::Number(1.0));
    }

    #[test]
    fn test_lexer_unknown_hash_literal_errors() {
        let mut lexer = Lexer::new("#NOTREAL!");
        assert!(lexer.next_token().is_err());
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

}
