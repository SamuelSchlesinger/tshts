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
