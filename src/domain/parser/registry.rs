//! Submodule of `parser` — see parser/mod.rs.

#![allow(unused_imports)]
use super::*;

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
