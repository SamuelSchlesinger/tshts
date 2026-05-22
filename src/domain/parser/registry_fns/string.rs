//! `string` builtin functions.
//!
//! Each function registers itself on the provided `FunctionRegistry` via
//! `reg.register_function(NAME, |args| ...)`. The category-level `register`
//! function is called from `registry::FunctionRegistry::register_builtin_functions`.

#![allow(unused_imports)]
use crate::domain::parser::{FunctionRegistry, Value, ErrorKind, flatten_args, shape_of, broadcast_binary, criteria_matches, add_commas, glob_match, date_to_serial, serial_to_date, parse_iso_date, days_in_month, today_serial, now_serial};

/// Register all `string` builtin functions on `reg`.
pub(in crate::domain::parser) fn register(reg: &mut FunctionRegistry) {
        reg.register_function("CONCAT", |args| {
            let flat = flatten_args(args);
            let result = flat.iter().map(|v| v.to_string()).collect::<String>();
            Ok(Value::String(result))
        });
        reg.register_function("LEN", |args| {
            if args.len() != 1 {
                Err("LEN requires exactly 1 argument".to_string())
            } else {
                let len = args[0].to_string().chars().count() as f64;
                Ok(Value::Number(len))
            }
        });
        reg.register_function("LEFT", |args| {
            if args.len() != 2 {
                Err("LEFT requires exactly 2 arguments".to_string())
            } else {
                let text = args[0].to_string();
                let num_chars = args[1].to_number() as usize;
                let result = text.chars().take(num_chars).collect::<String>();
                Ok(Value::String(result))
            }
        });
        reg.register_function("RIGHT", |args| {
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
        reg.register_function("MID", |args| {
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
        reg.register_function("FIND", |args| {
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
        reg.register_function("UPPER", |args| {
            if args.len() != 1 {
                Err("UPPER requires exactly 1 argument".to_string())
            } else {
                Ok(Value::String(args[0].to_string().to_uppercase()))
            }
        });
        reg.register_function("LOWER", |args| {
            if args.len() != 1 {
                Err("LOWER requires exactly 1 argument".to_string())
            } else {
                Ok(Value::String(args[0].to_string().to_lowercase()))
            }
        });
        reg.register_function("TRIM", |args| {
            if args.len() != 1 {
                Err("TRIM requires exactly 1 argument".to_string())
            } else {
                Ok(Value::String(args[0].to_string().trim().to_string()))
            }
        });
        reg.register_function("SUBSTITUTE", |args| {
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
        reg.register_function("REPLACE", |args| {
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
        reg.register_function("REPT", |args| {
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
        reg.register_function("EXACT", |args| {
            if args.len() != 2 {
                Err("EXACT requires exactly 2 arguments".to_string())
            } else {
                let a = args[0].to_string();
                let b = args[1].to_string();
                Ok(Value::Number(if a == b { 1.0 } else { 0.0 }))
            }
        });
        reg.register_function("PROPER", |args| {
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
        reg.register_function("CLEAN", |args| {
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
        reg.register_function("CHAR", |args| {
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
        reg.register_function("CODE", |args| {
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
        reg.register_function("TEXT", |args| {
            if args.is_empty() || args.len() > 2 {
                Err("TEXT requires 1 or 2 arguments".to_string())
            } else {
                Ok(Value::String(args[0].to_string()))
            }
        });
        reg.register_function("VALUE", |args| {
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
        reg.register_function("NUMBERVALUE", |args| {
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
        reg.register_function("TEXTJOIN", |args| {
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
        reg.register_function("SEARCH", |args| {
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
        reg.register_function("TEXTBEFORE", |args| {
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
        reg.register_function("TEXTAFTER", |args| {
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
        reg.register_function("REGEXMATCH", |args| {
            if args.len() != 2 {
                return Err("REGEXMATCH requires 2 arguments".to_string());
            }
            let re = regex::Regex::new(&args[1].to_string())
                .map_err(|e| format!("REGEXMATCH: bad pattern: {}", e))?;
            Ok(Value::Bool(re.is_match(&args[0].to_string())))
        });
        reg.register_function("REGEXEXTRACT", |args| {
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
        reg.register_function("REGEXREPLACE", |args| {
            if args.len() != 3 {
                return Err("REGEXREPLACE requires 3 arguments".to_string());
            }
            let re = regex::Regex::new(&args[1].to_string())
                .map_err(|e| format!("REGEXREPLACE: bad pattern: {}", e))?;
            let replacement = args[2].to_string();
            Ok(Value::String(re.replace_all(&args[0].to_string(), replacement.as_str()).into_owned()))
        });
        reg.register_function("UNICHAR", |args| {
            let n = args[0].to_number() as u32;
            match char::from_u32(n) {
                Some(c) => Ok(Value::String(c.to_string())),
                None => Ok(Value::Error(ErrorKind::Value)),
            }
        });
        reg.register_function("UNICODE", |args| {
            let s = args[0].to_string();
            match s.chars().next() {
                Some(c) => Ok(Value::Number(c as u32 as f64)),
                None => Ok(Value::Error(ErrorKind::Value)),
            }
        });
        reg.register_function("DOLLAR", |args| {
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
        reg.register_function("FIXED", |args| {
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
        reg.register_function("ARRAYTOTEXT", |args| {
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
}
