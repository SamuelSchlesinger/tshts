//! `lookup` builtin functions.
//!
//! Each function registers itself on the provided `FunctionRegistry` via
//! `reg.register_function(NAME, |args| ...)`. The category-level `register`
//! function is called from `registry::FunctionRegistry::register_builtin_functions`.

#![allow(unused_imports)]
use crate::domain::parser::{FunctionRegistry, Value, ErrorKind, flatten_args, shape_of, broadcast_binary, criteria_matches, add_commas, glob_match, date_to_serial, serial_to_date, parse_iso_date, days_in_month, today_serial, now_serial};

/// Register all `lookup` builtin functions on `reg`.
pub(in crate::domain::parser) fn register(reg: &mut FunctionRegistry) {
        reg.register_function("SUMIF", |args| {
            if args.len() != 2 && args.len() != 3 {
                return Err("SUMIF requires 2 or 3 arguments".to_string());
            }
            let range = args[0].flatten();
            let criteria = args[1].to_string();
            let sum_range = if args.len() == 3 { args[2].flatten() } else { range.clone() };
            let mut sum = 0.0;
            for (i, v) in range.iter().enumerate() {
                if criteria_matches(v, &criteria)
                    && let Some(target) = sum_range.get(i) {
                        sum += target.to_number();
                    }
            }
            Ok(Value::Number(sum))
        });
        reg.register_function("COUNTIF", |args| {
            if args.len() != 2 {
                return Err("COUNTIF requires 2 arguments".to_string());
            }
            let range = args[0].flatten();
            let criteria = args[1].to_string();
            let count = range.iter().filter(|v| criteria_matches(v, &criteria)).count();
            Ok(Value::Number(count as f64))
        });
        reg.register_function("AVERAGEIF", |args| {
            if args.len() != 2 && args.len() != 3 {
                return Err("AVERAGEIF requires 2 or 3 arguments".to_string());
            }
            let range = args[0].flatten();
            let criteria = args[1].to_string();
            let sum_range = if args.len() == 3 { args[2].flatten() } else { range.clone() };
            let mut sum = 0.0;
            let mut count = 0usize;
            for (i, v) in range.iter().enumerate() {
                if criteria_matches(v, &criteria)
                    && let Some(target) = sum_range.get(i) {
                        sum += target.to_number();
                        count += 1;
                    }
            }
            if count == 0 {
                Err("AVERAGEIF: no matching values".to_string())
            } else {
                Ok(Value::Number(sum / count as f64))
            }
        });
        reg.register_function("VLOOKUP", |args| {
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
        reg.register_function("HLOOKUP", |args| {
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
        reg.register_function("XLOOKUP", |args| {
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
                        -1 if n <= t
                            && next_smaller
                                .map(|si| keys[si].to_number())
                                .map(|sv| n > sv)
                                .unwrap_or(true)
                            => {
                                next_smaller = Some(*i);
                            }
                        1 if n >= t
                            && next_larger
                                .map(|li| keys[li].to_number())
                                .map(|lv| n < lv)
                                .unwrap_or(true)
                            => {
                                next_larger = Some(*i);
                            }
                        _ => {}
                    }
                }
            }
            if let Some(i) = exact_hit {
                return Ok(values[i].clone());
            }
            if match_mode == -1
                && let Some(i) = next_smaller {
                    return Ok(values[i].clone());
                }
            if match_mode == 1
                && let Some(i) = next_larger {
                    return Ok(values[i].clone());
                }
            args.get(3)
                .cloned()
                .ok_or_else(|| "XLOOKUP: value not found".to_string())
        });
        reg.register_function("INDEX", |args| {
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
        reg.register_function("MATCH", |args| {
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
}
