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
                return Ok(Value::Error(ErrorKind::Value));
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
                return Ok(Value::Error(ErrorKind::Value));
            }
            let range = args[0].flatten();
            let criteria = args[1].to_string();
            let count = range.iter().filter(|v| criteria_matches(v, &criteria)).count();
            Ok(Value::Number(count as f64))
        });
        reg.register_function("AVERAGEIF", |args| {
            if args.len() != 2 && args.len() != 3 {
                return Ok(Value::Error(ErrorKind::Value));
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
                Ok(Value::Error(ErrorKind::Div0))
            } else {
                Ok(Value::Number(sum / count as f64))
            }
        });
        reg.register_function("VLOOKUP", |args| {
            if args.len() < 3 || args.len() > 4 {
                return Ok(Value::Error(ErrorKind::Value));
            }
            let lookup = args[0].to_string();
            let col_index = args[2].to_number() as usize;
            let exact = args.get(3).map(|v| !v.is_truthy()).unwrap_or(false);
            let (rows, cols, data) = shape_of(&args[1]);
            if col_index == 0 || col_index > cols {
                return Ok(Value::Error(ErrorKind::Ref));
            }
            // Approximate-match mode requires sorted keys (Excel semantics):
            // we stop at the first key strictly greater than the target.
            // If keys are unsorted, the result is undefined per Excel docs.
            let target_num = lookup.parse::<f64>().ok();
            let mut last_match: Option<usize> = None;
            for r in 0..rows {
                let key = &data[r * cols];
                if exact {
                    // Excel exact-match VLOOKUP is case-insensitive on text.
                    if key.to_string().eq_ignore_ascii_case(&lookup) {
                        return Ok(data[r * cols + col_index - 1].clone());
                    }
                } else if let Some(t) = target_num {
                    let k_num = match key {
                        Value::Number(n) => Some(*n),
                        Value::String(s) => s.trim().parse::<f64>().ok(),
                        Value::Bool(b) => Some(if *b { 1.0 } else { 0.0 }),
                        _ => None,
                    };
                    let Some(k) = k_num else { continue };
                    if k > t {
                        break;
                    }
                    last_match = Some(r);
                } else if key.to_string().to_lowercase() <= lookup.to_lowercase() {
                    last_match = Some(r);
                }
            }
            if let Some(r) = last_match {
                return Ok(data[r * cols + col_index - 1].clone());
            }
            Ok(Value::Error(ErrorKind::NA))
        });
        reg.register_function("HLOOKUP", |args| {
            if args.len() < 3 || args.len() > 4 {
                return Ok(Value::Error(ErrorKind::Value));
            }
            let lookup = args[0].to_string();
            let row_index = args[2].to_number() as usize;
            let exact = args.get(3).map(|v| !v.is_truthy()).unwrap_or(false);
            let (rows, cols, data) = shape_of(&args[1]);
            if row_index == 0 || row_index > rows {
                return Ok(Value::Error(ErrorKind::Ref));
            }
            let target_num = lookup.parse::<f64>().ok();
            let mut last_match: Option<usize> = None;
            for c in 0..cols {
                let key = &data[c];
                if exact {
                    if key.to_string().eq_ignore_ascii_case(&lookup) {
                        return Ok(data[(row_index - 1) * cols + c].clone());
                    }
                } else if let Some(t) = target_num {
                    let k_num = match key {
                        Value::Number(n) => Some(*n),
                        Value::String(s) => s.trim().parse::<f64>().ok(),
                        Value::Bool(b) => Some(if *b { 1.0 } else { 0.0 }),
                        _ => None,
                    };
                    let Some(k) = k_num else { continue };
                    if k > t {
                        break;
                    }
                    last_match = Some(c);
                } else if key.to_string().to_lowercase() <= lookup.to_lowercase() {
                    last_match = Some(c);
                }
            }
            if let Some(c) = last_match {
                return Ok(data[(row_index - 1) * cols + c].clone());
            }
            Ok(Value::Error(ErrorKind::NA))
        });
        reg.register_function("XLOOKUP", |args| {
            if args.len() < 3 || args.len() > 6 {
                return Ok(Value::Error(ErrorKind::Value));
            }
            let lookup = args[0].to_string();
            let keys = args[1].flatten();
            let values = args[2].flatten();
            if keys.len() != values.len() {
                return Ok(Value::Error(ErrorKind::Value));
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
            let needle_lc = lookup.to_lowercase();
            let mut exact_hit: Option<usize> = None;
            // For match_mode = -1, "largest value ≤ target": track the
            // candidate whose comparison to needle is closest to Equal from
            // below (largest Less, or Equal). For match_mode = 1, "smallest
            // value ≥ target": closest to Equal from above (smallest Greater).
            // We carry the candidate index together with a numeric key for
            // the "how close" comparison; for text we use lexicographic order.
            let key_for_compare = |k: &Value| -> Option<f64> {
                let k_num = match k {
                    Value::Number(n) => Some(*n),
                    Value::Bool(b) => Some(if *b { 1.0 } else { 0.0 }),
                    Value::String(s) => s.parse::<f64>().ok(),
                    _ => None,
                };
                if needle_num.is_some() {
                    k_num
                } else if matches!(k, Value::Number(_) | Value::Bool(_)) {
                    None // numeric key, text needle: skip
                } else {
                    // Text-vs-text: project lowercase string compare against
                    // the needle to a signed integer ordering distance.
                    let kl = k.to_string().to_lowercase();
                    Some(match kl.cmp(&needle_lc) {
                        std::cmp::Ordering::Less => -1.0,
                        std::cmp::Ordering::Equal => 0.0,
                        std::cmp::Ordering::Greater => 1.0,
                    })
                }
            };
            let target = needle_num.unwrap_or(0.0);
            let mut best_smaller: Option<(usize, f64)> = None;
            let mut best_larger: Option<(usize, f64)> = None;

            for i in &indices {
                let k = &keys[*i];
                let matched = match match_mode {
                    2 => glob_match(&k.to_string(), &lookup),
                    _ => k.to_string().eq_ignore_ascii_case(&lookup),
                };
                if matched {
                    exact_hit = Some(*i);
                    break;
                }
                let Some(kn) = key_for_compare(k) else { continue };
                if match_mode == -1 && kn <= target {
                    let is_better = best_smaller.map(|(_, bv)| kn > bv).unwrap_or(true);
                    if is_better {
                        best_smaller = Some((*i, kn));
                    }
                } else if match_mode == 1 && kn >= target {
                    let is_better = best_larger.map(|(_, bv)| kn < bv).unwrap_or(true);
                    if is_better {
                        best_larger = Some((*i, kn));
                    }
                }
            }
            let next_smaller = best_smaller.map(|(i, _)| i);
            let next_larger = best_larger.map(|(i, _)| i);
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
            Ok(args.get(3).cloned().unwrap_or(Value::Error(ErrorKind::NA)))
        });
        reg.register_function("INDEX", |args| {
            if args.len() < 2 || args.len() > 3 {
                return Ok(Value::Error(ErrorKind::Value));
            }
            let row = args[1].to_number() as usize;
            let col = args.get(2).map(|v| v.to_number() as usize).unwrap_or(1);
            if row == 0 || col == 0 {
                return Ok(Value::Error(ErrorKind::Ref));
            }
            let (rows, cols, data) = shape_of(&args[0]);
            if row > rows || col > cols {
                return Ok(Value::Error(ErrorKind::Ref));
            }
            Ok(data[(row - 1) * cols + (col - 1)].clone())
        });
        reg.register_function("MATCH", |args| {
            if args.len() < 2 || args.len() > 3 {
                return Ok(Value::Error(ErrorKind::Value));
            }
            let needle = args[0].to_string();
            let range = args[1].flatten();
            let match_type = args.get(2).map(|v| v.to_number() as i64).unwrap_or(1);
            if match_type == 0 {
                let has_wild = needle.contains('*') || needle.contains('?');
                for (i, v) in range.iter().enumerate() {
                    let s = v.to_string();
                    let matched = if has_wild {
                        glob_match(&s, &needle)
                    } else {
                        s.eq_ignore_ascii_case(&needle)
                    };
                    if matched {
                        return Ok(Value::Number((i + 1) as f64));
                    }
                }
                Ok(Value::Error(ErrorKind::NA))
            } else {
                // match_type 1: ascending, find largest value <= needle.
                // match_type -1: descending, find smallest value >= needle.
                // Excel accepts text or numeric ranges; previously
                // `needle.parse::<f64>().unwrap_or(0.0)` collapsed text
                // needles to 0 and made every text cell match cell == 0.
                let needle_num = needle.parse::<f64>().ok();
                let needle_lc = needle.to_lowercase();
                let mut last_idx: Option<usize> = None;
                for (i, v) in range.iter().enumerate() {
                    let v_num = match v {
                        Value::Number(n) => Some(*n),
                        Value::Bool(b) => Some(if *b { 1.0 } else { 0.0 }),
                        Value::String(s) => s.parse::<f64>().ok(),
                        _ => None,
                    };
                    let ord = match (needle_num, v_num) {
                        (Some(t), Some(n)) => n.partial_cmp(&t),
                        (None, None) => Some(v.to_string().to_lowercase().cmp(&needle_lc)),
                        // Mixed types are skipped (Excel ignores them).
                        _ => None,
                    };
                    if let Some(o) = ord {
                        let hit = if match_type > 0 {
                            !o.is_gt() // n <= target
                        } else {
                            !o.is_lt() // n >= target
                        };
                        if hit {
                            last_idx = Some(i);
                        }
                    }
                }
                Ok(last_idx
                    .map(|i| Value::Number((i + 1) as f64))
                    .unwrap_or(Value::Error(ErrorKind::NA)))
            }
        });
}
