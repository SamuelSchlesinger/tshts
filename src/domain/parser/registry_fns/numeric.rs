//! `numeric` builtin functions.
//!
//! Each function registers itself on the provided `FunctionRegistry` via
//! `reg.register_function(NAME, |args| ...)`. The category-level `register`
//! function is called from `registry::FunctionRegistry::register_builtin_functions`.

#![allow(unused_imports)]
use crate::domain::parser::{FunctionRegistry, Value, ErrorKind, flatten_args, shape_of, broadcast_binary, criteria_matches, add_commas, glob_match, date_to_serial, serial_to_date, parse_iso_date, days_in_month, today_serial, now_serial};

/// Register all `numeric` builtin functions on `reg`.
pub(in crate::domain::parser) fn register(reg: &mut FunctionRegistry) {
        reg.register_function("SUM", |args| {
            let flat = flatten_args(args);
            // Excel: SUM propagates the first error it encounters.
            if let Some(e) = flat.iter().find_map(|v| v.first_error()) {
                return Ok(Value::Error(e));
            }
            let sum: f64 = flat.iter().map(|v| v.to_number()).sum();
            Ok(Value::Number(sum))
        });
        reg.register_function("AVERAGE", |args| {
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
        reg.register_function("MIN", |args| {
            let flat = flatten_args(args);
            if let Some(e) = flat.iter().find_map(|v| v.first_error()) {
                return Ok(Value::Error(e));
            }
            flat.iter().map(|v| v.to_number()).fold(None, |acc: Option<f64>, x| {
                Some(acc.map_or(x, |a| a.min(x)))
            }).map(Value::Number).ok_or_else(|| "MIN requires at least one argument".to_string())
        });
        reg.register_function("MAX", |args| {
            let flat = flatten_args(args);
            if let Some(e) = flat.iter().find_map(|v| v.first_error()) {
                return Ok(Value::Error(e));
            }
            flat.iter().map(|v| v.to_number()).fold(None, |acc: Option<f64>, x| {
                Some(acc.map_or(x, |a| a.max(x)))
            }).map(Value::Number).ok_or_else(|| "MAX requires at least one argument".to_string())
        });
        reg.register_function("ABS", |args| {
            if args.len() != 1 {
                Err("ABS requires exactly 1 argument".to_string())
            } else {
                Ok(Value::Number(args[0].to_number().abs()))
            }
        });
        reg.register_function("SQRT", |args| {
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
        reg.register_function("ROUND", |args| {
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
        reg.register_function("CEILING", |args| {
            if args.len() != 1 {
                Err("CEILING requires exactly 1 argument".to_string())
            } else {
                Ok(Value::Number(args[0].to_number().ceil()))
            }
        });
        reg.register_function("FLOOR", |args| {
            if args.len() != 1 {
                Err("FLOOR requires exactly 1 argument".to_string())
            } else {
                Ok(Value::Number(args[0].to_number().floor()))
            }
        });
        reg.register_function("INT", |args| {
            if args.len() != 1 {
                Err("INT requires exactly 1 argument".to_string())
            } else {
                // Excel `INT` is floor, not truncate (matters for negatives).
                Ok(Value::Number(args[0].to_number().floor()))
            }
        });
        reg.register_function("MOD", |args| {
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
        reg.register_function("LOG", |args| {
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
        reg.register_function("LN", |args| {
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
        reg.register_function("EXP", |args| {
            if args.len() != 1 {
                Err("EXP requires exactly 1 argument".to_string())
            } else {
                Ok(Value::Number(args[0].to_number().exp()))
            }
        });
        reg.register_function("PI", |args| {
            if !args.is_empty() {
                Err("PI takes no arguments".to_string())
            } else {
                Ok(Value::Number(std::f64::consts::PI))
            }
        });
        reg.register_function("RAND", |_args| {
            use std::time::SystemTime;
            let nanos = SystemTime::now()
                .duration_since(SystemTime::UNIX_EPOCH)
                .unwrap_or_default()
                .subsec_nanos();
            Ok(Value::Number(nanos as f64 / 1_000_000_000.0))
        });
        reg.register_function("RANDBETWEEN", |args| {
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
        reg.register_function("SIGN", |args| {
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
        reg.register_function("POWER", |args| {
            if args.len() != 2 {
                Err("POWER requires exactly 2 arguments".to_string())
            } else {
                Ok(Value::Number(args[0].to_number().powf(args[1].to_number())))
            }
        });
        reg.register_function("MEDIAN", |args| {
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
        reg.register_function("STDEV.S", |args| {
            let nums = collect_numbers(args);
            if nums.len() < 2 {
                return Err("STDEV.S requires at least 2 values".to_string());
            }
            let mean: f64 = nums.iter().sum::<f64>() / nums.len() as f64;
            let var: f64 = nums.iter().map(|n| (n - mean).powi(2)).sum::<f64>()
                / (nums.len() - 1) as f64;
            Ok(Value::Number(var.sqrt()))
        });
        reg.register_function("STDEV", |args| {
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
        reg.register_function("STDEV.P", |args| {
            let nums = collect_numbers(args);
            if nums.is_empty() {
                return Err("STDEV.P: empty".to_string());
            }
            let mean: f64 = nums.iter().sum::<f64>() / nums.len() as f64;
            let var: f64 = nums.iter().map(|n| (n - mean).powi(2)).sum::<f64>()
                / nums.len() as f64;
            Ok(Value::Number(var.sqrt()))
        });
        reg.register_function("VAR.S", |args| {
            let nums = collect_numbers(args);
            if nums.len() < 2 {
                return Err("VAR.S requires at least 2 values".to_string());
            }
            let mean: f64 = nums.iter().sum::<f64>() / nums.len() as f64;
            let var: f64 = nums.iter().map(|n| (n - mean).powi(2)).sum::<f64>()
                / (nums.len() - 1) as f64;
            Ok(Value::Number(var))
        });
        reg.register_function("VAR.P", |args| {
            let nums = collect_numbers(args);
            if nums.is_empty() {
                return Err("VAR.P: empty".to_string());
            }
            let mean: f64 = nums.iter().sum::<f64>() / nums.len() as f64;
            let var: f64 = nums.iter().map(|n| (n - mean).powi(2)).sum::<f64>()
                / nums.len() as f64;
            Ok(Value::Number(var))
        });
        reg.register_function("LARGE", |args| {
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
        reg.register_function("SMALL", |args| {
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
        reg.register_function("RANK.EQ", |args| {
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
        reg.register_function("PERCENTILE.INC", |args| {
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
        reg.register_function("CORREL", |args| {
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
        reg.register_function("TRUNC", |args| {
            if args.is_empty() || args.len() > 2 {
                return Err("TRUNC requires 1 or 2 arguments".to_string());
            }
            let n = args[0].to_number();
            let digits = args.get(1).map(|v| v.to_number() as i32).unwrap_or(0);
            let scale = 10f64.powi(digits);
            Ok(Value::Number((n * scale).trunc() / scale))
        });
        reg.register_function("ATAN", |args| Ok(Value::Number(args[0].to_number().atan())));
        reg.register_function("ASIN", |args| Ok(Value::Number(args[0].to_number().asin())));
        reg.register_function("ACOS", |args| Ok(Value::Number(args[0].to_number().acos())));
        reg.register_function("SINH", |args| Ok(Value::Number(args[0].to_number().sinh())));
        reg.register_function("COSH", |args| Ok(Value::Number(args[0].to_number().cosh())));
        reg.register_function("TANH", |args| Ok(Value::Number(args[0].to_number().tanh())));
        reg.register_function("SIN", |args| Ok(Value::Number(args[0].to_number().sin())));
        reg.register_function("COS", |args| Ok(Value::Number(args[0].to_number().cos())));
        reg.register_function("TAN", |args| Ok(Value::Number(args[0].to_number().tan())));
        reg.register_function("DEGREES", |args| {
            Ok(Value::Number(args[0].to_number().to_degrees()))
        });
        reg.register_function("RADIANS", |args| {
            Ok(Value::Number(args[0].to_number().to_radians()))
        });
        reg.register_function("FACT", |args| {
            let n = args[0].to_number() as u64;
            let mut r: f64 = 1.0;
            for i in 2..=n {
                r *= i as f64;
            }
            Ok(Value::Number(r))
        });
        reg.register_function("COMBIN", |args| {
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
        reg.register_function("GCD", |args| {
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
        reg.register_function("LCM", |args| {
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
        reg.register_function("ROUNDUP", |args| {
            if args.len() != 2 {
                return Err("ROUNDUP requires 2 arguments".to_string());
            }
            let n = args[0].to_number();
            let digits = args[1].to_number() as i32;
            let scale = 10f64.powi(digits);
            let v = (n * scale).abs().ceil() * n.signum() / scale;
            Ok(Value::Number(v))
        });
        reg.register_function("ROUNDDOWN", |args| {
            if args.len() != 2 {
                return Err("ROUNDDOWN requires 2 arguments".to_string());
            }
            let n = args[0].to_number();
            let digits = args[1].to_number() as i32;
            let scale = 10f64.powi(digits);
            let v = (n * scale).abs().floor() * n.signum() / scale;
            Ok(Value::Number(v))
        });
        reg.register_function("MROUND", |args| {
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
        reg.register_function("EVEN", |args| {
            let n = args[0].to_number();
            let v = if n >= 0.0 {
                (n / 2.0).ceil() * 2.0
            } else {
                (n / 2.0).floor() * 2.0
            };
            Ok(Value::Number(v))
        });
        reg.register_function("ODD", |args| {
            let n = args[0].to_number();
            let v = if n >= 0.0 {
                let c = ((n + 1.0) / 2.0).ceil() * 2.0 - 1.0;
                if c == -1.0 { 1.0 } else { c }
            } else {
                ((n - 1.0) / 2.0).floor() * 2.0 + 1.0
            };
            Ok(Value::Number(v))
        });
        reg.register_function("FREQUENCY", |args| {
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
}

/// Common helper for statistical functions: flattens args and parses each
/// value as f64 (numbers as-is, strings via parse).
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
