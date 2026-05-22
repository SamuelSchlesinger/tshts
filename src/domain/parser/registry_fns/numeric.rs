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
                Ok(Value::Error(ErrorKind::Value))
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
                    Ok(Value::Error(ErrorKind::Div0))
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
                Ok(Value::Error(ErrorKind::Value))
            } else {
                Ok(Value::Number(args[0].to_number().abs()))
            }
        });
        reg.register_function("SQRT", |args| {
            if args.len() != 1 {
                Ok(Value::Error(ErrorKind::Value))
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
                _ => Ok(Value::Error(ErrorKind::Value)),
            }
        });
        reg.register_function("CEILING", |args| {
            if args.len() != 1 {
                Ok(Value::Error(ErrorKind::Value))
            } else {
                Ok(Value::Number(args[0].to_number().ceil()))
            }
        });
        reg.register_function("FLOOR", |args| {
            if args.len() != 1 {
                Ok(Value::Error(ErrorKind::Value))
            } else {
                Ok(Value::Number(args[0].to_number().floor()))
            }
        });
        reg.register_function("INT", |args| {
            if args.len() != 1 {
                Ok(Value::Error(ErrorKind::Value))
            } else {
                // Excel `INT` is floor, not truncate (matters for negatives).
                Ok(Value::Number(args[0].to_number().floor()))
            }
        });
        reg.register_function("MOD", |args| {
            if args.len() != 2 {
                Ok(Value::Error(ErrorKind::Value))
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
                _ => Ok(Value::Error(ErrorKind::Value)),
            }
        });
        reg.register_function("LN", |args| {
            if args.len() != 1 {
                Ok(Value::Error(ErrorKind::Value))
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
                Ok(Value::Error(ErrorKind::Value))
            } else {
                Ok(Value::Number(args[0].to_number().exp()))
            }
        });
        reg.register_function("PI", |args| {
            if !args.is_empty() {
                Ok(Value::Error(ErrorKind::Value))
            } else {
                Ok(Value::Number(std::f64::consts::PI))
            }
        });
        reg.register_function("RAND", |_args| {
            // xorshift64 PRNG advanced per call. Seeded once per thread
            // from a high-entropy source. Two RAND() calls in the same
            // millisecond used to return identical or highly-correlated
            // values because the previous seed was just `subsec_nanos()`.
            Ok(Value::Number(next_rand_u64() as f64 / u64::MAX as f64))
        });
        reg.register_function("RANDBETWEEN", |args| {
            if args.len() != 2 {
                Ok(Value::Error(ErrorKind::Value))
            } else {
                let low = args[0].to_number();
                let high = args[1].to_number();
                if high < low {
                    return Ok(Value::Error(ErrorKind::Num));
                }
                let r = next_rand_u64() as f64 / u64::MAX as f64;
                let result = (low + r * (high - low + 1.0)).floor();
                Ok(Value::Number(result))
            }
        });
        reg.register_function("SIGN", |args| {
            if args.len() != 1 {
                Ok(Value::Error(ErrorKind::Value))
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
                Ok(Value::Error(ErrorKind::Value))
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
                return Ok(Value::Error(ErrorKind::Div0));
            }
            nums.sort_by(|a, b| a.partial_cmp(b).unwrap_or(std::cmp::Ordering::Equal));
            let n = nums.len();
            let m = if n.is_multiple_of(2) {
                (nums[n / 2 - 1] + nums[n / 2]) / 2.0
            } else {
                nums[n / 2]
            };
            Ok(Value::Number(m))
        });
        reg.register_function("STDEV.S", |args| {
            let nums = collect_numbers(args);
            if nums.len() < 2 {
                return Ok(Value::Error(ErrorKind::Value));
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
                return Ok(Value::Error(ErrorKind::Value));
            }
            let mean: f64 = nums.iter().sum::<f64>() / nums.len() as f64;
            let var: f64 = nums.iter().map(|n| (n - mean).powi(2)).sum::<f64>()
                / (nums.len() - 1) as f64;
            Ok(Value::Number(var.sqrt()))
        });
        reg.register_function("STDEV.P", |args| {
            let nums = collect_numbers(args);
            if nums.is_empty() {
                return Ok(Value::Error(ErrorKind::Div0));
            }
            let mean: f64 = nums.iter().sum::<f64>() / nums.len() as f64;
            let var: f64 = nums.iter().map(|n| (n - mean).powi(2)).sum::<f64>()
                / nums.len() as f64;
            Ok(Value::Number(var.sqrt()))
        });
        reg.register_function("VAR.S", |args| {
            let nums = collect_numbers(args);
            if nums.len() < 2 {
                return Ok(Value::Error(ErrorKind::Value));
            }
            let mean: f64 = nums.iter().sum::<f64>() / nums.len() as f64;
            let var: f64 = nums.iter().map(|n| (n - mean).powi(2)).sum::<f64>()
                / (nums.len() - 1) as f64;
            Ok(Value::Number(var))
        });
        reg.register_function("VAR.P", |args| {
            let nums = collect_numbers(args);
            if nums.is_empty() {
                return Ok(Value::Error(ErrorKind::Div0));
            }
            let mean: f64 = nums.iter().sum::<f64>() / nums.len() as f64;
            let var: f64 = nums.iter().map(|n| (n - mean).powi(2)).sum::<f64>()
                / nums.len() as f64;
            Ok(Value::Number(var))
        });
        reg.register_function("LARGE", |args| {
            if args.len() != 2 {
                return Ok(Value::Error(ErrorKind::Value));
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
                return Ok(Value::Error(ErrorKind::Value));
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
                return Ok(Value::Error(ErrorKind::Value));
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
                return Ok(Value::Error(ErrorKind::Value));
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
                return Ok(Value::Error(ErrorKind::Value));
            }
            let x = collect_numbers(&args[..1]);
            let y = collect_numbers(&args[1..2]);
            if x.len() != y.len() || x.len() < 2 {
                return Ok(Value::Error(ErrorKind::NA));
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
                return Ok(Value::Error(ErrorKind::Value));
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
            let n_raw = args[0].to_number();
            if n_raw < 0.0 || !n_raw.is_finite() {
                return Ok(Value::Error(ErrorKind::Num));
            }
            // 170! is the largest f64-representable factorial; 171! is Inf.
            // Reject up-front so callers get #NUM! instead of an Inf cell
            // that breaks downstream arithmetic silently.
            if n_raw > 170.0 {
                return Ok(Value::Error(ErrorKind::Num));
            }
            let n = n_raw as u64;
            let mut r: f64 = 1.0;
            for i in 2..=n {
                r *= i as f64;
            }
            Ok(Value::Number(r))
        });
        reg.register_function("COMBIN", |args| {
            if args.len() != 2 {
                return Ok(Value::Error(ErrorKind::Value));
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
            let raw = flatten_args(args);
            if raw.is_empty() {
                return Ok(Value::Error(ErrorKind::Num));
            }
            // Excel: any negative argument → #NUM!. Non-integer args are
            // truncated toward zero before computing.
            let mut nums: Vec<i64> = Vec::with_capacity(raw.len());
            for v in &raw {
                let n = v.to_number();
                if !n.is_finite() || n < 0.0 {
                    return Ok(Value::Error(ErrorKind::Num));
                }
                nums.push(n as i64);
            }
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
            // (a*b).abs() wraps silently on i64 overflow. Use checked_mul
            // and surface #NUM! when the product would wrap — much better
            // than a nonsense small or negative result.
            let mut acc: i64 = 1;
            for &b in &nums {
                if b == 0 {
                    return Ok(Value::Number(0.0));
                }
                let g = gcd(acc, b);
                let b_div = b / g;
                match acc.checked_mul(b_div) {
                    Some(v) => acc = v.abs(),
                    None => return Ok(Value::Error(ErrorKind::Num)),
                }
            }
            Ok(Value::Number(acc as f64))
        });
        reg.register_function("ROUNDUP", |args| {
            if args.len() != 2 {
                return Ok(Value::Error(ErrorKind::Value));
            }
            let n = args[0].to_number();
            let digits = args[1].to_number() as i32;
            let scale = 10f64.powi(digits);
            let v = (n * scale).abs().ceil() * n.signum() / scale;
            Ok(Value::Number(v))
        });
        reg.register_function("ROUNDDOWN", |args| {
            if args.len() != 2 {
                return Ok(Value::Error(ErrorKind::Value));
            }
            let n = args[0].to_number();
            let digits = args[1].to_number() as i32;
            let scale = 10f64.powi(digits);
            let v = (n * scale).abs().floor() * n.signum() / scale;
            Ok(Value::Number(v))
        });
        reg.register_function("MROUND", |args| {
            if args.len() != 2 {
                return Ok(Value::Error(ErrorKind::Value));
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
                return Ok(Value::Error(ErrorKind::Value));
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

// Thread-local xorshift64 PRNG state. Seeded once per thread from a
// high-entropy source (nanos XOR thread id, then mixed with the splitmix
// constants to deal out a non-zero starting value).
thread_local! {
    static RAND_STATE: std::cell::Cell<u64> = std::cell::Cell::new(seed_rand());
}

fn seed_rand() -> u64 {
    use std::time::SystemTime;
    let nanos = SystemTime::now()
        .duration_since(SystemTime::UNIX_EPOCH)
        .map(|d| d.as_nanos() as u64)
        .unwrap_or(0xCAFE_BABE);
    let mut s = nanos ^ 0x9E37_79B9_7F4A_7C15;
    // splitmix64 — one round is enough to decorrelate from epoch.
    s = (s ^ (s >> 30)).wrapping_mul(0xBF58_476D_1CE4_E5B9);
    s = (s ^ (s >> 27)).wrapping_mul(0x94D0_49BB_1331_11EB);
    s ^ (s >> 31)
}

/// Advance the thread-local PRNG and return the next 64-bit value.
fn next_rand_u64() -> u64 {
    RAND_STATE.with(|cell| {
        let mut s = cell.get();
        if s == 0 {
            s = seed_rand().max(1);
        }
        s ^= s << 13;
        s ^= s >> 7;
        s ^= s << 17;
        cell.set(s);
        s
    })
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
