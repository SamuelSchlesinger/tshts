//! `date` builtin functions.
//!
//! Each function registers itself on the provided `FunctionRegistry` via
//! `reg.register_function(NAME, |args| ...)`. The category-level `register`
//! function is called from `registry::FunctionRegistry::register_builtin_functions`.

#![allow(unused_imports)]
use crate::domain::parser::{FunctionRegistry, Value, ErrorKind, flatten_args, shape_of, broadcast_binary, criteria_matches, add_commas, glob_match, date_to_serial, serial_to_date, parse_iso_date, days_in_month, today_serial, now_serial};

/// Register all `date` builtin functions on `reg`.
pub(in crate::domain::parser) fn register(reg: &mut FunctionRegistry) {
        reg.register_function("TODAY", |args| {
            if !args.is_empty() {
                return Err("TODAY takes no arguments".to_string());
            }
            Ok(Value::Number(today_serial()))
        });
        reg.register_function("NOW", |args| {
            if !args.is_empty() {
                return Err("NOW takes no arguments".to_string());
            }
            Ok(Value::Number(now_serial()))
        });
        reg.register_function("DATE", |args| {
            if args.len() != 3 {
                return Err("DATE requires 3 arguments (year, month, day)".to_string());
            }
            let y = args[0].to_number() as i32;
            let m = args[1].to_number() as u32;
            let d = args[2].to_number() as u32;
            Ok(Value::Number(date_to_serial(y, m, d)))
        });
        reg.register_function("YEAR", |args| {
            if args.len() != 1 {
                return Err("YEAR requires 1 argument".to_string());
            }
            let (y, _, _) = serial_to_date(args[0].to_number());
            Ok(Value::Number(y as f64))
        });
        reg.register_function("MONTH", |args| {
            if args.len() != 1 {
                return Err("MONTH requires 1 argument".to_string());
            }
            let (_, m, _) = serial_to_date(args[0].to_number());
            Ok(Value::Number(m as f64))
        });
        reg.register_function("DAY", |args| {
            if args.len() != 1 {
                return Err("DAY requires 1 argument".to_string());
            }
            let (_, _, d) = serial_to_date(args[0].to_number());
            Ok(Value::Number(d as f64))
        });
        reg.register_function("TIME", |args| {
            if args.len() != 3 {
                return Err("TIME requires 3 arguments".to_string());
            }
            let h = args[0].to_number();
            let m = args[1].to_number();
            let s = args[2].to_number();
            Ok(Value::Number((h * 3600.0 + m * 60.0 + s) / 86400.0))
        });
        reg.register_function("HOUR", |args| {
            let frac = args[0].to_number().fract();
            let secs = (frac * 86400.0).round() as i64;
            Ok(Value::Number(((secs / 3600) % 24) as f64))
        });
        reg.register_function("MINUTE", |args| {
            let frac = args[0].to_number().fract();
            let secs = (frac * 86400.0).round() as i64;
            Ok(Value::Number(((secs / 60) % 60) as f64))
        });
        reg.register_function("SECOND", |args| {
            let frac = args[0].to_number().fract();
            let secs = (frac * 86400.0).round() as i64;
            Ok(Value::Number((secs % 60) as f64))
        });
        reg.register_function("DATEDIF", |args| {
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
        reg.register_function("WEEKDAY", |args| {
            if args.is_empty() {
                return Err("WEEKDAY requires 1 or 2 arguments".to_string());
            }
            let serial = args[0].to_number().floor() as i64;
            let ty = args.get(1).map(|v| v.to_number() as i64).unwrap_or(1);
            // 1899-12-30 (serial 0) was a Saturday → (0 + 6) % 7 = 6 (Sat in 0-based Mon=0).
            let mon_based = (serial + 5).rem_euclid(7); // Mon=0..Sun=6
            let v = match ty {
                1 => ((mon_based + 1) % 7) + 1, // Sun=1..Sat=7
                2 => mon_based + 1,             // Mon=1..Sun=7
                3 => mon_based,                 // Mon=0..Sun=6
                _ => return Err(format!("WEEKDAY: bad type {}", ty)),
            };
            Ok(Value::Number(v as f64))
        });
        reg.register_function("EDATE", |args| {
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
        reg.register_function("EOMONTH", |args| {
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
        reg.register_function("DAYS", |args| {
            if args.len() != 2 {
                return Err("DAYS requires 2 arguments".to_string());
            }
            Ok(Value::Number(args[0].to_number().floor() - args[1].to_number().floor()))
        });
        reg.register_function("NETWORKDAYS", |args| {
            if args.len() < 2 || args.len() > 3 {
                return Err("NETWORKDAYS requires 2 or 3 arguments".to_string());
            }
            let start = args[0].to_number().floor() as i64;
            let end = args[1].to_number().floor() as i64;
            let (lo, hi, sign) = if start <= end { (start, end, 1) } else { (end, start, -1) };
            // Cap the span at ~3 centuries (109,500 days) to keep an
            // accidental `=NETWORKDAYS(0, 1e9)` from hanging the UI.
            if hi.saturating_sub(lo) > 109_500 {
                return Ok(Value::Error(ErrorKind::Num));
            }
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
                let dow = (d + 5).rem_euclid(7);
                if dow < 5 && !holidays.contains(&d) {
                    count += 1;
                }
            }
            Ok(Value::Number((count * sign) as f64))
        });
        reg.register_function("WORKDAY", |args| {
            if args.len() < 2 || args.len() > 3 {
                return Err("WORKDAY requires 2 or 3 arguments".to_string());
            }
            let start = args[0].to_number().floor() as i64;
            let days = args[1].to_number() as i64;
            // Same span cap as NETWORKDAYS: refuse to step more than ~3
            // centuries' worth of business days.
            if days.unsigned_abs() > 109_500 {
                return Ok(Value::Error(ErrorKind::Num));
            }
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
        reg.register_function("DATEVALUE", |args| {
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
            if parts.len() == 3
                && let (Ok(m), Ok(d), Ok(y)) = (
                    parts[0].parse::<u32>(),
                    parts[1].parse::<u32>(),
                    parts[2].parse::<i32>(),
                ) {
                    let year = if y < 100 { 2000 + y } else { y };
                    return Ok(Value::Number(date_to_serial(year, m, d)));
                }
            Ok(Value::Error(ErrorKind::Value))
        });
        reg.register_function("TIMEVALUE", |args| {
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
        reg.register_function("YEARFRAC", |args| {
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
