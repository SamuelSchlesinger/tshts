//! `viz` builtin functions.
//!
//! Each function registers itself on the provided `FunctionRegistry` via
//! `reg.register_function(NAME, |args| ...)`. The category-level `register`
//! function is called from `registry::FunctionRegistry::register_builtin_functions`.

#![allow(unused_imports)]
use crate::domain::parser::{FunctionRegistry, Value, ErrorKind, flatten_args, shape_of, broadcast_binary, criteria_matches, add_commas, glob_match, date_to_serial, serial_to_date, parse_iso_date, days_in_month, today_serial, now_serial};

/// Register all `viz` builtin functions on `reg`.
pub(in crate::domain::parser) fn register(reg: &mut FunctionRegistry) {
        reg.register_function("SPARKLINE", |args| {
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
}
