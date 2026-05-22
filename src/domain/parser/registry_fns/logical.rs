//! `logical` builtin functions.
//!
//! Each function registers itself on the provided `FunctionRegistry` via
//! `reg.register_function(NAME, |args| ...)`. The category-level `register`
//! function is called from `registry::FunctionRegistry::register_builtin_functions`.

#![allow(unused_imports)]
use crate::domain::parser::{FunctionRegistry, Value, ErrorKind, flatten_args, shape_of, broadcast_binary, criteria_matches, add_commas, glob_match, date_to_serial, serial_to_date, parse_iso_date, days_in_month, today_serial, now_serial};

/// Register all `logical` builtin functions on `reg`.
pub(in crate::domain::parser) fn register(reg: &mut FunctionRegistry) {
        reg.register_function("IF", |args| {
            if args.len() != 3 {
                Err("IF requires exactly 3 arguments".to_string())
            } else {
                Ok(if args[0].is_truthy() { args[1].clone() } else { args[2].clone() })
            }
        });
        reg.register_function("AND", |args| {
            let flat = flatten_args(args);
            if let Some(e) = flat.iter().find_map(|v| v.first_error()) {
                return Ok(Value::Error(e));
            }
            let result = flat.iter().all(|v| v.is_truthy());
            Ok(Value::Bool(result))
        });
        reg.register_function("OR", |args| {
            let flat = flatten_args(args);
            if let Some(e) = flat.iter().find_map(|v| v.first_error()) {
                return Ok(Value::Error(e));
            }
            let result = flat.iter().any(|v| v.is_truthy());
            Ok(Value::Bool(result))
        });
        reg.register_function("NOT", |args| {
            if args.len() != 1 {
                Err("NOT requires exactly 1 argument".to_string())
            } else if let Some(e) = args[0].first_error() {
                Ok(Value::Error(e))
            } else {
                let result = !args[0].is_truthy();
                Ok(Value::Number(if result { 1.0 } else { 0.0 }))
            }
        });
        reg.register_function("IFERROR", |args| {
            if args.len() != 2 {
                return Err("IFERROR requires 2 arguments".to_string());
            }
            if args[0].is_error() {
                Ok(args[1].clone())
            } else {
                Ok(args[0].clone())
            }
        });
        reg.register_function("IFNA", |args| {
            if args.len() != 2 {
                return Err("IFNA requires 2 arguments".to_string());
            }
            if matches!(args[0].first_error(), Some(ErrorKind::NA)) {
                Ok(args[1].clone())
            } else {
                Ok(args[0].clone())
            }
        });
        reg.register_function("TRUE", |args| {
            if !args.is_empty() {
                return Err("TRUE takes no arguments".to_string());
            }
            Ok(Value::Bool(true))
        });
        reg.register_function("FALSE", |args| {
            if !args.is_empty() {
                return Err("FALSE takes no arguments".to_string());
            }
            Ok(Value::Bool(false))
        });
        reg.register_function("IFS", |args| {
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
        reg.register_function("SWITCH", |args| {
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
        reg.register_function("XOR", |args| {
            let flat = flatten_args(args);
            if let Some(e) = flat.iter().find_map(|v| v.first_error()) {
                return Ok(Value::Error(e));
            }
            let count_true = flat.iter().filter(|v| v.is_truthy()).count();
            Ok(Value::Bool(count_true % 2 == 1))
        });
}
