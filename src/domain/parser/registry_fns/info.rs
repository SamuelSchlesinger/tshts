//! `info` builtin functions.
//!
//! Each function registers itself on the provided `FunctionRegistry` via
//! `reg.register_function(NAME, |args| ...)`. The category-level `register`
//! function is called from `registry::FunctionRegistry::register_builtin_functions`.

#![allow(unused_imports)]
use crate::domain::parser::{FunctionRegistry, Value, ErrorKind, flatten_args, shape_of, broadcast_binary, criteria_matches, add_commas, glob_match, date_to_serial, serial_to_date, parse_iso_date, days_in_month, today_serial, now_serial};

/// Register all `info` builtin functions on `reg`.
pub(in crate::domain::parser) fn register(reg: &mut FunctionRegistry) {
        reg.register_function("ISERROR", |args| {
            if args.len() != 1 {
                return Ok(Value::Error(ErrorKind::Value));
            }
            Ok(Value::Bool(args[0].is_error()))
        });
        reg.register_function("ISERR", |args| {
            // ISERR: error EXCEPT #N/A.
            if args.len() != 1 {
                return Ok(Value::Error(ErrorKind::Value));
            }
            let result = match args[0].first_error() {
                Some(ErrorKind::NA) | None => false,
                _ => true,
            };
            Ok(Value::Bool(result))
        });
        reg.register_function("ISNA", |args| {
            if args.len() != 1 {
                return Ok(Value::Error(ErrorKind::Value));
            }
            Ok(Value::Bool(matches!(args[0].first_error(), Some(ErrorKind::NA))))
        });
        reg.register_function("NA", |args| {
            if !args.is_empty() {
                return Ok(Value::Error(ErrorKind::Value));
            }
            Ok(Value::Error(ErrorKind::NA))
        });
        reg.register_function("ERROR.TYPE", |args| {
            if args.len() != 1 {
                return Ok(Value::Error(ErrorKind::Value));
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
        reg.register_function("ISBLANK", |args| {
            if args.len() != 1 {
                Ok(Value::Error(ErrorKind::Value))
            } else {
                let is_blank = match &args[0] {
                    Value::String(s) => s.is_empty(),
                    Value::List(l) => l.is_empty(),
                    _ => false,
                };
                Ok(Value::Number(if is_blank { 1.0 } else { 0.0 }))
            }
        });
        reg.register_function("ISNUMBER", |args| {
            if args.len() != 1 {
                Ok(Value::Error(ErrorKind::Value))
            } else {
                let is_num = matches!(&args[0], Value::Number(_));
                Ok(Value::Number(if is_num { 1.0 } else { 0.0 }))
            }
        });
        reg.register_function("ISTEXT", |args| {
            if args.len() != 1 {
                Ok(Value::Error(ErrorKind::Value))
            } else {
                let is_text = matches!(&args[0], Value::String(_));
                Ok(Value::Number(if is_text { 1.0 } else { 0.0 }))
            }
        });
        reg.register_function("TYPE", |args| {
            if args.len() != 1 {
                Ok(Value::Error(ErrorKind::Value))
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
        reg.register_function("COUNT", |args| {
            let flat = flatten_args(args);
            let count = flat.iter().filter(|v| matches!(v, Value::Number(_))).count();
            Ok(Value::Number(count as f64))
        });
        reg.register_function("COUNTA", |args| {
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
}
