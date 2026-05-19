//! `finance` builtin functions.
//!
//! Each function registers itself on the provided `FunctionRegistry` via
//! `reg.register_function(NAME, |args| ...)`. The category-level `register`
//! function is called from `registry::FunctionRegistry::register_builtin_functions`.

#![allow(unused_imports)]
use crate::domain::parser::{FunctionRegistry, Value, ErrorKind, flatten_args, shape_of, broadcast_binary, criteria_matches, add_commas, glob_match, date_to_serial, serial_to_date, parse_iso_date, days_in_month, today_serial, now_serial};

/// Register all `finance` builtin functions on `reg`.
pub(in crate::domain::parser) fn register(reg: &mut FunctionRegistry) {
        reg.register_function("PMT", |args| {
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
        reg.register_function("FV", |args| {
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
        reg.register_function("PV", |args| {
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
        reg.register_function("NPV", |args| {
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
}
