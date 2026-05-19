//! `web` builtin functions.
//!
//! Each function registers itself on the provided `FunctionRegistry` via
//! `reg.register_function(NAME, |args| ...)`. The category-level `register`
//! function is called from `registry::FunctionRegistry::register_builtin_functions`.

#![allow(unused_imports)]
use crate::domain::parser::{FunctionRegistry, Value, ErrorKind, flatten_args, shape_of, broadcast_binary, criteria_matches, add_commas, glob_match, date_to_serial, serial_to_date, parse_iso_date, days_in_month, today_serial, now_serial};

/// Register all `web` builtin functions on `reg`.
pub(in crate::domain::parser) fn register(reg: &mut FunctionRegistry) {
        reg.register_function("GET", |args| {
            if args.len() != 1 {
                return Err("GET requires exactly 1 argument (URL)".to_string());
            }
            let url = args[0].to_string();
            if url.is_empty() {
                return Err("GET: empty URL".to_string());
            }
            use crate::infrastructure::fetcher::{fetch, FetchResult};
            match fetch(&url) {
                FetchResult::Value(body) => Ok(Value::String(body)),
                FetchResult::Loading => Ok(Value::String("Loading…".to_string())),
                FetchResult::Error(msg) => Err(msg),
            }
        });
}
