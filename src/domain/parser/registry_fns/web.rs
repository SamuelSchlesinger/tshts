//! `web` builtin functions.
//!
//! Each function registers itself on the provided `FunctionRegistry` via
//! `reg.register_function(NAME, |args| ...)`. The category-level `register`
//! function is called from `registry::FunctionRegistry::register_builtin_functions`.

#![allow(unused_imports)]
use crate::domain::parser::{FunctionRegistry, FunctionPurity, Value, ErrorKind, flatten_args, shape_of, broadcast_binary, criteria_matches, add_commas, glob_match, date_to_serial, serial_to_date, parse_iso_date, days_in_month, today_serial, now_serial};

/// Register all `web` builtin functions on `reg`.
pub(in crate::domain::parser) fn register(reg: &mut FunctionRegistry) {
        reg.register_function_with_purity("GET", |args| {
            if args.len() != 1 {
                return Ok(Value::Error(ErrorKind::Value));
            }
            let url = args[0].to_string();
            if url.is_empty() {
                return Ok(Value::Error(ErrorKind::Value));
            }
            // Route through the domain-level fetcher abstraction so this
            // function doesn't import `crate::infrastructure` directly —
            // infrastructure installs the impl at startup (see
            // `infrastructure::fetcher::install_as_http_fetcher`).
            use crate::domain::services::{http_fetch, HttpFetchResult};
            match http_fetch(&url) {
                HttpFetchResult::Value(body) => Ok(Value::String(body)),
                HttpFetchResult::Loading => Ok(Value::String("Loading…".to_string())),
                // Fetcher errors carry a human-readable string upstream;
                // surface as #VALUE! so the cell renders an error literal
                // and IFERROR can trap it. The original message is dropped
                // because Value::Error doesn't carry text — acceptable since
                // the cause is usually a transient network issue.
                HttpFetchResult::Error => Ok(Value::Error(ErrorKind::Value)),
            }
        }, FunctionPurity::SideEffecting);
}
