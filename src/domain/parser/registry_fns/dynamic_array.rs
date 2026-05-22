//! `dynamic_array` builtin functions.
//!
//! Each function registers itself on the provided `FunctionRegistry` via
//! `reg.register_function(NAME, |args| ...)`. The category-level `register`
//! function is called from `registry::FunctionRegistry::register_builtin_functions`.

#![allow(unused_imports)]
use crate::domain::parser::{FunctionRegistry, Value, ErrorKind, flatten_args, shape_of, broadcast_binary, criteria_matches, add_commas, glob_match, date_to_serial, serial_to_date, parse_iso_date, days_in_month, today_serial, now_serial};

/// Hard cap on the result size of dynamic-array constructors (SEQUENCE,
/// MAKEARRAY, etc.). Without this, `=SEQUENCE(1e6, 1e6)` allocates 8 TB
/// before the OS kills the process. One million cells is plenty for any
/// realistic use case.
pub(crate) const MAX_DYNAMIC_ARRAY_CELLS: usize = 1_000_000;

/// Register all `dynamic_array` builtin functions on `reg`.
pub(in crate::domain::parser) fn register(reg: &mut FunctionRegistry) {
        reg.register_function("SUMPRODUCT", |args| {
            if args.is_empty() {
                return Ok(Value::Error(ErrorKind::Value));
            }
            let mut acc: Vec<f64> = args[0]
                .flatten()
                .iter()
                .map(|v| v.to_number())
                .collect();
            for arg in &args[1..] {
                let next: Vec<f64> = arg.flatten().iter().map(|v| v.to_number()).collect();
                if next.len() != acc.len() {
                    return Ok(Value::Error(ErrorKind::Value));
                }
                for (i, n) in next.iter().enumerate() {
                    acc[i] *= n;
                }
            }
            Ok(Value::Number(acc.iter().sum()))
        });
        reg.register_function("TRANSPOSE", |args| {
            if args.len() != 1 {
                return Ok(Value::Error(ErrorKind::Value));
            }
            let (rows, cols, data) = shape_of(&args[0]);
            let mut out = Vec::with_capacity(rows * cols);
            for c in 0..cols {
                for r in 0..rows {
                    out.push(data[r * cols + c].clone());
                }
            }
            Ok(Value::Array { rows: cols, cols: rows, data: out })
        });
        reg.register_function("SEQUENCE", |args| {
            if args.is_empty() || args.len() > 4 {
                return Ok(Value::Error(ErrorKind::Value));
            }
            let rows_f = args[0].to_number();
            let cols_f = args.get(1).map(|v| v.to_number()).unwrap_or(1.0);
            if rows_f < 1.0 || cols_f < 1.0 || !rows_f.is_finite() || !cols_f.is_finite() {
                return Ok(Value::Error(ErrorKind::Num));
            }
            let rows = rows_f as usize;
            let cols = cols_f as usize;
            let total = rows.checked_mul(cols).unwrap_or(usize::MAX);
            if total > MAX_DYNAMIC_ARRAY_CELLS {
                return Ok(Value::Error(ErrorKind::Num));
            }
            let start = args.get(2).map(|v| v.to_number()).unwrap_or(1.0);
            let step = args.get(3).map(|v| v.to_number()).unwrap_or(1.0);
            let mut data = Vec::with_capacity(total);
            for i in 0..total {
                data.push(Value::Number(start + step * i as f64));
            }
            Ok(Value::Array { rows, cols, data })
        });
        reg.register_function("FILTER", |args| {
            if args.len() < 2 || args.len() > 3 {
                return Ok(Value::Error(ErrorKind::Value));
            }
            let (rows, cols, data) = shape_of(&args[0]);
            let mask = args[1].flatten();
            if mask.len() != rows && mask.len() != rows * cols {
                return Ok(Value::Error(ErrorKind::Value));
            }
            let mut out_rows: Vec<Value> = Vec::new();
            let mut kept = 0;
            for r in 0..rows {
                let keep = mask.get(r).map(|v| v.is_truthy()).unwrap_or(false);
                if keep {
                    for c in 0..cols {
                        out_rows.push(data[r * cols + c].clone());
                    }
                    kept += 1;
                }
            }
            if kept == 0 {
                if let Some(fallback) = args.get(2) {
                    return Ok(fallback.clone());
                }
                return Ok(Value::Error(ErrorKind::NA));
            }
            Ok(Value::Array {
                rows: kept,
                cols,
                data: out_rows,
            })
        });
        reg.register_function("SORT", |args| {
            if args.is_empty() || args.len() > 3 {
                return Ok(Value::Error(ErrorKind::Value));
            }
            let (rows, cols, data) = shape_of(&args[0]);
            let sort_col = args.get(1).map(|v| v.to_number() as usize).unwrap_or(1);
            let order = args.get(2).map(|v| v.to_number() as i64).unwrap_or(1);
            if sort_col == 0 || sort_col > cols {
                return Ok(Value::Error(ErrorKind::Ref));
            }
            let mut row_indices: Vec<usize> = (0..rows).collect();
            // Comparator follows Excel's mixed-type rule: numbers come first
            // (ordered numerically among themselves), then text (ordered
            // alphabetically), then bools/errors/empty. The previous version
            // ran `to_number()` unconditionally which collapses every text
            // cell to 0 and silently scrambles columns that mix types.
            fn sort_kind(v: &Value) -> u8 {
                match v {
                    Value::Number(_) => 0,
                    Value::Bool(_) => 0, // sorted with numbers per Excel
                    Value::String(s) if s.parse::<f64>().is_ok() => 0,
                    Value::String(_) => 1,
                    Value::Error(_) => 2,
                    _ => 3,
                }
            }
            row_indices.sort_by(|a, b| {
                let av = &data[*a * cols + sort_col - 1];
                let bv = &data[*b * cols + sort_col - 1];
                let ak = sort_kind(av);
                let bk = sort_kind(bv);
                let cmp = if ak != bk {
                    ak.cmp(&bk)
                } else if ak == 0 {
                    av.to_number()
                        .partial_cmp(&bv.to_number())
                        .unwrap_or(std::cmp::Ordering::Equal)
                } else {
                    av.to_string().to_lowercase().cmp(&bv.to_string().to_lowercase())
                };
                if order < 0 { cmp.reverse() } else { cmp }
            });
            let mut out = Vec::with_capacity(rows * cols);
            for r in &row_indices {
                for c in 0..cols {
                    out.push(data[r * cols + c].clone());
                }
            }
            Ok(Value::Array { rows, cols, data: out })
        });
        reg.register_function("UNIQUE", |args| {
            if args.len() != 1 {
                return Ok(Value::Error(ErrorKind::Value));
            }
            let (rows, cols, data) = shape_of(&args[0]);
            let mut seen: std::collections::HashSet<String> = std::collections::HashSet::new();
            let mut out = Vec::new();
            let mut kept = 0;
            for r in 0..rows {
                let row_key: String = (0..cols)
                    .map(|c| data[r * cols + c].to_string())
                    .collect::<Vec<_>>()
                    .join("\x1f");
                if seen.insert(row_key) {
                    for c in 0..cols {
                        out.push(data[r * cols + c].clone());
                    }
                    kept += 1;
                }
            }
            Ok(Value::Array { rows: kept, cols, data: out })
        });
}
