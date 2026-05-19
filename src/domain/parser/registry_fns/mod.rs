//! Builtin function definitions, organized by category.
//!
//! Each submodule registers a related set of functions on the
//! `FunctionRegistry`. The top-level dispatcher
//! `registry::FunctionRegistry::register_builtin_functions` calls each
//! `category::register(self)` in turn.
//!
//! Categories (146 functions total):
//!
//!   - `numeric`       — SUM/AVERAGE/MIN/MAX, math, trig, statistics (52 fns)
//!   - `string`        — LEN/UPPER/LOWER/CONCAT, REGEX*, text formatting (32 fns)
//!   - `logical`       — IF/AND/OR/NOT/XOR, IFS/SWITCH/IFERROR/IFNA (11 fns)
//!   - `lookup`        — VLOOKUP/HLOOKUP/INDEX/MATCH/XLOOKUP, SUMIF/COUNTIF/AVERAGEIF (8 fns)
//!   - `date`          — DATE/TIME, calendar arithmetic, weekday helpers (20 fns)
//!   - `info`          — type predicates (IS*), COUNT/COUNTA, error helpers (11 fns)
//!   - `web`           — GET (HTTP fetch) (1 fn)
//!   - `dynamic_array` — SEQUENCE/FILTER/SORT/UNIQUE/TRANSPOSE/SUMPRODUCT (6 fns)
//!   - `finance`       — FV/PV/NPV/PMT (4 fns)
//!   - `viz`           — SPARKLINE (1 fn)
//!
//! LAMBDA/MAP/REDUCE/BYROW/BYCOL/SCAN/MAKEARRAY are NOT here — they need the
//! evaluator's lambda machinery and live in `parser/evaluator.rs`.

pub(super) mod numeric;
pub(super) mod string;
pub(super) mod logical;
pub(super) mod lookup;
pub(super) mod date;
pub(super) mod info;
pub(super) mod web;
pub(super) mod dynamic_array;
pub(super) mod finance;
pub(super) mod viz;
