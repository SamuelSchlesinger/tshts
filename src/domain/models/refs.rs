//! Submodule of `models` — see models/mod.rs.

#![allow(unused_imports)]
use serde::{Deserialize, Serialize};
use std::collections::{HashMap, HashSet, VecDeque};
use super::*;

pub fn rewrite_sheet_refs(formula: &str, old: &str, new: &str) -> String {
    if !formula.starts_with('=') {
        return formula.to_string();
    }
    let needs_quotes = new.chars().any(|c| !c.is_ascii_alphanumeric() && c != '_');
    let quoted_new = if needs_quotes {
        format!("'{}'!", new.replace('\'', "''"))
    } else {
        format!("{}!", new)
    };
    // Walk char-by-char, replacing matches at boundaries. A "match" is an
    // identifier-like token immediately followed by `!` whose name equals
    // `old` (case-insensitively). Skip the inside of `"..."` literals so we
    // don't rewrite sheet names that appear inside cell-value strings.
    let chars: Vec<char> = formula.chars().collect();
    let mut out = String::with_capacity(formula.len());
    let mut i = 0;
    let mut in_string = false;
    while i < chars.len() {
        let c = chars[i];
        if in_string {
            // Inside "..."; just copy. Handle doubled "" as an escaped quote.
            out.push(c);
            if c == '"' {
                if i + 1 < chars.len() && chars[i + 1] == '"' {
                    out.push('"');
                    i += 2;
                    continue;
                }
                in_string = false;
            }
            i += 1;
            continue;
        }
        if c == '"' {
            in_string = true;
            out.push(c);
            i += 1;
            continue;
        }
        // Try to match `'<name>'!` first.
        if c == '\'' {
            let end = chars[i + 1..]
                .iter()
                .position(|&x| x == '\'')
                .map(|p| i + 1 + p);
            if let Some(close) = end {
                if close + 1 < chars.len() && chars[close + 1] == '!' {
                    let name: String = chars[i + 1..close].iter().collect();
                    if name.eq_ignore_ascii_case(old) {
                        out.push_str(&quoted_new);
                        i = close + 2;
                        continue;
                    }
                }
            }
        }
        // Try to match a bare identifier followed by `!`.
        if c.is_ascii_alphabetic() || c == '_' {
            let mut j = i;
            while j < chars.len() && (chars[j].is_ascii_alphanumeric() || chars[j] == '_') {
                j += 1;
            }
            if j < chars.len() && chars[j] == '!' {
                let name: String = chars[i..j].iter().collect();
                if name.eq_ignore_ascii_case(old) {
                    out.push_str(&quoted_new);
                    i = j + 1;
                    continue;
                }
            }
        }
        out.push(c);
        i += 1;
    }
    out
}

/// Variant of `rewrite_sheet_refs` for named-range *values* (not formulas).
/// Named-range values look like `Sheet1!A1:B10` (no leading `=`), so we
/// prepend `=` to satisfy the formula-shape check, then strip it.
pub fn rewrite_sheet_refs_for_name_value(value: &str, old: &str, new: &str) -> String {
    let s = format!("={}", value);
    let rewritten = rewrite_sheet_refs(&s, old, new);
    rewritten.strip_prefix('=').map(|x| x.to_string()).unwrap_or(rewritten)
}
