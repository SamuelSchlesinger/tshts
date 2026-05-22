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
            if let Some(close) = end
                && close + 1 < chars.len() && chars[close + 1] == '!' {
                    let name: String = chars[i + 1..close].iter().collect();
                    if name.eq_ignore_ascii_case(old) {
                        out.push_str(&quoted_new);
                        i = close + 2;
                        continue;
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

/// Walk `formula` and replace every reference to `removed_sheet` (whether
/// bare `Sheet1!A1` or quoted `'My Sheet'!A1:B10`) with the literal `#REF!`.
/// Excel does this when a referenced sheet is deleted — the existing
/// formula structure is preserved, but the dangling refs surface as
/// `#REF!` at the call site (and propagate via error semantics).
pub fn replace_sheet_refs_with_ref_error(formula: &str, removed_sheet: &str) -> String {
    if !formula.starts_with('=') {
        return formula.to_string();
    }
    let chars: Vec<char> = formula.chars().collect();
    let mut out = String::with_capacity(formula.len());
    let mut i = 0;
    let mut in_string = false;
    while i < chars.len() {
        let c = chars[i];
        if in_string {
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
        // Try `'<name>'!<cell_or_range>` first.
        if c == '\''
            && let Some(close) = chars[i + 1..].iter().position(|&x| x == '\'').map(|p| i + 1 + p)
            && close + 1 < chars.len() && chars[close + 1] == '!'
        {
            let name: String = chars[i + 1..close].iter().collect();
            if name.eq_ignore_ascii_case(removed_sheet) {
                // Consume `'name'!<cell>[:<cell>]` and emit `#REF!`.
                let after_bang = close + 2;
                let consumed = skip_cell_or_range(&chars, after_bang);
                out.push_str("#REF!");
                i = consumed;
                continue;
            }
        }
        // Bare `<name>!<cell_or_range>`.
        if c.is_ascii_alphabetic() || c == '_' {
            let mut j = i;
            while j < chars.len() && (chars[j].is_ascii_alphanumeric() || chars[j] == '_') {
                j += 1;
            }
            if j < chars.len() && chars[j] == '!' {
                let name: String = chars[i..j].iter().collect();
                if name.eq_ignore_ascii_case(removed_sheet) {
                    let after_bang = j + 1;
                    let consumed = skip_cell_or_range(&chars, after_bang);
                    out.push_str("#REF!");
                    i = consumed;
                    continue;
                }
            }
        }
        out.push(c);
        i += 1;
    }
    out
}

/// Skip an A1-style cell reference (optionally absolute) starting at `start`.
/// If a `:` follows, skip the second endpoint too. Returns the index past the
/// reference. Used by `replace_sheet_refs_with_ref_error` to consume the
/// `A1` / `A1:B10` / `$A$1:$B$10` that follows a sheet bang.
fn skip_cell_or_range(chars: &[char], start: usize) -> usize {
    let after_first = skip_one_cell(chars, start);
    if after_first < chars.len() && chars[after_first] == ':' {
        skip_one_cell(chars, after_first + 1)
    } else {
        after_first
    }
}

fn skip_one_cell(chars: &[char], start: usize) -> usize {
    let mut i = start;
    if i < chars.len() && chars[i] == '$' {
        i += 1;
    }
    while i < chars.len() && chars[i].is_ascii_alphabetic() {
        i += 1;
    }
    if i < chars.len() && chars[i] == '$' {
        i += 1;
    }
    while i < chars.len() && chars[i].is_ascii_digit() {
        i += 1;
    }
    i
}

#[cfg(test)]
mod replace_tests {
    use super::*;

    #[test]
    fn replaces_bare_sheet_ref() {
        assert_eq!(
            replace_sheet_refs_with_ref_error("=Sheet2!A1+1", "Sheet2"),
            "=#REF!+1"
        );
    }

    #[test]
    fn replaces_quoted_sheet_ref() {
        assert_eq!(
            replace_sheet_refs_with_ref_error("='My Sheet'!A1+1", "My Sheet"),
            "=#REF!+1"
        );
    }

    #[test]
    fn replaces_sheet_qualified_range() {
        assert_eq!(
            replace_sheet_refs_with_ref_error("=SUM(Data!A1:B10)", "Data"),
            "=SUM(#REF!)"
        );
    }

    #[test]
    fn handles_absolute_refs() {
        assert_eq!(
            replace_sheet_refs_with_ref_error("=Sheet2!$A$1+Sheet2!$B$5", "Sheet2"),
            "=#REF!+#REF!"
        );
    }

    #[test]
    fn case_insensitive() {
        assert_eq!(
            replace_sheet_refs_with_ref_error("=SHEET2!A1", "sheet2"),
            "=#REF!"
        );
    }

    #[test]
    fn leaves_other_sheet_refs_alone() {
        assert_eq!(
            replace_sheet_refs_with_ref_error("=Sheet1!A1+Sheet3!B2", "Sheet2"),
            "=Sheet1!A1+Sheet3!B2"
        );
    }

    #[test]
    fn skips_inside_string_literal() {
        assert_eq!(
            replace_sheet_refs_with_ref_error("=CONCAT(\"Sheet2!A1\",B5)", "Sheet2"),
            "=CONCAT(\"Sheet2!A1\",B5)"
        );
    }

    #[test]
    fn preserves_local_refs() {
        // No sheet qualifier → no rewrite.
        assert_eq!(
            replace_sheet_refs_with_ref_error("=A1+B2", "Sheet2"),
            "=A1+B2"
        );
    }
}
