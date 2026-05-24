//! Search matcher used by find-replace mode and the `:s/find/replace/` command.
//! Honors regex and case-sensitive flags; falls back to substring matching if
//! the regex is invalid.

/// Performs case-insensitive string replacement, preserving the replacement text as-is.
fn case_insensitive_replace(text: &str, search: &str, replacement: &str) -> String {
    let lower_text = text.to_lowercase();
    let lower_search = search.to_lowercase();
    let mut result = String::new();
    let mut start = 0;
    while let Some(pos) = lower_text[start..].find(&lower_search) {
        result.push_str(&text[start..start + pos]);
        result.push_str(replacement);
        start += pos + search.len();
    }
    result.push_str(&text[start..]);
    result
}

/// Search matcher that honors the regex and case-sensitive flags.
/// Falls back to substring matching if the regex is invalid.
pub struct TextMatcher {
    regex: Option<regex::Regex>,
    needle: String,
    case_sensitive: bool,
}

impl TextMatcher {
    pub fn new(query: &str, use_regex: bool, case_sensitive: bool) -> Self {
        let regex = if use_regex {
            let pattern = if case_sensitive {
                query.to_string()
            } else {
                format!("(?i){}", query)
            };
            regex::Regex::new(&pattern).ok()
        } else {
            None
        };
        Self {
            regex,
            needle: if case_sensitive { query.to_string() } else { query.to_lowercase() },
            case_sensitive,
        }
    }

    pub fn is_match(&self, hay: &str) -> bool {
        if let Some(ref r) = self.regex {
            return r.is_match(hay);
        }
        if self.case_sensitive {
            hay.contains(&self.needle)
        } else {
            hay.to_lowercase().contains(&self.needle)
        }
    }

    /// Replace all matches in `hay` with `replacement`. For regex mode,
    /// captures via `$1` style are supported.
    pub fn replace_all(&self, hay: &str, replacement: &str) -> String {
        if let Some(ref r) = self.regex {
            return r.replace_all(hay, replacement).into_owned();
        }
        if self.case_sensitive {
            hay.replace(&self.needle, replacement)
        } else {
            case_insensitive_replace(hay, &self.needle, replacement)
        }
    }
}
