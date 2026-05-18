//! Submodule of `services` — see services/mod.rs.

#![allow(unused_imports)]
use super::*;
use crate::domain::models::Spreadsheet;


// Known sequences for autofill pattern recognition
const DAYS_SHORT: &[&str] = &["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"];
const DAYS_FULL: &[&str] = &["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"];
const MONTHS_SHORT: &[&str] = &["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
const MONTHS_FULL: &[&str] = &["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
const QUARTERS: &[&str] = &["Q1", "Q2", "Q3", "Q4"];

/// Detected autofill pattern for smart sequence continuation.
#[derive(Debug, Clone, PartialEq)]
pub enum AutofillPattern {
    /// Arithmetic sequence: start + step * index
    Arithmetic { start: f64, step: f64 },
    /// Text prefix with numeric suffix following arithmetic pattern
    PrefixedNumber { prefix: String, suffix: String, start: f64, step: f64 },
    /// Known sequence (days, months, quarters) with wrap-around
    KnownSequence { sequence: Vec<String>, start_index: usize },
    /// Simple copy of single value
    Copy { value: String },
}

impl AutofillPattern {
    /// Detection priority: Arithmetic > KnownSequence > PrefixedNumber > Copy.
    /// KnownSequence runs before PrefixedNumber so "Q1, Q2" picks quarters
    /// instead of prefix "Q" + numbers.
    pub fn detect(values: &[String]) -> Self {
        if values.is_empty() {
            return AutofillPattern::Copy { value: String::new() };
        }

        if values.len() == 1 {
            return AutofillPattern::Copy { value: values[0].clone() };
        }

        // Try arithmetic pattern first (all numeric values)
        if let Some((start, step)) = Self::try_parse_arithmetic(values) {
            return AutofillPattern::Arithmetic { start, step };
        }

        // Try known sequence pattern (days, months, quarters) BEFORE prefixed numbers
        // This ensures Q1, Q2 is detected as quarters, not "Q" + 1, 2
        if let Some((sequence, start_index)) = Self::try_match_known_sequence(values) {
            return AutofillPattern::KnownSequence { sequence, start_index };
        }

        // Try prefixed number pattern (e.g., "Item1", "Item2")
        if let Some((prefix, suffix, start, step)) = Self::try_parse_prefixed_number(values) {
            return AutofillPattern::PrefixedNumber { prefix, suffix, start, step };
        }

        // Fallback to copy
        AutofillPattern::Copy { value: values[0].clone() }
    }

    /// Generate value at given index (0-based from start of pattern).
    pub fn generate(&self, index: usize) -> String {
        match self {
            AutofillPattern::Arithmetic { start, step } => {
                let value = start + step * (index as f64);
                Self::format_number(value)
            }
            AutofillPattern::PrefixedNumber { prefix, suffix, start, step } => {
                let num = start + step * (index as f64);
                format!("{}{}{}", prefix, Self::format_number(num), suffix)
            }
            AutofillPattern::KnownSequence { sequence, start_index } => {
                let idx = (start_index + index) % sequence.len();
                sequence[idx].clone()
            }
            AutofillPattern::Copy { value } => value.clone(),
        }
    }

    /// Return description for status message.
    pub fn description(&self) -> String {
        match self {
            AutofillPattern::Arithmetic { step, .. } => {
                if *step >= 0.0 {
                    format!("arithmetic sequence (+{})", Self::format_number(*step))
                } else {
                    format!("arithmetic sequence ({})", Self::format_number(*step))
                }
            }
            AutofillPattern::PrefixedNumber { prefix, step, .. } => {
                format!("\"{}...\" sequence (+{})", prefix, Self::format_number(*step))
            }
            AutofillPattern::KnownSequence { sequence, .. } => {
                if sequence == &DAYS_SHORT.iter().map(|s| s.to_string()).collect::<Vec<_>>()
                    || sequence == &DAYS_FULL.iter().map(|s| s.to_string()).collect::<Vec<_>>() {
                    "days sequence".to_string()
                } else if sequence == &MONTHS_SHORT.iter().map(|s| s.to_string()).collect::<Vec<_>>()
                    || sequence == &MONTHS_FULL.iter().map(|s| s.to_string()).collect::<Vec<_>>() {
                    "months sequence".to_string()
                } else if sequence == &QUARTERS.iter().map(|s| s.to_string()).collect::<Vec<_>>() {
                    "quarters sequence".to_string()
                } else {
                    "known sequence".to_string()
                }
            }
            AutofillPattern::Copy { .. } => "copy".to_string(),
        }
    }

    /// Try to parse values as an arithmetic sequence.
    /// Returns (start, step) if successful.
    fn try_parse_arithmetic(values: &[String]) -> Option<(f64, f64)> {
        if values.len() < 2 {
            return None;
        }

        // Try to parse all values as numbers
        let nums: Option<Vec<f64>> = values.iter()
            .map(|s| s.trim().parse::<f64>().ok())
            .collect();

        let nums = nums?;

        // Check if differences are constant (within epsilon)
        let step = nums[1] - nums[0];
        const EPSILON: f64 = 1e-9;

        for i in 2..nums.len() {
            let diff = nums[i] - nums[i - 1];
            if (diff - step).abs() > EPSILON {
                return None;
            }
        }

        Some((nums[0], step))
    }

    /// Try to parse values as prefixed numbers (e.g., "Item1", "Item2").
    /// Returns (prefix, suffix, start, step) if successful.
    fn try_parse_prefixed_number(values: &[String]) -> Option<(String, String, f64, f64)> {
        if values.len() < 2 {
            return None;
        }

        // Extract prefix, number, and suffix from each value
        let mut parsed: Vec<(String, f64, String)> = Vec::new();

        for val in values {
            if let Some((prefix, num, suffix)) = Self::split_prefixed_number(val) {
                parsed.push((prefix, num, suffix));
            } else {
                return None;
            }
        }

        // Check that all prefixes and suffixes are the same
        let first_prefix = &parsed[0].0;
        let first_suffix = &parsed[0].2;

        for (prefix, _, suffix) in &parsed {
            if prefix != first_prefix || suffix != first_suffix {
                return None;
            }
        }

        // Extract numbers and check for arithmetic pattern
        let nums: Vec<f64> = parsed.iter().map(|(_, n, _)| *n).collect();
        let step = nums[1] - nums[0];
        const EPSILON: f64 = 1e-9;

        for i in 2..nums.len() {
            let diff = nums[i] - nums[i - 1];
            if (diff - step).abs() > EPSILON {
                return None;
            }
        }

        Some((first_prefix.clone(), first_suffix.clone(), nums[0], step))
    }

    /// Split a string into (prefix, number, suffix).
    /// E.g., "Item10" -> ("Item", 10.0, ""), "Row_5_data" -> ("Row_", 5.0, "_data")
    fn split_prefixed_number(s: &str) -> Option<(String, f64, String)> {
        // Find the first digit
        let first_digit = s.chars().position(|c| c.is_ascii_digit())?;

        // Find where the number ends
        let after_number = s[first_digit..].chars()
            .position(|c| !c.is_ascii_digit() && c != '.' && c != '-')
            .map(|i| first_digit + i)
            .unwrap_or(s.len());

        let prefix = &s[..first_digit];
        let num_str = &s[first_digit..after_number];
        let suffix = &s[after_number..];

        let num = num_str.parse::<f64>().ok()?;

        Some((prefix.to_string(), num, suffix.to_string()))
    }

    /// Try to match values against known sequences (days, months, quarters).
    /// Returns (sequence, start_index) if successful.
    fn try_match_known_sequence(values: &[String]) -> Option<(Vec<String>, usize)> {
        let all_sequences: &[&[&str]] = &[
            DAYS_SHORT, DAYS_FULL, MONTHS_SHORT, MONTHS_FULL, QUARTERS
        ];

        for seq in all_sequences {
            if let Some(start_idx) = Self::match_sequence(values, seq) {
                let owned_seq: Vec<String> = seq.iter().map(|s| s.to_string()).collect();
                return Some((owned_seq, start_idx));
            }
        }

        None
    }

    /// Check if values match a sequence starting at some index.
    /// Returns the starting index if matched.
    fn match_sequence(values: &[String], sequence: &[&str]) -> Option<usize> {
        if values.is_empty() || values.len() > sequence.len() {
            return None;
        }

        // Find where the first value matches in the sequence (case-insensitive)
        let first_lower = values[0].to_lowercase();
        let start_idx = sequence.iter()
            .position(|s| s.to_lowercase() == first_lower)?;

        // Check if all subsequent values match the sequence
        for (i, val) in values.iter().enumerate() {
            let seq_idx = (start_idx + i) % sequence.len();
            if val.to_lowercase() != sequence[seq_idx].to_lowercase() {
                return None;
            }
        }

        Some(start_idx)
    }

    /// Format a number smartly: show as integer if whole, otherwise as decimal.
    pub(super) fn format_number(n: f64) -> String {
        if n.fract().abs() < 1e-9 {
            if n.abs() < (i64::MAX as f64) {
                format!("{}", n as i64)
            } else {
                format!("{:.0}", n)
            }
        } else {
            // Remove trailing zeros
            let s = format!("{}", n);
            s.trim_end_matches('0').trim_end_matches('.').to_string()
        }
    }
}
