//! Submodule of `models` — see models/mod.rs.

#![allow(unused_imports)]
use serde::{Deserialize, Serialize};
use std::collections::{HashMap, HashSet, VecDeque};
use super::*;

/// Number format for cell display.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub enum NumberFormat {
    /// Default rendering (no formatting)
    General,
    /// Fixed decimal places with optional thousands separator
    Number { decimals: u32, thousands_sep: bool },
    /// Currency with symbol and decimal places
    Currency { symbol: String, decimals: u32 },
    /// Percentage (multiply by 100 and add %)
    Percentage { decimals: u32 },
}

/// Terminal color for cell styling.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub enum TerminalColor {
    Black,
    Red,
    Green,
    Yellow,
    Blue,
    Magenta,
    Cyan,
    White,
    DarkGray,
    LightRed,
    LightGreen,
    LightYellow,
    LightBlue,
    LightMagenta,
    LightCyan,
}

impl TerminalColor {
    /// Parses a color name string into a TerminalColor.
    pub fn from_name(name: &str) -> Option<Self> {
        match name.to_lowercase().as_str() {
            "black" => Some(Self::Black),
            "red" => Some(Self::Red),
            "green" => Some(Self::Green),
            "yellow" => Some(Self::Yellow),
            "blue" => Some(Self::Blue),
            "magenta" => Some(Self::Magenta),
            "cyan" => Some(Self::Cyan),
            "white" => Some(Self::White),
            "darkgray" | "dark_gray" => Some(Self::DarkGray),
            "lightred" | "light_red" => Some(Self::LightRed),
            "lightgreen" | "light_green" => Some(Self::LightGreen),
            "lightyellow" | "light_yellow" => Some(Self::LightYellow),
            "lightblue" | "light_blue" => Some(Self::LightBlue),
            "lightmagenta" | "light_magenta" => Some(Self::LightMagenta),
            "lightcyan" | "light_cyan" => Some(Self::LightCyan),
            _ => None,
        }
    }
}

/// Visual style for a cell.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[derive(Default)]
pub struct CellStyle {
    pub bold: bool,
    pub underline: bool,
    pub fg_color: Option<TerminalColor>,
    pub bg_color: Option<TerminalColor>,
}


/// Cell formatting options.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct CellFormat {
    /// Number format
    pub number_format: NumberFormat,
    /// Cell visual style
    pub style: CellStyle,
}

impl Default for CellFormat {
    fn default() -> Self {
        Self {
            number_format: NumberFormat::General,
            style: CellStyle::default(),
        }
    }
}

/// Formats a cell value according to the given number format.
pub fn format_cell_value(value: &str, format: &CellFormat) -> String {
    match &format.number_format {
        NumberFormat::General => value.to_string(),
        NumberFormat::Number { decimals, thousands_sep } => {
            if let Ok(n) = value.parse::<f64>() {
                let formatted = format!("{:.prec$}", n, prec = *decimals as usize);
                if *thousands_sep {
                    add_thousands_separator(&formatted)
                } else {
                    formatted
                }
            } else {
                value.to_string()
            }
        }
        NumberFormat::Currency { symbol, decimals } => {
            if let Ok(n) = value.parse::<f64>() {
                // Excel convention: sign goes BEFORE the currency symbol
                // ("-$42.50"), not between the symbol and the magnitude
                // ("$-42.50"). Format the absolute value, then prepend
                // the sign manually so the symbol always sits next to
                // the digits.
                let abs_formatted = format!("{:.prec$}", n.abs(), prec = *decimals as usize);
                let body = add_thousands_separator(&abs_formatted);
                if n < 0.0 {
                    format!("-{}{}", symbol, body)
                } else {
                    format!("{}{}", symbol, body)
                }
            } else {
                value.to_string()
            }
        }
        NumberFormat::Percentage { decimals } => {
            if let Ok(n) = value.parse::<f64>() {
                format!("{:.prec$}%", n * 100.0, prec = *decimals as usize)
            } else {
                value.to_string()
            }
        }
    }
}

pub(super) fn add_thousands_separator(s: &str) -> String {
    let parts: Vec<&str> = s.splitn(2, '.').collect();
    let int_part = parts[0];
    let negative = int_part.starts_with('-');
    let digits = if negative { &int_part[1..] } else { int_part };

    let mut result = String::new();
    for (i, c) in digits.chars().rev().enumerate() {
        if i > 0 && i % 3 == 0 {
            result.push(',');
        }
        result.push(c);
    }
    let int_formatted: String = result.chars().rev().collect();
    let prefix = if negative { "-" } else { "" };

    if parts.len() > 1 {
        format!("{}{}.{}", prefix, int_formatted, parts[1])
    } else {
        format!("{}{}", prefix, int_formatted)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::domain::{CellData};
    #[test]
    fn has_volatile_cf_predicate_detects_now_in_rule() {
        let mut sheet = Spreadsheet::default();
        // Pure predicate — not volatile.
        sheet.conditional_formats.push(ConditionalFormat {
            column: 0,
            predicate: "_ > 100".to_string(),
            style: CellStyle::default(),
        });
        assert!(!sheet.has_volatile_cf_predicate());
        // Add a NOW()-based rule — sheet is now volatile.
        sheet.conditional_formats.push(ConditionalFormat {
            column: 1,
            predicate: "NOW() > _".to_string(),
            style: CellStyle::default(),
        });
        assert!(sheet.has_volatile_cf_predicate());
    }

    #[test]
    fn recalc_clears_cf_cache_when_predicate_is_volatile() {
        use crate::domain::Workbook;
        let mut wb = Workbook::default();
        wb.sheets[0].set_cell(0, 0, CellData {
            value: "1".to_string(), formula: None, format: None, comment: None,
            spill_anchor: None,
        });
        wb.sheets[0].conditional_formats.push(ConditionalFormat {
            column: 0,
            predicate: "NOW() > _".to_string(),
            style: CellStyle { bold: true, ..CellStyle::default() },
        });
        // Prime the cache.
        let _ = wb.sheets[0].conditional_style_for(0, 0);
        assert!(!wb.sheets[0].cf_cache.lock().unwrap().is_empty(),
            "predicate eval should have populated the cache");
        // Force a recalc — even with no dirty cells, the post-recalc hook
        // must still clear the cache because the predicate is volatile.
        wb.mark_all_formula_cells_dirty();
        let _ = wb.recalc_via_graph_result();
        assert!(wb.sheets[0].cf_cache.lock().unwrap().is_empty(),
            "volatile CF predicate must invalidate cf_cache on recalc");
    }

    #[test]
    fn recalc_keeps_cf_cache_when_predicates_are_pure() {
        use crate::domain::Workbook;
        let mut wb = Workbook::default();
        wb.sheets[0].set_cell(0, 0, CellData {
            value: "150".to_string(), formula: None, format: None, comment: None,
            spill_anchor: None,
        });
        wb.sheets[0].conditional_formats.push(ConditionalFormat {
            column: 0,
            predicate: "_ > 100".to_string(),
            style: CellStyle { bold: true, ..CellStyle::default() },
        });
        // Prime the cache.
        let _ = wb.sheets[0].conditional_style_for(0, 0);
        assert!(!wb.sheets[0].cf_cache.lock().unwrap().is_empty());
        // Recalc with no dirty cells touching this sheet: cache stays
        // because all rules are pure. (A cell mutation would still
        // invalidate via `set_cell_internal`.)
        let _ = wb.recalc_via_graph_result();
        assert!(!wb.sheets[0].cf_cache.lock().unwrap().is_empty(),
            "pure CF predicates must not trigger spurious cache invalidation");
    }

    #[test]
    fn test_conditional_format_fires_on_truthy_predicate() {
        let mut sheet = Spreadsheet::default();
        sheet.set_cell(0, 0, CellData {
            value: "150".to_string(), formula: None, format: None, comment: None,
        spill_anchor: None,
        });
        sheet.set_cell(1, 0, CellData {
            value: "50".to_string(), formula: None, format: None, comment: None,
        spill_anchor: None,
        });
        sheet.conditional_formats.push(ConditionalFormat {
            column: 0,
            predicate: "_ > 100".to_string(),
            style: CellStyle {
                bold: true,
                underline: false,
                fg_color: Some(TerminalColor::Red),
                bg_color: None,
            },
        });
        let s0 = sheet.conditional_style_for(0, 0);
        assert!(s0.is_some());
        assert!(s0.as_ref().unwrap().bold);
        assert_eq!(s0.unwrap().fg_color, Some(TerminalColor::Red));
        // Row 1 doesn't satisfy the predicate.
        assert!(sheet.conditional_style_for(1, 0).is_none());
    }

    #[test]
    fn test_thousands_separator_edge_cases() {
        let fmt = CellFormat {
            number_format: NumberFormat::Number { decimals: 2, thousands_sep: true },
            style: CellStyle::default(),
        };
        assert_eq!(format_cell_value("1234567.89", &fmt), "1,234,567.89");
        assert_eq!(format_cell_value("-1234.5", &fmt), "-1,234.50");
        assert_eq!(format_cell_value("999.99", &fmt), "999.99");
        assert_eq!(format_cell_value("0", &fmt), "0.00");
        assert_eq!(format_cell_value("-0.5", &fmt), "-0.50");

        // Whole-million boundary
        let fmt0 = CellFormat {
            number_format: NumberFormat::Number { decimals: 0, thousands_sep: true },
            style: CellStyle::default(),
        };
        assert_eq!(format_cell_value("1000000", &fmt0), "1,000,000");
        assert_eq!(format_cell_value("-1000000", &fmt0), "-1,000,000");
    }

    #[test]
    fn test_format_cell_value_general() {
        let fmt = CellFormat { number_format: NumberFormat::General, ..CellFormat::default() };
        assert_eq!(super::format_cell_value("42.5", &fmt), "42.5");
        assert_eq!(super::format_cell_value("hello", &fmt), "hello");
    }

    #[test]
    fn test_format_cell_value_number() {
        let fmt = CellFormat { number_format: NumberFormat::Number { decimals: 2, thousands_sep: false }, ..CellFormat::default() };
        assert_eq!(super::format_cell_value("42", &fmt), "42.00");
        assert_eq!(super::format_cell_value("3.14159", &fmt), "3.14");
        assert_eq!(super::format_cell_value("hello", &fmt), "hello"); // non-numeric passthrough
    }

    #[test]
    fn test_format_cell_value_number_thousands() {
        let fmt = CellFormat { number_format: NumberFormat::Number { decimals: 2, thousands_sep: true }, ..CellFormat::default() };
        assert_eq!(super::format_cell_value("1234567.89", &fmt), "1,234,567.89");
        assert_eq!(super::format_cell_value("42", &fmt), "42.00");
    }

    #[test]
    fn test_format_cell_value_currency() {
        let fmt = CellFormat { number_format: NumberFormat::Currency { symbol: "$".to_string(), decimals: 2 }, ..CellFormat::default() };
        assert_eq!(super::format_cell_value("1234.5", &fmt), "$1,234.50");
        assert_eq!(super::format_cell_value("42", &fmt), "$42.00");
        assert_eq!(super::format_cell_value("hello", &fmt), "hello");
    }

    #[test]
    fn test_format_cell_value_percentage() {
        let fmt = CellFormat { number_format: NumberFormat::Percentage { decimals: 1 }, ..CellFormat::default() };
        assert_eq!(super::format_cell_value("0.75", &fmt), "75.0%");
        assert_eq!(super::format_cell_value("1", &fmt), "100.0%");
        assert_eq!(super::format_cell_value("0.123", &fmt), "12.3%");
    }

    #[test]
    fn test_thousands_separator() {
        assert_eq!(super::style::add_thousands_separator("1234567"), "1,234,567");
        assert_eq!(super::style::add_thousands_separator("123"), "123");
        assert_eq!(super::style::add_thousands_separator("1234.56"), "1,234.56");
        assert_eq!(super::style::add_thousands_separator("-1234567"), "-1,234,567");
    }

    #[test]
    fn test_cell_data_format_serialization() {
        let cell = CellData {
            value: "100".to_string(),
            formula: None,
            format: Some(CellFormat {
                number_format: NumberFormat::Percentage { decimals: 1 },
                ..CellFormat::default()
            }),
            comment: None,
        spill_anchor: None,
        };
        let json = serde_json::to_string(&cell).unwrap();
        let deserialized: CellData = serde_json::from_str(&json).unwrap();
        assert_eq!(deserialized.value, "100");
        assert!(deserialized.format.is_some());
        assert!(matches!(deserialized.format.unwrap().number_format, NumberFormat::Percentage { decimals: 1 }));
    }

}
