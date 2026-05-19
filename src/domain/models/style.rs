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
pub struct CellStyle {
    pub bold: bool,
    pub underline: bool,
    pub fg_color: Option<TerminalColor>,
    pub bg_color: Option<TerminalColor>,
}

impl Default for CellStyle {
    fn default() -> Self {
        Self { bold: false, underline: false, fg_color: None, bg_color: None }
    }
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
                let formatted = format!("{:.prec$}", n, prec = *decimals as usize);
                format!("{}{}", symbol, add_thousands_separator(&formatted))
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
