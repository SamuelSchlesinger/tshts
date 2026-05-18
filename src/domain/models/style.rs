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
