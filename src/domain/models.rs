//! Domain models for the terminal spreadsheet application.
//!
//! This module contains the core data structures that represent
//! spreadsheet cells and the spreadsheet itself.

use std::collections::HashMap;
use serde::{Deserialize, Serialize};

/// Represents the data contained within a single spreadsheet cell.
///
/// Each cell can contain both a display value and an optional formula.
/// When a formula is present, the value represents the evaluated result
/// of that formula.
///
/// # Examples
///
/// ```
/// use tshts::domain::CellData;
///
/// // Simple value cell
/// let cell = CellData {
///     value: "42".to_string(),
///     formula: None,
/// };
///
/// // Formula cell
/// let formula_cell = CellData {
///     value: "84".to_string(),
///     formula: Some("=A1*2".to_string()),
/// };
/// ```
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CellData {
    /// The display value of the cell (either user input or formula result)
    pub value: String,
    /// Optional formula that generates the value (starts with '=')
    pub formula: Option<String>,
}

impl Default for CellData {
    fn default() -> Self {
        Self {
            value: String::new(),
            formula: None,
        }
    }
}

/// The main spreadsheet data structure containing cells and metadata.
///
/// A spreadsheet is organized as a grid of cells with configurable dimensions
/// and column widths. It supports formulas, cell references, and automatic
/// column width adjustment.
///
/// # Examples
///
/// ```
/// use tshts::domain::{Spreadsheet, CellData};
///
/// let mut sheet = Spreadsheet::default();
/// let cell = CellData {
///     value: "Hello".to_string(),
///     formula: None,
/// };
/// sheet.set_cell(0, 0, cell);
/// ```
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Spreadsheet {
    /// Cell data stored as a sparse matrix using (row, col) coordinates
    #[serde(serialize_with = "serialize_cells", deserialize_with = "deserialize_cells")]
    pub cells: HashMap<(usize, usize), CellData>,
    /// Maximum number of rows in the spreadsheet
    pub rows: usize,
    /// Maximum number of columns in the spreadsheet
    pub cols: usize,
    /// Custom column widths for specific columns
    pub column_widths: HashMap<usize, usize>,
    /// Default width for columns without custom widths
    pub default_column_width: usize,
}

impl Default for Spreadsheet {
    fn default() -> Self {
        Self {
            cells: HashMap::new(),
            rows: 100,
            cols: 26,
            column_widths: HashMap::new(),
            default_column_width: 8,
        }
    }
}

impl Spreadsheet {
    /// Retrieves the cell data at the specified coordinates.
    ///
    /// If no cell exists at the given position, returns a default empty cell.
    ///
    /// # Arguments
    ///
    /// * `row` - Zero-based row index
    /// * `col` - Zero-based column index
    ///
    /// # Examples
    ///
    /// ```
    /// use tshts::domain::Spreadsheet;
    ///
    /// let sheet = Spreadsheet::default();
    /// let cell = sheet.get_cell(0, 0);
    /// assert!(cell.value.is_empty());
    /// ```
    pub fn get_cell(&self, row: usize, col: usize) -> CellData {
        self.cells.get(&(row, col)).cloned().unwrap_or_default()
    }

    /// Sets the cell data at the specified coordinates.
    ///
    /// This method also automatically adjusts the column width if the new
    /// content is wider than the current column width.
    ///
    /// # Arguments
    ///
    /// * `row` - Zero-based row index
    /// * `col` - Zero-based column index
    /// * `data` - Cell data to store
    ///
    /// # Examples
    ///
    /// ```
    /// use tshts::domain::{Spreadsheet, CellData};
    ///
    /// let mut sheet = Spreadsheet::default();
    /// let cell = CellData {
    ///     value: "Hello World".to_string(),
    ///     formula: None,
    /// };
    /// sheet.set_cell(0, 0, cell);
    /// ```
    pub fn set_cell(&mut self, row: usize, col: usize, data: CellData) {
        self.cells.insert((row, col), data.clone());
        
        let current_width = self.get_column_width(col);
        let value_width = data.value.len();
        let formula_width = data.formula.as_ref().map(|f| f.len()).unwrap_or(0);
        let content_width = value_width.max(formula_width);
        let header_width = Self::column_label(col).len();
        let needed_width = content_width.max(header_width).max(3).min(50);
        
        if needed_width > current_width {
            self.set_column_width(col, needed_width);
        }
    }

    /// Retrieves the numeric value of a cell for use in formula calculations.
    ///
    /// Attempts to parse the cell's value as a floating-point number.
    /// Returns 0.0 if the cell is empty or contains non-numeric data.
    ///
    /// # Arguments
    ///
    /// * `row` - Zero-based row index
    /// * `col` - Zero-based column index
    ///
    /// # Returns
    ///
    /// The numeric value of the cell, or 0.0 if parsing fails
    pub fn get_cell_value_for_formula(&self, row: usize, col: usize) -> f64 {
        let cell = self.get_cell(row, col);
        cell.value.parse::<f64>().unwrap_or(0.0)
    }

    /// Converts a zero-based column index to an Excel-style column label.
    ///
    /// Uses the standard spreadsheet convention: A, B, C, ..., Z, AA, AB, etc.
    ///
    /// # Arguments
    ///
    /// * `col` - Zero-based column index
    ///
    /// # Returns
    ///
    /// String representation of the column (e.g., "A", "B", "AA")
    ///
    /// # Examples
    ///
    /// ```
    /// use tshts::domain::Spreadsheet;
    ///
    /// assert_eq!(Spreadsheet::column_label(0), "A");
    /// assert_eq!(Spreadsheet::column_label(25), "Z");
    /// assert_eq!(Spreadsheet::column_label(26), "AA");
    /// ```
    pub fn column_label(col: usize) -> String {
        let mut result = String::new();
        let mut c = col;
        loop {
            result = char::from(b'A' + (c % 26) as u8).to_string() + &result;
            if c < 26 {
                break;
            }
            c = c / 26 - 1;
        }
        result
    }

    /// Parses a cell reference string into row and column coordinates.
    ///
    /// Accepts Excel-style cell references like "A1", "B2", "AA123", etc.
    /// Returns None if the reference format is invalid.
    ///
    /// # Arguments
    ///
    /// * `cell_ref` - Cell reference string (e.g., "A1", "B2")
    ///
    /// # Returns
    ///
    /// Option containing (row, col) tuple with zero-based indices,
    /// or None if parsing fails
    ///
    /// # Examples
    ///
    /// ```
    /// use tshts::domain::Spreadsheet;
    ///
    /// assert_eq!(Spreadsheet::parse_cell_reference("A1"), Some((0, 0)));
    /// assert_eq!(Spreadsheet::parse_cell_reference("B2"), Some((1, 1)));
    /// assert_eq!(Spreadsheet::parse_cell_reference("invalid"), None);
    /// ```
    pub fn parse_cell_reference(cell_ref: &str) -> Option<(usize, usize)> {
        if cell_ref.is_empty() {
            return None;
        }
        
        let mut chars = cell_ref.chars();
        let mut col_str = String::new();
        let mut row_str = String::new();
        
        for ch in chars.by_ref() {
            if ch.is_ascii_alphabetic() {
                col_str.push(ch.to_ascii_uppercase());
            } else if ch.is_ascii_digit() {
                row_str.push(ch);
                break;
            } else {
                return None;
            }
        }
        
        for ch in chars {
            if ch.is_ascii_digit() {
                row_str.push(ch);
            } else {
                return None;
            }
        }
        
        if col_str.is_empty() || row_str.is_empty() {
            return None;
        }
        
        let col = Self::column_str_to_index(&col_str)?;
        let row = row_str.parse::<usize>().ok()?.checked_sub(1)?;
        
        Some((row, col))
    }
    
    /// Converts a column label string to a zero-based column index.
    ///
    /// Helper function for parsing cell references.
    ///
    /// # Arguments
    ///
    /// * `col_str` - Column label (e.g., "A", "B", "AA")
    ///
    /// # Returns
    ///
    /// Option containing zero-based column index, or None if invalid
    fn column_str_to_index(col_str: &str) -> Option<usize> {
        if col_str.is_empty() {
            return None;
        }
        
        let mut result = 0;
        for ch in col_str.chars() {
            if !ch.is_ascii_alphabetic() {
                return None;
            }
            result = result * 26 + (ch as usize - 'A' as usize + 1);
        }
        Some(result - 1)
    }

    /// Gets the display width for a specific column.
    ///
    /// Returns the custom width if set, otherwise returns the default width.
    ///
    /// # Arguments
    ///
    /// * `col` - Zero-based column index
    ///
    /// # Returns
    ///
    /// Width in characters for the column
    pub fn get_column_width(&self, col: usize) -> usize {
        self.column_widths.get(&col).copied().unwrap_or(self.default_column_width)
    }

    /// Sets the display width for a specific column.
    ///
    /// # Arguments
    ///
    /// * `col` - Zero-based column index
    /// * `width` - Width in characters
    pub fn set_column_width(&mut self, col: usize, width: usize) {
        self.column_widths.insert(col, width);
    }

    /// Automatically resizes a column to fit its content.
    ///
    /// Examines all cells in the column and adjusts the width to accommodate
    /// the longest content, with a minimum of 3 characters and maximum of 50.
    ///
    /// # Arguments
    ///
    /// * `col` - Zero-based column index
    pub fn auto_resize_column(&mut self, col: usize) {
        let current_width = self.get_column_width(col);
        let mut max_width = Self::column_label(col).len().max(current_width);
        
        for row in 0..self.rows {
            let cell = self.get_cell(row, col);
            let value_width = cell.value.len();
            let formula_width = cell.formula.as_ref().map(|f| f.len()).unwrap_or(0);
            let content_width = value_width.max(formula_width);
            max_width = max_width.max(content_width);
        }
        
        max_width = max_width.max(3).min(50);
        if max_width > current_width {
            self.set_column_width(col, max_width);
        }
    }

    /// Automatically resizes all columns to fit their content.
    ///
    /// Calls `auto_resize_column` for each column in the spreadsheet.
    pub fn auto_resize_all_columns(&mut self) {
        for col in 0..self.cols {
            self.auto_resize_column(col);
        }
    }
}

fn serialize_cells<S>(cells: &HashMap<(usize, usize), CellData>, serializer: S) -> Result<S::Ok, S::Error>
where
    S: serde::Serializer,
{
    use serde::ser::SerializeSeq;
    let mut seq = serializer.serialize_seq(Some(cells.len()))?;
    for (key, value) in cells {
        seq.serialize_element(&(key.0, key.1, value))?;
    }
    seq.end()
}

fn deserialize_cells<'de, D>(deserializer: D) -> Result<HashMap<(usize, usize), CellData>, D::Error>
where
    D: serde::Deserializer<'de>,
{
    use serde::de::{SeqAccess, Visitor};
    use std::fmt;

    struct CellsVisitor;

    impl<'de> Visitor<'de> for CellsVisitor {
        type Value = HashMap<(usize, usize), CellData>;

        fn expecting(&self, formatter: &mut fmt::Formatter) -> fmt::Result {
            formatter.write_str("a sequence of cell data")
        }

        fn visit_seq<A>(self, mut seq: A) -> Result<Self::Value, A::Error>
        where
            A: SeqAccess<'de>,
        {
            let mut cells = HashMap::new();
            while let Some((row, col, data)) = seq.next_element::<(usize, usize, CellData)>()? {
                cells.insert((row, col), data);
            }
            Ok(cells)
        }
    }

    deserializer.deserialize_seq(CellsVisitor)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_cell_data_default() {
        let cell = CellData::default();
        assert!(cell.value.is_empty());
        assert!(cell.formula.is_none());
    }

    #[test]
    fn test_cell_data_creation() {
        let cell = CellData {
            value: "42".to_string(),
            formula: Some("=6*7".to_string()),
        };
        assert_eq!(cell.value, "42");
        assert_eq!(cell.formula.unwrap(), "=6*7");
    }

    #[test]
    fn test_spreadsheet_default() {
        let sheet = Spreadsheet::default();
        assert_eq!(sheet.rows, 100);
        assert_eq!(sheet.cols, 26);
        assert_eq!(sheet.default_column_width, 8);
        assert!(sheet.cells.is_empty());
        assert!(sheet.column_widths.is_empty());
    }

    #[test]
    fn test_get_cell_empty() {
        let sheet = Spreadsheet::default();
        let cell = sheet.get_cell(0, 0);
        assert!(cell.value.is_empty());
        assert!(cell.formula.is_none());
    }

    #[test]
    fn test_set_and_get_cell() {
        let mut sheet = Spreadsheet::default();
        let cell_data = CellData {
            value: "Hello".to_string(),
            formula: None,
        };
        sheet.set_cell(0, 0, cell_data.clone());
        
        let retrieved = sheet.get_cell(0, 0);
        assert_eq!(retrieved.value, "Hello");
        assert!(retrieved.formula.is_none());
    }

    #[test]
    fn test_set_cell_adjusts_column_width() {
        let mut sheet = Spreadsheet::default();
        let initial_width = sheet.get_column_width(0);
        
        let long_cell = CellData {
            value: "This is a very long cell value".to_string(),
            formula: None,
        };
        sheet.set_cell(0, 0, long_cell);
        
        let new_width = sheet.get_column_width(0);
        assert!(new_width > initial_width);
    }

    #[test]
    fn test_get_cell_value_for_formula() {
        let mut sheet = Spreadsheet::default();
        
        // Test numeric value
        let numeric_cell = CellData {
            value: "42.5".to_string(),
            formula: None,
        };
        sheet.set_cell(0, 0, numeric_cell);
        assert_eq!(sheet.get_cell_value_for_formula(0, 0), 42.5);
        
        // Test non-numeric value
        let text_cell = CellData {
            value: "hello".to_string(),
            formula: None,
        };
        sheet.set_cell(1, 0, text_cell);
        assert_eq!(sheet.get_cell_value_for_formula(1, 0), 0.0);
        
        // Test empty cell
        assert_eq!(sheet.get_cell_value_for_formula(2, 0), 0.0);
    }

    #[test]
    fn test_column_label() {
        assert_eq!(Spreadsheet::column_label(0), "A");
        assert_eq!(Spreadsheet::column_label(1), "B");
        assert_eq!(Spreadsheet::column_label(25), "Z");
        assert_eq!(Spreadsheet::column_label(26), "AA");
        assert_eq!(Spreadsheet::column_label(27), "AB");
        assert_eq!(Spreadsheet::column_label(51), "AZ");
        assert_eq!(Spreadsheet::column_label(52), "BA");
        assert_eq!(Spreadsheet::column_label(701), "ZZ");
        assert_eq!(Spreadsheet::column_label(702), "AAA");
    }

    #[test]
    fn test_parse_cell_reference() {
        // Valid references
        assert_eq!(Spreadsheet::parse_cell_reference("A1"), Some((0, 0)));
        assert_eq!(Spreadsheet::parse_cell_reference("B2"), Some((1, 1)));
        assert_eq!(Spreadsheet::parse_cell_reference("Z26"), Some((25, 25)));
        assert_eq!(Spreadsheet::parse_cell_reference("AA1"), Some((0, 26)));
        assert_eq!(Spreadsheet::parse_cell_reference("AB100"), Some((99, 27)));
        
        // Case insensitive
        assert_eq!(Spreadsheet::parse_cell_reference("a1"), Some((0, 0)));
        assert_eq!(Spreadsheet::parse_cell_reference("b2"), Some((1, 1)));
        
        // Invalid references
        assert_eq!(Spreadsheet::parse_cell_reference(""), None);
        assert_eq!(Spreadsheet::parse_cell_reference("A"), None);
        assert_eq!(Spreadsheet::parse_cell_reference("1"), None);
        assert_eq!(Spreadsheet::parse_cell_reference("A0"), None); // Row 0 doesn't exist in Excel notation
        assert_eq!(Spreadsheet::parse_cell_reference("1A"), None);
        assert_eq!(Spreadsheet::parse_cell_reference("A1B"), None);
        assert_eq!(Spreadsheet::parse_cell_reference("A-1"), None);
    }

    #[test]
    fn test_column_width_management() {
        let mut sheet = Spreadsheet::default();
        
        // Test default width
        assert_eq!(sheet.get_column_width(0), 8);
        
        // Test setting custom width
        sheet.set_column_width(0, 15);
        assert_eq!(sheet.get_column_width(0), 15);
        
        // Test other columns still use default
        assert_eq!(sheet.get_column_width(1), 8);
    }

    #[test]
    fn test_auto_resize_column() {
        let mut sheet = Spreadsheet::default();
        
        // Add cells with varying lengths
        sheet.set_cell(0, 0, CellData { value: "Hi".to_string(), formula: None });
        sheet.set_cell(1, 0, CellData { value: "Medium length".to_string(), formula: None });
        sheet.set_cell(2, 0, CellData { value: "Very long content here".to_string(), formula: None });
        
        sheet.auto_resize_column(0);
        let width = sheet.get_column_width(0);
        
        // Should be at least as wide as the longest content
        assert!(width >= "Very long content here".len());
        // But not more than the maximum of 50
        assert!(width <= 50);
    }

    #[test]
    fn test_auto_resize_all_columns() {
        let mut sheet = Spreadsheet::default();
        
        // Add content to multiple columns
        sheet.set_cell(0, 0, CellData { value: "Short".to_string(), formula: None });
        sheet.set_cell(0, 1, CellData { value: "Much longer content".to_string(), formula: None });
        sheet.set_cell(0, 2, CellData { value: "X".to_string(), formula: None });
        
        sheet.auto_resize_all_columns();
        
        // Each column should be sized appropriately
        assert!(sheet.get_column_width(0) >= 5); // "Short".len()
        assert!(sheet.get_column_width(1) >= 19); // "Much longer content".len()
        assert!(sheet.get_column_width(2) >= 3); // Minimum width
    }

    #[test]
    fn test_formula_cell_with_auto_resize() {
        let mut sheet = Spreadsheet::default();
        
        let formula_cell = CellData {
            value: "42".to_string(),
            formula: Some("=SUM(A1:A10)".to_string()),
        };
        
        sheet.set_cell(0, 0, formula_cell);
        let width = sheet.get_column_width(0);
        
        // Width should account for both value and formula length
        assert!(width >= "=SUM(A1:A10)".len());
    }

    #[test]
    fn test_serialization_roundtrip() {
        let mut original = Spreadsheet::default();
        original.set_cell(0, 0, CellData {
            value: "test".to_string(),
            formula: Some("=1+1".to_string()),
        });
        original.set_cell(1, 1, CellData {
            value: "42".to_string(),
            formula: None,
        });
        original.set_column_width(0, 15);
        
        // Serialize to JSON
        let json = serde_json::to_string(&original).expect("Serialization failed");
        
        // Deserialize back
        let deserialized: Spreadsheet = serde_json::from_str(&json).expect("Deserialization failed");
        
        // Verify data integrity
        assert_eq!(deserialized.rows, original.rows);
        assert_eq!(deserialized.cols, original.cols);
        assert_eq!(deserialized.default_column_width, original.default_column_width);
        
        let cell_0_0 = deserialized.get_cell(0, 0);
        assert_eq!(cell_0_0.value, "test");
        assert_eq!(cell_0_0.formula.unwrap(), "=1+1");
        
        let cell_1_1 = deserialized.get_cell(1, 1);
        assert_eq!(cell_1_1.value, "42");
        assert!(cell_1_1.formula.is_none());
        
        assert_eq!(deserialized.get_column_width(0), 15);
    }
}