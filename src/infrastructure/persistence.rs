//! File persistence layer for spreadsheet data.
//!
//! This module handles saving and loading spreadsheet data to/from JSON files.
//! It provides a simple file-based storage mechanism for the application.

use crate::domain::{Spreadsheet, Workbook};
use std::fs;

/// Repository for file-based spreadsheet persistence.
///
/// Handles serialization and deserialization of spreadsheet data using JSON format.
/// All operations are synchronous and use the standard filesystem.
///
/// # Examples
///
/// ```
/// use tshts::infrastructure::FileRepository;
/// use tshts::domain::Spreadsheet;
///
/// let sheet = Spreadsheet::default();
/// let result = FileRepository::save_spreadsheet(&sheet, "test.tshts");
/// ```
pub struct FileRepository;

impl FileRepository {
    /// Saves a spreadsheet to a JSON file.
    ///
    /// Serializes the spreadsheet data to JSON format and writes it to the specified file.
    /// The file extension .tshts is recommended but not enforced.
    ///
    /// # Arguments
    ///
    /// * `spreadsheet` - Reference to the spreadsheet to save
    /// * `filename` - Path where the file should be saved
    ///
    /// # Returns
    ///
    /// Result containing the filename on success, or error message on failure
    ///
    /// # Examples
    ///
    /// ```
    /// use tshts::infrastructure::FileRepository;
    /// use tshts::domain::Spreadsheet;
    ///
    /// let sheet = Spreadsheet::default();
    /// match FileRepository::save_spreadsheet(&sheet, "my_sheet.tshts") {
    ///     Ok(filename) => println!("Saved to {}", filename),
    ///     Err(error) => println!("Save failed: {}", error),
    /// }
    /// ```
    #[allow(dead_code)] // Used in tests for single-sheet persistence
    pub fn save_spreadsheet(spreadsheet: &Spreadsheet, filename: &str) -> Result<String, String> {
        match serde_json::to_string_pretty(spreadsheet) {
            Ok(json) => {
                match fs::write(filename, &json) {
                    Ok(_) => Ok(filename.to_string()),
                    Err(e) => Err(e.to_string()),
                }
            }
            Err(e) => Err(format!("Serialization failed: {}", e)),
        }
    }

    /// Loads a spreadsheet from a JSON file.
    ///
    /// Reads and deserializes spreadsheet data from the specified file.
    /// The file must contain valid JSON in the expected spreadsheet format.
    ///
    /// # Arguments
    ///
    /// * `filename` - Path to the file to load
    ///
    /// # Returns
    ///
    /// Result containing (spreadsheet, filename) tuple on success, or error message on failure
    ///
    /// # Examples
    ///
    /// ```
    /// use tshts::infrastructure::FileRepository;
    ///
    /// match FileRepository::load_spreadsheet("my_sheet.tshts") {
    ///     Ok((sheet, filename)) => {
    ///         println!("Loaded {} rows from {}", sheet.rows, filename);
    ///     }
    ///     Err(error) => println!("Load failed: {}", error),
    /// }
    /// ```
    #[allow(dead_code)] // Used in tests for single-sheet persistence
    pub fn load_spreadsheet(filename: &str) -> Result<(Spreadsheet, String), String> {
        match fs::read_to_string(filename) {
            Ok(content) => {
                match serde_json::from_str::<Spreadsheet>(&content) {
                    Ok(mut spreadsheet) => {
                        // Rebuild dependencies since they're not serialized
                        spreadsheet.rebuild_dependencies();
                        Ok((spreadsheet, filename.to_string()))
                    }
                    Err(e) => Err(format!("Invalid file format - {}", e)),
                }
            }
            Err(e) => Err(e.to_string()),
        }
    }

    /// Saves a workbook to a JSON file.
    pub fn save_workbook(workbook: &Workbook, filename: &str) -> Result<String, String> {
        match serde_json::to_string_pretty(workbook) {
            Ok(json) => {
                match fs::write(filename, &json) {
                    Ok(_) => Ok(filename.to_string()),
                    Err(e) => Err(e.to_string()),
                }
            }
            Err(e) => Err(format!("Serialization failed: {}", e)),
        }
    }

    /// Loads a workbook from a JSON file.
    /// Handles both old Spreadsheet format and new Workbook format.
    pub fn load_workbook(filename: &str) -> Result<(Workbook, String), String> {
        match fs::read_to_string(filename) {
            Ok(content) => {
                // Try workbook format first
                if let Ok(mut workbook) = serde_json::from_str::<Workbook>(&content) {
                    for sheet in &mut workbook.sheets {
                        sheet.rebuild_dependencies();
                    }
                    return Ok((workbook, filename.to_string()));
                }
                // Fall back to single spreadsheet format
                match serde_json::from_str::<Spreadsheet>(&content) {
                    Ok(mut spreadsheet) => {
                        spreadsheet.rebuild_dependencies();
                        Ok((Workbook::from_spreadsheet(spreadsheet), filename.to_string()))
                    }
                    Err(e) => Err(format!("Invalid file format - {}", e)),
                }
            }
            Err(e) => Err(e.to_string()),
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::domain::CellData;
    use tempfile::NamedTempFile;
    use std::io::Write;

    fn create_test_spreadsheet() -> Spreadsheet {
        let mut sheet = Spreadsheet::default();
        
        // Add some test data
        sheet.set_cell(0, 0, CellData {
            value: "Hello".to_string(),
            formula: None,
            format: None,
            comment: None,
        });
        
        sheet.set_cell(1, 1, CellData {
            value: "42".to_string(),
            formula: Some("=6*7".to_string()),
            format: None,
            comment: None,
        });
        
        sheet.set_cell(2, 0, CellData {
            value: "World".to_string(),
            formula: None,
            format: None,
            comment: None,
        });
        
        // Set custom column width
        sheet.set_column_width(0, 15);
        sheet.set_column_width(1, 10);
        
        sheet
    }

    #[test]
    fn test_save_and_load_roundtrip() {
        let original_sheet = create_test_spreadsheet();
        let temp_file = NamedTempFile::new().expect("Failed to create temp file");
        let file_path = temp_file.path().to_str().unwrap();
        
        // Save the spreadsheet
        let save_result = FileRepository::save_spreadsheet(&original_sheet, file_path);
        assert!(save_result.is_ok());
        assert_eq!(save_result.unwrap(), file_path);
        
        // Load the spreadsheet back
        let load_result = FileRepository::load_spreadsheet(file_path);
        assert!(load_result.is_ok());
        
        let (loaded_sheet, loaded_filename) = load_result.unwrap();
        assert_eq!(loaded_filename, file_path);
        
        // Verify the data matches
        assert_eq!(loaded_sheet.rows, original_sheet.rows);
        assert_eq!(loaded_sheet.cols, original_sheet.cols);
        assert_eq!(loaded_sheet.default_column_width, original_sheet.default_column_width);
        
        // Check specific cells
        let cell_0_0 = loaded_sheet.get_cell(0, 0);
        assert_eq!(cell_0_0.value, "Hello");
        assert!(cell_0_0.formula.is_none());
        
        let cell_1_1 = loaded_sheet.get_cell(1, 1);
        assert_eq!(cell_1_1.value, "42");
        assert_eq!(cell_1_1.formula.unwrap(), "=6*7");
        
        let cell_2_0 = loaded_sheet.get_cell(2, 0);
        assert_eq!(cell_2_0.value, "World");
        assert!(cell_2_0.formula.is_none());
        
        // Check column widths
        assert_eq!(loaded_sheet.get_column_width(0), 15);
        assert_eq!(loaded_sheet.get_column_width(1), 10);
        assert_eq!(loaded_sheet.get_column_width(2), original_sheet.default_column_width);
    }

    #[test]
    fn test_save_empty_spreadsheet() {
        let empty_sheet = Spreadsheet::default();
        let temp_file = NamedTempFile::new().expect("Failed to create temp file");
        let file_path = temp_file.path().to_str().unwrap();
        
        let save_result = FileRepository::save_spreadsheet(&empty_sheet, file_path);
        assert!(save_result.is_ok());
        
        let load_result = FileRepository::load_spreadsheet(file_path);
        assert!(load_result.is_ok());
        
        let (loaded_sheet, _) = load_result.unwrap();
        assert_eq!(loaded_sheet.rows, empty_sheet.rows);
        assert_eq!(loaded_sheet.cols, empty_sheet.cols);
        assert!(loaded_sheet.cells.is_empty());
    }

    #[test]
    fn test_save_invalid_path() {
        let sheet = create_test_spreadsheet();
        let invalid_path = "/nonexistent/directory/file.tshts";
        
        let result = FileRepository::save_spreadsheet(&sheet, invalid_path);
        assert!(result.is_err());
        let error_msg = result.unwrap_err();
        assert!(error_msg.contains("No such file or directory") ||
                error_msg.contains("cannot find the path"));
    }

    #[test]
    fn test_load_nonexistent_file() {
        let result = FileRepository::load_spreadsheet("/nonexistent/file.tshts");
        assert!(result.is_err());
        let error_msg = result.unwrap_err();
        assert!(error_msg.contains("No such file or directory") ||
                error_msg.contains("cannot find the path"));
    }

    #[test]
    fn test_load_invalid_json() {
        let mut temp_file = NamedTempFile::new().expect("Failed to create temp file");
        writeln!(temp_file, "{{invalid json content}}").expect("Failed to write to temp file");
        let file_path = temp_file.path().to_str().unwrap();
        
        let result = FileRepository::load_spreadsheet(file_path);
        assert!(result.is_err());
        assert!(result.unwrap_err().contains("Invalid file format"));
    }

    #[test]
    fn test_load_valid_json_wrong_format() {
        let mut temp_file = NamedTempFile::new().expect("Failed to create temp file");
        writeln!(temp_file, r#"{{"wrong": "format"}}"#).expect("Failed to write to temp file");
        let file_path = temp_file.path().to_str().unwrap();
        
        let result = FileRepository::load_spreadsheet(file_path);
        assert!(result.is_err());
        assert!(result.unwrap_err().contains("Invalid file format"));
    }

    #[test]
    fn test_save_with_special_characters() {
        let mut sheet = Spreadsheet::default();
        
        // Add cells with special characters
        sheet.set_cell(0, 0, CellData {
            value: "Héllo Wörld! 🌍".to_string(),
            formula: None,
            format: None,
            comment: None,
        });
        
        sheet.set_cell(1, 0, CellData {
            value: "100".to_string(),
            formula: Some("=SUM(\"quotes\", 'apostrophes', `backticks`)".to_string()),
            format: None,
            comment: None,
        });
        
        let temp_file = NamedTempFile::new().expect("Failed to create temp file");
        let file_path = temp_file.path().to_str().unwrap();
        
        let save_result = FileRepository::save_spreadsheet(&sheet, file_path);
        assert!(save_result.is_ok());
        
        let load_result = FileRepository::load_spreadsheet(file_path);
        assert!(load_result.is_ok());
        
        let (loaded_sheet, _) = load_result.unwrap();
        
        let cell_0_0 = loaded_sheet.get_cell(0, 0);
        assert_eq!(cell_0_0.value, "Héllo Wörld! 🌍");
        
        let cell_1_0 = loaded_sheet.get_cell(1, 0);
        assert_eq!(cell_1_0.value, "100");
        assert_eq!(cell_1_0.formula.unwrap(), "=SUM(\"quotes\", 'apostrophes', `backticks`)");
    }

    #[test]
    fn test_save_large_spreadsheet() {
        let mut sheet = Spreadsheet::default();
        
        // Fill a decent number of cells
        for row in 0..10 {
            for col in 0..10 {
                sheet.set_cell(row, col, CellData {
                    value: format!("R{}C{}", row + 1, col + 1),
                    formula: if row == 0 && col > 0 {
                        Some(format!("=A{}", col + 1))
                    } else {
                        None
                    },
                    format: None,
                    comment: None,
                });
            }
        }
        
        let temp_file = NamedTempFile::new().expect("Failed to create temp file");
        let file_path = temp_file.path().to_str().unwrap();
        
        let save_result = FileRepository::save_spreadsheet(&sheet, file_path);
        assert!(save_result.is_ok());
        
        let load_result = FileRepository::load_spreadsheet(file_path);
        assert!(load_result.is_ok());
        
        let (loaded_sheet, _) = load_result.unwrap();
        
        // Verify a few cells
        assert_eq!(loaded_sheet.get_cell(0, 0).value, "R1C1");
        assert_eq!(loaded_sheet.get_cell(9, 9).value, "R10C10");
        assert_eq!(loaded_sheet.get_cell(0, 5).formula.unwrap(), "=A6");
    }

    #[test]
    fn test_save_custom_dimensions() {
        let mut sheet = Spreadsheet::default();
        sheet.rows = 50;
        sheet.cols = 15;
        sheet.default_column_width = 12;
        
        let temp_file = NamedTempFile::new().expect("Failed to create temp file");
        let file_path = temp_file.path().to_str().unwrap();
        
        let save_result = FileRepository::save_spreadsheet(&sheet, file_path);
        assert!(save_result.is_ok());
        
        let load_result = FileRepository::load_spreadsheet(file_path);
        assert!(load_result.is_ok());
        
        let (loaded_sheet, _) = load_result.unwrap();
        assert_eq!(loaded_sheet.rows, 50);
        assert_eq!(loaded_sheet.cols, 15);
        assert_eq!(loaded_sheet.default_column_width, 12);
    }

    #[test]
    fn test_json_format_structure() {
        let sheet = create_test_spreadsheet();
        let temp_file = NamedTempFile::new().expect("Failed to create temp file");
        let file_path = temp_file.path().to_str().unwrap();
        
        FileRepository::save_spreadsheet(&sheet, file_path).expect("Save failed");
        
        // Read the raw JSON and verify it has the expected structure
        let json_content = std::fs::read_to_string(file_path).expect("Failed to read file");
        let json_value: serde_json::Value = serde_json::from_str(&json_content).expect("Invalid JSON");
        
        // Check that main fields exist
        assert!(json_value.get("cells").is_some());
        assert!(json_value.get("rows").is_some());
        assert!(json_value.get("cols").is_some());
        assert!(json_value.get("column_widths").is_some());
        assert!(json_value.get("default_column_width").is_some());
        
        // Verify cells are stored as array format (due to custom serialization)
        let cells = json_value.get("cells").unwrap();
        assert!(cells.is_array());
    }

    // === Workbook Persistence Tests ===

    #[test]
    fn test_save_load_workbook_roundtrip() {
        let mut workbook = Workbook::default();
        workbook.current_sheet_mut().set_cell(0, 0, CellData {
            value: "Sheet1".to_string(), formula: None, format: None, comment: None,
        });
        workbook.add_sheet("Data".to_string());
        workbook.active_sheet = 1;
        workbook.current_sheet_mut().set_cell(0, 0, CellData {
            value: "Sheet2".to_string(), formula: None, format: None, comment: None,
        });
        workbook.active_sheet = 0; // Reset before save

        let temp_file = NamedTempFile::new().expect("Failed to create temp file");
        let file_path = temp_file.path().to_str().unwrap();

        let save_result = FileRepository::save_workbook(&workbook, file_path);
        assert!(save_result.is_ok());

        let (loaded, _) = FileRepository::load_workbook(file_path).unwrap();
        assert_eq!(loaded.sheets.len(), 2);
        assert_eq!(loaded.sheet_names[0], "Sheet1");
        assert_eq!(loaded.sheet_names[1], "Data");
        assert_eq!(loaded.sheets[0].get_cell(0, 0).value, "Sheet1");
        assert_eq!(loaded.sheets[1].get_cell(0, 0).value, "Sheet2");
    }

    #[test]
    fn test_load_workbook_from_old_spreadsheet_format() {
        // Save a single spreadsheet in old format
        let mut sheet = Spreadsheet::default();
        sheet.set_cell(0, 0, CellData {
            value: "Legacy".to_string(), formula: None, format: None, comment: None,
        });

        let temp_file = NamedTempFile::new().expect("Failed to create temp file");
        let file_path = temp_file.path().to_str().unwrap();

        FileRepository::save_spreadsheet(&sheet, file_path).unwrap();

        // Load as workbook - should wrap in a single-sheet workbook
        let (loaded, _) = FileRepository::load_workbook(file_path).unwrap();
        assert_eq!(loaded.sheets.len(), 1);
        assert_eq!(loaded.sheet_names[0], "Sheet1");
        assert_eq!(loaded.sheets[0].get_cell(0, 0).value, "Legacy");
    }
}