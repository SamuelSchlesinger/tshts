//! JSON serialization for Workbook/Spreadsheet.

use crate::domain::{Spreadsheet, Workbook};
use crate::infrastructure::atomic::atomic_write;
use std::fs;

pub struct FileRepository;

impl FileRepository {
    /// Save a single sheet. Public lib API; main loop uses `save_workbook`.
    #[allow(dead_code)]
    pub fn save_spreadsheet(spreadsheet: &Spreadsheet, filename: &str) -> Result<String, String> {
        match serde_json::to_string_pretty(spreadsheet) {
            Ok(json) => {
                match atomic_write(filename, json.as_bytes()) {
                    Ok(_) => Ok(filename.to_string()),
                    Err(e) => Err(e.to_string()),
                }
            }
            Err(e) => Err(format!("Serialization failed: {}", e)),
        }
    }

    /// Load a single sheet. Public lib API; main loop uses `load_workbook`.
    #[allow(dead_code)]
    pub fn load_spreadsheet(filename: &str) -> Result<(Spreadsheet, String), String> {
        match fs::read_to_string(filename) {
            Ok(content) => {
                match serde_json::from_str::<Spreadsheet>(&content) {
                    Ok(spreadsheet) => {
                        // Dep graph is workbook-level and rebuilt lazily
                        // on first `recalc_via_graph_result()`.
                        Ok((spreadsheet, filename.to_string()))
                    }
                    Err(e) => Err(format!("Invalid file format - {}", e)),
                }
            }
            Err(e) => Err(e.to_string()),
        }
    }

    pub fn save_workbook(workbook: &Workbook, filename: &str) -> Result<String, String> {
        match serde_json::to_string_pretty(workbook) {
            Ok(json) => {
                match atomic_write(filename, json.as_bytes()) {
                    Ok(_) => Ok(filename.to_string()),
                    Err(e) => Err(e.to_string()),
                }
            }
            Err(e) => Err(format!("Serialization failed: {}", e)),
        }
    }

    /// Loads a workbook. Falls back to single-sheet Spreadsheet format for
    /// backward compatibility with files saved before the workbook tabs feature.
    pub fn load_workbook(filename: &str) -> Result<(Workbook, String), String> {
        match fs::read_to_string(filename) {
            Ok(content) => {
                // Pre-parse to detect "looks like a tshts file" rather than
                // "happens to be valid JSON that serde fills in with all
                // defaults". A bare `{}` would otherwise deserialize as an
                // empty Workbook and we'd silently overwrite the file with
                // a default state on the next save — destroying data if
                // the source was actually a corrupted or unrelated file.
                let mut raw: serde_json::Value = serde_json::from_str(&content)
                    .map_err(|e| format!("Invalid file format - {}", e))?;
                let obj = raw.as_object().ok_or_else(|| {
                    "Invalid file format - top-level value is not an object".to_string()
                })?;
                let looks_like_workbook = obj.contains_key("sheets");
                let looks_like_spreadsheet = obj.contains_key("cells") && obj.contains_key("rows");
                if !looks_like_workbook && !looks_like_spreadsheet {
                    return Err(
                        "Invalid file format - missing 'sheets' or 'cells' field; \
                         not a tshts file".to_string(),
                    );
                }
                // Try workbook format first. Run the schema migrator before
                // typed deserialize so future schema bumps don't panic on
                // load — see `domain::models::workbook::migrate_workbook_json`.
                if looks_like_workbook {
                    crate::domain::models::migrate_workbook_json(&mut raw)?;
                }
                if let Ok(mut workbook) = serde_json::from_value::<Workbook>(raw.clone()) {
                    // Validate invariants: a file with `active_sheet` past
                    // `sheets.len()` or mismatched `sheet_names.len()` would
                    // panic on the first UI read. The order here matters:
                    // an empty-sheets file gets a single default Sheet1 BEFORE
                    // any sheet-relative bookkeeping, so the later clamps run
                    // against a non-degenerate state. We preserve the loaded
                    // `named_ranges`, `version`, etc. — only the sheets and
                    // sheet_names arrays are reset.
                    if workbook.sheets.is_empty() {
                        workbook.sheets.push(Spreadsheet::default());
                        workbook.sheet_names.clear();
                        workbook.sheet_names.push("Sheet1".to_string());
                        workbook.active_sheet = 0;
                    }
                    if workbook.sheet_names.len() < workbook.sheets.len() {
                        let needed = workbook.sheets.len() - workbook.sheet_names.len();
                        for i in 0..needed {
                            workbook
                                .sheet_names
                                .push(format!("Sheet{}", workbook.sheet_names.len() + i + 1));
                        }
                    } else if workbook.sheet_names.len() > workbook.sheets.len() {
                        workbook.sheet_names.truncate(workbook.sheets.len());
                    }
                    if workbook.active_sheet >= workbook.sheets.len() {
                        workbook.active_sheet = 0;
                    }
                    // Pre-PR-1 files have no sheet_ids; allocate now so
                    // the unified graph has stable identities. Files saved
                    // by PR-1+ carry their own IDs and ensure_* no-ops.
                    workbook.build_dep_graph_from_scratch();
                    return Ok((workbook, filename.to_string()));
                }
                // Fall back to single spreadsheet format
                match serde_json::from_str::<Spreadsheet>(&content) {
                    Ok(spreadsheet) => {
                        let mut wb = Workbook::from_spreadsheet(spreadsheet);
                        wb.build_dep_graph_from_scratch();
                        Ok((wb, filename.to_string()))
                    }
                    Err(e) => Err(format!("Invalid file format - {}", e)),
                }
            }
            Err(e) => Err(e.to_string()),
        }
    }
}

#[cfg(test)]
#[allow(clippy::field_reassign_with_default)]
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
        spill_anchor: None,
        });
        
        sheet.set_cell(1, 1, CellData {
            value: "42".to_string(),
            formula: Some("=6*7".to_string()),
            format: None,
            comment: None,
        spill_anchor: None,
        });
        
        sheet.set_cell(2, 0, CellData {
            value: "World".to_string(),
            formula: None,
            format: None,
            comment: None,
        spill_anchor: None,
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
        spill_anchor: None,
        });
        
        sheet.set_cell(1, 0, CellData {
            value: "100".to_string(),
            formula: Some("=SUM(\"quotes\", 'apostrophes', `backticks`)".to_string()),
            format: None,
            comment: None,
        spill_anchor: None,
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
                spill_anchor: None,
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
        spill_anchor: None,
        });
        workbook.add_sheet("Data".to_string());
        workbook.active_sheet = 1;
        workbook.current_sheet_mut().set_cell(0, 0, CellData {
            value: "Sheet2".to_string(), formula: None, format: None, comment: None,
        spill_anchor: None,
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
        spill_anchor: None,
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