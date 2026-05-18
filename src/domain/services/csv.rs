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
        
        let file = File::create(filename).map_err(|e| format!("Failed to create file: {}", e))?;
        let mut writer = ::csv::Writer::from_writer(file);
        
        // Export data row by row
        for row in 0..=max_row {
            let mut record = Vec::new();
            for col in 0..=max_col {
                let cell = spreadsheet.get_cell(row, col);
                record.push(cell.value.clone());
            }
            writer.write_record(&record).map_err(|e| format!("Failed to write row: {}", e))?;
        }
        
        writer.flush().map_err(|e| format!("Failed to flush CSV writer: {}", e))?;
        Ok(filename.to_string())
    }
    
    /// Reads `filename` into a fresh `Spreadsheet`; no header row is assumed.
    ///
    /// ```no_run
    /// use tshts::domain::CsvExporter;
    /// let _ = CsvExporter::import_from_csv("data.csv");
    /// ```
    pub fn import_from_csv(filename: &str) -> Result<Spreadsheet, String> {
        let file = File::open(filename).map_err(|e| format!("Failed to open file: {}", e))?;
        let mut reader = ::csv::ReaderBuilder::new()
            .has_headers(false) // Don't treat first row as headers
            .flexible(true)     // Tolerate rows with varying numbers of fields
            .from_reader(file);

        let mut spreadsheet = Spreadsheet::default();
        let mut max_row = 0;
        let mut max_col = 0;

        for (row_index, result) in reader.records().enumerate() {
            let record = result.map_err(|e| format!("Failed to read CSV row {}: {}", row_index + 1, e))?;

            for (col_index, field) in record.iter().enumerate() {
                if !field.is_empty() {
                    let formula = if field.starts_with('=') {
                        Some(field.to_string())
                    } else {
                        None
                    };
                    let cell_data = crate::domain::models::CellData {
                        value: field.to_string(),
                        formula,
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

        // Update spreadsheet dimensions based on imported data
        if max_row > 0 || max_col > 0 {
            spreadsheet.rows = spreadsheet.rows.max(max_row + 5);
            spreadsheet.cols = spreadsheet.cols.max(max_col + 5);
        }

        // Rebuild dependencies in case any imported cells contain formulas
        spreadsheet.rebuild_dependencies();

        Ok(spreadsheet)
    }

    /// Append CSV rows beneath the existing data in `dest`, starting one row
    /// below the last non-empty cell. Used by `:import-append`.
    pub fn append_from_csv(dest: &mut Spreadsheet, filename: &str) -> Result<usize, String> {
        let file = std::fs::File::open(filename).map_err(|e| e.to_string())?;
        let mut reader = ::csv::ReaderBuilder::new()
            .has_headers(false)
            .flexible(true)
            .from_reader(file);
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
                    let formula = if field.starts_with('=') {
                        Some(field.to_string())
                    } else {
                        None
                    };
                    let cell_data = crate::domain::models::CellData {
                        value: field.to_string(),
                        formula,
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
