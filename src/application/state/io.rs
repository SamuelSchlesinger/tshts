//! File and CSV import/export state transitions.

use crate::domain::{Spreadsheet, Workbook};
use super::{App, AppMode};

impl App {
    /// Switches to save-as mode to prompt for a filename.
    ///
    /// Initializes the filename input with the current filename or default.
    pub fn start_save_as(&mut self) {
        self.mode = AppMode::SaveAs;
        self.filename_input = self.filename.clone().unwrap_or_else(|| "spreadsheet.tshts".to_string());
        self.cursor_position = self.filename_input.len();
        self.status_message = None;
    }

    /// Switches to load-file mode to prompt for a filename.
    ///
    /// Initializes the filename input with the current filename or default.
    pub fn start_load_file(&mut self) {
        self.mode = AppMode::LoadFile;
        self.filename_input = self.filename.clone().unwrap_or_else(|| "spreadsheet.tshts".to_string());
        self.cursor_position = self.filename_input.len();
        self.status_message = None;
    }

    /// Cancels filename input and returns to normal mode.
    ///
    /// Clears the filename input buffer and resets cursor position.
    pub fn cancel_filename_input(&mut self) {
        self.mode = AppMode::Normal;
        self.filename_input.clear();
        self.cursor_position = 0;
    }

    /// Processes the result of a save operation.
    pub fn set_save_result(&mut self, result: Result<String, String>) {
        match result {
            Ok(filename) => {
                self.filename = Some(filename.clone());
                self.status_message = Some(format!("Saved to {}", filename));
            }
            Err(error) => {
                self.status_message = Some(format!("Save failed: {}", error));
            }
        }

        self.mode = AppMode::Normal;
        self.filename_input.clear();
        self.cursor_position = 0;
    }

    /// Processes the result of a workbook load operation.
    pub fn set_load_workbook_result(&mut self, result: Result<(Workbook, String), String>) {
        match result {
            Ok((workbook, filename)) => {
                self.workbook = workbook;
                self.filename = Some(filename.clone());
                self.selected_row = 0;
                self.selected_col = 0;
                self.scroll_row = 0;
                self.scroll_col = 0;
                self.status_message = Some(format!("Loaded from {}", filename));
            }
            Err(error) => {
                self.status_message = Some(format!("Load failed: {}", error));
            }
        }

        self.mode = AppMode::Normal;
        self.filename_input.clear();
        self.cursor_position = 0;
    }

    /// Gets the filename to use for saving.
    pub fn get_save_filename(&self) -> String {
        if self.filename_input.is_empty() {
            "spreadsheet.tshts".to_string()
        } else {
            self.filename_input.clone()
        }
    }

    /// Gets the filename to use for loading.
    pub fn get_load_filename(&self) -> String {
        if self.filename_input.is_empty() {
            "spreadsheet.tshts".to_string()
        } else {
            self.filename_input.clone()
        }
    }

    /// Switches to CSV export mode to prompt for a filename.
    pub fn start_csv_export(&mut self) {
        self.mode = AppMode::ExportCsv;
        self.filename_input = self.filename
            .as_ref()
            .map(|f| f.replace(".tshts", ".csv"))
            .unwrap_or_else(|| "spreadsheet.csv".to_string());
        self.cursor_position = self.filename_input.len();
        self.status_message = None;
    }

    /// Gets the filename to use for CSV export.
    pub fn get_csv_export_filename(&self) -> String {
        if self.filename_input.is_empty() {
            "spreadsheet.csv".to_string()
        } else {
            self.filename_input.clone()
        }
    }

    /// Processes the result of a CSV export operation.
    pub fn set_csv_export_result(&mut self, result: Result<String, String>) {
        match result {
            Ok(filename) => {
                self.status_message = Some(format!("Exported to {}", filename));
            }
            Err(error) => {
                self.status_message = Some(format!("Export failed: {}", error));
            }
        }

        self.mode = AppMode::Normal;
        self.filename_input.clear();
        self.cursor_position = 0;
    }

    /// Switches to CSV import mode to prompt for a filename.
    pub fn start_csv_import(&mut self) {
        self.mode = AppMode::ImportCsv;
        self.filename_input = "data.csv".to_string();
        self.cursor_position = self.filename_input.len();
        self.status_message = None;
    }

    /// Gets the filename to use for CSV import.
    pub fn get_csv_import_filename(&self) -> String {
        if self.filename_input.is_empty() {
            "data.csv".to_string()
        } else {
            self.filename_input.clone()
        }
    }

    /// Processes the result of a CSV import operation.
    pub fn set_csv_import_result(&mut self, result: Result<Spreadsheet, String>) {
        match result {
            Ok(spreadsheet) => {
                *self.workbook.current_sheet_mut() = spreadsheet;
                self.selected_row = 0;
                self.selected_col = 0;
                self.scroll_row = 0;
                self.scroll_col = 0;
                self.status_message = Some("CSV data imported successfully".to_string());
            }
            Err(error) => {
                self.status_message = Some(format!("Import failed: {}", error));
            }
        }

        self.mode = AppMode::Normal;
        self.filename_input.clear();
        self.cursor_position = 0;
    }
}
