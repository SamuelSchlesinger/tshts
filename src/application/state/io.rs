//! Submodule of `state` — see state/mod.rs.

use super::*;

impl App {
    pub fn save_in_place_or_prompt(&mut self) {
        if let Some(filename) = self.filename.clone() {
            let result = if filename.to_lowercase().ends_with(".xlsx") {
                crate::infrastructure::xlsx::save_xlsx(&self.workbook, &filename)
                    .map(|_| filename.clone())
            } else {
                crate::infrastructure::FileRepository::save_workbook(&self.workbook, &filename)
            };
            self.set_save_result(result);
        } else {
            self.start_save_as();
        }
    }

    pub fn start_save_as(&mut self) {
        self.mode = AppMode::SaveAs;
        self.filename_input = self.filename.clone().unwrap_or_else(|| "spreadsheet.tshts".to_string());
        self.cursor_position = self.filename_input.len();
        self.status_message = None;
    }

    pub fn start_load_file(&mut self) {
        self.mode = AppMode::LoadFile;
        self.filename_input = self
            .filename
            .clone()
            .unwrap_or_else(|| "spreadsheet.tshts".to_string());
        self.cursor_position = self.filename_input.len();
        self.status_message = None;
    }

    /// Tab-completion for filename input dialogs. Currently exposed for
    /// tests and future Tab-key wiring; the input handler does not yet call it.
    #[allow(dead_code)]
    pub fn complete_filename(&mut self) {
        let input = self.filename_input.clone();
        // Split into (dir, prefix). `dir` is the directory to read; `prefix`
        // is matched against entry names. With no `/`, both default to `.`
        // and the whole input.
        let (dir_part, name_prefix): (String, String) = match input.rfind('/') {
            Some(i) => (input[..=i].to_string(), input[i + 1..].to_string()),
            None => ("".to_string(), input.clone()),
        };
        let read_root = if dir_part.is_empty() { "." } else { dir_part.as_str() };
        let mut candidates: Vec<String> = Vec::new();
        // Recent files only matter when typing from scratch (no dir prefix).
        if dir_part.is_empty() {
            for r in crate::infrastructure::recent::load() {
                if r.starts_with(&name_prefix) && r != name_prefix {
                    candidates.push(r);
                }
            }
        }
        if let Ok(read_dir) = std::fs::read_dir(read_root) {
            for entry in read_dir.flatten() {
                if let Some(name) = entry.file_name().to_str() {
                    if !name.starts_with(&name_prefix) {
                        continue;
                    }
                    // Append `/` to directory matches so a follow-up Tab
                    // descends into them.
                    let is_dir = entry
                        .file_type()
                        .map(|t| t.is_dir())
                        .unwrap_or(false);
                    let full = if is_dir {
                        format!("{}{}/", dir_part, name)
                    } else {
                        format!("{}{}", dir_part, name)
                    };
                    if full != input && !candidates.iter().any(|c| c == &full) {
                        candidates.push(full);
                    }
                }
            }
        }
        if let Some(first) = candidates.into_iter().next() {
            self.filename_input = first;
            self.cursor_position = self.filename_input.chars().count();
        }
    }

    pub fn cancel_filename_input(&mut self) {
        self.mode = AppMode::Normal;
        self.filename_input.clear();
        self.cursor_position = 0;
    }

    pub fn set_save_result(&mut self, result: Result<String, String>) {
        match result {
            Ok(filename) => {
                self.filename = Some(filename.clone());
                self.status_message = Some(format!("Saved to {}", filename));
                self.dirty = false;
                crate::infrastructure::recent::add(&filename);
            }
            Err(error) => {
                self.status_message = Some(format!("Save failed: {}", error));
            }
        }

        self.mode = AppMode::Normal;
        self.filename_input.clear();
        self.cursor_position = 0;
    }

    pub fn set_load_workbook_result(&mut self, result: Result<(Workbook, String), String>) {
        match result {
            Ok((workbook, filename)) => {
                self.workbook = workbook;
                self.filename = Some(filename.clone());
                self.selected_row = 0;
                self.selected_col = 0;
                self.scroll_row = 0;
                self.scroll_col = 0;
                self.undo_stack.clear();
                self.redo_stack.clear();
                self.dirty = false;
                crate::infrastructure::recent::add(&filename);
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

    pub fn get_save_filename(&self) -> String {
        if self.filename_input.is_empty() {
            "spreadsheet.tshts".to_string()
        } else {
            self.filename_input.clone()
        }
    }

    pub fn get_load_filename(&self) -> String {
        if self.filename_input.is_empty() {
            "spreadsheet.tshts".to_string()
        } else {
            self.filename_input.clone()
        }
    }

    pub fn start_csv_export(&mut self) {
        self.mode = AppMode::ExportCsv;
        self.filename_input = self.filename
            .as_ref()
            .map(|f| f.replace(".tshts", ".csv"))
            .unwrap_or_else(|| "spreadsheet.csv".to_string());
        self.cursor_position = self.filename_input.len();
        self.status_message = None;
    }

    pub fn get_csv_export_filename(&self) -> String {
        if self.filename_input.is_empty() {
            "spreadsheet.csv".to_string()
        } else {
            self.filename_input.clone()
        }
    }

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

    pub fn start_csv_import(&mut self) {
        self.mode = AppMode::ImportCsv;
        self.filename_input = "data.csv".to_string();
        self.cursor_position = self.filename_input.len();
        self.status_message = None;
    }

    pub fn get_csv_import_filename(&self) -> String {
        if self.filename_input.is_empty() {
            "data.csv".to_string()
        } else {
            self.filename_input.clone()
        }
    }

    pub fn set_csv_import_result(&mut self, result: Result<Spreadsheet, String>) {
        match result {
            Ok(spreadsheet) => {
                *self.workbook.current_sheet_mut() = spreadsheet;
                self.selected_row = 0;
                self.selected_col = 0;
                self.scroll_row = 0;
                self.scroll_col = 0;
                self.undo_stack.clear();
                self.redo_stack.clear();
                self.dirty = true; // CSV data is in-memory only until saved
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
