use crate::domain::{Spreadsheet, CellData, FormulaEvaluator};

#[derive(Debug)]
pub enum AppMode {
    Normal,
    Editing,
    Help,
    SaveAs,
    LoadFile,
}

#[derive(Debug)]
pub struct App {
    pub spreadsheet: Spreadsheet,
    pub selected_row: usize,
    pub selected_col: usize,
    pub scroll_row: usize,
    pub scroll_col: usize,
    pub mode: AppMode,
    pub input: String,
    pub cursor_position: usize,
    pub filename: Option<String>,
    pub help_scroll: usize,
    pub status_message: Option<String>,
    pub filename_input: String,
}

impl Default for App {
    fn default() -> Self {
        Self {
            spreadsheet: Spreadsheet::default(),
            selected_row: 0,
            selected_col: 0,
            scroll_row: 0,
            scroll_col: 0,
            mode: AppMode::Normal,
            input: String::new(),
            cursor_position: 0,
            filename: None,
            help_scroll: 0,
            status_message: None,
            filename_input: String::new(),
        }
    }
}

impl App {
    pub fn start_editing(&mut self) {
        self.mode = AppMode::Editing;
        let cell = self.spreadsheet.get_cell(self.selected_row, self.selected_col);
        self.input = cell.formula.unwrap_or(cell.value);
        self.cursor_position = self.input.len();
    }

    pub fn finish_editing(&mut self) {
        let mut cell_data = CellData::default();
        
        if self.input.starts_with('=') {
            let evaluator = FormulaEvaluator::new(&self.spreadsheet);
            if evaluator.would_create_circular_reference(&self.input, (self.selected_row, self.selected_col)) {
                return;
            }
            cell_data.formula = Some(self.input.clone());
            cell_data.value = evaluator.evaluate_formula(&self.input);
        } else {
            cell_data.value = self.input.clone();
        }

        self.spreadsheet.set_cell(self.selected_row, self.selected_col, cell_data);
        self.mode = AppMode::Normal;
        self.input.clear();
        self.cursor_position = 0;
    }

    pub fn cancel_editing(&mut self) {
        self.mode = AppMode::Normal;
        self.input.clear();
        self.cursor_position = 0;
    }

    pub fn start_save_as(&mut self) {
        self.mode = AppMode::SaveAs;
        self.filename_input = self.filename.clone().unwrap_or_else(|| "spreadsheet.tshts".to_string());
        self.cursor_position = self.filename_input.len();
        self.status_message = None;
    }

    pub fn start_load_file(&mut self) {
        self.mode = AppMode::LoadFile;
        self.filename_input = self.filename.clone().unwrap_or_else(|| "spreadsheet.tshts".to_string());
        self.cursor_position = self.filename_input.len();
        self.status_message = None;
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
            }
            Err(error) => {
                self.status_message = Some(format!("Save failed: {}", error));
            }
        }
        
        self.mode = AppMode::Normal;
        self.filename_input.clear();
        self.cursor_position = 0;
    }

    pub fn set_load_result(&mut self, result: Result<(Spreadsheet, String), String>) {
        match result {
            Ok((spreadsheet, filename)) => {
                self.spreadsheet = spreadsheet;
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
}