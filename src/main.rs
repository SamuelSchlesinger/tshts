use std::io;
use std::collections::HashMap;
use std::fs;
use crossterm::{
    event::{self, DisableMouseCapture, EnableMouseCapture, Event, KeyCode, KeyEventKind, KeyModifiers},
    execute,
    terminal::{disable_raw_mode, enable_raw_mode, EnterAlternateScreen, LeaveAlternateScreen},
};
use ratatui::{
    backend::{Backend, CrosstermBackend},
    layout::{Constraint, Direction, Layout, Rect},
    style::{Color, Style},
    widgets::{Block, Borders, Cell, Clear, Paragraph, Row, Table},
    Frame, Terminal,
};
use serde::{Deserialize, Serialize};

#[derive(Debug, Clone, Serialize, Deserialize)]
struct CellData {
    value: String,
    formula: Option<String>,
}

impl Default for CellData {
    fn default() -> Self {
        Self {
            value: String::new(),
            formula: None,
        }
    }
}

#[derive(Debug, Clone, Serialize, Deserialize)]
struct Spreadsheet {
    #[serde(serialize_with = "serialize_cells", deserialize_with = "deserialize_cells")]
    cells: HashMap<(usize, usize), CellData>,
    rows: usize,
    cols: usize,
    column_widths: HashMap<usize, usize>,
    default_column_width: usize,
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
    fn get_cell(&self, row: usize, col: usize) -> CellData {
        self.cells.get(&(row, col)).cloned().unwrap_or_default()
    }

    fn parse_cell_reference(cell_ref: &str) -> Option<(usize, usize)> {
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

    fn get_cell_value_for_formula(&self, row: usize, col: usize) -> f64 {
        let cell = self.get_cell(row, col);
        cell.value.parse::<f64>().unwrap_or(0.0)
    }

    fn set_cell(&mut self, row: usize, col: usize, data: CellData) {
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

    fn column_label(col: usize) -> String {
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

    fn get_column_width(&self, col: usize) -> usize {
        self.column_widths.get(&col).copied().unwrap_or(self.default_column_width)
    }

    fn set_column_width(&mut self, col: usize, width: usize) {
        self.column_widths.insert(col, width);
    }

    fn auto_resize_column(&mut self, col: usize) {
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

    fn auto_resize_all_columns(&mut self) {
        for col in 0..self.cols {
            self.auto_resize_column(col);
        }
    }
}

#[derive(Debug)]
enum AppMode {
    Normal,
    Editing,
    Help,
    SaveAs,
    LoadFile,
}

#[derive(Debug)]
struct App {
    spreadsheet: Spreadsheet,
    selected_row: usize,
    selected_col: usize,
    scroll_row: usize,
    scroll_col: usize,
    mode: AppMode,
    input: String,
    cursor_position: usize,
    filename: Option<String>,
    help_scroll: usize,
    status_message: Option<String>,
    filename_input: String,
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
    fn handle_key_event(&mut self, key: KeyCode, modifiers: KeyModifiers) {
        match self.mode {
            AppMode::Normal => self.handle_normal_mode(key, modifiers),
            AppMode::Editing => self.handle_editing_mode(key),
            AppMode::Help => self.handle_help_mode(key),
            AppMode::SaveAs => self.handle_filename_input_mode(key, true),
            AppMode::LoadFile => self.handle_filename_input_mode(key, false),
        }
    }

    fn handle_normal_mode(&mut self, key: KeyCode, modifiers: KeyModifiers) {
        if modifiers.contains(KeyModifiers::CONTROL) {
            match key {
                KeyCode::Char('s') => {
                    self.start_save_as();
                    return;
                }
                KeyCode::Char('o') => {
                    self.start_load_file();
                    return;
                }
                _ => {}
            }
        }
        
        self.status_message = None; // Clear status on any navigation
        
        match key {
            KeyCode::Up | KeyCode::Char('k') => {
                if self.selected_row > 0 {
                    self.selected_row -= 1;
                    if self.selected_row < self.scroll_row {
                        self.scroll_row = self.selected_row;
                    }
                }
            }
            KeyCode::Down | KeyCode::Char('j') => {
                if self.selected_row < self.spreadsheet.rows - 1 {
                    self.selected_row += 1;
                }
            }
            KeyCode::Left | KeyCode::Char('h') => {
                if self.selected_col > 0 {
                    self.selected_col -= 1;
                    if self.selected_col < self.scroll_col {
                        self.scroll_col = self.selected_col;
                    }
                }
            }
            KeyCode::Right | KeyCode::Char('l') => {
                if self.selected_col < self.spreadsheet.cols - 1 {
                    self.selected_col += 1;
                }
            }
            KeyCode::Enter | KeyCode::F(2) => {
                self.start_editing();
            }
            KeyCode::Char('=') => {
                self.spreadsheet.auto_resize_column(self.selected_col);
            }
            KeyCode::Char('+') => {
                self.spreadsheet.auto_resize_all_columns();
            }
            KeyCode::Char('-') => {
                let current_width = self.spreadsheet.get_column_width(self.selected_col);
                if current_width > 3 {
                    self.spreadsheet.set_column_width(self.selected_col, current_width - 1);
                }
            }
            KeyCode::Char('_') => {
                let current_width = self.spreadsheet.get_column_width(self.selected_col);
                self.spreadsheet.set_column_width(self.selected_col, current_width + 1);
            }
            KeyCode::F(1) | KeyCode::Char('?') => {
                self.mode = AppMode::Help;
                self.help_scroll = 0;
            }
            KeyCode::Char('q') => {
                // Will be handled by main loop
            }
            _ => {}
        }
    }

    fn handle_editing_mode(&mut self, key: KeyCode) {
        match key {
            KeyCode::Enter => {
                self.finish_editing();
            }
            KeyCode::Esc => {
                self.cancel_editing();
            }
            KeyCode::Backspace => {
                if self.cursor_position > 0 {
                    self.input.remove(self.cursor_position - 1);
                    self.cursor_position -= 1;
                }
            }
            KeyCode::Delete => {
                if self.cursor_position < self.input.len() {
                    self.input.remove(self.cursor_position);
                }
            }
            KeyCode::Left => {
                if self.cursor_position > 0 {
                    self.cursor_position -= 1;
                }
            }
            KeyCode::Right => {
                if self.cursor_position < self.input.len() {
                    self.cursor_position += 1;
                }
            }
            KeyCode::Home => {
                self.cursor_position = 0;
            }
            KeyCode::End => {
                self.cursor_position = self.input.len();
            }
            KeyCode::Char(c) => {
                self.input.insert(self.cursor_position, c);
                self.cursor_position += 1;
            }
            _ => {}
        }
    }

    fn handle_help_mode(&mut self, key: KeyCode) {
        match key {
            KeyCode::Esc | KeyCode::F(1) | KeyCode::Char('?') | KeyCode::Char('q') => {
                self.mode = AppMode::Normal;
            }
            KeyCode::Up | KeyCode::Char('k') => {
                if self.help_scroll > 0 {
                    self.help_scroll -= 1;
                }
            }
            KeyCode::Down | KeyCode::Char('j') => {
                self.help_scroll += 1;
            }
            KeyCode::PageUp => {
                self.help_scroll = self.help_scroll.saturating_sub(5);
            }
            KeyCode::PageDown => {
                self.help_scroll += 5;
            }
            KeyCode::Home => {
                self.help_scroll = 0;
            }
            _ => {}
        }
    }

    fn handle_filename_input_mode(&mut self, key: KeyCode, is_save: bool) {
        match key {
            KeyCode::Enter => {
                if is_save {
                    self.save_spreadsheet_with_filename();
                } else {
                    self.load_spreadsheet_with_filename();
                }
            }
            KeyCode::Esc => {
                self.cancel_filename_input();
            }
            KeyCode::Backspace => {
                if self.cursor_position > 0 {
                    self.filename_input.remove(self.cursor_position - 1);
                    self.cursor_position -= 1;
                }
            }
            KeyCode::Delete => {
                if self.cursor_position < self.filename_input.len() {
                    self.filename_input.remove(self.cursor_position);
                }
            }
            KeyCode::Left => {
                if self.cursor_position > 0 {
                    self.cursor_position -= 1;
                }
            }
            KeyCode::Right => {
                if self.cursor_position < self.filename_input.len() {
                    self.cursor_position += 1;
                }
            }
            KeyCode::Home => {
                self.cursor_position = 0;
            }
            KeyCode::End => {
                self.cursor_position = self.filename_input.len();
            }
            KeyCode::Char(c) => {
                self.filename_input.insert(self.cursor_position, c);
                self.cursor_position += 1;
            }
            _ => {}
        }
    }

    fn start_save_as(&mut self) {
        self.mode = AppMode::SaveAs;
        self.filename_input = self.filename.clone().unwrap_or_else(|| "spreadsheet.tshts".to_string());
        self.cursor_position = self.filename_input.len();
        self.status_message = None;
    }

    fn start_load_file(&mut self) {
        self.mode = AppMode::LoadFile;
        self.filename_input = self.filename.clone().unwrap_or_else(|| "spreadsheet.tshts".to_string());
        self.cursor_position = self.filename_input.len();
        self.status_message = None;
    }

    fn cancel_filename_input(&mut self) {
        self.mode = AppMode::Normal;
        self.filename_input.clear();
        self.cursor_position = 0;
    }

    fn save_spreadsheet_with_filename(&mut self) {
        let filename = if self.filename_input.is_empty() {
            "spreadsheet.tshts".to_string()
        } else {
            self.filename_input.clone()
        };
        
        match serde_json::to_string_pretty(&self.spreadsheet) {
            Ok(json) => {
                match fs::write(&filename, &json) {
                    Ok(_) => {
                        self.filename = Some(filename.clone());
                        self.status_message = Some(format!("Saved to {}", filename));
                    }
                    Err(e) => {
                        self.status_message = Some(format!("Save failed: {}", e));
                    }
                }
            }
            Err(e) => {
                self.status_message = Some(format!("Serialization failed: {}", e));
            }
        }
        
        self.mode = AppMode::Normal;
        self.filename_input.clear();
        self.cursor_position = 0;
    }

    fn load_spreadsheet_with_filename(&mut self) {
        let filename = if self.filename_input.is_empty() {
            "spreadsheet.tshts".to_string()
        } else {
            self.filename_input.clone()
        };
        
        match fs::read_to_string(&filename) {
            Ok(content) => {
                match serde_json::from_str::<Spreadsheet>(&content) {
                    Ok(spreadsheet) => {
                        self.spreadsheet = spreadsheet;
                        self.filename = Some(filename.clone());
                        self.selected_row = 0;
                        self.selected_col = 0;
                        self.scroll_row = 0;
                        self.scroll_col = 0;
                        self.status_message = Some(format!("Loaded from {}", filename));
                    }
                    Err(e) => {
                        self.status_message = Some(format!("Load failed: Invalid file format - {}", e));
                    }
                }
            }
            Err(e) => {
                self.status_message = Some(format!("Load failed: {}", e));
            }
        }
        
        self.mode = AppMode::Normal;
        self.filename_input.clear();
        self.cursor_position = 0;
    }

    fn start_editing(&mut self) {
        self.mode = AppMode::Editing;
        let cell = self.spreadsheet.get_cell(self.selected_row, self.selected_col);
        self.input = cell.formula.unwrap_or(cell.value);
        self.cursor_position = self.input.len();
    }

    fn finish_editing(&mut self) {
        let mut cell_data = CellData::default();
        
        if self.input.starts_with('=') {
            if self.would_create_circular_reference(&self.input) {
                return;
            }
            cell_data.formula = Some(self.input.clone());
            cell_data.value = self.evaluate_formula(&self.input);
        } else {
            cell_data.value = self.input.clone();
        }

        self.spreadsheet.set_cell(self.selected_row, self.selected_col, cell_data);
        self.mode = AppMode::Normal;
        self.input.clear();
        self.cursor_position = 0;
    }

    fn would_create_circular_reference(&self, formula: &str) -> bool {
        let current_cell = (self.selected_row, self.selected_col);
        self.check_circular_reference_recursive(formula, current_cell, &mut std::collections::HashSet::new())
    }

    fn check_circular_reference_recursive(&self, formula: &str, target_cell: (usize, usize), visited: &mut std::collections::HashSet<(usize, usize)>) -> bool {
        if !formula.starts_with('=') {
            return false;
        }

        let expr = &formula[1..];
        let referenced_cells = self.extract_cell_references(expr);
        
        for (row, col) in referenced_cells {
            if (row, col) == target_cell {
                return true;
            }
            
            if visited.contains(&(row, col)) {
                continue;
            }
            
            visited.insert((row, col));
            
            let cell = self.spreadsheet.get_cell(row, col);
            if let Some(ref cell_formula) = cell.formula {
                if self.check_circular_reference_recursive(cell_formula, target_cell, visited) {
                    return true;
                }
            }
            
            visited.remove(&(row, col));
        }
        
        false
    }

    fn extract_cell_references(&self, expr: &str) -> Vec<(usize, usize)> {
        let mut references = Vec::new();
        let mut current_token = String::new();
        
        for ch in expr.chars() {
            if ch.is_alphanumeric() {
                current_token.push(ch);
            } else {
                if !current_token.is_empty() {
                    if let Some((row, col)) = Spreadsheet::parse_cell_reference(&current_token) {
                        references.push((row, col));
                    }
                    current_token.clear();
                }
            }
        }
        
        if !current_token.is_empty() {
            if let Some((row, col)) = Spreadsheet::parse_cell_reference(&current_token) {
                references.push((row, col));
            }
        }
        
        references
    }

    fn cancel_editing(&mut self) {
        self.mode = AppMode::Normal;
        self.input.clear();
        self.cursor_position = 0;
    }

    fn evaluate_formula(&self, formula: &str) -> String {
        if formula.starts_with('=') {
            let expr = &formula[1..];
            
            match self.evaluate_expression(expr) {
                Ok(result) => result.to_string(),
                Err(_) => "#ERROR".to_string(),
            }
        } else {
            formula.to_string()
        }
    }

    fn evaluate_expression(&self, expr: &str) -> Result<f64, String> {
        let expr = expr.trim();
        
        if let Ok(result) = expr.parse::<f64>() {
            return Ok(result);
        }
        
        if let Some((row, col)) = Spreadsheet::parse_cell_reference(expr) {
            return Ok(self.spreadsheet.get_cell_value_for_formula(row, col));
        }
        
        if expr.to_uppercase().starts_with("SUM(") || expr.to_uppercase().starts_with("AVERAGE(") || 
           expr.to_uppercase().starts_with("MIN(") || expr.to_uppercase().starts_with("MAX(") ||
           expr.to_uppercase().starts_with("IF(") || expr.to_uppercase().starts_with("AND(") ||
           expr.to_uppercase().starts_with("OR(") || expr.to_uppercase().starts_with("NOT(") {
            return self.evaluate_function(expr);
        }
        
        if expr.contains("**") {
            let parts: Vec<&str> = expr.split("**").collect();
            if parts.len() == 2 {
                let a = self.evaluate_operand(parts[0].trim())?;
                let b = self.evaluate_operand(parts[1].trim())?;
                return Ok(a.powf(b));
            }
        }
        
        if expr.contains("^") {
            let parts: Vec<&str> = expr.split("^").collect();
            if parts.len() == 2 {
                let a = self.evaluate_operand(parts[0].trim())?;
                let b = self.evaluate_operand(parts[1].trim())?;
                return Ok(a.powf(b));
            }
        }
        
        if expr.contains("%") {
            let parts: Vec<&str> = expr.split("%").collect();
            if parts.len() == 2 {
                let a = self.evaluate_operand(parts[0].trim())?;
                let b = self.evaluate_operand(parts[1].trim())?;
                if b == 0.0 {
                    return Err("Modulo by zero".to_string());
                }
                return Ok(a % b);
            }
        }
        
        if expr.contains("<=") {
            let parts: Vec<&str> = expr.split("<=").collect();
            if parts.len() == 2 {
                let a = self.evaluate_operand(parts[0].trim())?;
                let b = self.evaluate_operand(parts[1].trim())?;
                return Ok(if a <= b { 1.0 } else { 0.0 });
            }
        }
        
        if expr.contains(">=") {
            let parts: Vec<&str> = expr.split(">=").collect();
            if parts.len() == 2 {
                let a = self.evaluate_operand(parts[0].trim())?;
                let b = self.evaluate_operand(parts[1].trim())?;
                return Ok(if a >= b { 1.0 } else { 0.0 });
            }
        }
        
        if expr.contains("<>") {
            let parts: Vec<&str> = expr.split("<>").collect();
            if parts.len() == 2 {
                let a = self.evaluate_operand(parts[0].trim())?;
                let b = self.evaluate_operand(parts[1].trim())?;
                return Ok(if (a - b).abs() > f64::EPSILON { 1.0 } else { 0.0 });
            }
        }
        
        if expr.contains("<") {
            let parts: Vec<&str> = expr.split("<").collect();
            if parts.len() == 2 {
                let a = self.evaluate_operand(parts[0].trim())?;
                let b = self.evaluate_operand(parts[1].trim())?;
                return Ok(if a < b { 1.0 } else { 0.0 });
            }
        }
        
        if expr.contains(">") {
            let parts: Vec<&str> = expr.split(">").collect();
            if parts.len() == 2 {
                let a = self.evaluate_operand(parts[0].trim())?;
                let b = self.evaluate_operand(parts[1].trim())?;
                return Ok(if a > b { 1.0 } else { 0.0 });
            }
        }
        
        if expr.contains('+') {
            let parts: Vec<&str> = expr.split('+').collect();
            if parts.len() == 2 {
                let a = self.evaluate_operand(parts[0].trim())?;
                let b = self.evaluate_operand(parts[1].trim())?;
                return Ok(a + b);
            }
        }
        
        if expr.contains('-') && !expr.starts_with('-') {
            let parts: Vec<&str> = expr.split('-').collect();
            if parts.len() == 2 {
                let a = self.evaluate_operand(parts[0].trim())?;
                let b = self.evaluate_operand(parts[1].trim())?;
                return Ok(a - b);
            }
        }
        
        if expr.contains('*') && !expr.contains("**") {
            let parts: Vec<&str> = expr.split('*').collect();
            if parts.len() == 2 {
                let a = self.evaluate_operand(parts[0].trim())?;
                let b = self.evaluate_operand(parts[1].trim())?;
                return Ok(a * b);
            }
        }
        
        if expr.contains('/') {
            let parts: Vec<&str> = expr.split('/').collect();
            if parts.len() == 2 {
                let a = self.evaluate_operand(parts[0].trim())?;
                let b = self.evaluate_operand(parts[1].trim())?;
                if b == 0.0 {
                    return Err("Division by zero".to_string());
                }
                return Ok(a / b);
            }
        }
        
        Err("Invalid expression".to_string())
    }

    fn evaluate_function(&self, expr: &str) -> Result<f64, String> {
        let expr = expr.trim().to_uppercase();
        
        if expr.starts_with("SUM(") && expr.ends_with(')') {
            let args = &expr[4..expr.len()-1];
            let values = self.parse_function_args(args)?;
            return Ok(values.iter().sum());
        }
        
        if expr.starts_with("AVERAGE(") && expr.ends_with(')') {
            let args = &expr[8..expr.len()-1];
            let values = self.parse_function_args(args)?;
            if values.is_empty() {
                return Err("No values for average".to_string());
            }
            return Ok(values.iter().sum::<f64>() / values.len() as f64);
        }
        
        if expr.starts_with("MIN(") && expr.ends_with(')') {
            let args = &expr[4..expr.len()-1];
            let values = self.parse_function_args(args)?;
            if values.is_empty() {
                return Err("No values for min".to_string());
            }
            return Ok(values.iter().fold(f64::INFINITY, |a, &b| a.min(b)));
        }
        
        if expr.starts_with("MAX(") && expr.ends_with(')') {
            let args = &expr[4..expr.len()-1];
            let values = self.parse_function_args(args)?;
            if values.is_empty() {
                return Err("No values for max".to_string());
            }
            return Ok(values.iter().fold(f64::NEG_INFINITY, |a, &b| a.max(b)));
        }
        
        if expr.starts_with("IF(") && expr.ends_with(')') {
            let args = &expr[3..expr.len()-1];
            let parts: Vec<&str> = args.split(',').collect();
            if parts.len() == 3 {
                let condition = self.evaluate_operand(parts[0].trim())?;
                let true_val = self.evaluate_operand(parts[1].trim())?;
                let false_val = self.evaluate_operand(parts[2].trim())?;
                return Ok(if condition != 0.0 { true_val } else { false_val });
            }
        }
        
        if expr.starts_with("AND(") && expr.ends_with(')') {
            let args = &expr[4..expr.len()-1];
            let values = self.parse_function_args(args)?;
            return Ok(if values.iter().all(|&x| x != 0.0) { 1.0 } else { 0.0 });
        }
        
        if expr.starts_with("OR(") && expr.ends_with(')') {
            let args = &expr[3..expr.len()-1];
            let values = self.parse_function_args(args)?;
            return Ok(if values.iter().any(|&x| x != 0.0) { 1.0 } else { 0.0 });
        }
        
        if expr.starts_with("NOT(") && expr.ends_with(')') {
            let args = &expr[4..expr.len()-1];
            let value = self.evaluate_operand(args.trim())?;
            return Ok(if value == 0.0 { 1.0 } else { 0.0 });
        }
        
        Err("Unknown function".to_string())
    }

    fn parse_function_args(&self, args: &str) -> Result<Vec<f64>, String> {
        let mut values = Vec::new();
        for arg in args.split(',') {
            let arg = arg.trim();
            if arg.contains(':') {
                let range_values = self.parse_range(arg)?;
                values.extend(range_values);
            } else {
                values.push(self.evaluate_operand(arg)?);
            }
        }
        Ok(values)
    }

    fn parse_range(&self, range: &str) -> Result<Vec<f64>, String> {
        let parts: Vec<&str> = range.split(':').collect();
        if parts.len() != 2 {
            return Err("Invalid range".to_string());
        }
        
        let start = Spreadsheet::parse_cell_reference(parts[0].trim())
            .ok_or("Invalid start cell")?;
        let end = Spreadsheet::parse_cell_reference(parts[1].trim())
            .ok_or("Invalid end cell")?;
        
        let mut values = Vec::new();
        for row in start.0..=end.0 {
            for col in start.1..=end.1 {
                values.push(self.spreadsheet.get_cell_value_for_formula(row, col));
            }
        }
        
        Ok(values)
    }


    fn evaluate_operand(&self, operand: &str) -> Result<f64, String> {
        let operand = operand.trim();
        
        if let Ok(num) = operand.parse::<f64>() {
            return Ok(num);
        }
        
        if let Some((row, col)) = Spreadsheet::parse_cell_reference(operand) {
            return Ok(self.spreadsheet.get_cell_value_for_formula(row, col));
        }
        
        Err("Invalid operand".to_string())
    }
}

fn ui(f: &mut Frame, app: &App) {
    let chunks = Layout::default()
        .direction(Direction::Vertical)
        .constraints([
            Constraint::Length(1),
            Constraint::Min(0),
            Constraint::Length(3),
        ])
        .split(f.area());

    let header = Paragraph::new(format!(
        "tshts - Terminal Spreadsheet | Cell: {}{}",
        Spreadsheet::column_label(app.selected_col),
        app.selected_row + 1
    ))
    .style(Style::default().fg(Color::Cyan));
    f.render_widget(header, chunks[0]);

    let visible_rows = chunks[1].height as usize - 1;
    
    let mut total_width = 4;
    let mut visible_cols = 0;
    let available_width = chunks[1].width as usize;
    
    for col in app.scroll_col..app.spreadsheet.cols {
        let col_width = app.spreadsheet.get_column_width(col);
        if total_width + col_width + 1 > available_width {
            break;
        }
        total_width += col_width + 1;
        visible_cols += 1;
    }
    
    let mut headers = vec![Cell::from("")];
    for col in app.scroll_col..app.scroll_col + visible_cols {
        headers.push(Cell::from(Spreadsheet::column_label(col)).style(Style::default().fg(Color::Yellow)));
    }

    let header_row = Row::new(headers).height(1);
    
    let mut rows = vec![header_row];
    
    for row in app.scroll_row..std::cmp::min(app.scroll_row + visible_rows, app.spreadsheet.rows) {
        let mut cells = vec![Cell::from(format!("{}", row + 1)).style(Style::default().fg(Color::Yellow))];
        
        for col in app.scroll_col..app.scroll_col + visible_cols {
            let cell_data = app.spreadsheet.get_cell(row, col);
            let cell_value = if cell_data.value.is_empty() { " ".to_string() } else { cell_data.value };
            
            let style = if row == app.selected_row && col == app.selected_col {
                Style::default().bg(Color::Blue).fg(Color::White)
            } else {
                Style::default()
            };
            
            cells.push(Cell::from(cell_value).style(style));
        }
        
        rows.push(Row::new(cells).height(1));
    }

    let mut widths = vec![Constraint::Length(4)];
    for col in app.scroll_col..app.scroll_col + visible_cols {
        widths.push(Constraint::Length(app.spreadsheet.get_column_width(col) as u16));
    }
    let table = Table::new(rows, widths)
        .block(Block::default().borders(Borders::ALL).title("Spreadsheet"))
        .column_spacing(1);

    f.render_widget(table, chunks[1]);

    let input_text = match app.mode {
        AppMode::Normal => {
            if let Some(ref status) = app.status_message {
                status.clone()
            } else {
                let filename = app.filename.as_ref().map(|f| f.as_str()).unwrap_or("unsaved");
                format!("File: {} | Ctrl+S: save | Ctrl+O: load | F1/?: help | q: quit", filename)
            }
        }
        AppMode::Editing => format!("Editing: {} (Enter to save, Esc to cancel)", app.input),
        AppMode::Help => "↑↓/jk: scroll | PgUp/PgDn: fast scroll | Home: top | Esc/q: close help".to_string(),
        AppMode::SaveAs => format!("Save as: {} (Enter to save, Esc to cancel)", app.filename_input),
        AppMode::LoadFile => format!("Load file: {} (Enter to load, Esc to cancel)", app.filename_input),
    };

    let input = Paragraph::new(input_text)
        .block(Block::default().borders(Borders::ALL).title("Status"))
        .style(match app.mode {
            AppMode::Normal => Style::default(),
            AppMode::Editing => Style::default().fg(Color::Green),
            AppMode::Help => Style::default().fg(Color::Cyan),
            AppMode::SaveAs => Style::default().fg(Color::Yellow),
            AppMode::LoadFile => Style::default().fg(Color::Yellow),
        });
    f.render_widget(input, chunks[2]);

    if matches!(app.mode, AppMode::Help) {
        render_help_popup(f, app.help_scroll);
    }
}

fn render_help_popup(f: &mut Frame, scroll: usize) {
    let area = f.area();
    let popup_area = Rect {
        x: area.width / 10,
        y: area.height / 10,
        width: area.width * 4 / 5,
        height: area.height * 4 / 5,
    };

    f.render_widget(Clear, popup_area);
    
    let help_text = get_help_text();
    let help_lines: Vec<&str> = help_text.lines().collect();
    let visible_height = popup_area.height.saturating_sub(2) as usize; // Account for borders
    
    let start_line = scroll.min(help_lines.len().saturating_sub(visible_height));
    let end_line = (start_line + visible_height).min(help_lines.len());
    
    let visible_text = help_lines[start_line..end_line].join("\n");
    
    let help_widget = Paragraph::new(visible_text)
        .block(Block::default()
            .borders(Borders::ALL)
            .title(format!("tshts Expression Language Help (Line {}/{})", start_line + 1, help_lines.len()))
            .style(Style::default().fg(Color::Cyan)))
        .style(Style::default().fg(Color::White));
    
    f.render_widget(help_widget, popup_area);
}

fn get_help_text() -> String {
    r#"TSHTS EXPRESSION LANGUAGE REFERENCE

=== BASIC CONCEPTS ===
• All formulas start with = (equals sign)
• Cell references use column letter + row number (A1, B2, Z99, AA1, etc.)
• Numbers can be integers or decimals (42, 3.14, -5.5)
• Case insensitive for functions and cell references

=== ARITHMETIC OPERATORS ===
+       Addition                    =5+3 → 8, =A1+B1
-       Subtraction                 =10-3 → 7, =A1-5
*       Multiplication              =4*3 → 12, =A1*B1
/       Division                    =15/3 → 5, =A1/B1
**      Exponentiation              =2**3 → 8, =A1**2
^       Power (same as **)          =3^2 → 9, =A1^B1
%       Modulo (remainder)          =10%3 → 1, =A1%B1

=== COMPARISON OPERATORS ===
<       Less than                   =A1<B1 → 1 or 0
>       Greater than                =A1>B1 → 1 or 0
<=      Less than or equal          =A1<=B1 → 1 or 0
>=      Greater than or equal       =A1>=B1 → 1 or 0
<>      Not equal                   =A1<>B1 → 1 or 0

Note: Comparisons return 1 for true, 0 for false

=== BASIC FUNCTIONS ===
SUM(...)        Sum of values           =SUM(A1,B1,C1) or =SUM(A1:C1)
AVERAGE(...)    Average of values       =AVERAGE(A1:A10)
MIN(...)        Minimum value           =MIN(A1,B1,C1,5)
MAX(...)        Maximum value           =MAX(A1:C3)

=== LOGICAL FUNCTIONS ===
IF(cond,true,false) Conditional         =IF(A1>5,100,0)
AND(...)        All values true         =AND(A1>0,B1<10)
OR(...)         Any value true          =OR(A1=0,B1=0)
NOT(value)      Logical not             =NOT(A1>5)

Note: 0 is false, anything else is true

=== CELL RANGES ===
A1:C3           Rectangle from A1 to C3
A1:A10          Column A, rows 1-10
B2:D2           Row 2, columns B-D

=== EXAMPLES ===
=A1+B1*2        Math with precedence
=IF(A1>0,A1*2,0) Conditional calculation
=SUM(A1:A5)/5   Same as AVERAGE(A1:A5)
=MAX(A1:C3)     Largest in 3x3 range
=A1**2+B1**2    Pythagorean calculation

=== FILE OPERATIONS ===
Ctrl+S          Save spreadsheet to file
Ctrl+O          Load spreadsheet from file
                Files are saved as "spreadsheet.tshts" in JSON format

=== NAVIGATION SHORTCUTS ===
F1 or ?         Show this help (scroll with ↑↓, PgUp/PgDn, Home)
Enter/F2        Edit selected cell
Arrow keys      Navigate cells (hjkl also work)
= key           Auto-resize column to fit content
+ key           Auto-resize all columns to fit content
- / _ keys      Manually shrink/grow column width
q               Quit application

=== HELP NAVIGATION ===
↑↓ or j/k       Scroll help text up/down one line
Page Up/Down    Scroll help text up/down 5 lines
Home            Jump to top of help text
Esc/F1/?/q      Close this help window

Note: Your spreadsheet is automatically saved when you use Ctrl+S.
Use Ctrl+O to load the saved spreadsheet on next session."#.to_string()
}

fn main() -> Result<(), Box<dyn std::error::Error>> {
    // Test save functionality first
    if std::env::args().any(|arg| arg == "--test-save") {
        test_save_functionality();
        return Ok(());
    }

    enable_raw_mode()?;
    let mut stdout = io::stdout();
    execute!(stdout, EnterAlternateScreen, EnableMouseCapture)?;
    let backend = CrosstermBackend::new(stdout);
    let mut terminal = Terminal::new(backend)?;

    let mut app = App::default();
    let res = run_app(&mut terminal, &mut app);

    disable_raw_mode()?;
    execute!(
        terminal.backend_mut(),
        LeaveAlternateScreen,
        DisableMouseCapture
    )?;
    terminal.show_cursor()?;

    if let Err(err) = res {
        println!("{err:?}");
    }

    Ok(())
}

fn test_save_functionality() {
    println!("Testing save functionality...");
    
    let mut app = App::default();
    
    // Add some test data
    let test_cell = CellData {
        value: "42".to_string(),
        formula: Some("=6*7".to_string()),
    };
    app.spreadsheet.set_cell(0, 0, test_cell);
    
    let test_cell2 = CellData {
        value: "Hello World".to_string(),
        formula: None,
    };
    app.spreadsheet.set_cell(1, 2, test_cell2);
    
    // Test save (simulate the new workflow)
    app.filename_input = "test_spreadsheet.tshts".to_string();
    app.save_spreadsheet_with_filename();
    
    if let Some(ref msg) = app.status_message {
        println!("Save result: {}", msg);
    } else {
        println!("No status message after save");
    }
    
    // Check if file exists
    if std::path::Path::new("test_spreadsheet.tshts").exists() {
        println!("✓ Save file created successfully");
        
        // Test load
        let mut app2 = App::default();
        app2.filename_input = "test_spreadsheet.tshts".to_string();
        app2.load_spreadsheet_with_filename();
        
        if let Some(ref msg) = app2.status_message {
            println!("Load result: {}", msg);
        }
        
        // Verify data
        let loaded_cell = app2.spreadsheet.get_cell(0, 0);
        println!("Loaded cell (0,0): value='{}', formula={:?}", loaded_cell.value, loaded_cell.formula);
        
        let loaded_cell2 = app2.spreadsheet.get_cell(1, 2);
        println!("Loaded cell (1,2): value='{}', formula={:?}", loaded_cell2.value, loaded_cell2.formula);
        
        if loaded_cell.value == "42" && loaded_cell2.value == "Hello World" {
            println!("✓ Save/load functionality working correctly");
        } else {
            println!("✗ Data mismatch after load");
        }
        
    } else {
        println!("✗ Save file was not created");
    }
}

fn run_app<B: Backend>(terminal: &mut Terminal<B>, app: &mut App) -> io::Result<()> {
    loop {
        terminal.draw(|f| ui(f, app))?;

        if let Event::Key(key) = event::read()? {
            if key.kind == KeyEventKind::Press {
                match key.code {
                    KeyCode::Char('q') if matches!(app.mode, AppMode::Normal) => return Ok(()),
                    _ => app.handle_key_event(key.code, key.modifiers),
                }
            }
        }
    }
}