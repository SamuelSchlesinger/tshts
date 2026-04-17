//! Command palette mode and command dispatch.

use crate::domain::{NumberFormat, TerminalColor, Spreadsheet};
use super::{App, AppMode};

impl App {
    /// Starts command palette mode.
    pub fn start_command_palette(&mut self) {
        self.mode = AppMode::CommandPalette;
        self.command_input.clear();
        self.cursor_position = 0;
        self.status_message = None;
    }

    /// Executes a command from the command palette.
    pub fn execute_command(&mut self) {
        let trimmed = self.command_input.trim().to_string();
        if trimmed.starts_with("rename ") || trimmed.starts_with("RENAME ") {
            let name = trimmed[7..].trim().to_string();
            if !name.is_empty() {
                self.workbook.rename_sheet(name.clone());
                self.status_message = Some(format!("Renamed sheet to '{}'", name));
            } else {
                self.status_message = Some("Usage: rename <name>".to_string());
            }
            self.mode = AppMode::Normal;
            self.command_input.clear();
            self.cursor_position = 0;
            return;
        }

        let cmd = trimmed.to_lowercase();
        let parts: Vec<&str> = cmd.split_whitespace().collect();
        match parts.as_slice() {
            ["ir"] | ["insert", "row"] => self.insert_row(),
            ["dr"] | ["delete", "row"] => self.delete_row(),
            ["ic"] | ["insert", "col"] | ["insert", "column"] => self.insert_col(),
            ["dc"] | ["delete", "col"] | ["delete", "column"] => self.delete_col(),
            ["sort", "asc"] => self.sort_column_asc(),
            ["sort", "desc"] => self.sort_column_desc(),
            ["freeze"] => {
                self.frozen_rows = self.selected_row;
                self.frozen_cols = self.selected_col;
                self.status_message = Some(format!("Frozen {} rows, {} cols", self.frozen_rows, self.frozen_cols));
            }
            ["unfreeze"] => {
                self.frozen_rows = 0;
                self.frozen_cols = 0;
                self.status_message = Some("Unfrozen all panes".to_string());
            }
            ["format", "general"] => self.set_selection_format(NumberFormat::General),
            ["format", "number"] => self.set_selection_format(NumberFormat::Number { decimals: 2, thousands_sep: false }),
            ["format", "number", d] => {
                if let Ok(decimals) = d.parse::<u32>() {
                    self.set_selection_format(NumberFormat::Number { decimals, thousands_sep: false });
                } else {
                    self.status_message = Some("Invalid decimal count".to_string());
                }
            }
            ["format", "currency"] => self.set_selection_format(NumberFormat::Currency { symbol: "$".to_string(), decimals: 2 }),
            ["format", "currency", sym] => self.set_selection_format(NumberFormat::Currency { symbol: sym.to_string(), decimals: 2 }),
            ["format", "percent"] | ["format", "percentage"] => self.set_selection_format(NumberFormat::Percentage { decimals: 1 }),
            ["format", "percent", d] | ["format", "percentage", d] => {
                if let Ok(decimals) = d.parse::<u32>() {
                    self.set_selection_format(NumberFormat::Percentage { decimals });
                } else {
                    self.status_message = Some("Invalid decimal count".to_string());
                }
            }
            ["bold"] => self.toggle_bold(),
            ["underline"] => self.toggle_underline(),
            ["color", color_name] => {
                if *color_name == "none" || *color_name == "default" {
                    self.set_selection_fg_color(None);
                } else if let Some(c) = TerminalColor::from_name(color_name) {
                    self.set_selection_fg_color(Some(c));
                } else {
                    self.status_message = Some(format!("Unknown color: {}", color_name));
                }
            }
            ["bg", color_name] => {
                if *color_name == "none" || *color_name == "default" {
                    self.set_selection_bg_color(None);
                } else if let Some(c) = TerminalColor::from_name(color_name) {
                    self.set_selection_bg_color(Some(c));
                } else {
                    self.status_message = Some(format!("Unknown color: {}", color_name));
                }
            }
            ["sheet", "new"] | ["new", "sheet"] | ["addsheet"] => {
                let name = format!("Sheet{}", self.workbook.sheets.len() + 1);
                self.workbook.add_sheet(name.clone());
                self.workbook.active_sheet = self.workbook.sheets.len() - 1;
                self.selected_row = 0;
                self.selected_col = 0;
                self.scroll_row = 0;
                self.scroll_col = 0;
                self.status_message = Some(format!("Added sheet '{}'", name));
            }
            ["sheet", "delete"] | ["delsheet"] => {
                let name = self.workbook.sheet_names[self.workbook.active_sheet].clone();
                if self.workbook.remove_sheet(self.workbook.active_sheet) {
                    self.selected_row = 0;
                    self.selected_col = 0;
                    self.scroll_row = 0;
                    self.scroll_col = 0;
                    self.status_message = Some(format!("Deleted sheet '{}'", name));
                } else {
                    self.status_message = Some("Cannot delete the last sheet".to_string());
                }
            }
            ["sheet", "next"] | ["sn"] => {
                self.switch_next_sheet();
            }
            ["sheet", "prev"] | ["sp"] => {
                self.switch_prev_sheet();
            }
            ["comment", ..] => {
                let text = parts[1..].join(" ");
                if text == "clear" || text == "none" {
                    self.set_cell_comment(None);
                } else {
                    self.set_cell_comment(Some(text));
                }
            }
            ["filter", column_name] => {
                if let Some(col) = Spreadsheet::parse_column_label(column_name) {
                    self.apply_filter(col, None);
                } else {
                    self.status_message = Some(format!("Invalid column: {}", column_name));
                }
            }
            ["filter", column_name, ..] => {
                if let Some(col) = Spreadsheet::parse_column_label(column_name) {
                    let criteria = parts[2..].join(" ");
                    self.apply_filter(col, Some(criteria));
                } else {
                    self.status_message = Some(format!("Invalid column: {}", column_name));
                }
            }
            ["unfilter"] | ["clearfilter"] | ["clear", "filter"] => {
                self.clear_filter();
            }
            _ => {
                self.status_message = Some(format!("Unknown command: {}", self.command_input));
            }
        }
        self.mode = AppMode::Normal;
        self.command_input.clear();
        self.cursor_position = 0;
    }
}
