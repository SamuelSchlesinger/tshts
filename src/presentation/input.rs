use crate::application::{App, AppMode};
use crate::infrastructure::FileRepository;
use crossterm::event::{KeyCode, KeyModifiers};

pub struct InputHandler;

impl InputHandler {
    pub fn handle_key_event(app: &mut App, key: KeyCode, modifiers: KeyModifiers) {
        match app.mode {
            AppMode::Normal => Self::handle_normal_mode(app, key, modifiers),
            AppMode::Editing => Self::handle_editing_mode(app, key),
            AppMode::Help => Self::handle_help_mode(app, key),
            AppMode::SaveAs => Self::handle_filename_input_mode(app, key, true),
            AppMode::LoadFile => Self::handle_filename_input_mode(app, key, false),
        }
    }

    fn handle_normal_mode(app: &mut App, key: KeyCode, modifiers: KeyModifiers) {
        if modifiers.contains(KeyModifiers::CONTROL) {
            match key {
                KeyCode::Char('s') => {
                    app.start_save_as();
                    return;
                }
                KeyCode::Char('o') => {
                    app.start_load_file();
                    return;
                }
                _ => {}
            }
        }
        
        app.status_message = None;
        
        match key {
            KeyCode::Up | KeyCode::Char('k') => {
                if app.selected_row > 0 {
                    app.selected_row -= 1;
                    if app.selected_row < app.scroll_row {
                        app.scroll_row = app.selected_row;
                    }
                }
            }
            KeyCode::Down | KeyCode::Char('j') => {
                if app.selected_row < app.spreadsheet.rows - 1 {
                    app.selected_row += 1;
                }
            }
            KeyCode::Left | KeyCode::Char('h') => {
                if app.selected_col > 0 {
                    app.selected_col -= 1;
                    if app.selected_col < app.scroll_col {
                        app.scroll_col = app.selected_col;
                    }
                }
            }
            KeyCode::Right | KeyCode::Char('l') => {
                if app.selected_col < app.spreadsheet.cols - 1 {
                    app.selected_col += 1;
                }
            }
            KeyCode::Enter | KeyCode::F(2) => {
                app.start_editing();
            }
            KeyCode::Char('=') => {
                app.spreadsheet.auto_resize_column(app.selected_col);
            }
            KeyCode::Char('+') => {
                app.spreadsheet.auto_resize_all_columns();
            }
            KeyCode::Char('-') => {
                let current_width = app.spreadsheet.get_column_width(app.selected_col);
                if current_width > 3 {
                    app.spreadsheet.set_column_width(app.selected_col, current_width - 1);
                }
            }
            KeyCode::Char('_') => {
                let current_width = app.spreadsheet.get_column_width(app.selected_col);
                app.spreadsheet.set_column_width(app.selected_col, current_width + 1);
            }
            KeyCode::F(1) | KeyCode::Char('?') => {
                app.mode = AppMode::Help;
                app.help_scroll = 0;
            }
            KeyCode::Char('q') => {
                // Will be handled by main loop
            }
            _ => {}
        }
    }

    fn handle_editing_mode(app: &mut App, key: KeyCode) {
        match key {
            KeyCode::Enter => {
                app.finish_editing();
            }
            KeyCode::Esc => {
                app.cancel_editing();
            }
            KeyCode::Backspace => {
                if app.cursor_position > 0 {
                    app.input.remove(app.cursor_position - 1);
                    app.cursor_position -= 1;
                }
            }
            KeyCode::Delete => {
                if app.cursor_position < app.input.len() {
                    app.input.remove(app.cursor_position);
                }
            }
            KeyCode::Left => {
                if app.cursor_position > 0 {
                    app.cursor_position -= 1;
                }
            }
            KeyCode::Right => {
                if app.cursor_position < app.input.len() {
                    app.cursor_position += 1;
                }
            }
            KeyCode::Home => {
                app.cursor_position = 0;
            }
            KeyCode::End => {
                app.cursor_position = app.input.len();
            }
            KeyCode::Char(c) => {
                app.input.insert(app.cursor_position, c);
                app.cursor_position += 1;
            }
            _ => {}
        }
    }

    fn handle_help_mode(app: &mut App, key: KeyCode) {
        match key {
            KeyCode::Esc | KeyCode::F(1) | KeyCode::Char('?') | KeyCode::Char('q') => {
                app.mode = AppMode::Normal;
            }
            KeyCode::Up | KeyCode::Char('k') => {
                if app.help_scroll > 0 {
                    app.help_scroll -= 1;
                }
            }
            KeyCode::Down | KeyCode::Char('j') => {
                app.help_scroll += 1;
            }
            KeyCode::PageUp => {
                app.help_scroll = app.help_scroll.saturating_sub(5);
            }
            KeyCode::PageDown => {
                app.help_scroll += 5;
            }
            KeyCode::Home => {
                app.help_scroll = 0;
            }
            _ => {}
        }
    }

    fn handle_filename_input_mode(app: &mut App, key: KeyCode, is_save: bool) {
        match key {
            KeyCode::Enter => {
                if is_save {
                    let filename = app.get_save_filename();
                    let result = FileRepository::save_spreadsheet(&app.spreadsheet, &filename);
                    app.set_save_result(result);
                } else {
                    let filename = app.get_load_filename();
                    let result = FileRepository::load_spreadsheet(&filename);
                    app.set_load_result(result);
                }
            }
            KeyCode::Esc => {
                app.cancel_filename_input();
            }
            KeyCode::Backspace => {
                if app.cursor_position > 0 {
                    app.filename_input.remove(app.cursor_position - 1);
                    app.cursor_position -= 1;
                }
            }
            KeyCode::Delete => {
                if app.cursor_position < app.filename_input.len() {
                    app.filename_input.remove(app.cursor_position);
                }
            }
            KeyCode::Left => {
                if app.cursor_position > 0 {
                    app.cursor_position -= 1;
                }
            }
            KeyCode::Right => {
                if app.cursor_position < app.filename_input.len() {
                    app.cursor_position += 1;
                }
            }
            KeyCode::Home => {
                app.cursor_position = 0;
            }
            KeyCode::End => {
                app.cursor_position = app.filename_input.len();
            }
            KeyCode::Char(c) => {
                app.filename_input.insert(app.cursor_position, c);
                app.cursor_position += 1;
            }
            _ => {}
        }
    }
}