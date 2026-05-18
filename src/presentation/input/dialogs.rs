//! Submodule of `input` — see input/mod.rs.

use super::*;
use crate::application::{App, AppMode};
use crossterm::event::{KeyCode, KeyModifiers};

impl InputHandler {
    pub(super) fn handle_confirm_discard_mode(app: &mut App, key: KeyCode) {
        match key {
            KeyCode::Char('y') | KeyCode::Char('Y') => app.confirm_pending_action(),
            KeyCode::Char('s') | KeyCode::Char('S') => {
                let pending = app.pending_action.take();
                app.save_in_place_or_prompt();
                if let Some(action) = pending {
                    if !app.dirty {
                        app.pending_action = Some(action);
                        app.confirm_pending_action();
                    } else {
                        app.pending_action = Some(action);
                    }
                }
            }
            KeyCode::Esc | KeyCode::Char('n') | KeyCode::Char('N') => app.cancel_pending_action(),
            _ => {}
        }
    }

    pub(super) fn handle_help_mode(app: &mut App, key: KeyCode) {
        // When the help-search prompt is active, route most keys to it.
        if app.help_search_active {
            match key {
                KeyCode::Esc => {
                    app.help_search_active = false;
                    app.help_search.clear();
                }
                KeyCode::Enter => {
                    // Commit the query; keep it so `n`/`N` can navigate.
                    app.help_search_active = false;
                }
                KeyCode::Backspace => {
                    app.help_search.pop();
                    if let Some(line) = crate::presentation::ui::find_help_match(
                        &app.help_search,
                        0,
                    ) {
                        app.help_scroll = line;
                    }
                }
                KeyCode::Char(c) => {
                    app.help_search.push(c);
                    if let Some(line) = crate::presentation::ui::find_help_match(
                        &app.help_search,
                        0,
                    ) {
                        app.help_scroll = line;
                    }
                }
                _ => {}
            }
            return;
        }

        match key {
            KeyCode::Esc | KeyCode::F(1) | KeyCode::Char('?') | KeyCode::Char('q') => {
                app.mode = AppMode::Normal;
                app.help_search.clear();
                app.help_search_active = false;
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
            // Digit jumps — anchor keys defined in HELP_SECTIONS (1-9 plus 0
            // for Vim Mode). The help text intro promises this works.
            KeyCode::Char(c) if c.is_ascii_digit() => {
                if let Some(line) = crate::presentation::ui::help_section_offset(c) {
                    app.help_scroll = line;
                }
            }
            // `/` enters help-search; `n` / `N` navigate prior matches.
            KeyCode::Char('/') => {
                app.help_search_active = true;
                app.help_search.clear();
            }
            KeyCode::Char('n') => {
                if !app.help_search.is_empty() {
                    if let Some(line) = crate::presentation::ui::find_help_match(
                        &app.help_search,
                        app.help_scroll + 1,
                    ) {
                        app.help_scroll = line;
                    }
                }
            }
            KeyCode::Char('N') => {
                if !app.help_search.is_empty() {
                    // Search backward from current position, wrapping.
                    let q = app.help_search.to_lowercase();
                    let text = crate::presentation::ui::get_help_text();
                    let lines: Vec<&str> = text.lines().collect();
                    let len = lines.len();
                    if len > 0 {
                        let start = if app.help_scroll == 0 { len - 1 } else { app.help_scroll - 1 };
                        for i in 0..len {
                            let idx = (start + len - i) % len;
                            if lines[idx].to_lowercase().contains(&q) {
                                app.help_scroll = idx;
                                break;
                            }
                        }
                    }
                }
            }
            _ => {}
        }
    }

    pub(super) fn handle_filename_input_mode(app: &mut App, key: KeyCode, mode: &str) {
        match key {
            KeyCode::Enter => {
                match mode {
                    "save" => {
                        let filename = app.get_save_filename();
                        let result = FileRepository::save_workbook(&app.workbook, &filename);
                        app.set_save_result(result);
                    }
                    "load" => {
                        let filename = app.get_load_filename();
                        let result = FileRepository::load_workbook(&filename);
                        app.set_load_workbook_result(result);
                    }
                    "csv_export" => {
                        let filename = app.get_csv_export_filename();
                        let result = CsvExporter::export_to_csv(app.workbook.current_sheet(), &filename);
                        app.set_csv_export_result(result);
                    }
                    "csv_import" => {
                        let filename = app.get_csv_import_filename();
                        let result = CsvExporter::import_from_csv(&filename);
                        app.set_csv_import_result(result);
                    }
                    _ => {}
                }
            }
            KeyCode::Esc => {
                app.cancel_filename_input();
            }
            KeyCode::Backspace => {
                if app.cursor_position > 0 {
                    app.filename_input.remove(char_to_byte_pos(&app.filename_input, app.cursor_position - 1));
                    app.cursor_position -= 1;
                }
            }
            KeyCode::Delete => {
                if app.cursor_position < char_count(&app.filename_input) {
                    app.filename_input.remove(char_to_byte_pos(&app.filename_input, app.cursor_position));
                }
            }
            KeyCode::Left => {
                if app.cursor_position > 0 {
                    app.cursor_position -= 1;
                }
            }
            KeyCode::Right => {
                if app.cursor_position < char_count(&app.filename_input) {
                    app.cursor_position += 1;
                }
            }
            KeyCode::Home => {
                app.cursor_position = 0;
            }
            KeyCode::End => {
                app.cursor_position = char_count(&app.filename_input);
            }
            KeyCode::Char(c) => {
                app.filename_input.insert(char_to_byte_pos(&app.filename_input, app.cursor_position), c);
                app.cursor_position += 1;
            }
            _ => {}
        }
    }

    pub(super) fn handle_search_mode(app: &mut App, key: KeyCode) {
        match key {
            KeyCode::Enter => {
                app.perform_search();
                app.finish_search();
            }
            KeyCode::Esc => {
                app.cancel_search();
            }
            KeyCode::Backspace => {
                if app.cursor_position > 0 {
                    app.search_query.remove(char_to_byte_pos(&app.search_query, app.cursor_position - 1));
                    app.cursor_position -= 1;
                    // Perform live search as user types
                    app.perform_search();
                }
            }
            KeyCode::Delete => {
                if app.cursor_position < char_count(&app.search_query) {
                    app.search_query.remove(char_to_byte_pos(&app.search_query, app.cursor_position));
                    // Perform live search as user types
                    app.perform_search();
                }
            }
            KeyCode::Left => {
                if app.cursor_position > 0 {
                    app.cursor_position -= 1;
                }
            }
            KeyCode::Right => {
                if app.cursor_position < char_count(&app.search_query) {
                    app.cursor_position += 1;
                }
            }
            KeyCode::Home => {
                app.cursor_position = 0;
            }
            KeyCode::End => {
                app.cursor_position = char_count(&app.search_query);
            }
            KeyCode::Down => {
                // Navigate to next search result while searching
                app.next_search_result();
            }
            KeyCode::Up => {
                // Navigate to previous search result while searching
                app.previous_search_result();
            }
            KeyCode::Char(c) => {
                app.search_query.insert(char_to_byte_pos(&app.search_query, app.cursor_position), c);
                app.cursor_position += 1;
                // Perform live search as user types
                app.perform_search();
            }
            _ => {}
        }
    }

    pub(super) fn handle_goto_cell_mode(app: &mut App, key: KeyCode) {
        match key {
            KeyCode::Enter => {
                app.finish_goto_cell();
            }
            KeyCode::Esc => {
                app.cancel_goto_cell();
            }
            KeyCode::Backspace => {
                if app.cursor_position > 0 {
                    app.goto_cell_input.remove(char_to_byte_pos(&app.goto_cell_input, app.cursor_position - 1));
                    app.cursor_position -= 1;
                }
            }
            KeyCode::Char(c) => {
                app.goto_cell_input.insert(char_to_byte_pos(&app.goto_cell_input, app.cursor_position), c);
                app.cursor_position += 1;
            }
            _ => {}
        }
    }

    pub(super) fn handle_find_replace_mode(app: &mut App, key: KeyCode, modifiers: KeyModifiers) {
        // Ctrl+A: replace all
        if modifiers.contains(KeyModifiers::CONTROL) {
            if let KeyCode::Char('a') = key {
                app.replace_all();
                return;
            }
        }
        match key {
            KeyCode::Enter => {
                if app.find_replace_on_replace {
                    app.replace_current();
                } else {
                    // Perform search, then switch to replace field
                    app.find_replace_search();
                    app.find_replace_on_replace = true;
                    app.cursor_position = app.find_replace_replace.chars().count();
                }
            }
            KeyCode::Esc => {
                app.finish_find_replace();
            }
            KeyCode::Tab => {
                // Toggle between search and replace fields
                app.find_replace_on_replace = !app.find_replace_on_replace;
                if app.find_replace_on_replace {
                    app.cursor_position = app.find_replace_replace.chars().count();
                } else {
                    app.cursor_position = app.find_replace_search.chars().count();
                }
            }
            KeyCode::Backspace => {
                if app.find_replace_on_replace {
                    if app.cursor_position > 0 {
                        app.find_replace_replace.remove(char_to_byte_pos(&app.find_replace_replace, app.cursor_position - 1));
                        app.cursor_position -= 1;
                    }
                } else {
                    if app.cursor_position > 0 {
                        app.find_replace_search.remove(char_to_byte_pos(&app.find_replace_search, app.cursor_position - 1));
                        app.cursor_position -= 1;
                        app.find_replace_search();
                    }
                }
            }
            KeyCode::Down => {
                // Next result
                if !app.find_replace_results.is_empty() {
                    app.find_replace_index = (app.find_replace_index + 1) % app.find_replace_results.len();
                    let (row, col) = app.find_replace_results[app.find_replace_index];
                    app.selected_row = row;
                    app.selected_col = col;
                    app.ensure_cursor_visible();
                }
            }
            KeyCode::Up => {
                // Previous result
                if !app.find_replace_results.is_empty() {
                    if app.find_replace_index == 0 {
                        app.find_replace_index = app.find_replace_results.len() - 1;
                    } else {
                        app.find_replace_index -= 1;
                    }
                    let (row, col) = app.find_replace_results[app.find_replace_index];
                    app.selected_row = row;
                    app.selected_col = col;
                    app.ensure_cursor_visible();
                }
            }
            KeyCode::Char(c) => {
                if app.find_replace_on_replace {
                    app.find_replace_replace.insert(char_to_byte_pos(&app.find_replace_replace, app.cursor_position), c);
                    app.cursor_position += 1;
                } else {
                    app.find_replace_search.insert(char_to_byte_pos(&app.find_replace_search, app.cursor_position), c);
                    app.cursor_position += 1;
                    app.find_replace_search();
                }
            }
            _ => {}
        }
    }

    pub(super) fn handle_command_palette_mode(app: &mut App, key: KeyCode) {
        match key {
            KeyCode::Enter => {
                app.execute_command();
            }
            KeyCode::Esc => {
                app.mode = AppMode::Normal;
                app.command_input.clear();
                app.cursor_position = 0;
            }
            KeyCode::Backspace => {
                if app.cursor_position > 0 {
                    app.command_input.remove(char_to_byte_pos(&app.command_input, app.cursor_position - 1));
                    app.cursor_position -= 1;
                }
            }
            KeyCode::Left => {
                if app.cursor_position > 0 {
                    app.cursor_position -= 1;
                }
            }
            KeyCode::Right => {
                if app.cursor_position < char_count(&app.command_input) {
                    app.cursor_position += 1;
                }
            }
            KeyCode::Char(c) => {
                app.command_input.insert(char_to_byte_pos(&app.command_input, app.cursor_position), c);
                app.cursor_position += 1;
            }
            _ => {}
        }
    }

}
