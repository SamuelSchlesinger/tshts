use crate::application::{App, AppMode};
use crate::infrastructure::FileRepository;
use crate::domain::CsvExporter;
use crossterm::event::{KeyCode, KeyModifiers};

fn char_to_byte_pos(s: &str, char_pos: usize) -> usize {
    s.char_indices()
        .nth(char_pos)
        .map(|(byte_idx, _)| byte_idx)
        .unwrap_or(s.len())
}

fn char_count(s: &str) -> usize {
    s.chars().count()
}

pub struct InputHandler;

impl InputHandler {
    pub fn handle_key_event(app: &mut App, key: KeyCode, modifiers: KeyModifiers) {
        match app.mode {
            AppMode::Normal => Self::handle_normal_mode(app, key, modifiers),
            AppMode::Editing => Self::handle_editing_mode(app, key),
            AppMode::Help => Self::handle_help_mode(app, key),
            AppMode::SaveAs => Self::handle_filename_input_mode(app, key, "save"),
            AppMode::LoadFile => Self::handle_filename_input_mode(app, key, "load"),
            AppMode::ExportCsv => Self::handle_filename_input_mode(app, key, "csv_export"),
            AppMode::ImportCsv => Self::handle_filename_input_mode(app, key, "csv_import"),
            AppMode::Search => Self::handle_search_mode(app, key),
            AppMode::GoToCell => Self::handle_goto_cell_mode(app, key),
            AppMode::FindReplace => Self::handle_find_replace_mode(app, key, modifiers),
            AppMode::CommandPalette => Self::handle_command_palette_mode(app, key),
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
                KeyCode::Char('e') => {
                    app.start_csv_export();
                    return;
                }
                KeyCode::Char('i') => {
                    app.start_csv_import();
                    return;
                }
                KeyCode::Char('l') => {
                    app.start_csv_import();
                    return;
                }
                KeyCode::Char('z') => {
                    app.undo();
                    return;
                }
                KeyCode::Char('y') => {
                    app.redo();
                    return;
                }
                KeyCode::Char('d') => {
                    app.autofill_selection();
                    return;
                }
                KeyCode::Char('g') => {
                    app.start_goto_cell();
                    return;
                }
                KeyCode::Char('c') => {
                    app.copy_selection();
                    return;
                }
                KeyCode::Char('x') => {
                    app.cut_selection();
                    return;
                }
                KeyCode::Char('v') => {
                    app.paste();
                    return;
                }
                KeyCode::Char('h') => {
                    app.start_find_replace();
                    return;
                }
                KeyCode::Char('b') => {
                    app.toggle_bold();
                    return;
                }
                KeyCode::Char('u') => {
                    app.toggle_underline();
                    return;
                }
                KeyCode::Home => {
                    app.jump_to_home();
                    return;
                }
                KeyCode::End => {
                    app.jump_to_end();
                    return;
                }
                KeyCode::PageDown => {
                    app.switch_next_sheet();
                    return;
                }
                KeyCode::PageUp => {
                    app.switch_prev_sheet();
                    return;
                }
                _ => {}
            }
        }
        
        // Handle navigation with optional selection
        let is_shift = modifiers.contains(KeyModifiers::SHIFT);
        
        // Clear status message if not doing something that should preserve it
        if !matches!(key, KeyCode::Char('d')) || !modifiers.contains(KeyModifiers::CONTROL) {
            app.status_message = None;
        }
        
        match key {
            KeyCode::Up | KeyCode::Char('k') => {
                if !is_shift {
                    app.clear_selection();
                }
                
                if app.selected_row > 0 {
                    if is_shift && !app.selecting {
                        app.start_selection();
                    }

                    app.selected_row -= 1;
                    // Skip hidden rows
                    while app.selected_row > 0 && app.hidden_rows.contains(&app.selected_row) {
                        app.selected_row -= 1;
                    }
                    app.ensure_cursor_visible();

                    if is_shift {
                        app.update_selection(app.selected_row, app.selected_col);
                    }
                }
            }
            KeyCode::Down | KeyCode::Char('j') => {
                if !is_shift {
                    app.clear_selection();
                }

                if app.selected_row < app.workbook.current_sheet().rows - 1 {
                    if is_shift && !app.selecting {
                        app.start_selection();
                    }

                    app.selected_row += 1;
                    // Skip hidden rows
                    while app.selected_row < app.workbook.current_sheet().rows - 1 && app.hidden_rows.contains(&app.selected_row) {
                        app.selected_row += 1;
                    }
                    app.ensure_cursor_visible();

                    if is_shift {
                        app.update_selection(app.selected_row, app.selected_col);
                    }
                }
            }
            KeyCode::Left | KeyCode::Char('h') => {
                if !is_shift {
                    app.clear_selection();
                }
                
                if app.selected_col > 0 {
                    if is_shift && !app.selecting {
                        app.start_selection();
                    }
                    
                    app.selected_col -= 1;
                    app.ensure_cursor_visible();
                    
                    if is_shift {
                        app.update_selection(app.selected_row, app.selected_col);
                    }
                }
            }
            KeyCode::Right | KeyCode::Char('l') => {
                if !is_shift {
                    app.clear_selection();
                }
                
                if app.selected_col < app.workbook.current_sheet().cols - 1 {
                    if is_shift && !app.selecting {
                        app.start_selection();
                    }
                    
                    app.selected_col += 1;
                    app.ensure_cursor_visible();
                    
                    if is_shift {
                        app.update_selection(app.selected_row, app.selected_col);
                    }
                }
            }
            KeyCode::Enter | KeyCode::F(2) => {
                app.start_editing();
            }
            KeyCode::Char('+') => {
                app.workbook.current_sheet_mut().auto_resize_all_columns();
            }
            KeyCode::Char('-') => {
                let current_width = app.workbook.current_sheet().get_column_width(app.selected_col);
                if current_width > 3 {
                    app.workbook.current_sheet_mut().set_column_width(app.selected_col, current_width - 1);
                }
            }
            KeyCode::Char('_') => {
                let current_width = app.workbook.current_sheet().get_column_width(app.selected_col);
                app.workbook.current_sheet_mut().set_column_width(app.selected_col, current_width + 1);
            }
            KeyCode::F(1) | KeyCode::Char('?') => {
                app.mode = AppMode::Help;
                app.help_scroll = 0;
            }
            KeyCode::Backspace => {
                app.clear_cell_with_undo(app.selected_row, app.selected_col);
            }
            KeyCode::Char('/') => {
                app.start_search();
            }
            KeyCode::Char('n') => {
                // Next search result (only if we have previous search results)
                if !app.search_results.is_empty() {
                    app.next_search_result();
                }
            }
            KeyCode::Char('N') => {
                // Previous search result (only if we have previous search results)
                if !app.search_results.is_empty() {
                    app.previous_search_result();
                }
            }
            KeyCode::Char('s') => {
                app.sort_column_asc();
            }
            KeyCode::Char('S') => {
                app.sort_column_desc();
            }
            KeyCode::Char(':') => {
                app.start_command_palette();
            }
            KeyCode::Char('q') => {
                // Will be handled by main loop
            }
            KeyCode::Tab => {
                app.clear_selection();
                if app.selected_col < app.workbook.current_sheet().cols - 1 {
                    app.selected_col += 1;
                    app.ensure_cursor_visible();
                }
            }
            KeyCode::BackTab => {
                app.clear_selection();
                if app.selected_col > 0 {
                    app.selected_col -= 1;
                    app.ensure_cursor_visible();
                }
            }
            KeyCode::PageDown => {
                app.clear_selection();
                let jump = app.viewport_rows.max(1);
                app.selected_row = (app.selected_row + jump).min(app.workbook.current_sheet().rows - 1);
                app.ensure_cursor_visible();
            }
            KeyCode::PageUp => {
                app.clear_selection();
                let jump = app.viewport_rows.max(1);
                app.selected_row = app.selected_row.saturating_sub(jump);
                app.ensure_cursor_visible();
            }
            KeyCode::Esc => {
                app.clear_selection();
            }
            _ => {}
        }
    }

    fn handle_editing_mode(app: &mut App, key: KeyCode) {
        match key {
            KeyCode::Enter => {
                app.finish_editing();
            }
            KeyCode::Tab => {
                // Finish editing and move right instead of down
                app.finish_editing_move_right();
            }
            KeyCode::Esc => {
                app.cancel_editing();
            }
            KeyCode::Backspace => {
                if app.cursor_position > 0 {
                    app.input.remove(char_to_byte_pos(&app.input, app.cursor_position - 1));
                    app.cursor_position -= 1;
                }
            }
            KeyCode::Delete => {
                if app.cursor_position < char_count(&app.input) {
                    app.input.remove(char_to_byte_pos(&app.input, app.cursor_position));
                }
            }
            KeyCode::Left => {
                if app.cursor_position > 0 {
                    app.cursor_position -= 1;
                }
            }
            KeyCode::Right => {
                if app.cursor_position < char_count(&app.input) {
                    app.cursor_position += 1;
                }
            }
            KeyCode::Home => {
                app.cursor_position = 0;
            }
            KeyCode::End => {
                app.cursor_position = char_count(&app.input);
            }
            KeyCode::Char(c) => {
                app.input.insert(char_to_byte_pos(&app.input, app.cursor_position), c);
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

    fn handle_filename_input_mode(app: &mut App, key: KeyCode, mode: &str) {
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

    fn handle_search_mode(app: &mut App, key: KeyCode) {
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
    fn handle_goto_cell_mode(app: &mut App, key: KeyCode) {
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

    fn handle_find_replace_mode(app: &mut App, key: KeyCode, modifiers: KeyModifiers) {
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

    fn handle_command_palette_mode(app: &mut App, key: KeyCode) {
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

#[cfg(test)]
mod tests {
    use super::*;
    use crate::application::{App, AppMode};

    #[test]
    fn test_csv_import_key_binding() {
        let mut app = App::default();
        
        // Initially in normal mode
        assert!(matches!(app.mode, AppMode::Normal));
        
        // Simulate Ctrl+I key press
        InputHandler::handle_key_event(&mut app, KeyCode::Char('i'), KeyModifiers::CONTROL);
        
        // Should switch to ImportCsv mode
        assert!(matches!(app.mode, AppMode::ImportCsv));
        assert_eq!(app.filename_input, "data.csv");
    }

    #[test]
    fn test_csv_import_alternative_key_binding() {
        let mut app = App::default();
        
        // Initially in normal mode
        assert!(matches!(app.mode, AppMode::Normal));
        
        // Simulate Ctrl+L key press (alternative binding)
        InputHandler::handle_key_event(&mut app, KeyCode::Char('l'), KeyModifiers::CONTROL);
        
        // Should switch to ImportCsv mode
        assert!(matches!(app.mode, AppMode::ImportCsv));
        assert_eq!(app.filename_input, "data.csv");
    }

    #[test]
    fn test_csv_export_key_binding() {
        let mut app = App::default();
        
        // Initially in normal mode
        assert!(matches!(app.mode, AppMode::Normal));
        
        // Simulate Ctrl+E key press
        InputHandler::handle_key_event(&mut app, KeyCode::Char('e'), KeyModifiers::CONTROL);
        
        // Should switch to ExportCsv mode
        assert!(matches!(app.mode, AppMode::ExportCsv));
        assert_eq!(app.filename_input, "spreadsheet.csv");
    }

    #[test]
    fn test_import_csv_filename_input() {
        let mut app = App::default();
        app.start_csv_import();
        
        // Test typing a character
        InputHandler::handle_key_event(&mut app, KeyCode::Char('m'), KeyModifiers::NONE);
        assert_eq!(app.filename_input, "data.csvm");
        
        // Test backspace
        InputHandler::handle_key_event(&mut app, KeyCode::Backspace, KeyModifiers::NONE);
        assert_eq!(app.filename_input, "data.csv");
        
        // Test escape to cancel
        InputHandler::handle_key_event(&mut app, KeyCode::Esc, KeyModifiers::NONE);
        assert!(matches!(app.mode, AppMode::Normal));
        assert!(app.filename_input.is_empty());
    }

    #[test]
    fn test_ctrl_c_copies() {
        let mut app = App::default();
        app.workbook.current_sheet_mut().set_cell(0, 0, crate::domain::CellData {
            value: "test".to_string(), formula: None, format: None, comment: None,
        });
        InputHandler::handle_key_event(&mut app, KeyCode::Char('c'), KeyModifiers::CONTROL);
        assert!(app.clipboard.is_some());
    }

    #[test]
    fn test_ctrl_x_cuts() {
        let mut app = App::default();
        app.workbook.current_sheet_mut().set_cell(0, 0, crate::domain::CellData {
            value: "test".to_string(), formula: None, format: None, comment: None,
        });
        InputHandler::handle_key_event(&mut app, KeyCode::Char('x'), KeyModifiers::CONTROL);
        assert!(app.clipboard.is_some());
        assert!(app.workbook.current_sheet().get_cell(0, 0).value.is_empty()); // Cut clears
    }

    #[test]
    fn test_ctrl_v_pastes() {
        let mut app = App::default();
        app.workbook.current_sheet_mut().set_cell(0, 0, crate::domain::CellData {
            value: "test".to_string(), formula: None, format: None, comment: None,
        });
        // Copy first
        InputHandler::handle_key_event(&mut app, KeyCode::Char('c'), KeyModifiers::CONTROL);
        // Move to B2
        app.selected_row = 1;
        app.selected_col = 1;
        // Paste
        InputHandler::handle_key_event(&mut app, KeyCode::Char('v'), KeyModifiers::CONTROL);
        assert_eq!(app.workbook.current_sheet().get_cell(1, 1).value, "test");
    }

    #[test]
    fn test_ctrl_h_starts_find_replace() {
        let mut app = App::default();
        InputHandler::handle_key_event(&mut app, KeyCode::Char('h'), KeyModifiers::CONTROL);
        assert!(matches!(app.mode, AppMode::FindReplace));
    }

    #[test]
    fn test_ctrl_g_starts_goto() {
        let mut app = App::default();
        InputHandler::handle_key_event(&mut app, KeyCode::Char('g'), KeyModifiers::CONTROL);
        assert!(matches!(app.mode, AppMode::GoToCell));
    }

    #[test]
    fn test_colon_starts_command_palette() {
        let mut app = App::default();
        InputHandler::handle_key_event(&mut app, KeyCode::Char(':'), KeyModifiers::NONE);
        assert!(matches!(app.mode, AppMode::CommandPalette));
    }

    #[test]
    fn test_tab_moves_right_in_normal_mode() {
        let mut app = App::default();
        assert_eq!(app.selected_col, 0);
        InputHandler::handle_key_event(&mut app, KeyCode::Tab, KeyModifiers::NONE);
        assert_eq!(app.selected_col, 1);
    }

    #[test]
    fn test_backtab_moves_left_in_normal_mode() {
        let mut app = App::default();
        app.selected_col = 3;
        InputHandler::handle_key_event(&mut app, KeyCode::BackTab, KeyModifiers::NONE);
        assert_eq!(app.selected_col, 2);
    }

    #[test]
    fn test_tab_in_editing_finishes_and_moves_right() {
        let mut app = App::default();
        app.start_editing();
        app.input = "42".to_string();
        InputHandler::handle_key_event(&mut app, KeyCode::Tab, KeyModifiers::NONE);
        assert!(matches!(app.mode, AppMode::Normal));
        assert_eq!(app.selected_col, 1); // Moved right
        assert_eq!(app.workbook.current_sheet().get_cell(0, 0).value, "42");
    }

    #[test]
    fn test_page_down() {
        let mut app = App::default();
        app.viewport_rows = 10;
        assert_eq!(app.selected_row, 0);
        InputHandler::handle_key_event(&mut app, KeyCode::PageDown, KeyModifiers::NONE);
        assert_eq!(app.selected_row, 10);
    }

    #[test]
    fn test_page_up() {
        let mut app = App::default();
        app.viewport_rows = 10;
        app.selected_row = 15;
        InputHandler::handle_key_event(&mut app, KeyCode::PageUp, KeyModifiers::NONE);
        assert_eq!(app.selected_row, 5);
    }

    #[test]
    fn test_find_replace_mode_tab_toggles_fields() {
        let mut app = App::default();
        app.start_find_replace();
        assert!(!app.find_replace_on_replace);

        // Tab should switch to replace field
        InputHandler::handle_key_event(&mut app, KeyCode::Tab, KeyModifiers::NONE);
        assert!(app.find_replace_on_replace);

        // Tab again should switch back to search field
        InputHandler::handle_key_event(&mut app, KeyCode::Tab, KeyModifiers::NONE);
        assert!(!app.find_replace_on_replace);
    }

    #[test]
    fn test_find_replace_mode_escape_exits() {
        let mut app = App::default();
        app.start_find_replace();
        assert!(matches!(app.mode, AppMode::FindReplace));

        InputHandler::handle_key_event(&mut app, KeyCode::Esc, KeyModifiers::NONE);
        assert!(matches!(app.mode, AppMode::Normal));
    }

    #[test]
    fn test_command_palette_mode_typing_and_execute() {
        let mut app = App::default();
        app.start_command_palette();

        // Type "ir"
        InputHandler::handle_key_event(&mut app, KeyCode::Char('i'), KeyModifiers::NONE);
        InputHandler::handle_key_event(&mut app, KeyCode::Char('r'), KeyModifiers::NONE);
        assert_eq!(app.command_input, "ir");

        // Execute
        let orig_rows = app.workbook.current_sheet().rows;
        InputHandler::handle_key_event(&mut app, KeyCode::Enter, KeyModifiers::NONE);
        assert!(matches!(app.mode, AppMode::Normal));
        assert_eq!(app.workbook.current_sheet().rows, orig_rows + 1);
    }

    #[test]
    fn test_command_palette_escape_cancels() {
        let mut app = App::default();
        app.start_command_palette();
        InputHandler::handle_key_event(&mut app, KeyCode::Char('x'), KeyModifiers::NONE);
        InputHandler::handle_key_event(&mut app, KeyCode::Esc, KeyModifiers::NONE);

        assert!(matches!(app.mode, AppMode::Normal));
        assert!(app.command_input.is_empty());
    }

    #[test]
    fn test_sort_key_s_ascending() {
        let mut app = App::default();
        app.workbook.current_sheet_mut().set_cell(0, 0, crate::domain::CellData {
            value: "20".to_string(), formula: None, format: None, comment: None,
        });
        app.workbook.current_sheet_mut().set_cell(1, 0, crate::domain::CellData {
            value: "10".to_string(), formula: None, format: None, comment: None,
        });

        InputHandler::handle_key_event(&mut app, KeyCode::Char('s'), KeyModifiers::NONE);

        assert_eq!(app.workbook.current_sheet().get_cell(0, 0).value, "10");
        assert_eq!(app.workbook.current_sheet().get_cell(1, 0).value, "20");
    }

    #[test]
    fn test_sort_key_shift_s_descending() {
        let mut app = App::default();
        app.workbook.current_sheet_mut().set_cell(0, 0, crate::domain::CellData {
            value: "10".to_string(), formula: None, format: None, comment: None,
        });
        app.workbook.current_sheet_mut().set_cell(1, 0, crate::domain::CellData {
            value: "20".to_string(), formula: None, format: None, comment: None,
        });

        InputHandler::handle_key_event(&mut app, KeyCode::Char('S'), KeyModifiers::NONE);

        assert_eq!(app.workbook.current_sheet().get_cell(0, 0).value, "20");
        assert_eq!(app.workbook.current_sheet().get_cell(1, 0).value, "10");
    }

    #[test]
    fn test_ctrl_b_toggles_bold() {
        let mut app = App::default();
        app.workbook.current_sheet_mut().set_cell(0, 0, crate::domain::CellData {
            value: "test".to_string(), formula: None, format: None, comment: None,
        });
        InputHandler::handle_key_event(&mut app, KeyCode::Char('b'), KeyModifiers::CONTROL);
        assert!(app.workbook.current_sheet().get_cell(0, 0).format.as_ref().unwrap().style.bold);
    }

    #[test]
    fn test_ctrl_u_toggles_underline() {
        let mut app = App::default();
        app.workbook.current_sheet_mut().set_cell(0, 0, crate::domain::CellData {
            value: "test".to_string(), formula: None, format: None, comment: None,
        });
        InputHandler::handle_key_event(&mut app, KeyCode::Char('u'), KeyModifiers::CONTROL);
        assert!(app.workbook.current_sheet().get_cell(0, 0).format.as_ref().unwrap().style.underline);
    }
}