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
            KeyCode::Up | KeyCode::Char('k')
                if app.help_scroll > 0 => {
                    app.help_scroll -= 1;
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
                if !app.help_search.is_empty()
                    && let Some(line) = crate::presentation::ui::find_help_match(
                        &app.help_search,
                        app.help_scroll + 1,
                    ) {
                        app.help_scroll = line;
                    }
            }
            KeyCode::Char('N')
                if !app.help_search.is_empty() => {
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
            KeyCode::Backspace
                if app.cursor_position > 0 => {
                    app.filename_input.remove(char_to_byte_pos(&app.filename_input, app.cursor_position - 1));
                    app.cursor_position -= 1;
                }
            KeyCode::Delete
                if app.cursor_position < char_count(&app.filename_input) => {
                    app.filename_input.remove(char_to_byte_pos(&app.filename_input, app.cursor_position));
                }
            KeyCode::Left
                if app.cursor_position > 0 => {
                    app.cursor_position -= 1;
                }
            KeyCode::Right
                if app.cursor_position < char_count(&app.filename_input) => {
                    app.cursor_position += 1;
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
            KeyCode::Backspace
                if app.cursor_position > 0 => {
                    app.search_query.remove(char_to_byte_pos(&app.search_query, app.cursor_position - 1));
                    app.cursor_position -= 1;
                    // Perform live search as user types
                    app.perform_search();
                }
            KeyCode::Delete
                if app.cursor_position < char_count(&app.search_query) => {
                    app.search_query.remove(char_to_byte_pos(&app.search_query, app.cursor_position));
                    // Perform live search as user types
                    app.perform_search();
                }
            KeyCode::Left
                if app.cursor_position > 0 => {
                    app.cursor_position -= 1;
                }
            KeyCode::Right
                if app.cursor_position < char_count(&app.search_query) => {
                    app.cursor_position += 1;
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
            KeyCode::Backspace
                if app.cursor_position > 0 => {
                    app.goto_cell_input.remove(char_to_byte_pos(&app.goto_cell_input, app.cursor_position - 1));
                    app.cursor_position -= 1;
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
        if modifiers.contains(KeyModifiers::CONTROL)
            && let KeyCode::Char('a') = key {
                app.replace_all();
                return;
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
            KeyCode::Down
                // Next result
                if !app.find_replace_results.is_empty() => {
                    app.find_replace_index = (app.find_replace_index + 1) % app.find_replace_results.len();
                    let (row, col) = app.find_replace_results[app.find_replace_index];
                    app.selected_row = row;
                    app.selected_col = col;
                    app.ensure_cursor_visible();
                }
            KeyCode::Up
                // Previous result
                if !app.find_replace_results.is_empty() => {
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
            KeyCode::Backspace
                if app.cursor_position > 0 => {
                    app.command_input.remove(char_to_byte_pos(&app.command_input, app.cursor_position - 1));
                    app.cursor_position -= 1;
                }
            KeyCode::Left
                if app.cursor_position > 0 => {
                    app.cursor_position -= 1;
                }
            KeyCode::Right
                if app.cursor_position < char_count(&app.command_input) => {
                    app.cursor_position += 1;
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
    use crate::application::{App, AppMode, VimOperator, VisualKind};
    use crossterm::event::{KeyCode, KeyModifiers};


    fn typestr(app: &mut App, s: &str) {
        for c in s.chars() {
            InputHandler::handle_key_event(app, KeyCode::Char(c), KeyModifiers::NONE);
        }
    }
    fn key(app: &mut App, code: KeyCode) {
        InputHandler::handle_key_event(app, code, KeyModifiers::NONE);
    }
    fn ctrl(app: &mut App, c: char) {
        InputHandler::handle_key_event(app, KeyCode::Char(c), KeyModifiers::CONTROL);
    }
    fn set_text(app: &mut App, row: usize, col: usize, s: &str) {
        app.workbook.current_sheet_mut().set_cell(row, col, crate::domain::CellData {
            value: s.to_string(), formula: None, format: None, comment: None,
            spill_anchor: None,
        });
    }
    fn put(app: &mut App, row: usize, col: usize, s: &str) { set_text(app, row, col, s); }

    fn make_dirty(app: &mut App) {
        app.workbook.current_sheet_mut().set_cell(0, 0, crate::domain::CellData {
            value: "x".to_string(), formula: None, format: None, comment: None,
            spill_anchor: None,
        });
        app.dirty = true;
    }
    fn run_palette(app: &mut App, cmd: &str) {
        app.start_command_palette();
        typestr(app, cmd);
        key(app, KeyCode::Enter);
    }
    fn unique_tmp(name: &str) -> String {
        let mut p = std::env::temp_dir();
        let now = std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH).map(|d| d.as_nanos()).unwrap_or(0);
        p.push(format!("tshts_quit_{}_{}_{}.tshts", std::process::id(), now, name));
        p.to_string_lossy().to_string()
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
    fn test_command_palette_escape_cancels() {
        let mut app = App::default();
        app.start_command_palette();
        InputHandler::handle_key_event(&mut app, KeyCode::Char('x'), KeyModifiers::NONE);
        InputHandler::handle_key_event(&mut app, KeyCode::Esc, KeyModifiers::NONE);

        assert!(matches!(app.mode, AppMode::Normal));
        assert!(app.command_input.is_empty());
    }

    #[test]
    fn agent_quit_q_clean_quits_immediately() {
        let mut app = App::default();
        typestr(&mut app, "q");
        assert!(app.should_quit, "clean `q` must set should_quit");
        assert!(!matches!(app.mode, AppMode::ConfirmDiscard));
    }

    #[test]
    fn agent_quit_q_dirty_enters_confirm_discard() {
        let mut app = App::default();
        make_dirty(&mut app);
        typestr(&mut app, "q");
        assert!(matches!(app.mode, AppMode::ConfirmDiscard));
        assert!(!app.should_quit);
    }

    #[test]
    fn agent_quit_confirm_y_quits() {
        let mut app = App::default();
        make_dirty(&mut app);
        typestr(&mut app, "q");
        key(&mut app, KeyCode::Char('y'));
        assert!(app.should_quit);
    }

    #[test]
    fn agent_quit_confirm_n_cancels() {
        let mut app = App::default();
        make_dirty(&mut app);
        typestr(&mut app, "q");
        key(&mut app, KeyCode::Char('n'));
        assert!(!app.should_quit);
        assert!(matches!(app.mode, AppMode::Normal));
        assert!(app.dirty, "n must not lose dirty");
    }

    #[test]
    fn agent_quit_confirm_s_saves_and_quits() {
        let mut app = App::default();
        let path = unique_tmp("confirm_s");
        app.filename = Some(path.clone());
        make_dirty(&mut app);
        typestr(&mut app, "q");
        assert!(matches!(app.mode, AppMode::ConfirmDiscard));
        key(&mut app, KeyCode::Char('s'));
        assert!(!app.dirty, "after 's' the workbook must be saved (dirty=false)");
        assert!(app.should_quit, "after 's' the deferred quit must fire");
        let _ = std::fs::remove_file(&path);
    }

    #[test]
    fn agent_quit_palette_q_clean_quits() {
        let mut app = App::default();
        run_palette(&mut app, "q");
        assert!(app.should_quit);
    }

    #[test]
    fn agent_quit_palette_qbang_forces_quit() {
        let mut app = App::default();
        make_dirty(&mut app);
        run_palette(&mut app, "q!");
        assert!(app.should_quit);
        assert!(!matches!(app.mode, AppMode::ConfirmDiscard));
    }

    #[test]
    fn agent_quit_w_then_q_quits() {
        let mut app = App::default();
        let path = unique_tmp("w_then_q");
        app.filename = Some(path.clone());
        make_dirty(&mut app);
        run_palette(&mut app, "w");
        assert!(!app.dirty, ":w must clear dirty");
        run_palette(&mut app, "q");
        assert!(app.should_quit);
        let _ = std::fs::remove_file(&path);
    }

    #[test]
    fn agent_quit_w_no_filename_opens_saveas() {
        let mut app = App::default();
        assert!(app.filename.is_none());
        make_dirty(&mut app);
        run_palette(&mut app, "w");
        assert!(matches!(app.mode, AppMode::SaveAs),
            ":w with no filename should open SaveAs dialog; mode={:?}", app.mode);
    }

    #[test]
    fn agent_quit_w_with_filename_saves_and_sets_filename() {
        let mut app = App::default();
        let path = unique_tmp("w_with_file");
        make_dirty(&mut app);
        run_palette(&mut app, &format!("w {}", path));
        assert!(!app.dirty, ":w <file> must clear dirty");
        assert_eq!(app.filename.as_deref(), Some(path.as_str()),
            ":w <file> must update app.filename");
        assert!(std::path::Path::new(&path).exists(), "file must be written");
        let _ = std::fs::remove_file(&path);
    }

    #[test]
    fn agent_quit_wq_no_filename_opens_saveas_no_quit() {
        let mut app = App::default();
        make_dirty(&mut app);
        run_palette(&mut app, "wq");
        // With no filename, save_in_place_or_prompt opens SaveAs; dirty stays true.
        // The code path: save_in_place_or_prompt (no filename → SaveAs), then
        // `if !self.dirty { should_quit = true }`. Since dirty is still true,
        // we should NOT quit.
        assert!(!app.should_quit, ":wq without filename should not silently quit");
        assert!(matches!(app.mode, AppMode::SaveAs));
    }

    #[test]
    fn agent_quit_wq_with_filename_saves_and_quits() {
        let mut app = App::default();
        let path = unique_tmp("wq");
        app.filename = Some(path.clone());
        make_dirty(&mut app);
        run_palette(&mut app, "wq");
        assert!(!app.dirty);
        assert!(app.should_quit);
        let _ = std::fs::remove_file(&path);
    }

    #[test]
    fn agent_quit_wqbang_force_quits_even_when_save_deferred() {
        let mut app = App::default();
        make_dirty(&mut app);
        // No filename — save_in_place_or_prompt opens SaveAs. Then `!` path force quits.
        run_palette(&mut app, "wq!");
        assert!(app.should_quit, ":wq! must force-quit");
    }

    #[test]
    fn agent_quit_xbang_force_quits() {
        let mut app = App::default();
        make_dirty(&mut app);
        run_palette(&mut app, "x!");
        assert!(app.should_quit);
    }

    #[test]
    fn agent_quit_x_with_filename_saves_and_quits() {
        let mut app = App::default();
        let path = unique_tmp("x");
        app.filename = Some(path.clone());
        make_dirty(&mut app);
        run_palette(&mut app, "x");
        assert!(!app.dirty);
        assert!(app.should_quit);
        let _ = std::fs::remove_file(&path);
    }

    #[test]
    fn agent_quit_e_no_arg_opens_loadfile() {
        let mut app = App::default();
        run_palette(&mut app, "e");
        assert!(matches!(app.mode, AppMode::LoadFile),
            ":e with no arg should open LoadFile; mode={:?}", app.mode);
    }

    #[test]
    fn agent_quit_e_with_filename_loads_file() {
        let mut app = App::default();
        // First, write a file via :w <file>
        let path = unique_tmp("e_load");
        run_palette(&mut app, &format!("w {}", path));
        assert!(std::path::Path::new(&path).exists());
        // Now create a fresh app and :e it
        let mut app2 = App::default();
        run_palette(&mut app2, &format!("e {}", path));
        assert_eq!(app2.filename.as_deref(), Some(path.as_str()),
            ":e <file> should load and set filename");
        assert!(!app2.dirty);
        let _ = std::fs::remove_file(&path);
    }

    #[test]
    fn agent_quit_e_with_filename_when_other_open() {
        // Set up file A.
        let mut app = App::default();
        let path_a = unique_tmp("e_other_a");
        let path_b = unique_tmp("e_other_b");
        run_palette(&mut app, &format!("w {}", path_a));
        // Write a different file B by mutating a cell, then :w <B>.
        app.workbook.current_sheet_mut().set_cell(2, 2, crate::domain::CellData {
            value: "B-marker".to_string(), formula: None, format: None, comment: None,
            spill_anchor: None,
        });
        app.dirty = true;
        run_palette(&mut app, &format!("w {}", path_b));
        // Now app.filename is path_b; open path_a via :e
        let mut app2 = App::default();
        app2.filename = Some(path_b.clone());
        // Load it so dirty=false and consistent
        let _ = crate::infrastructure::FileRepository::save_workbook(&app2.workbook, &path_b);
        assert!(!app2.dirty);
        run_palette(&mut app2, &format!("e {}", path_a));
        assert_eq!(app2.filename.as_deref(), Some(path_a.as_str()),
            ":e <file> with another file open should load typed file, not currently-open one");
        let _ = std::fs::remove_file(&path_a);
        let _ = std::fs::remove_file(&path_b);
    }

    #[test]
    fn agent_quit_palette_uppercase_q_works() {
        let mut app = App::default();
        run_palette(&mut app, "Q");
        // execute_command lowercases — :Q should behave like :q.
        assert!(app.should_quit, ":Q must work like :q (lowercased)");
    }

    #[test]
    fn agent_quit_palette_uppercase_wq_works() {
        let mut app = App::default();
        let path = unique_tmp("wq_upper");
        app.filename = Some(path.clone());
        make_dirty(&mut app);
        run_palette(&mut app, "WQ");
        assert!(app.should_quit, ":WQ must work like :wq");
        let _ = std::fs::remove_file(&path);
    }

    #[test]
    fn agent_quit_palette_q_with_trailing_space() {
        let mut app = App::default();
        run_palette(&mut app, "q ");
        assert!(app.should_quit, ":q with trailing space should trim and quit");
    }

    #[test]
    fn agent_quit_palette_q_with_leading_space() {
        let mut app = App::default();
        run_palette(&mut app, " q");
        assert!(app.should_quit, ":q with leading space should trim and quit");
    }

}
