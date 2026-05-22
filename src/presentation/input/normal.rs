//! Submodule of `input` — see input/mod.rs.

use super::*;
use crate::application::{App, AppMode, VimOperator, VisualKind};
use crossterm::event::{KeyCode, KeyModifiers};

impl InputHandler {
    pub(super) fn handle_normal_mode(app: &mut App, key: KeyCode, modifiers: KeyModifiers) {
        // ----- Ctrl-shortcut layer (Excel/Sheets familiar) -----
        // Ctrl-bindings are standalone actions; they never compose with a
        // pending vim count/operator. Clear any pending vim state BEFORE
        // firing the action so the count doesn't leak across the shortcut.
        if modifiers.contains(KeyModifiers::CONTROL) {
            // Only clear if there IS pending state — otherwise we'd no-op anyway.
            if app.vim_count.is_some()
                || app.vim_pending_op.is_some()
                || app.vim_awaiting_g
            {
                app.vim_reset_pending();
            }
            match key {
                KeyCode::Char('s') => { app.start_save_as(); return; }
                KeyCode::Char('o') => { app.start_load_file(); return; }
                KeyCode::Char('e') => { app.start_csv_export(); return; }
                // Ctrl+I and Ctrl+H removed — terminals send them as Tab and
                // Backspace, causing dual-bindings. Use Ctrl+L for CSV import
                // and `/`+palette for find/replace instead. We swallow them
                // here as no-ops so that on modern terminals which DO
                // distinguish Ctrl+I from Tab, the user doesn't accidentally
                // enter Insert mode via the `i` fall-through.
                KeyCode::Char('i') | KeyCode::Char('h') => return,
                KeyCode::Char('l') => { app.start_csv_import(); return; }
                KeyCode::Char('z') => { app.undo(); return; }
                KeyCode::Char('y') => { app.redo(); return; }
                KeyCode::Char('r') => { app.redo(); return; } // vim-style Ctrl-R
                KeyCode::Char('d') => { app.autofill_selection(); return; }
                KeyCode::Char('g') => { app.start_goto_cell(); return; }
                KeyCode::Char('c') => { app.copy_selection(); return; }
                KeyCode::Char('x') => { app.cut_selection(); return; }
                KeyCode::Char('v') => {
                    // Ctrl+V: in normal mode, enter visual block. The "paste"
                    // function is now mapped to `p` (vim) and remains
                    // available via the command palette as needed.
                    app.vim_enter_visual(VisualKind::Block);
                    return;
                }
                KeyCode::Char('b') => { app.toggle_bold(); return; }
                KeyCode::Char('u') => { app.toggle_underline(); return; }
                KeyCode::Home => { app.jump_to_home(); return; }
                KeyCode::End => { app.jump_to_end(); return; }
                KeyCode::PageDown => { app.switch_next_sheet(); return; }
                KeyCode::PageUp => { app.switch_prev_sheet(); return; }
                _ => {}
            }
        }

        let is_shift = modifiers.contains(KeyModifiers::SHIFT);
        let take_count = app.vim_count.unwrap_or(1).max(1);
        let pending_op = app.vim_pending_op;
        let awaiting_g = app.vim_awaiting_g;

        // Status-message lifecycle:
        // Don't auto-clear on plain navigation keys — the user has a moment
        // to read "Yanked 3 cells", "Saved", "Circular reference rejected",
        // etc. New actions overwrite the message naturally; this match only
        // clears on Esc and explicit mode-entry keys so stale notes don't
        // linger across a context switch.
        let is_context_switch = matches!(
            key,
            KeyCode::Esc | KeyCode::Char(':') | KeyCode::Char('/') | KeyCode::Char('?')
        );
        if is_context_switch
            && app.vim_count.is_none()
            && app.vim_pending_op.is_none()
            && !app.vim_awaiting_g
        {
            app.status_message = None;
        }

        // ----- Pending 'g' (we saw one 'g', looking for second 'g' = gg) -----
        if awaiting_g {
            app.vim_awaiting_g = false;
            if let KeyCode::Char('g') = key {
                if let Some(op) = pending_op {
                    let start = app.selected_row;
                    app.vim_motion_top();
                    let end = app.selected_row;
                    // Apply over the row range [min..=max] across all columns.
                    let last_col = app.workbook.current_sheet().cols.saturating_sub(1);
                    app.vim_apply_operator(op, start, 0, end, last_col);
                    return;
                }
                app.vim_motion_top();
                app.vim_reset_pending();
                return;
            }
            // Any other key cancels the pending 'g' and falls through normally.
        }

        // ----- Count-prefix accumulation -----
        // Digits (other than a leading 0) build up a count.
        if let KeyCode::Char(c) = key
            && c.is_ascii_digit() {
                let d = c.to_digit(10).unwrap() as usize;
                // '0' at the start is the row-start motion, not a count.
                if !(d == 0 && app.vim_count.is_none()) {
                    app.vim_count = Some(app.vim_count.unwrap_or(0) * 10 + d);
                    return;
                }
            }

        // ----- Pending operator (d/y/c waiting for a motion) -----
        if let Some(op) = pending_op {
            match key {
                // Line operations: dd, yy, cc
                KeyCode::Char('d') if matches!(op, VimOperator::Delete) => {
                    app.vim_apply_line_op(op, take_count); return;
                }
                KeyCode::Char('y') if matches!(op, VimOperator::Yank) => {
                    app.vim_apply_line_op(op, take_count); return;
                }
                KeyCode::Char('c') if matches!(op, VimOperator::Change) => {
                    app.vim_apply_line_op(op, take_count); return;
                }
                // Motion + operator: compute target, then apply.
                KeyCode::Char('j') | KeyCode::Down => {
                    let r0 = app.selected_row;
                    let r1 = (r0 + take_count)
                        .min(app.workbook.current_sheet().rows.saturating_sub(1));
                    let c = app.selected_col;
                    app.vim_apply_operator(op, r0, c, r1, c);
                    return;
                }
                KeyCode::Char('k') | KeyCode::Up => {
                    let r0 = app.selected_row;
                    let r1 = r0.saturating_sub(take_count);
                    let c = app.selected_col;
                    app.vim_apply_operator(op, r0, c, r1, c);
                    return;
                }
                KeyCode::Char('h') | KeyCode::Left => {
                    let c0 = app.selected_col;
                    let c1 = c0.saturating_sub(take_count);
                    let r = app.selected_row;
                    app.vim_apply_operator(op, r, c0, r, c1);
                    return;
                }
                KeyCode::Char('l') | KeyCode::Right => {
                    let c0 = app.selected_col;
                    let c1 = (c0 + take_count)
                        .min(app.workbook.current_sheet().cols.saturating_sub(1));
                    let r = app.selected_row;
                    app.vim_apply_operator(op, r, c0, r, c1);
                    return;
                }
                KeyCode::Char('G') => {
                    let r0 = app.selected_row;
                    let target_row = app.vim_count;
                    if let Some(n) = target_row {
                        app.vim_motion_goto_row(n);
                    } else {
                        app.vim_motion_bottom();
                    }
                    let r1 = app.selected_row;
                    let last_col = app.workbook.current_sheet().cols.saturating_sub(1);
                    app.vim_apply_operator(op, r0, 0, r1, last_col);
                    return;
                }
                KeyCode::Char('g') => {
                    // First g of dgg — wait for second
                    app.vim_awaiting_g = true;
                    return;
                }
                KeyCode::Char('0') | KeyCode::Home => {
                    let c0 = app.selected_col;
                    let r = app.selected_row;
                    app.vim_apply_operator(op, r, 0, r, c0);
                    return;
                }
                KeyCode::Char('$') | KeyCode::End => {
                    let r = app.selected_row;
                    let c0 = app.selected_col;
                    app.vim_motion_row_end();
                    let c1 = app.selected_col;
                    app.vim_apply_operator(op, r, c0, r, c1);
                    return;
                }
                KeyCode::Esc => {
                    app.vim_reset_pending();
                    return;
                }
                _ => {
                    // Unknown motion under pending op — cancel
                    app.vim_reset_pending();
                    return;
                }
            }
        }

        // ----- Normal-mode commands (no pending operator) -----
        match key {
            // Movement
            KeyCode::Up | KeyCode::Char('k') => {
                let n = take_count;
                if !is_shift { app.clear_selection(); }
                for _ in 0..n {
                    if app.selected_row == 0 { break; }
                    if is_shift && !app.selecting { app.start_selection(); }
                    app.selected_row -= 1;
                    while app.selected_row > 0 && app.hidden_rows.contains(&app.selected_row) {
                        app.selected_row -= 1;
                    }
                }
                app.ensure_cursor_visible();
                if is_shift { app.update_selection(app.selected_row, app.selected_col); }
                app.vim_count = None;
            }
            KeyCode::Down | KeyCode::Char('j') => {
                let n = take_count;
                if !is_shift { app.clear_selection(); }
                let max_row = app.workbook.current_sheet().rows.saturating_sub(1);
                for _ in 0..n {
                    if app.selected_row >= max_row { break; }
                    if is_shift && !app.selecting { app.start_selection(); }
                    app.selected_row += 1;
                    while app.selected_row < max_row && app.hidden_rows.contains(&app.selected_row) {
                        app.selected_row += 1;
                    }
                }
                app.ensure_cursor_visible();
                if is_shift { app.update_selection(app.selected_row, app.selected_col); }
                app.vim_count = None;
            }
            KeyCode::Left | KeyCode::Char('h') => {
                let n = take_count;
                if !is_shift { app.clear_selection(); }
                for _ in 0..n {
                    if app.selected_col == 0 { break; }
                    if is_shift && !app.selecting { app.start_selection(); }
                    app.selected_col -= 1;
                }
                app.ensure_cursor_visible();
                if is_shift { app.update_selection(app.selected_row, app.selected_col); }
                app.vim_count = None;
            }
            KeyCode::Right | KeyCode::Char('l') => {
                let n = take_count;
                if !is_shift { app.clear_selection(); }
                let max_col = app.workbook.current_sheet().cols.saturating_sub(1);
                for _ in 0..n {
                    if app.selected_col >= max_col { break; }
                    if is_shift && !app.selecting { app.start_selection(); }
                    app.selected_col += 1;
                }
                app.ensure_cursor_visible();
                if is_shift { app.update_selection(app.selected_row, app.selected_col); }
                app.vim_count = None;
            }
            // Insert-mode entries. Vim semantics for a cell collapse to two
            // distinct points: cursor at start (`i`/`I`) or end (`a`/`A`).
            KeyCode::Char('i') | KeyCode::Char('I') => { app.vim_enter_insert(); app.vim_count = None; }
            KeyCode::Char('a') | KeyCode::Char('A') => { app.vim_enter_insert_at_end(); app.vim_count = None; }
            KeyCode::Char('o') => { app.vim_open_row_below(); app.vim_count = None; }
            KeyCode::Char('O') => { app.vim_open_row_above(); app.vim_count = None; }
            KeyCode::Char('s') => { app.vim_substitute_cell(); app.vim_count = None; }
            KeyCode::Char('S') => { app.vim_substitute_row(); app.vim_count = None; }
            KeyCode::Enter | KeyCode::F(2) => { app.start_editing(); app.vim_count = None; }
            // Visual-mode entries
            KeyCode::Char('v') => { app.vim_enter_visual(VisualKind::Cell); app.vim_count = None; }
            KeyCode::Char('V') => { app.vim_enter_visual(VisualKind::Row); app.vim_count = None; }
            // Operators (set pending state)
            KeyCode::Char('d') => { app.vim_pending_op = Some(VimOperator::Delete); }
            KeyCode::Char('y') => { app.vim_pending_op = Some(VimOperator::Yank); }
            KeyCode::Char('c') => { app.vim_pending_op = Some(VimOperator::Change); }
            // x = delete current cell
            KeyCode::Char('x') | KeyCode::Backspace | KeyCode::Delete => {
                app.clear_cell_with_undo(app.selected_row, app.selected_col);
                app.vim_count = None;
            }
            // Paste
            KeyCode::Char('p') => { app.vim_paste_below(); app.vim_count = None; }
            KeyCode::Char('P') => { app.vim_paste_above(); app.vim_count = None; }
            // Undo
            KeyCode::Char('u') => { app.undo(); app.vim_count = None; }
            // Motions (no pending op)
            KeyCode::Char('g') => { app.vim_awaiting_g = true; }
            KeyCode::Char('G') => {
                if let Some(n) = app.vim_count {
                    app.vim_motion_goto_row(n);
                } else {
                    app.vim_motion_bottom();
                }
                app.vim_count = None;
            }
            KeyCode::Char('0') | KeyCode::Home if app.vim_count.is_none() => {
                app.vim_motion_row_start();
            }
            KeyCode::Char('$') | KeyCode::End => { app.vim_motion_row_end(); app.vim_count = None; }
            KeyCode::Char('^') => { app.vim_motion_row_first_data(); app.vim_count = None; }
            // Column-width keys (punctuation, not letters — doesn't collide with vim).
            KeyCode::Char('+') => {
                app.workbook.current_sheet_mut().auto_resize_all_columns();
                app.vim_count = None;
            }
            KeyCode::Char('-') => {
                let cur = app.workbook.current_sheet().get_column_width(app.selected_col);
                if cur > 3 {
                    app.workbook.current_sheet_mut().set_column_width(app.selected_col, cur - 1);
                }
                app.vim_count = None;
            }
            KeyCode::Char('_') => {
                let cur = app.workbook.current_sheet().get_column_width(app.selected_col);
                app.workbook.current_sheet_mut().set_column_width(app.selected_col, cur + 1);
                app.vim_count = None;
            }
            // Help, search, command-line
            KeyCode::F(1) | KeyCode::Char('?') => {
                app.mode = AppMode::Help;
                app.help_scroll = 0;
                app.vim_count = None;
            }
            KeyCode::Char('/') => { app.start_search(); app.vim_count = None; }
            KeyCode::Char('n') if !app.search_results.is_empty() => {
                app.next_search_result(); app.vim_count = None;
            }
            KeyCode::Char('N') if !app.search_results.is_empty() => {
                app.previous_search_result(); app.vim_count = None;
            }
            KeyCode::Char(':') => { app.start_command_palette(); app.vim_count = None; }
            // Quit
            KeyCode::Char('q') => { app.request_quit(); app.vim_count = None; }
            // Tab / BackTab keep cursor-style move-right/left
            KeyCode::Tab => {
                app.clear_selection();
                let max_col = app.workbook.current_sheet().cols.saturating_sub(1);
                if app.selected_col < max_col { app.selected_col += 1; app.ensure_cursor_visible(); }
                app.vim_count = None;
            }
            KeyCode::BackTab => {
                app.clear_selection();
                if app.selected_col > 0 { app.selected_col -= 1; app.ensure_cursor_visible(); }
                app.vim_count = None;
            }
            KeyCode::PageDown => {
                app.clear_selection();
                let jump = app.viewport_rows.max(1) * take_count;
                let max_row = app.workbook.current_sheet().rows.saturating_sub(1);
                app.selected_row = (app.selected_row + jump).min(max_row);
                app.ensure_cursor_visible();
                app.vim_count = None;
            }
            KeyCode::PageUp => {
                app.clear_selection();
                let jump = app.viewport_rows.max(1) * take_count;
                app.selected_row = app.selected_row.saturating_sub(jump);
                app.ensure_cursor_visible();
                app.vim_count = None;
            }
            KeyCode::Esc => {
                app.dismiss_transients();
                app.vim_reset_pending();
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

    fn fill_col(app: &mut App, col: usize, n: usize) {
        for r in 0..n {
            app.workbook.current_sheet_mut().set_cell(r, col, crate::domain::CellData {
                value: format!("r{}c{}", r, col), formula: None, format: None, comment: None,
                spill_anchor: None,
            });
        }
    }
    fn fill_row(app: &mut App, row: usize, n: usize) {
        for c in 0..n {
            app.workbook.current_sheet_mut().set_cell(row, c, crate::domain::CellData {
                value: format!("r{}c{}", row, c), formula: None, format: None, comment: None,
                spill_anchor: None,
            });
        }
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
    fn test_ctrl_c_copies() {
        let mut app = App::default();
        app.workbook.current_sheet_mut().set_cell(0, 0, crate::domain::CellData {
            value: "test".to_string(), formula: None, format: None, comment: None,
        spill_anchor: None,
        });
        InputHandler::handle_key_event(&mut app, KeyCode::Char('c'), KeyModifiers::CONTROL);
        assert!(app.clipboard.is_some());
    }

    #[test]
    fn test_ctrl_x_cuts() {
        let mut app = App::default();
        app.workbook.current_sheet_mut().set_cell(0, 0, crate::domain::CellData {
            value: "test".to_string(), formula: None, format: None, comment: None,
        spill_anchor: None,
        });
        InputHandler::handle_key_event(&mut app, KeyCode::Char('x'), KeyModifiers::CONTROL);
        assert!(app.clipboard.is_some());
        assert!(app.workbook.current_sheet().get_cell(0, 0).value.is_empty()); // Cut clears
    }

    #[test]
    fn test_p_pastes_after_yank() {
        // Vim flow: yank cell with yy, move, paste with p.
        let mut app = App::default();
        app.workbook.current_sheet_mut().set_cell(0, 0, crate::domain::CellData {
            value: "test".to_string(), formula: None, format: None, comment: None,
            spill_anchor: None,
        });
        // Ctrl+C still copies (Excel/Sheets familiar).
        InputHandler::handle_key_event(&mut app, KeyCode::Char('c'), KeyModifiers::CONTROL);
        app.selected_row = 1;
        app.selected_col = 1;
        // p pastes
        InputHandler::handle_key_event(&mut app, KeyCode::Char('p'), KeyModifiers::NONE);
        // p pastes "below/after" — for a single-cell clipboard that means at col+1.
        assert_eq!(app.workbook.current_sheet().get_cell(1, 2).value, "test");
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
    fn test_s_substitutes_cell() {
        // Vim `s` clears the current cell and enters Insert mode.
        let mut app = App::default();
        app.workbook.current_sheet_mut().set_cell(0, 0, crate::domain::CellData {
            value: "old".to_string(), formula: None, format: None, comment: None,
            spill_anchor: None,
        });
        InputHandler::handle_key_event(&mut app, KeyCode::Char('s'), KeyModifiers::NONE);
        assert!(matches!(app.mode, AppMode::Editing));
        assert_eq!(app.input, "");
        assert!(app.workbook.current_sheet().get_cell(0, 0).value.is_empty());
    }

    #[test]
    fn test_shift_s_substitutes_row() {
        // Vim `S` clears the current row and enters Insert mode at col 0.
        let mut app = App::default();
        app.workbook.current_sheet_mut().set_cell(0, 0, crate::domain::CellData {
            value: "a".to_string(), formula: None, format: None, comment: None,
            spill_anchor: None,
        });
        app.workbook.current_sheet_mut().set_cell(0, 1, crate::domain::CellData {
            value: "b".to_string(), formula: None, format: None, comment: None,
            spill_anchor: None,
        });
        app.selected_col = 1;
        InputHandler::handle_key_event(&mut app, KeyCode::Char('S'), KeyModifiers::NONE);
        assert!(matches!(app.mode, AppMode::Editing));
        assert_eq!(app.selected_col, 0);
        assert!(app.workbook.current_sheet().get_cell(0, 0).value.is_empty());
        assert!(app.workbook.current_sheet().get_cell(0, 1).value.is_empty());
    }

    #[test]
    fn test_sort_via_command_palette() {
        // Sort moved off single-letter `s` to the palette.
        let mut app = App::default();
        app.workbook.current_sheet_mut().set_cell(0, 0, crate::domain::CellData {
            value: "20".to_string(), formula: None, format: None, comment: None,
            spill_anchor: None,
        });
        app.workbook.current_sheet_mut().set_cell(1, 0, crate::domain::CellData {
            value: "10".to_string(), formula: None, format: None, comment: None,
            spill_anchor: None,
        });
        app.start_command_palette();
        for c in "sort asc".chars() {
            InputHandler::handle_key_event(&mut app, KeyCode::Char(c), KeyModifiers::NONE);
        }
        InputHandler::handle_key_event(&mut app, KeyCode::Enter, KeyModifiers::NONE);
        assert_eq!(app.workbook.current_sheet().get_cell(0, 0).value, "10");
        assert_eq!(app.workbook.current_sheet().get_cell(1, 0).value, "20");
    }

    #[test]
    fn test_ctrl_b_toggles_bold() {
        let mut app = App::default();
        app.workbook.current_sheet_mut().set_cell(0, 0, crate::domain::CellData {
            value: "test".to_string(), formula: None, format: None, comment: None,
        spill_anchor: None,
        });
        InputHandler::handle_key_event(&mut app, KeyCode::Char('b'), KeyModifiers::CONTROL);
        assert!(app.workbook.current_sheet().get_cell(0, 0).format.as_ref().unwrap().style.bold);
    }

    #[test]
    fn test_ctrl_u_toggles_underline() {
        let mut app = App::default();
        app.workbook.current_sheet_mut().set_cell(0, 0, crate::domain::CellData {
            value: "test".to_string(), formula: None, format: None, comment: None,
        spill_anchor: None,
        });
        InputHandler::handle_key_event(&mut app, KeyCode::Char('u'), KeyModifiers::CONTROL);
        assert!(app.workbook.current_sheet().get_cell(0, 0).format.as_ref().unwrap().style.underline);
    }

    #[test]
    fn test_q_quits_when_clean() {
        let mut app = App::default();
        assert!(!app.should_quit);
        typestr(&mut app, "q");
        assert!(app.should_quit);
    }

    #[test]
    fn test_q_prompts_when_dirty() {
        let mut app = App::default();
        app.dirty = true;
        typestr(&mut app, "q");
        // Should NOT immediately quit; instead enter ConfirmDiscard mode.
        assert!(!app.should_quit);
        assert!(matches!(app.mode, AppMode::ConfirmDiscard));
    }

    #[test]
    fn test_colon_q_quits() {
        let mut app = App::default();
        typestr(&mut app, ":");
        assert!(matches!(app.mode, AppMode::CommandPalette));
        typestr(&mut app, "q");
        key(&mut app, KeyCode::Enter);
        assert!(app.should_quit);
    }

    #[test]
    fn test_colon_q_bang_force_quits() {
        let mut app = App::default();
        app.dirty = true;
        typestr(&mut app, ":");
        typestr(&mut app, "q!");
        key(&mut app, KeyCode::Enter);
        assert!(app.should_quit);
    }

    #[test]
    fn test_o_opens_row_below() {
        let mut app = App::default();
        assert_eq!(app.selected_row, 0);
        typestr(&mut app, "o");
        assert_eq!(app.selected_row, 1);
        assert!(matches!(app.mode, AppMode::Editing));
    }

    #[test]
    fn test_h_j_k_l_navigation() {
        let mut app = App::default();
        typestr(&mut app, "jjl");
        assert_eq!(app.selected_row, 2);
        assert_eq!(app.selected_col, 1);
        typestr(&mut app, "h");
        assert_eq!(app.selected_col, 0);
        typestr(&mut app, "k");
        assert_eq!(app.selected_row, 1);
    }

    #[test]
    fn test_count_prefix_motion() {
        let mut app = App::default();
        typestr(&mut app, "5j");
        assert_eq!(app.selected_row, 5);
        typestr(&mut app, "3l");
        assert_eq!(app.selected_col, 3);
    }

    #[test]
    fn test_dd_deletes_row() {
        let mut app = App::default();
        app.workbook.current_sheet_mut().set_cell(0, 0, crate::domain::CellData {
            value: "a".to_string(), formula: None, format: None, comment: None,
            spill_anchor: None,
        });
        app.workbook.current_sheet_mut().set_cell(0, 1, crate::domain::CellData {
            value: "b".to_string(), formula: None, format: None, comment: None,
            spill_anchor: None,
        });
        typestr(&mut app, "dd");
        assert!(app.workbook.current_sheet().get_cell(0, 0).value.is_empty());
        assert!(app.workbook.current_sheet().get_cell(0, 1).value.is_empty());
    }

    #[test]
    fn test_yy_then_p_pastes_row_below() {
        let mut app = App::default();
        app.workbook.current_sheet_mut().set_cell(0, 0, crate::domain::CellData {
            value: "a".to_string(), formula: None, format: None, comment: None,
            spill_anchor: None,
        });
        app.workbook.current_sheet_mut().set_cell(0, 1, crate::domain::CellData {
            value: "b".to_string(), formula: None, format: None, comment: None,
            spill_anchor: None,
        });
        typestr(&mut app, "yy");
        // Move down then paste — p should drop the yanked row at next row.
        typestr(&mut app, "p");
        assert_eq!(app.workbook.current_sheet().get_cell(1, 0).value, "a");
        assert_eq!(app.workbook.current_sheet().get_cell(1, 1).value, "b");
    }

    #[test]
    fn test_x_deletes_current_cell() {
        let mut app = App::default();
        app.workbook.current_sheet_mut().set_cell(0, 0, crate::domain::CellData {
            value: "v".to_string(), formula: None, format: None, comment: None,
            spill_anchor: None,
        });
        typestr(&mut app, "x");
        assert!(app.workbook.current_sheet().get_cell(0, 0).value.is_empty());
    }

    #[test]
    fn test_u_undoes() {
        let mut app = App::default();
        app.workbook.current_sheet_mut().set_cell(0, 0, crate::domain::CellData {
            value: "v".to_string(), formula: None, format: None, comment: None,
            spill_anchor: None,
        });
        typestr(&mut app, "x"); // delete
        assert!(app.workbook.current_sheet().get_cell(0, 0).value.is_empty());
        typestr(&mut app, "u"); // undo
        assert_eq!(app.workbook.current_sheet().get_cell(0, 0).value, "v");
    }

    #[test]
    fn test_gg_jumps_to_top() {
        let mut app = App::default();
        app.selected_row = 5;
        typestr(&mut app, "gg");
        assert_eq!(app.selected_row, 0);
    }

    #[test]
    fn test_G_with_count_jumps_to_row() {
        let mut app = App::default();
        typestr(&mut app, "3");
        InputHandler::handle_key_event(&mut app, KeyCode::Char('G'), KeyModifiers::NONE);
        assert_eq!(app.selected_row, 2); // 1-based 3 = 0-based 2
    }

    #[test]
    fn test_dollar_motion_to_last_data_in_row() {
        let mut app = App::default();
        app.workbook.current_sheet_mut().set_cell(0, 4, crate::domain::CellData {
            value: "x".to_string(), formula: None, format: None, comment: None,
            spill_anchor: None,
        });
        InputHandler::handle_key_event(&mut app, KeyCode::Char('$'), KeyModifiers::NONE);
        assert_eq!(app.selected_col, 4);
    }

    #[test]
    fn test_zero_motion_to_first_col() {
        let mut app = App::default();
        app.selected_col = 4;
        InputHandler::handle_key_event(&mut app, KeyCode::Char('0'), KeyModifiers::NONE);
        assert_eq!(app.selected_col, 0);
    }

    #[test]
    fn test_ctrl_r_redoes() {
        let mut app = App::default();
        app.workbook.current_sheet_mut().set_cell(0, 0, crate::domain::CellData {
            value: "v".to_string(), formula: None, format: None, comment: None,
            spill_anchor: None,
        });
        typestr(&mut app, "x");
        typestr(&mut app, "u");
        assert_eq!(app.workbook.current_sheet().get_cell(0, 0).value, "v");
        ctrl(&mut app, 'r');
        assert!(app.workbook.current_sheet().get_cell(0, 0).value.is_empty());
    }

    #[test]
    fn test_count_with_dd() {
        // 3dd should delete 3 rows.
        let mut app = App::default();
        for r in 0..5 {
            app.workbook.current_sheet_mut().set_cell(r, 0, crate::domain::CellData {
                value: format!("r{}", r), formula: None, format: None, comment: None,
                spill_anchor: None,
            });
        }
        typestr(&mut app, "3dd");
        assert!(app.workbook.current_sheet().get_cell(0, 0).value.is_empty());
        assert!(app.workbook.current_sheet().get_cell(1, 0).value.is_empty());
        assert!(app.workbook.current_sheet().get_cell(2, 0).value.is_empty());
        assert_eq!(app.workbook.current_sheet().get_cell(3, 0).value, "r3");
    }

    #[test]
    fn test_esc_clears_pending_op() {
        let mut app = App::default();
        typestr(&mut app, "5d");
        assert!(app.vim_pending_op.is_some());
        key(&mut app, KeyCode::Esc);
        assert!(app.vim_pending_op.is_none());
        assert!(app.vim_count.is_none());
    }

    #[test]
    fn agent_pending_count_state_after_5() {
        let mut app = App::default();
        typestr(&mut app, "5");
        assert_eq!(app.vim_count, Some(5));
        assert!(app.vim_pending_op.is_none());
        assert!(!app.vim_awaiting_g);
    }

    #[test]
    fn agent_pending_5d_5dg_5dgg_fires_action() {
        let mut app = App::default();
        app.workbook.current_sheet_mut().rows = 20;
        app.selected_row = 5;
        typestr(&mut app, "5");
        assert_eq!(app.vim_count, Some(5));
        typestr(&mut app, "d");
        assert_eq!(app.vim_pending_op, Some(VimOperator::Delete));
        assert_eq!(app.vim_count, Some(5));
        typestr(&mut app, "g");
        assert!(app.vim_awaiting_g, "5dg should set awaiting_g");
        typestr(&mut app, "g");
        assert!(app.vim_count.is_none());
        assert!(app.vim_pending_op.is_none());
        assert!(!app.vim_awaiting_g);
    }

    #[test]
    fn agent_yank_status_message_keeps_normal_mode() {
        let mut app = App::default();
        put(&mut app, 0, 0, "x");
        typestr(&mut app, "yy");
        assert!(matches!(app.mode, AppMode::Normal));
        let msg = app.status_message.clone().unwrap_or_default();
        assert!(msg.contains("Yanked"), "expected 'Yanked' status, got {:?}", msg);
    }

    #[test]
    fn agent_pending_count_cleared_by_ctrl_b() {
        // Ctrl-shortcuts are standalone actions; they clear pending vim
        // state so the count doesn't leak into subsequent commands.
        let mut app = App::default();
        put(&mut app, 0, 0, "x");
        typestr(&mut app, "5");
        assert_eq!(app.vim_count, Some(5));
        ctrl(&mut app, 'b');
        assert!(app.vim_count.is_none(),
            "Ctrl-shortcuts must clear pending vim_count.");
    }

    #[test]
    fn agent_pending_d_then_ctrl_s_clears_pending_op() {
        let mut app = App::default();
        typestr(&mut app, "5d");
        assert_eq!(app.vim_pending_op, Some(VimOperator::Delete));
        ctrl(&mut app, 's');
        assert!(matches!(app.mode, AppMode::SaveAs));
        assert!(app.vim_pending_op.is_none(),
            "Ctrl+S must clear pending Delete (Ctrl shortcuts don't compose).");
        assert!(app.vim_count.is_none(),
            "Ctrl+S must clear vim_count too.");
    }

    #[test]
    fn agent_status_message_survives_motion_until_context_switch() {
        // Status messages persist across plain motion keys (so the user can
        // actually read them) and only clear on a context switch like Esc /
        // : / / / ? — or when a new action overwrites them.
        let mut app = App::default();
        put(&mut app, 0, 0, "a");
        typestr(&mut app, "dd");
        let msg = app.status_message.clone();
        assert!(msg.as_deref().unwrap_or("").contains("Deleted"));
        typestr(&mut app, "j");
        assert_eq!(app.status_message, msg, "motion should not wipe status");
        // Esc clears it.
        key(&mut app, KeyCode::Esc);
        assert!(app.status_message.is_none());
    }

    #[test]
    fn agent_status_message_preserved_under_pending_count() {
        let mut app = App::default();
        put(&mut app, 0, 0, "a");
        typestr(&mut app, "yy");
        let yanked = app.status_message.clone();
        assert!(yanked.is_some());
        typestr(&mut app, "5");
        assert_eq!(app.status_message, yanked,
            "Status message should be preserved across pending-count state");
    }

    #[test]
    fn agent_help_mode_q_exits_to_normal() {
        let mut app = App::default();
        typestr(&mut app, "?");
        assert!(matches!(app.mode, AppMode::Help));
        typestr(&mut app, "q");
        assert!(matches!(app.mode, AppMode::Normal));
    }

    #[test]
    fn agent_help_mode_esc_exits_to_normal() {
        let mut app = App::default();
        typestr(&mut app, "?");
        assert!(matches!(app.mode, AppMode::Help));
        key(&mut app, KeyCode::Esc);
        assert!(matches!(app.mode, AppMode::Normal));
    }

    #[test]
    fn agent_confirm_discard_status_message_set_by_request_quit() {
        let mut app = App::default();
        app.dirty = true;
        app.request_quit();
        assert!(matches!(app.mode, AppMode::ConfirmDiscard));
        let msg = app.status_message.clone().unwrap_or_default();
        assert!(msg.to_lowercase().contains("unsaved"),
            "ConfirmDiscard status should explain — got {:?}", msg);
    }

    #[test]
    fn agent_pending_g_alone_then_unrelated_key_falls_through() {
        let mut app = App::default();
        app.workbook.current_sheet_mut().rows = 5;
        typestr(&mut app, "g");
        assert!(app.vim_awaiting_g);
        let r0 = app.selected_row;
        typestr(&mut app, "j");
        assert!(!app.vim_awaiting_g);
        assert_eq!(app.selected_row, r0 + 1,
            "After g+unrelated, the next key should execute normally (g dropped)");
    }

    #[test]
    fn agent_pending_d_then_question_drops_keystroke() {
        let mut app = App::default();
        typestr(&mut app, "d");
        typestr(&mut app, "?");
        assert!(app.vim_pending_op.is_none());
        assert!(matches!(app.mode, AppMode::Normal),
            "BUG: `d?` cancels d but drops `?` — user expects Help, gets nothing.");
    }

    #[test]
    fn agent_grammar_5dd_clears_five_rows() {
        let mut app = App::default();
        for r in 0..8 { fill_row(&mut app, r, 3); }
        typestr(&mut app, "5dd");
        for r in 0..5 {
            for c in 0..3 {
                assert!(app.workbook.current_sheet().get_cell(r, c).value.is_empty(),
                    "row {} col {} should be cleared", r, c);
            }
        }
        // Row 5 must still have data.
        assert_eq!(app.workbook.current_sheet().get_cell(5, 0).value, "r5c0");
    }

    #[test]
    fn agent_grammar_dj_deletes_two_cells_in_column() {
        let mut app = App::default();
        fill_col(&mut app, 0, 4);
        // cursor at (0,0); dj covers (0,0)..=(1,0)
        typestr(&mut app, "dj");
        assert!(app.workbook.current_sheet().get_cell(0, 0).value.is_empty());
        assert!(app.workbook.current_sheet().get_cell(1, 0).value.is_empty());
        assert_eq!(app.workbook.current_sheet().get_cell(2, 0).value, "r2c0");
    }

    #[test]
    fn agent_grammar_dk_deletes_two_cells_upward() {
        let mut app = App::default();
        fill_col(&mut app, 0, 4);
        app.selected_row = 2;
        typestr(&mut app, "dk");
        assert!(app.workbook.current_sheet().get_cell(1, 0).value.is_empty());
        assert!(app.workbook.current_sheet().get_cell(2, 0).value.is_empty());
        assert_eq!(app.workbook.current_sheet().get_cell(0, 0).value, "r0c0");
        assert_eq!(app.workbook.current_sheet().get_cell(3, 0).value, "r3c0");
    }

    #[test]
    fn agent_grammar_dl_deletes_horizontal_range() {
        let mut app = App::default();
        fill_row(&mut app, 0, 4);
        typestr(&mut app, "dl");
        assert!(app.workbook.current_sheet().get_cell(0, 0).value.is_empty());
        assert!(app.workbook.current_sheet().get_cell(0, 1).value.is_empty());
        assert_eq!(app.workbook.current_sheet().get_cell(0, 2).value, "r0c2");
    }

    #[test]
    fn agent_grammar_dgg_deletes_rows_up_to_top_all_columns() {
        let mut app = App::default();
        for r in 0..5 { fill_row(&mut app, r, 3); }
        app.selected_row = 2;
        app.selected_col = 1; // pick a non-zero col to verify cross-column behavior
        typestr(&mut app, "dgg");
        for r in 0..=2 {
            for c in 0..3 {
                assert!(app.workbook.current_sheet().get_cell(r, c).value.is_empty(),
                    "row {} col {} should be cleared after dgg", r, c);
            }
        }
        assert_eq!(app.workbook.current_sheet().get_cell(3, 0).value, "r3c0");
    }

    #[test]
    fn agent_grammar_dG_deletes_rows_to_bottom_data() {
        let mut app = App::default();
        for r in 0..5 { fill_row(&mut app, r, 3); }
        app.selected_row = 2;
        typestr(&mut app, "dG");
        for r in 2..5 {
            for c in 0..3 {
                assert!(app.workbook.current_sheet().get_cell(r, c).value.is_empty(),
                    "row {} col {} should be cleared after dG", r, c);
            }
        }
        assert_eq!(app.workbook.current_sheet().get_cell(0, 0).value, "r0c0");
        assert_eq!(app.workbook.current_sheet().get_cell(1, 0).value, "r1c0");
    }

    #[test]
    fn agent_grammar_5dj_count_applies_to_motion() {
        let mut app = App::default();
        fill_col(&mut app, 0, 10);
        typestr(&mut app, "5dj");
        // Should clear rows 0..=5 (6 cells)
        for r in 0..=5 {
            assert!(app.workbook.current_sheet().get_cell(r, 0).value.is_empty(),
                "row {} should be cleared after 5dj", r);
        }
        assert_eq!(app.workbook.current_sheet().get_cell(6, 0).value, "r6c0");
    }

    #[test]
    fn agent_grammar_d_dollar_deletes_to_row_end() {
        let mut app = App::default();
        fill_row(&mut app, 0, 5);
        app.selected_col = 1;
        typestr(&mut app, "d$");
        assert_eq!(app.workbook.current_sheet().get_cell(0, 0).value, "r0c0");
        for c in 1..5 {
            assert!(app.workbook.current_sheet().get_cell(0, c).value.is_empty(),
                "col {} should be cleared after d$", c);
        }
    }

    #[test]
    fn agent_grammar_d0_deletes_from_row_start_to_cursor() {
        let mut app = App::default();
        fill_row(&mut app, 0, 5);
        app.selected_col = 3;
        typestr(&mut app, "d0");
        for c in 0..=3 {
            assert!(app.workbook.current_sheet().get_cell(0, c).value.is_empty(),
                "col {} should be cleared after d0", c);
        }
        assert_eq!(app.workbook.current_sheet().get_cell(0, 4).value, "r0c4");
    }

    #[test]
    fn agent_grammar_yj_clipboard_two_rows_one_col() {
        let mut app = App::default();
        fill_col(&mut app, 0, 4);
        typestr(&mut app, "yj");
        let cb = app.clipboard.as_ref().expect("clipboard");
        // 2 cells: (0,0,r0c0) and (1,0,r1c0)
        assert_eq!(cb.cells.len(), 2);
        let v0 = cb.cells.iter().find(|(r, c, _)| *r == 0 && *c == 0).map(|(_, _, cd)| cd.value.clone());
        let v1 = cb.cells.iter().find(|(r, c, _)| *r == 1 && *c == 0).map(|(_, _, cd)| cd.value.clone());
        assert_eq!(v0.as_deref(), Some("r0c0"));
        assert_eq!(v1.as_deref(), Some("r1c0"));
        // Source row anchored at top-left of yank
        assert_eq!(cb.source_row, 0);
        assert_eq!(cb.source_col, 0);
        // No data changed
        assert_eq!(app.workbook.current_sheet().get_cell(0, 0).value, "r0c0");
    }

    #[test]
    fn esc_after_d_then_y_does_not_inherit_delete() {
        // d-Esc-y must not delete: the Esc cancels the pending delete and
        // the y starts fresh as a yank operator. Regression test for the
        // claim that Esc leaves vim_pending_op set.
        let mut app = App::default();
        app.workbook.current_sheet_mut().rows = 5;
        fill_col(&mut app, 0, 3);
        let original = app.workbook.current_sheet().get_cell(0, 0).value.clone();
        typestr(&mut app, "d");
        assert!(app.vim_pending_op.is_some(), "d should leave a pending delete");
        key(&mut app, KeyCode::Esc);
        assert!(app.vim_pending_op.is_none(), "Esc must clear the pending operator");
        typestr(&mut app, "yy");
        // The cell must still have its original value — yy yanks, doesn't delete.
        assert_eq!(
            app.workbook.current_sheet().get_cell(0, 0).value,
            original,
            "yy after d-Esc must yank, not inherit the cancelled delete"
        );
    }

    #[test]
    fn agent_grammar_esc_after_5d_then_j_moves_one() {
        let mut app = App::default();
        app.workbook.current_sheet_mut().rows = 50;
        typestr(&mut app, "5d");
        key(&mut app, KeyCode::Esc);
        assert!(app.vim_pending_op.is_none());
        assert!(app.vim_count.is_none());
        typestr(&mut app, "j");
        assert_eq!(app.selected_row, 1,
            "after Esc canceling 5d, j must move exactly 1 row, got {}", app.selected_row);
    }

    #[test]
    fn agent_grammar_dk_at_top_no_panic_clears_only_current_cell() {
        let mut app = App::default();
        fill_col(&mut app, 0, 3);
        // cursor at row 0
        typestr(&mut app, "dk");
        // saturating sub: motion target = 0, so range is (0,0)..(0,0) — clears current cell only.
        assert!(app.workbook.current_sheet().get_cell(0, 0).value.is_empty());
        assert_eq!(app.workbook.current_sheet().get_cell(1, 0).value, "r1c0");
        assert_eq!(app.workbook.current_sheet().get_cell(2, 0).value, "r2c0");
    }

    #[test]
    fn agent_grammar_dh_at_left_edge_no_panic() {
        let mut app = App::default();
        fill_row(&mut app, 0, 3);
        typestr(&mut app, "dh");
        // Range (0,0)..(0,0) — only current cell cleared.
        assert!(app.workbook.current_sheet().get_cell(0, 0).value.is_empty());
        assert_eq!(app.workbook.current_sheet().get_cell(0, 1).value, "r0c1");
    }

    #[test]
    fn agent_grammar_3yy_then_p_pastes_three_rows_below() {
        let mut app = App::default();
        for r in 0..5 { fill_row(&mut app, r, 3); }
        // Capture state
        typestr(&mut app, "3yy");
        // After yy, clipboard exists, cursor still at (0,0)
        let cb = app.clipboard.as_ref().expect("clipboard set");
        // 3 rows × 3 cols = 9 entries (all non-empty)
        assert_eq!(cb.cells.len(), 9, "expected 9 cells in clipboard for 3yy");
        // Move cursor down 4 to an area we can verify paste lands correctly.
        app.selected_row = 0;
        typestr(&mut app, "p");
        // After p with a multi-row clipboard, was_row is false (rows aren't all 0),
        // so it shifts col right and pastes there. This may NOT be 'paste 3 rows
        // below' — observe and assert. We measure what actually happened.
        // Expected vim-ish behavior: paste should land at row 1 column 0.
        // Probe the actual landing point:
        let landed_at_row_1_col_0 = app.workbook.current_sheet().get_cell(1, 0).value == "r0c0";
        let landed_at_row_0_col_1 = app.workbook.current_sheet().get_cell(0, 1).value == "r0c0";
        assert!(landed_at_row_1_col_0,
            "BUG: 3yy then p should paste 3-row clipboard BELOW (row 1, col 0). \
             Instead landed_at_row_1_col_0={} landed_at_row_0_col_1={}",
            landed_at_row_1_col_0, landed_at_row_0_col_1);
    }

    #[test]
    fn agent_grammar_count_does_not_leak_past_motion_into_next_command() {
        let mut app = App::default();
        app.workbook.current_sheet_mut().rows = 50;
        fill_col(&mut app, 0, 20);
        typestr(&mut app, "5dj");
        assert!(app.vim_count.is_none(), "vim_count should be None after operator+motion completes");
        // Now cursor still at (0,0). A bare `j` should move down 1.
        let r_before = app.selected_row;
        typestr(&mut app, "j");
        assert_eq!(app.selected_row, r_before + 1);
    }

}
