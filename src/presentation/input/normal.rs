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
        if let KeyCode::Char(c) = key {
            if c.is_ascii_digit() {
                let d = c.to_digit(10).unwrap() as usize;
                // '0' at the start is the row-start motion, not a count.
                if !(d == 0 && app.vim_count.is_none()) {
                    app.vim_count = Some(app.vim_count.unwrap_or(0) * 10 + d);
                    return;
                }
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
