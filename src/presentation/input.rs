use crate::application::{App, AppMode, VimOperator, VisualKind};
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
            AppMode::Visual { .. } => Self::handle_visual_mode(app, key, modifiers),
            AppMode::Help => Self::handle_help_mode(app, key),
            AppMode::SaveAs => Self::handle_filename_input_mode(app, key, "save"),
            AppMode::LoadFile => Self::handle_filename_input_mode(app, key, "load"),
            AppMode::ExportCsv => Self::handle_filename_input_mode(app, key, "csv_export"),
            AppMode::ImportCsv => Self::handle_filename_input_mode(app, key, "csv_import"),
            AppMode::Search => Self::handle_search_mode(app, key),
            AppMode::GoToCell => Self::handle_goto_cell_mode(app, key),
            AppMode::FindReplace => Self::handle_find_replace_mode(app, key, modifiers),
            AppMode::CommandPalette => Self::handle_command_palette_mode(app, key),
            AppMode::ConfirmDiscard => Self::handle_confirm_discard_mode(app, key),
        }
    }

    fn handle_confirm_discard_mode(app: &mut App, key: KeyCode) {
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

    /// Normal mode dispatch — vim-style. Single letters trigger commands;
    /// typing into a cell requires entering Insert mode first (i/a/o/O/s/S
    /// or Enter/F2). Ctrl-shortcuts retain their familiar
    /// Excel/Sheets meanings (Ctrl+S save, Ctrl+C copy, etc.).
    fn handle_normal_mode(app: &mut App, key: KeyCode, modifiers: KeyModifiers) {
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

    /// Visual mode dispatch. Motions extend the selection; y/d/c/x operate on
    /// the selection then return to Normal mode. Esc exits without action;
    /// v/V/Ctrl-V either toggle the granularity or (if pressed in the same
    /// variant they entered) exit back to Normal.
    fn handle_visual_mode(app: &mut App, key: KeyCode, modifiers: KeyModifiers) {
        let kind = if let AppMode::Visual { kind } = app.mode { kind } else { return; };

        // Ctrl shortcuts in visual mode. Ctrl-V toggles to Visual Block (or
        // back to Cell if already in Block), matching the symmetry that V
        // gives between Cell and Row. Other Ctrl-shortcuts that compose
        // naturally with a selection (Ctrl-D autofill, Ctrl-B/U formatting,
        // Ctrl-S save, Ctrl-Z undo) are handled here too — without this
        // guard they would fall through to the bare-key match below and
        // accidentally trigger the same letters' vim semantics (e.g.
        // Ctrl-D in visual would otherwise do a `d` delete).
        if modifiers.contains(KeyModifiers::CONTROL) {
            match key {
                KeyCode::Char('c') => { app.copy_selection(); app.vim_exit_visual(); return; }
                KeyCode::Char('x') => { app.cut_selection(); app.vim_exit_visual(); return; }
                KeyCode::Char('d') => { app.autofill_selection(); app.vim_exit_visual(); return; }
                KeyCode::Char('b') => { app.toggle_bold(); return; }
                KeyCode::Char('u') => { app.toggle_underline(); return; }
                KeyCode::Char('s') => { app.start_save_as(); app.vim_exit_visual(); return; }
                KeyCode::Char('z') => { app.undo(); app.vim_exit_visual(); return; }
                KeyCode::Char('r') => { app.redo(); app.vim_exit_visual(); return; }
                KeyCode::Char('v') => {
                    if matches!(kind, VisualKind::Block) {
                        app.vim_exit_visual();
                    } else {
                        let cursor = (app.selected_row, app.selected_col);
                        let anchor = app.selection_start.unwrap_or(cursor);
                        app.vim_exit_visual();
                        app.selected_row = anchor.0;
                        app.selected_col = anchor.1;
                        app.vim_enter_visual(VisualKind::Block);
                        app.selected_row = cursor.0;
                        app.selected_col = cursor.1;
                        app.update_selection(cursor.0, cursor.1);
                    }
                    return;
                }
                _ => {
                    // Other Ctrl-shortcuts: swallow so they don't fall through
                    // to bare-key dispatch and trigger unintended vim ops.
                    return;
                }
            }
        }
        let take_count = app.vim_count.unwrap_or(1).max(1);
        let extend = |app: &mut App| {
            // In Row visual, snap the selection back to span the full row.
            if matches!(kind, VisualKind::Row) {
                let last_col = app.workbook.current_sheet().cols.saturating_sub(1);
                let anchor_row = app.selection_start.map(|(r, _)| r).unwrap_or(app.selected_row);
                app.selection_start = Some((anchor_row, 0));
                app.selection_end = Some((app.selected_row, last_col));
            } else {
                app.update_selection(app.selected_row, app.selected_col);
            }
        };

        // Handle pending 'g' (for `gg`) FIRST so the `Char('g')` arm below
        // doesn't keep re-setting the flag on a second press.
        if app.vim_awaiting_g {
            app.vim_awaiting_g = false;
            if let KeyCode::Char('g') = key {
                app.vim_motion_top();
                extend(app);
                app.vim_count = None;
                return;
            }
            // Any other key cancels the pending g and continues to the
            // normal dispatch below.
        }

        match key {
            KeyCode::Esc => {
                app.vim_exit_visual();
                app.dismiss_transients();
                return;
            }
            KeyCode::Char('v') => {
                // Toggle: in Cell visual `v` exits; from Row/Block it swaps
                // back to Cell (vim convention).
                if matches!(kind, VisualKind::Cell) {
                    app.vim_exit_visual();
                } else {
                    let cursor = (app.selected_row, app.selected_col);
                    let anchor = app.selection_start.unwrap_or(cursor);
                    app.vim_exit_visual();
                    app.selected_row = anchor.0;
                    app.selected_col = anchor.1;
                    app.vim_enter_visual(VisualKind::Cell);
                    app.selected_row = cursor.0;
                    app.selected_col = cursor.1;
                    app.update_selection(cursor.0, cursor.1);
                }
                return;
            }
            KeyCode::Char('V') => {
                if matches!(kind, VisualKind::Row) {
                    app.vim_exit_visual();
                } else {
                    let anchor_row = app.selection_start.map(|(r, _)| r)
                        .unwrap_or(app.selected_row);
                    let cursor_row = app.selected_row;
                    app.vim_exit_visual();
                    app.selected_row = anchor_row;
                    app.vim_enter_visual(VisualKind::Row);
                    app.selected_row = cursor_row;
                    let last_col = app.workbook.current_sheet().cols.saturating_sub(1);
                    app.selection_end = Some((cursor_row, last_col));
                }
                return;
            }
            KeyCode::Up | KeyCode::Char('k') => {
                for _ in 0..take_count {
                    if app.selected_row == 0 { break; }
                    app.selected_row -= 1;
                }
                app.ensure_cursor_visible();
                extend(app);
                app.vim_count = None;
            }
            KeyCode::Down | KeyCode::Char('j') => {
                let max_row = app.workbook.current_sheet().rows.saturating_sub(1);
                for _ in 0..take_count {
                    if app.selected_row >= max_row { break; }
                    app.selected_row += 1;
                }
                app.ensure_cursor_visible();
                extend(app);
                app.vim_count = None;
            }
            KeyCode::Left | KeyCode::Char('h') if !matches!(kind, VisualKind::Row) => {
                for _ in 0..take_count {
                    if app.selected_col == 0 { break; }
                    app.selected_col -= 1;
                }
                app.ensure_cursor_visible();
                extend(app);
                app.vim_count = None;
            }
            KeyCode::Right | KeyCode::Char('l') if !matches!(kind, VisualKind::Row) => {
                let max_col = app.workbook.current_sheet().cols.saturating_sub(1);
                for _ in 0..take_count {
                    if app.selected_col >= max_col { break; }
                    app.selected_col += 1;
                }
                app.ensure_cursor_visible();
                extend(app);
                app.vim_count = None;
            }
            KeyCode::Char('y') => {
                if let Some(((r0, c0), (r1, c1))) = app.get_selection_range() {
                    app.vim_apply_operator(VimOperator::Yank, r0, c0, r1, c1);
                }
                app.vim_exit_visual();
            }
            KeyCode::Char('d') | KeyCode::Char('x') => {
                if let Some(((r0, c0), (r1, c1))) = app.get_selection_range() {
                    app.vim_apply_operator(VimOperator::Delete, r0, c0, r1, c1);
                }
                app.vim_exit_visual();
            }
            KeyCode::Char('c') => {
                if let Some(((r0, c0), (r1, c1))) = app.get_selection_range() {
                    app.vim_apply_operator(VimOperator::Change, r0, c0, r1, c1);
                }
                // Change leaves us in Editing mode; don't exit_visual (would override).
            }
            KeyCode::Char('p') => {
                // Paste replaces the selection.
                if let Some(((r0, c0), _)) = app.get_selection_range() {
                    app.selected_row = r0;
                    app.selected_col = c0;
                }
                app.paste();
                app.vim_exit_visual();
            }
            KeyCode::Char('0') | KeyCode::Home if app.vim_count.is_none() && !matches!(kind, VisualKind::Row) => {
                app.vim_motion_row_start(); extend(app);
            }
            KeyCode::Char('$') | KeyCode::End if !matches!(kind, VisualKind::Row) => {
                app.vim_motion_row_end(); extend(app); app.vim_count = None;
            }
            KeyCode::Char('g') => { app.vim_awaiting_g = true; }
            KeyCode::Char('G') => {
                if let Some(n) = app.vim_count { app.vim_motion_goto_row(n); }
                else { app.vim_motion_bottom(); }
                extend(app);
                app.vim_count = None;
            }
            KeyCode::Char(':') => { app.start_command_palette(); }
            KeyCode::Char('q') => {
                // Exit visual first, then request_quit (with dirty-check).
                app.vim_exit_visual();
                app.request_quit();
                return;
            }
            KeyCode::Char(c) if c.is_ascii_digit() && !(c == '0' && app.vim_count.is_none()) => {
                let d = c.to_digit(10).unwrap() as usize;
                app.vim_count = Some(app.vim_count.unwrap_or(0) * 10 + d);
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
    fn test_ctrl_i_no_longer_starts_csv_import() {
        // Ctrl+I === Tab on most terminals, so we removed the binding.
        // It should now be a no-op in normal mode (Ctrl+L still imports CSV).
        let mut app = App::default();
        assert!(matches!(app.mode, AppMode::Normal));
        InputHandler::handle_key_event(&mut app, KeyCode::Char('i'), KeyModifiers::CONTROL);
        assert!(matches!(app.mode, AppMode::Normal));
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
    fn test_ctrl_v_enters_visual_block() {
        // Ctrl+V now enters Visual Block mode (vim convention).
        // Paste is on `p` or the command palette.
        let mut app = App::default();
        InputHandler::handle_key_event(&mut app, KeyCode::Char('v'), KeyModifiers::CONTROL);
        assert!(matches!(app.mode, AppMode::Visual { kind: VisualKind::Block }));
    }

    #[test]
    fn test_ctrl_h_no_longer_starts_find_replace() {
        // Ctrl+H === Backspace on most terminals; binding removed.
        let mut app = App::default();
        InputHandler::handle_key_event(&mut app, KeyCode::Char('h'), KeyModifiers::CONTROL);
        assert!(matches!(app.mode, AppMode::Normal));
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

    // -------------------------------------------------------------------
    // Vim-mode tests
    // -------------------------------------------------------------------

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
    fn test_i_enters_insert_mode() {
        let mut app = App::default();
        typestr(&mut app, "i");
        assert!(matches!(app.mode, AppMode::Editing));
    }

    #[test]
    fn test_A_enters_insert_at_end() {
        let mut app = App::default();
        app.workbook.current_sheet_mut().set_cell(0, 0, crate::domain::CellData {
            value: "hi".to_string(), formula: None, format: None, comment: None,
            spill_anchor: None,
        });
        InputHandler::handle_key_event(&mut app, KeyCode::Char('A'), KeyModifiers::NONE);
        assert!(matches!(app.mode, AppMode::Editing));
        assert_eq!(app.cursor_position, 2);
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
    fn test_v_enters_visual() {
        let mut app = App::default();
        typestr(&mut app, "v");
        assert!(matches!(app.mode, AppMode::Visual { kind: VisualKind::Cell }));
    }

    #[test]
    fn test_visual_motion_extends_then_d() {
        let mut app = App::default();
        for col in 0..3 {
            app.workbook.current_sheet_mut().set_cell(0, col, crate::domain::CellData {
                value: format!("v{}", col), formula: None, format: None, comment: None,
                spill_anchor: None,
            });
        }
        typestr(&mut app, "v");
        typestr(&mut app, "ll"); // extend two cells right
        typestr(&mut app, "d");  // delete selection
        assert!(matches!(app.mode, AppMode::Normal));
        assert!(app.workbook.current_sheet().get_cell(0, 0).value.is_empty());
        assert!(app.workbook.current_sheet().get_cell(0, 1).value.is_empty());
        assert!(app.workbook.current_sheet().get_cell(0, 2).value.is_empty());
    }

    #[test]
    fn test_V_enters_visual_line() {
        let mut app = App::default();
        typestr(&mut app, "V");
        assert!(matches!(app.mode, AppMode::Visual { kind: VisualKind::Row }));
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

    // -------------------------------------------------------------------
    // Agent probe tests: insert-mode entries
    // -------------------------------------------------------------------

    fn set_text(app: &mut App, row: usize, col: usize, v: &str) {
        app.workbook.current_sheet_mut().set_cell(row, col, crate::domain::CellData {
            value: v.to_string(), formula: None, format: None, comment: None,
            spill_anchor: None,
        });
    }

    #[test]
    fn agent_insert_i_cursor_position_on_existing_cell() {
        let mut app = App::default();
        set_text(&mut app, 0, 0, "hello");
        typestr(&mut app, "i");
        assert!(matches!(app.mode, AppMode::Editing));
        assert_eq!(app.input, "hello");
        // Vim `i` places the cursor at the start (before the first character).
        assert_eq!(app.cursor_position, 0,
            "`i` should place cursor at start of cell text (vim semantics).");
    }

    #[test]
    fn agent_insert_a_cursor_position_on_existing_cell() {
        let mut app = App::default();
        set_text(&mut app, 0, 0, "hello");
        typestr(&mut app, "a");
        assert!(matches!(app.mode, AppMode::Editing));
        assert_eq!(app.input, "hello");
        // `a` is vim's "append after cursor" — for a cell, cursor at end.
        assert_eq!(app.cursor_position, 5);
    }

    #[test]
    fn agent_insert_i_and_a_differ() {
        // `i` puts cursor at start (0); `a` puts cursor at end (text length).
        // They must behave differently on a non-empty cell.
        let mut app1 = App::default();
        let mut app2 = App::default();
        set_text(&mut app1, 0, 0, "hello");
        set_text(&mut app2, 0, 0, "hello");
        typestr(&mut app1, "i");
        typestr(&mut app2, "a");
        assert_eq!(app1.cursor_position, 0);
        assert_eq!(app2.cursor_position, 5);
    }

    #[test]
    fn agent_insert_I_cursor_at_start() {
        let mut app = App::default();
        set_text(&mut app, 0, 0, "hello");
        InputHandler::handle_key_event(&mut app, KeyCode::Char('I'), KeyModifiers::SHIFT);
        assert!(matches!(app.mode, AppMode::Editing));
        assert_eq!(app.cursor_position, 0);
        assert_eq!(app.input, "hello");
    }

    #[test]
    fn agent_insert_A_cursor_at_end() {
        let mut app = App::default();
        set_text(&mut app, 0, 0, "hello");
        InputHandler::handle_key_event(&mut app, KeyCode::Char('A'), KeyModifiers::SHIFT);
        assert!(matches!(app.mode, AppMode::Editing));
        assert_eq!(app.cursor_position, 5);
        assert_eq!(app.input, "hello");
    }

    #[test]
    fn agent_insert_o_grows_sheet_rows_when_at_last_row() {
        let mut app = App::default();
        let orig_rows = app.workbook.current_sheet().rows;
        app.selected_row = orig_rows - 1; // sit at last row
        typestr(&mut app, "o");
        assert!(matches!(app.mode, AppMode::Editing));
        // Expected: sheet rows grow by 1 to accommodate the new row below.
        assert_eq!(app.workbook.current_sheet().rows, orig_rows + 1,
            "`o` at last row should extend the sheet.");
        assert_eq!(app.selected_row, orig_rows);
    }

    #[test]
    fn agent_insert_O_at_row_zero_inserts_above() {
        let mut app = App::default();
        set_text(&mut app, 0, 0, "row0");
        assert_eq!(app.selected_row, 0);
        typestr(&mut app, "O");
        assert!(matches!(app.mode, AppMode::Editing));
        // O at row 0 inserts a new blank row at index 0 and shifts existing
        // content down. Cursor stays at row 0 (now the blank row).
        assert_eq!(app.selected_row, 0);
        assert_eq!(app.workbook.current_sheet().get_cell(1, 0).value, "row0",
            "Existing row should be shifted down by 1.");
        assert_eq!(app.input, "",
            "Edit buffer should be empty for the freshly-inserted row.");
    }

    #[test]
    fn agent_insert_s_clears_cell_and_enters_editing() {
        let mut app = App::default();
        set_text(&mut app, 0, 0, "old");
        typestr(&mut app, "s");
        assert!(matches!(app.mode, AppMode::Editing));
        assert_eq!(app.input, "",
            "After `s`, input should be empty (cell was cleared first).");
        assert!(app.workbook.current_sheet().get_cell(0, 0).value.is_empty());
        // Undo stack should record the clear.
        assert!(!app.undo_stack.is_empty(),
            "`s` should record an undo entry for the clear.");
    }

    #[test]
    fn agent_insert_S_clears_row_and_restores_undo() {
        let mut app = App::default();
        set_text(&mut app, 0, 0, "a");
        set_text(&mut app, 0, 1, "b");
        set_text(&mut app, 0, 2, "c");
        app.selected_col = 2;
        typestr(&mut app, "S");
        assert!(matches!(app.mode, AppMode::Editing));
        assert_eq!(app.selected_col, 0);
        assert!(app.workbook.current_sheet().get_cell(0, 0).value.is_empty());
        assert!(app.workbook.current_sheet().get_cell(0, 1).value.is_empty());
        assert!(app.workbook.current_sheet().get_cell(0, 2).value.is_empty());
        // Cancel out of editing, then undo.
        InputHandler::handle_key_event(&mut app, KeyCode::Esc, KeyModifiers::NONE);
        app.undo();
        assert_eq!(app.workbook.current_sheet().get_cell(0, 0).value, "a",
            "Undo after `S` should restore col 0.");
        assert_eq!(app.workbook.current_sheet().get_cell(0, 1).value, "b",
            "Undo after `S` should restore col 1.");
        assert_eq!(app.workbook.current_sheet().get_cell(0, 2).value, "c",
            "Undo after `S` should restore col 2.");
    }

    #[test]
    fn agent_insert_type_text_then_enter_saves_and_moves_down() {
        let mut app = App::default();
        typestr(&mut app, "i");
        typestr(&mut app, "test");
        key(&mut app, KeyCode::Enter);
        assert!(matches!(app.mode, AppMode::Normal));
        assert_eq!(app.workbook.current_sheet().get_cell(0, 0).value, "test");
        assert_eq!(app.selected_row, 1, "Enter after edit should move down.");
    }

    #[test]
    fn agent_insert_type_text_then_esc_discards() {
        let mut app = App::default();
        set_text(&mut app, 0, 0, "kept");
        typestr(&mut app, "i");
        typestr(&mut app, "X");
        key(&mut app, KeyCode::Esc);
        assert!(matches!(app.mode, AppMode::Normal));
        assert_eq!(app.workbook.current_sheet().get_cell(0, 0).value, "kept",
            "Esc cancels edits.");
    }

    #[test]
    fn agent_insert_s_then_esc_does_not_restore_cell() {
        let mut app = App::default();
        set_text(&mut app, 0, 0, "old");
        typestr(&mut app, "s");
        key(&mut app, KeyCode::Esc);
        assert!(matches!(app.mode, AppMode::Normal));
        // `s` already cleared the cell; Esc cancels the EDIT, not the clear.
        assert!(app.workbook.current_sheet().get_cell(0, 0).value.is_empty(),
            "Esc after `s` does not restore (cell was already cleared). Undo is the recovery path.");
        // Undo should restore.
        app.undo();
        assert_eq!(app.workbook.current_sheet().get_cell(0, 0).value, "old");
    }

    #[test]
    fn agent_insert_multibyte_cursor_is_char_based() {
        let mut app = App::default();
        set_text(&mut app, 0, 0, "café"); // 4 chars, 5 bytes
        // Use `a` to land cursor at end of the cell.
        typestr(&mut app, "a");
        assert!(matches!(app.mode, AppMode::Editing));
        assert_eq!(app.cursor_position, 4,
            "Cursor should be char-based at end of multibyte 'café' (4 chars, not 5 bytes).");
        key(&mut app, KeyCode::Backspace);
        assert_eq!(app.input, "caf");
        assert_eq!(app.cursor_position, 3);
    }

    #[test]
    fn agent_insert_tab_finishes_and_moves_right() {
        let mut app = App::default();
        typestr(&mut app, "i");
        typestr(&mut app, "v");
        key(&mut app, KeyCode::Tab);
        assert!(matches!(app.mode, AppMode::Normal));
        assert_eq!(app.selected_col, 1);
        assert_eq!(app.workbook.current_sheet().get_cell(0, 0).value, "v");
    }

    #[test]
    fn agent_insert_spill_anchor_blocks_editing() {
        let mut app = App::default();
        app.workbook.current_sheet_mut().set_cell(0, 0, crate::domain::CellData {
            value: "spilled".to_string(), formula: None, format: None, comment: None,
            spill_anchor: Some((5, 5)),
        });
        typestr(&mut app, "i");
        // Should stay in Normal mode and surface a hint.
        assert!(matches!(app.mode, AppMode::Normal),
            "Editing a spill ghost should be blocked.");
        assert!(app.status_message.as_ref().map(|s| s.contains("Read-only")).unwrap_or(false),
            "Should show read-only spill hint, got: {:?}", app.status_message);
    }

    #[test]
    fn agent_insert_i_in_visual_mode_does_not_enter_editing() {
        let mut app = App::default();
        typestr(&mut app, "v"); // enter visual
        assert!(matches!(app.mode, AppMode::Visual { .. }));
        typestr(&mut app, "i");
        // In visual mode, `i` is NOT handled — it's not a documented arm.
        // It currently is a no-op (or falls into the awaiting-g default branch).
        assert!(!matches!(app.mode, AppMode::Editing),
            "Visual+i should not jump straight to Editing.");
    }

    #[test]
    fn agent_insert_repeated_o_grows_sheet() {
        let mut app = App::default();
        let orig_rows = app.workbook.current_sheet().rows;
        // Sit at last row.
        app.selected_row = orig_rows - 1;
        for _ in 0..100 {
            // Each `o` enters editing; commit with Enter so we return to Normal.
            typestr(&mut app, "o");
            key(&mut app, KeyCode::Enter);
        }
        // Selected row advanced by 100; sheet should have grown to accommodate.
        assert!(app.workbook.current_sheet().rows >= orig_rows + 100,
            "After 100 `o` presses, sheet should grow. Got {} rows.",
            app.workbook.current_sheet().rows);
    }

    #[test]
    fn agent_insert_O_shifts_existing_rows_down() {
        // Vim's `O` inserts a new blank line above and shifts existing content
        // down. Tshts now matches this.
        let mut app = App::default();
        set_text(&mut app, 0, 0, "row0");
        set_text(&mut app, 1, 0, "row1");
        set_text(&mut app, 2, 0, "row2");
        app.selected_row = 2;
        typestr(&mut app, "O");
        assert!(matches!(app.mode, AppMode::Editing));
        // The row that was at index 2 is now at index 3; cursor stays at 2.
        assert_eq!(app.selected_row, 2);
        assert_eq!(app.workbook.current_sheet().get_cell(3, 0).value, "row2",
            "row2 should be shifted to index 3.");
        assert!(app.workbook.current_sheet().get_cell(2, 0).value.is_empty(),
            "Row at cursor index should be the freshly-inserted blank row.");
        assert_eq!(app.input, "", "Edit buffer should be empty.");
    }

    #[test]
    fn agent_insert_o_then_esc_leaves_sheet_grown() {
        // BUG: `o` grows the sheet but cancelling the edit leaves the sheet grown.
        let mut app = App::default();
        let orig_rows = app.workbook.current_sheet().rows;
        app.selected_row = orig_rows - 1;
        typestr(&mut app, "o");
        key(&mut app, KeyCode::Esc);
        assert!(matches!(app.mode, AppMode::Normal));
        assert_eq!(app.workbook.current_sheet().rows, orig_rows + 1,
            "Cancelling `o` does NOT rollback the row growth — observable side-effect.");
    }

    #[test]
    fn agent_insert_s_on_empty_cell_records_no_undo() {
        // `clear_cell_with_undo` only records if cell existed. On a blank cell `s`
        // should be a no-op for undo.
        let mut app = App::default();
        assert!(app.undo_stack.is_empty());
        typestr(&mut app, "s");
        assert!(matches!(app.mode, AppMode::Editing));
        // Undo stack should remain empty since there was nothing to clear.
        assert!(app.undo_stack.is_empty(),
            "`s` on an empty cell should not push an undo entry.");
    }

    // -------------------------------------------------------------------
    // Agent probe tests: visual mode
    // -------------------------------------------------------------------

    use crate::application::VisualKind;

    fn put(app: &mut App, r: usize, c: usize, v: &str) {
        app.workbook.current_sheet_mut().set_cell(r, c, crate::domain::CellData {
            value: v.to_string(), formula: None, format: None, comment: None,
            spill_anchor: None,
        });
    }

    #[test]
    fn agent_visual_enter_cell_anchors_selection() {
        let mut app = App::default();
        app.selected_row = 2;
        app.selected_col = 3;
        typestr(&mut app, "v");
        assert!(matches!(app.mode, AppMode::Visual { kind: VisualKind::Cell }));
        assert_eq!(app.selection_start, Some((2, 3)));
        assert_eq!(app.selection_end, Some((2, 3)));
    }

    #[test]
    fn agent_visual_enter_row_spans_full_row() {
        let mut app = App::default();
        app.selected_row = 4;
        app.selected_col = 2;
        let last_col = app.workbook.current_sheet().cols - 1;
        typestr(&mut app, "V");
        assert!(matches!(app.mode, AppMode::Visual { kind: VisualKind::Row }));
        assert_eq!(app.selection_start, Some((4, 0)));
        assert_eq!(app.selection_end, Some((4, last_col)));
    }

    #[test]
    fn agent_visual_enter_block_anchors_at_cursor() {
        let mut app = App::default();
        app.selected_row = 1;
        app.selected_col = 1;
        ctrl(&mut app, 'v');
        assert!(matches!(app.mode, AppMode::Visual { kind: VisualKind::Block }));
        assert_eq!(app.selection_start, Some((1, 1)));
        assert_eq!(app.selection_end, Some((1, 1)));
    }

    #[test]
    fn agent_visual_v3jl_produces_4x2_selection() {
        // Spec said: "v3jl should produce a 4x2 selection."
        let mut app = App::default();
        typestr(&mut app, "v");
        // 3j extends down 3 rows
        typestr(&mut app, "3j");
        // l extends right 1 col
        typestr(&mut app, "l");
        let range = app.get_selection_range().expect("selection should exist");
        let ((r0, c0), (r1, c1)) = range;
        assert_eq!((r0, c0), (0, 0));
        assert_eq!((r1, c1), (3, 1));
        let rows = r1 - r0 + 1;
        let cols = c1 - c0 + 1;
        assert_eq!((rows, cols), (4, 2), "expected 4 rows x 2 cols selection");
    }

    #[test]
    fn agent_visual_row_motion_snaps_full_row_width() {
        let mut app = App::default();
        let last_col = app.workbook.current_sheet().cols - 1;
        typestr(&mut app, "V");
        // Move down 2 rows
        typestr(&mut app, "jj");
        // After motion the selection should still span columns 0..=last_col,
        // and rows 0..=2.
        let ((r0, c0), (r1, c1)) = app.get_selection_range().unwrap();
        assert_eq!(r0, 0);
        assert_eq!(r1, 2);
        assert_eq!(c0, 0);
        assert_eq!(c1, last_col, "V-mode motion should keep full row width");
        // Then up 1 row: end should be at row 1.
        typestr(&mut app, "k");
        let ((r0, _), (r1, c1)) = app.get_selection_range().unwrap();
        assert_eq!(r0, 0);
        assert_eq!(r1, 1);
        assert_eq!(c1, last_col);
    }

    #[test]
    fn agent_visual_row_ignores_h_and_l() {
        // Spec #15: in Row visual, h and l should be ignored due to the
        // `if !matches!(kind, VisualKind::Row)` guard. They should NOT move
        // the cursor or change selection columns.
        let mut app = App::default();
        app.selected_col = 5;
        let last_col = app.workbook.current_sheet().cols - 1;
        typestr(&mut app, "V");
        let col_before = app.selected_col;
        typestr(&mut app, "h");
        assert_eq!(app.selected_col, col_before, "V-mode should ignore `h`");
        typestr(&mut app, "l");
        assert_eq!(app.selected_col, col_before, "V-mode should ignore `l`");
        let ((_, c0), (_, c1)) = app.get_selection_range().unwrap();
        assert_eq!(c0, 0);
        assert_eq!(c1, last_col);
    }

    #[test]
    fn agent_visual_block_motion_makes_rectangular_selection() {
        let mut app = App::default();
        ctrl(&mut app, 'v');
        typestr(&mut app, "jjll");
        let ((r0, c0), (r1, c1)) = app.get_selection_range().unwrap();
        assert_eq!((r0, c0, r1, c1), (0, 0, 2, 2));
    }

    #[test]
    fn agent_visual_switch_cell_to_row_via_V() {
        let mut app = App::default();
        typestr(&mut app, "v");
        assert!(matches!(app.mode, AppMode::Visual { kind: VisualKind::Cell }));
        typestr(&mut app, "V");
        assert!(matches!(app.mode, AppMode::Visual { kind: VisualKind::Row }));
    }

    #[test]
    fn agent_visual_switch_row_to_cell_via_v() {
        // Spec #5: From `V` hit `v` — should swap mode. BUT: per source,
        // matching `v` simply calls vim_exit_visual and returns to Normal,
        // it does NOT swap to cell visual. This test documents current
        // behaviour vs. spec.
        let mut app = App::default();
        typestr(&mut app, "V");
        assert!(matches!(app.mode, AppMode::Visual { kind: VisualKind::Row }));
        typestr(&mut app, "v");
        // Expected by spec: AppMode::Visual { kind: Cell }
        // Actual: AppMode::Normal (bug — `v` exits instead of swapping).
        assert!(matches!(app.mode, AppMode::Visual { kind: VisualKind::Cell }),
            "Expected V→v to swap to Cell visual; actually exits to Normal.");
    }

    #[test]
    fn agent_visual_y_yanks_and_exits_to_normal() {
        let mut app = App::default();
        put(&mut app, 0, 0, "a");
        put(&mut app, 0, 1, "b");
        typestr(&mut app, "vl");
        typestr(&mut app, "y");
        assert!(matches!(app.mode, AppMode::Normal));
        assert!(app.clipboard.is_some(), "y should populate clipboard");
        // Source cells should be unchanged after yank.
        assert_eq!(app.workbook.current_sheet().get_cell(0, 0).value, "a");
        assert_eq!(app.workbook.current_sheet().get_cell(0, 1).value, "b");
    }

    #[test]
    fn agent_visual_d_deletes_selection_and_exits() {
        let mut app = App::default();
        put(&mut app, 0, 0, "a");
        put(&mut app, 0, 1, "b");
        put(&mut app, 1, 0, "c");
        put(&mut app, 1, 1, "d");
        typestr(&mut app, "vjl"); // 2x2 selection
        typestr(&mut app, "d");
        assert!(matches!(app.mode, AppMode::Normal));
        for r in 0..2 {
            for c in 0..2 {
                assert!(app.workbook.current_sheet().get_cell(r, c).value.is_empty(),
                    "({},{}) should be empty after visual delete", r, c);
            }
        }
    }

    #[test]
    fn agent_visual_x_deletes_selection_and_exits() {
        let mut app = App::default();
        put(&mut app, 0, 0, "a");
        put(&mut app, 0, 1, "b");
        typestr(&mut app, "vl");
        typestr(&mut app, "x");
        assert!(matches!(app.mode, AppMode::Normal));
        assert!(app.workbook.current_sheet().get_cell(0, 0).value.is_empty());
        assert!(app.workbook.current_sheet().get_cell(0, 1).value.is_empty());
    }

    #[test]
    fn agent_visual_c_deletes_and_enters_editing_at_top_left() {
        // Spec #7: `c` in visual should enter Editing at top-left AFTER deleting.
        let mut app = App::default();
        put(&mut app, 0, 0, "a");
        put(&mut app, 0, 1, "b");
        put(&mut app, 1, 1, "d");
        // Anchor at (1,1), then extend up-left to (0,0).
        app.selected_row = 1;
        app.selected_col = 1;
        typestr(&mut app, "v");
        typestr(&mut app, "kh");
        typestr(&mut app, "c");
        assert!(matches!(app.mode, AppMode::Editing),
            "c should leave the user in Editing mode");
        // After deletion all selected cells should be empty.
        assert!(app.workbook.current_sheet().get_cell(0, 0).value.is_empty());
        assert!(app.workbook.current_sheet().get_cell(0, 1).value.is_empty());
        assert!(app.workbook.current_sheet().get_cell(1, 1).value.is_empty());
        // Cursor should be at top-left of the selection.
        assert_eq!((app.selected_row, app.selected_col), (0, 0),
            "c should park cursor at top-left of selection");
    }

    #[test]
    fn agent_visual_p_paste_replaces_selection_and_exits() {
        let mut app = App::default();
        // Build a single-cell clipboard by yanking (0,0).
        put(&mut app, 0, 0, "Z");
        typestr(&mut app, "vy");
        // Now go to a 2x2 region and paste.
        app.selected_row = 5;
        app.selected_col = 5;
        typestr(&mut app, "vjl"); // selection covers (5,5)..(6,6)
        typestr(&mut app, "p");
        assert!(matches!(app.mode, AppMode::Normal),
            "p in visual should return to Normal");
        // Paste from a single-cell clipboard at the top-left of the selection.
        assert_eq!(app.workbook.current_sheet().get_cell(5, 5).value, "Z");
    }

    #[test]
    fn agent_visual_esc_exits_and_clears_selection() {
        let mut app = App::default();
        typestr(&mut app, "vjl");
        assert!(app.selection_start.is_some());
        key(&mut app, KeyCode::Esc);
        assert!(matches!(app.mode, AppMode::Normal));
        assert!(app.selection_start.is_none(),
            "Esc should clear selection (via vim_exit_visual)");
        assert!(app.selection_end.is_none());
    }

    #[test]
    fn agent_visual_v_toggles_out_to_normal() {
        let mut app = App::default();
        typestr(&mut app, "v");
        typestr(&mut app, "v");
        assert!(matches!(app.mode, AppMode::Normal));
        assert!(app.selection_start.is_none());
    }

    #[test]
    fn agent_visual_count_5j_extends_5_rows() {
        // Spec #9: 5j in visual should extend by 5 rows.
        let mut app = App::default();
        typestr(&mut app, "v");
        typestr(&mut app, "5j");
        let ((r0, _), (r1, _)) = app.get_selection_range().unwrap();
        assert_eq!(r0, 0);
        assert_eq!(r1, 5, "5j in visual should extend exactly 5 rows");
        assert_eq!(app.selected_row, 5);
    }

    #[test]
    fn agent_visual_large_range_scrolls_with_cursor() {
        let mut app = App::default();
        app.viewport_rows = 5;
        app.viewport_cols = 5;
        typestr(&mut app, "v");
        typestr(&mut app, "20j");
        // Cursor should be visible — scroll_row should have advanced.
        assert!(app.scroll_row > 0,
            "selecting beyond viewport should scroll cursor into view");
        assert!(app.selected_row >= app.scroll_row);
        assert!(app.selected_row < app.scroll_row + app.viewport_rows);
    }

    #[test]
    fn agent_visual_cross_sheet_paste() {
        // Spec #11: enter visual, paste from another sheet's clipboard.
        let mut app = App::default();
        put(&mut app, 0, 0, "from_sheet_1");
        typestr(&mut app, "vy"); // clipboard now contains "from_sheet_1"
        // Add and switch to a second sheet via the workbook API.
        app.workbook.add_sheet("Sheet2".to_string());
        app.workbook.active_sheet = 1;
        app.selected_row = 2;
        app.selected_col = 2;
        typestr(&mut app, "v");
        typestr(&mut app, "p");
        assert!(matches!(app.mode, AppMode::Normal));
        assert_eq!(app.workbook.current_sheet().get_cell(2, 2).value, "from_sheet_1",
            "paste should drop clipboard content on the new sheet");
    }

    #[test]
    fn agent_visual_single_row_sheet_extension() {
        // Spec #12: in a sheet with only 1 row, j shouldn't extend past it.
        let mut app = App::default();
        // Forcibly shrink the sheet.
        app.workbook.current_sheet_mut().rows = 1;
        typestr(&mut app, "v");
        typestr(&mut app, "j");
        // Selection should remain at row 0 — j must not panic and must
        // saturate at the last available row.
        let ((r0, _), (r1, _)) = app.get_selection_range().unwrap();
        assert_eq!(r0, 0);
        assert_eq!(r1, 0);
        assert_eq!(app.selected_row, 0);
    }

    #[test]
    fn agent_visual_status_bar_shows_stats() {
        // Spec #13: SUM/AVG/COUNT for the selection should be available.
        let mut app = App::default();
        put(&mut app, 0, 0, "10");
        put(&mut app, 0, 1, "20");
        put(&mut app, 0, 2, "30");
        typestr(&mut app, "vll"); // (0,0)..(0,2)
        let stats = app.get_selection_stats();
        assert!(stats.is_some(), "stats should be present for multi-cell selection");
        let (sum, avg, count) = stats.unwrap();
        assert_eq!(count, 3);
        assert!((sum - 60.0).abs() < 1e-9);
        assert!((avg - 20.0).abs() < 1e-9);
    }

    #[test]
    fn agent_visual_esc_invokes_dismiss_transients() {
        // Esc in visual should behave like Esc in normal: clear selection
        // AND dismiss transient state (status, search highlights, etc.).
        let mut app = App::default();
        typestr(&mut app, "v");
        typestr(&mut app, "jl");
        app.status_message = Some("hello".to_string());
        app.search_results.push((9, 9));
        key(&mut app, KeyCode::Esc);
        assert!(matches!(app.mode, AppMode::Normal));
        assert!(app.selection_start.is_none(),
            "Esc in visual must clear selection");
        assert!(app.status_message.is_none(),
            "Esc in visual must dismiss the status message");
        assert!(app.search_results.is_empty(),
            "Esc in visual must clear search results");
    }

    #[test]
    fn agent_visual_d_on_empty_selection_undo_invariant() {
        // d/x on a fully-empty selection should not push spurious undo entries
        // (matches the existing pattern with `s`/`clear_cell_with_undo`).
        let mut app = App::default();
        let undo_before = app.undo_stack.len();
        typestr(&mut app, "vll");
        typestr(&mut app, "d");
        assert!(matches!(app.mode, AppMode::Normal));
        assert_eq!(app.undo_stack.len(), undo_before,
            "deleting all-empty selection should not add undo entries");
    }

    #[test]
    fn agent_visual_count_zero_dollar_motion() {
        // The `0` key is guarded by `app.vim_count.is_none()` so that `0` after
        // a digit acts as a digit, not a motion. Confirm `$` jumps to last
        // populated col in visual.
        let mut app = App::default();
        put(&mut app, 0, 7, "x");
        typestr(&mut app, "v");
        InputHandler::handle_key_event(&mut app, KeyCode::Char('$'), KeyModifiers::NONE);
        let ((_, c0), (_, c1)) = app.get_selection_range().unwrap();
        assert_eq!(c0, 0);
        assert_eq!(c1, 7);
        // Then 0 back to start.
        InputHandler::handle_key_event(&mut app, KeyCode::Char('0'), KeyModifiers::NONE);
        let (_, c1b) = app.selection_end.unwrap();
        assert_eq!(c1b, 0);
    }

    #[test]
    fn agent_visual_gg_extends_to_top() {
        let mut app = App::default();
        app.selected_row = 5;
        typestr(&mut app, "v");
        typestr(&mut app, "gg");
        let ((r0, _), (r1, _)) = app.get_selection_range().unwrap();
        // Anchored at row 5, then moved up to row 0: normalized to (0, _)..(5, _).
        assert_eq!(r0, 0);
        assert_eq!(r1, 5);
    }

    #[test]
    fn agent_visual_G_extends_to_bottom_data() {
        let mut app = App::default();
        put(&mut app, 9, 0, "bottom");
        typestr(&mut app, "v");
        InputHandler::handle_key_event(&mut app, KeyCode::Char('G'), KeyModifiers::NONE);
        let ((r0, _), (r1, _)) = app.get_selection_range().unwrap();
        assert_eq!(r0, 0);
        assert_eq!(r1, 9);
    }

    #[test]
    fn agent_visual_ctrl_c_copies_and_exits() {
        let mut app = App::default();
        put(&mut app, 0, 0, "a");
        put(&mut app, 0, 1, "b");
        typestr(&mut app, "vl");
        ctrl(&mut app, 'c');
        assert!(matches!(app.mode, AppMode::Normal));
        assert!(app.clipboard.is_some());
    }

    #[test]
    fn agent_visual_ctrl_x_cuts_and_exits() {
        let mut app = App::default();
        put(&mut app, 0, 0, "a");
        put(&mut app, 0, 1, "b");
        typestr(&mut app, "vl");
        ctrl(&mut app, 'x');
        assert!(matches!(app.mode, AppMode::Normal));
        assert!(app.clipboard.is_some());
        assert!(app.workbook.current_sheet().get_cell(0, 0).value.is_empty());
        assert!(app.workbook.current_sheet().get_cell(0, 1).value.is_empty());
    }

    #[test]
    fn agent_visual_row_mode_d_clears_full_row() {
        let mut app = App::default();
        for c in 0..5 {
            put(&mut app, 0, c, "x");
        }
        typestr(&mut app, "V");
        typestr(&mut app, "d");
        assert!(matches!(app.mode, AppMode::Normal));
        let last_col = app.workbook.current_sheet().cols - 1;
        for c in 0..=last_col {
            assert!(app.workbook.current_sheet().get_cell(0, c).value.is_empty(),
                "row visual delete should clear full row, col {} still set", c);
        }
    }

    #[test]
    fn agent_visual_ctrl_v_swaps_to_block() {
        // Ctrl+V in Cell visual swaps to Block visual (vim-style toggle).
        // Pressing Ctrl+V again from Block visual exits back to Normal.
        let mut app = App::default();
        typestr(&mut app, "v");
        assert!(matches!(app.mode, AppMode::Visual { kind: VisualKind::Cell }));
        ctrl(&mut app, 'v');
        assert!(matches!(app.mode, AppMode::Visual { kind: VisualKind::Block }));
        ctrl(&mut app, 'v');
        assert!(matches!(app.mode, AppMode::Normal));
    }

    // -------------------------------------------------------------------
    // Agent probes: mode-chip / status-bar correctness
    // -------------------------------------------------------------------

    #[test]
    fn agent_mode_chip_cycles_normal_insert_visual_command_search_help() {
        let mut app = App::default();
        assert!(matches!(app.mode, AppMode::Normal));
        typestr(&mut app, "i");
        assert!(matches!(app.mode, AppMode::Editing));
        key(&mut app, KeyCode::Esc);
        assert!(matches!(app.mode, AppMode::Normal));

        typestr(&mut app, "v");
        assert!(matches!(app.mode, AppMode::Visual { .. }));
        key(&mut app, KeyCode::Esc);
        assert!(matches!(app.mode, AppMode::Normal));

        typestr(&mut app, ":");
        assert!(matches!(app.mode, AppMode::CommandPalette));
        key(&mut app, KeyCode::Esc);
        assert!(matches!(app.mode, AppMode::Normal));

        typestr(&mut app, "/");
        assert!(matches!(app.mode, AppMode::Search));
        key(&mut app, KeyCode::Esc);
        assert!(matches!(app.mode, AppMode::Normal));

        typestr(&mut app, "?");
        assert!(matches!(app.mode, AppMode::Help));
        key(&mut app, KeyCode::Esc);
        assert!(matches!(app.mode, AppMode::Normal));
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
    fn agent_pending_d_then_v_cancels_pending_no_visual_mode() {
        let mut app = App::default();
        typestr(&mut app, "d");
        assert_eq!(app.vim_pending_op, Some(VimOperator::Delete));
        typestr(&mut app, "v");
        assert!(matches!(app.mode, AppMode::Normal),
            "BUG: After pending d, `v` is silently dropped (mode={:?}).", app.mode);
        assert!(app.vim_pending_op.is_none());
    }

    #[test]
    fn agent_visual_op_clears_vim_pending_state() {
        let mut app = App::default();
        put(&mut app, 0, 0, "a");
        typestr(&mut app, "v");
        typestr(&mut app, "y");
        assert!(matches!(app.mode, AppMode::Normal));
        assert!(app.vim_count.is_none());
        assert!(app.vim_pending_op.is_none());
        assert!(!app.vim_awaiting_g);
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
    fn agent_visual_kinds_all_distinct_via_input_drivers() {
        let mut app = App::default();
        typestr(&mut app, "v");
        assert!(matches!(app.mode, AppMode::Visual { kind: VisualKind::Cell }));
        key(&mut app, KeyCode::Esc);
        typestr(&mut app, "V");
        assert!(matches!(app.mode, AppMode::Visual { kind: VisualKind::Row }));
        key(&mut app, KeyCode::Esc);
        ctrl(&mut app, 'v');
        assert!(matches!(app.mode, AppMode::Visual { kind: VisualKind::Block }));
    }

    #[test]
    fn agent_pending_count_in_visual_then_5j_clears_count() {
        let mut app = App::default();
        app.workbook.current_sheet_mut().rows = 20;
        typestr(&mut app, "v");
        typestr(&mut app, "5j");
        assert!(app.vim_count.is_none(), "vim_count must clear after motion");
    }

    // -------------------------------------------------------------------
    // Agent grammar probe tests: vim operator/motion/count grammar
    // -------------------------------------------------------------------

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

    // 1. `5dd` clears 5 rows from current row down.
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

    // 2a. `dj` deletes current row + 1 row down at current column.
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

    // 2b. `dk` at row 2 deletes current + 1 up.
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

    // 2c/d. `dl` horizontal range.
    #[test]
    fn agent_grammar_dl_deletes_horizontal_range() {
        let mut app = App::default();
        fill_row(&mut app, 0, 4);
        typestr(&mut app, "dl");
        assert!(app.workbook.current_sheet().get_cell(0, 0).value.is_empty());
        assert!(app.workbook.current_sheet().get_cell(0, 1).value.is_empty());
        assert_eq!(app.workbook.current_sheet().get_cell(0, 2).value, "r0c2");
    }

    // 3. `dgg` deletes from current row up to row 0 across all columns.
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

    // 4. `dG` deletes from current row to last data row across all columns.
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

    // 5. `5dj` deletes current + 5 rows down (count applies to motion).
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

    // 6a. `d$` deletes from cursor to end of row's data.
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

    // 6b. `d0` deletes from row start to cursor.
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

    // 7. `y{motion}` places content in clipboard.cells.
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

    // 8. `c{motion}` deletes then enters Editing at top-left of range.
    #[test]
    fn agent_grammar_cj_enters_editing_at_topleft() {
        let mut app = App::default();
        fill_col(&mut app, 0, 4);
        app.selected_row = 2;
        // cj covers (2,0) and (3,0); top-left is (2,0)
        typestr(&mut app, "cj");
        assert!(matches!(app.mode, AppMode::Editing), "should be Editing after cj");
        assert_eq!(app.selected_row, 2);
        assert_eq!(app.selected_col, 0);
        assert!(app.workbook.current_sheet().get_cell(2, 0).value.is_empty());
        assert!(app.workbook.current_sheet().get_cell(3, 0).value.is_empty());
    }

    // 8b. `ck` (upward) — Editing cursor should land at top-left, i.e. higher row.
    #[test]
    fn agent_grammar_ck_enters_editing_at_top_of_range() {
        let mut app = App::default();
        fill_col(&mut app, 0, 4);
        app.selected_row = 3;
        typestr(&mut app, "ck");
        assert!(matches!(app.mode, AppMode::Editing));
        // After normalization, top-left of (3,0)..(2,0) is (2,0)
        assert_eq!(app.selected_row, 2,
            "BUG: c<motion-upward> should leave cursor at top of range, got {}", app.selected_row);
    }

    // 9. Esc after `5d` clears pending; subsequent `j` moves down exactly 1.
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

    // 10a. `dk` at row 0 must not panic and must not corrupt unrelated cells.
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

    // 10b. `dh` at col 0 must not panic.
    #[test]
    fn agent_grammar_dh_at_left_edge_no_panic() {
        let mut app = App::default();
        fill_row(&mut app, 0, 3);
        typestr(&mut app, "dh");
        // Range (0,0)..(0,0) — only current cell cleared.
        assert!(app.workbook.current_sheet().get_cell(0, 0).value.is_empty());
        assert_eq!(app.workbook.current_sheet().get_cell(0, 1).value, "r0c1");
    }

    // 11. `c` on empty range still enters Editing.
    #[test]
    fn agent_grammar_c_on_empty_cell_still_enters_editing() {
        let mut app = App::default();
        // Empty sheet; `cl` covers (0,0)..(0,1), nothing to delete.
        typestr(&mut app, "cl");
        assert!(matches!(app.mode, AppMode::Editing),
            "c<motion> must enter Editing even when range is empty");
        assert_eq!(app.selected_row, 0);
        assert_eq!(app.selected_col, 0);
    }

    // 12. `3yy` then `p` pastes 3 rows below current cursor.
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

    // Extra: `5dj` is 5*1 cells down (motion count applies). Verify state after.
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

    // ====================================================================
    // agent_quit_* — audit of ex-command quit/save flows
    // ====================================================================

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
        // Use process id + nanos to avoid collisions across parallel tests.
        let now = std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH).map(|d| d.as_nanos()).unwrap_or(0);
        p.push(format!("tshts_quit_{}_{}_{}.tshts", std::process::id(), now, name));
        p.to_string_lossy().to_string()
    }

    // (1a) `q` in normal mode, clean → should_quit
    #[test]
    fn agent_quit_q_clean_quits_immediately() {
        let mut app = App::default();
        typestr(&mut app, "q");
        assert!(app.should_quit, "clean `q` must set should_quit");
        assert!(!matches!(app.mode, AppMode::ConfirmDiscard));
    }

    // (1b) `q` in normal mode, dirty → ConfirmDiscard
    #[test]
    fn agent_quit_q_dirty_enters_confirm_discard() {
        let mut app = App::default();
        make_dirty(&mut app);
        typestr(&mut app, "q");
        assert!(matches!(app.mode, AppMode::ConfirmDiscard));
        assert!(!app.should_quit);
    }

    // (1c) ConfirmDiscard 'y' → quit
    #[test]
    fn agent_quit_confirm_y_quits() {
        let mut app = App::default();
        make_dirty(&mut app);
        typestr(&mut app, "q");
        key(&mut app, KeyCode::Char('y'));
        assert!(app.should_quit);
    }

    // (1d) ConfirmDiscard 'n' → cancel
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

    // (1e) ConfirmDiscard 's' → save THEN execute pending quit
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

    // (2a) `:q` (palette), clean → quits
    #[test]
    fn agent_quit_palette_q_clean_quits() {
        let mut app = App::default();
        run_palette(&mut app, "q");
        assert!(app.should_quit);
    }

    // (2b) `:q!` always quits, even when dirty
    #[test]
    fn agent_quit_palette_qbang_forces_quit() {
        let mut app = App::default();
        make_dirty(&mut app);
        run_palette(&mut app, "q!");
        assert!(app.should_quit);
        assert!(!matches!(app.mode, AppMode::ConfirmDiscard));
    }

    // (2c) `:w` then `:q` — after save dirty=false so :q quits
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

    // (3) `:w` with no filename → SaveAs dialog
    #[test]
    fn agent_quit_w_no_filename_opens_saveas() {
        let mut app = App::default();
        assert!(app.filename.is_none());
        make_dirty(&mut app);
        run_palette(&mut app, "w");
        assert!(matches!(app.mode, AppMode::SaveAs),
            ":w with no filename should open SaveAs dialog; mode={:?}", app.mode);
    }

    // (4) `:w foo.tshts` writes to that file & app.filename updates
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

    // (5a) `:wq` clean (no dirty) — but no filename: opens SaveAs, no quit
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

    // (5b) `:wq` with known filename → save + quit
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

    // (6a) `:wq!` saves and force-quits even if save couldn't happen (no filename → SaveAs)
    #[test]
    fn agent_quit_wqbang_force_quits_even_when_save_deferred() {
        let mut app = App::default();
        make_dirty(&mut app);
        // No filename — save_in_place_or_prompt opens SaveAs. Then `!` path force quits.
        run_palette(&mut app, "wq!");
        assert!(app.should_quit, ":wq! must force-quit");
    }

    // (6b) `:x!` mirrors `:wq!`
    #[test]
    fn agent_quit_xbang_force_quits() {
        let mut app = App::default();
        make_dirty(&mut app);
        run_palette(&mut app, "x!");
        assert!(app.should_quit);
    }

    // (6c) `:x` with filename → save & quit
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

    // (7a) `:e` no arg → LoadFile dialog (when clean)
    #[test]
    fn agent_quit_e_no_arg_opens_loadfile() {
        let mut app = App::default();
        run_palette(&mut app, "e");
        assert!(matches!(app.mode, AppMode::LoadFile),
            ":e with no arg should open LoadFile; mode={:?}", app.mode);
    }

    // (7b) `:e foo.tshts` (existing file) → loads it, filename set
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

    // (7c) `:e <file>` when a different file is already open and clean —
    // does it load the typed file, or the currently-open one?
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

    // (8a) `:Q` uppercase — does it work like `:q`?
    #[test]
    fn agent_quit_palette_uppercase_q_works() {
        let mut app = App::default();
        run_palette(&mut app, "Q");
        // execute_command lowercases — :Q should behave like :q.
        assert!(app.should_quit, ":Q must work like :q (lowercased)");
    }

    // (8b) `:WQ` uppercase
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

    // (8c) `:q ` (trailing space) — trim works
    #[test]
    fn agent_quit_palette_q_with_trailing_space() {
        let mut app = App::default();
        run_palette(&mut app, "q ");
        assert!(app.should_quit, ":q with trailing space should trim and quit");
    }

    // (8d) ` :q` — note start_command_palette already strips ':' (palette typed AFTER :).
    // Leading-space test: type a leading space then 'q'.
    #[test]
    fn agent_quit_palette_q_with_leading_space() {
        let mut app = App::default();
        run_palette(&mut app, " q");
        assert!(app.should_quit, ":q with leading space should trim and quit");
    }

    // (10) `q` while in Visual mode — does it request_quit?
    #[test]
    fn agent_quit_q_in_visual_mode() {
        let mut app = App::default();
        // Enter visual cell mode
        typestr(&mut app, "v");
        assert!(matches!(app.mode, AppMode::Visual { .. }));
        typestr(&mut app, "q");
        // Expect either request_quit (should_quit true OR ConfirmDiscard) — or
        // documented no-op. Either way we record actual behavior.
        // Clean state, so request_quit would set should_quit.
        assert!(app.should_quit,
            "BUG candidate: `q` in Visual mode does not quit. mode={:?} should_quit={}",
            app.mode, app.should_quit);
    }
}