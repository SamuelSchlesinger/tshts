//! Submodule of `input` — see input/mod.rs.

use super::*;
use crate::application::{App, AppMode, VimOperator, VisualKind};
use crossterm::event::{KeyCode, KeyModifiers};

impl InputHandler {
    pub(super) fn handle_visual_mode(app: &mut App, key: KeyCode, modifiers: KeyModifiers) {
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

}
