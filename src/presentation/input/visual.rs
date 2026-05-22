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
            }
            KeyCode::Char(c) if c.is_ascii_digit() && !(c == '0' && app.vim_count.is_none()) => {
                let d = c.to_digit(10).unwrap() as usize;
                app.vim_count = Some(app.vim_count.unwrap_or(0) * 10 + d);
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

    #[test]
    fn test_ctrl_v_enters_visual_block() {
        // Ctrl+V now enters Visual Block mode (vim convention).
        // Paste is on `p` or the command palette.
        let mut app = App::default();
        InputHandler::handle_key_event(&mut app, KeyCode::Char('v'), KeyModifiers::CONTROL);
        assert!(matches!(app.mode, AppMode::Visual { kind: VisualKind::Block }));
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
