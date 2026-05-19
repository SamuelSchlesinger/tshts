//! Submodule of `input` — see input/mod.rs.

use super::*;
use crate::application::App;
use crossterm::event::KeyCode;

impl InputHandler {
    pub(super) fn handle_editing_mode(app: &mut App, key: KeyCode) {
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
    fn test_ctrl_i_no_longer_starts_csv_import() {
        // Ctrl+I === Tab on most terminals, so we removed the binding.
        // It should now be a no-op in normal mode (Ctrl+L still imports CSV).
        let mut app = App::default();
        assert!(matches!(app.mode, AppMode::Normal));
        InputHandler::handle_key_event(&mut app, KeyCode::Char('i'), KeyModifiers::CONTROL);
        assert!(matches!(app.mode, AppMode::Normal));
    }

    #[test]
    fn test_ctrl_h_no_longer_starts_find_replace() {
        // Ctrl+H === Backspace on most terminals; binding removed.
        let mut app = App::default();
        InputHandler::handle_key_event(&mut app, KeyCode::Char('h'), KeyModifiers::CONTROL);
        assert!(matches!(app.mode, AppMode::Normal));
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

}
