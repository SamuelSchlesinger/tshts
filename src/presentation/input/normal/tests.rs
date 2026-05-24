// Test names use uppercase letters to mirror the vim keys they exercise
// (e.g. `test_G_with_count`, `agent_grammar_dG_*`); the `Default::default()`
// + sequential field assignment pattern reads more clearly than struct
// literals for test setup with many fields.
#![allow(non_snake_case, clippy::field_reassign_with_default)]

use super::*;
use crate::application::{App, AppMode, VimOperator};
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
fn agent_pending_d_then_question_opens_help() {
    let mut app = App::default();
    typestr(&mut app, "d");
    typestr(&mut app, "?");
    assert!(app.vim_pending_op.is_none());
    // After fix: `d?` cancels d and re-dispatches `?` so Help opens.
    assert!(matches!(app.mode, AppMode::Help));
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
    // `3yy` is a row-shaped yank (line operator with count), so `p` must
    // paste BELOW the cursor (row 1, col 0) — not shift the paste a column
    // to the right as it would for a cell-shaped clipboard.
    let landed_at_row_1_col_0 = app.workbook.current_sheet().get_cell(1, 0).value == "r0c0";
    let landed_at_row_0_col_1 = app.workbook.current_sheet().get_cell(0, 1).value == "r0c0";
    assert!(
        landed_at_row_1_col_0,
        "3yy then p should paste 3-row clipboard BELOW (row 1, col 0); \
         landed_at_row_1_col_0={} landed_at_row_0_col_1={}",
        landed_at_row_1_col_0, landed_at_row_0_col_1
    );
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

