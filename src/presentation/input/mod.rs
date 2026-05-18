use crate::application::{App, AppMode};
#[allow(unused_imports)] // re-exported via super::* for submodules
use crate::application::{VimOperator, VisualKind};
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


mod normal;
mod visual;
mod editing;
mod dialogs;

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