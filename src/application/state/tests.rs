use super::*;
use crate::domain::CellData;

#[test]
fn test_app_default() {
    let app = App::default();
    assert_eq!(app.selected_row, 0);
    assert_eq!(app.selected_col, 0);
    assert_eq!(app.scroll_row, 0);
    assert_eq!(app.scroll_col, 0);
    assert!(matches!(app.mode, AppMode::Normal));
    assert!(app.input.is_empty());
    assert_eq!(app.cursor_position, 0);
    assert!(app.filename.is_none());
    assert_eq!(app.help_scroll, 0);
    assert!(app.status_message.is_none());
    assert!(app.filename_input.is_empty());
}

#[test]
fn test_cross_sheet_auto_recalc() {
    // Sheet2!A1 = Sheet1!A1 + 10. Editing Sheet1!A1 should auto-update
    // Sheet2!A1 without a manual F5.
    let mut app = App::default();
    app.workbook.add_sheet("Sheet2".to_string());
    // Sheet1!A1 = 5
    app.workbook.active_sheet = 0;
    app.set_cell_with_undo(0, 0, CellData {
        value: "5".to_string(),
        formula: None,
        format: None,
        comment: None,
    spill_anchor: None,
    });
    // Sheet2!A1 = =Sheet1!A1 + 10  (evaluates to 15)
    app.workbook.active_sheet = 1;
    app.set_cell_with_undo(0, 0, CellData {
        value: "15".to_string(),
        formula: Some("=Sheet1!A1 + 10".to_string()),
        format: None,
        comment: None,
    spill_anchor: None,
    });
    assert_eq!(app.workbook.sheets[1].get_cell(0, 0).value, "15");

    // Now change Sheet1!A1 to 20. Sheet2!A1 should auto-update to 30.
    app.workbook.active_sheet = 0;
    app.set_cell_with_undo(0, 0, CellData {
        value: "20".to_string(),
        formula: None,
        format: None,
        comment: None,
    spill_anchor: None,
    });
    assert_eq!(app.workbook.sheets[1].get_cell(0, 0).value, "30");
}

#[test]
fn test_cross_sheet_cycle_rejected() {
    let mut app = App::default();
    app.workbook.add_sheet("Sheet2".to_string());
    // Sheet1!A1 = =Sheet2!A1
    app.workbook.active_sheet = 0;
    app.start_editing();
    app.input = "=Sheet2!A1".to_string();
    app.cursor_position = app.input.chars().count();
    app.finish_editing();
    // Sheet2!A1 = =Sheet1!A1 — should be rejected (cross-sheet cycle).
    app.workbook.active_sheet = 1;
    app.start_editing();
    app.input = "=Sheet1!A1".to_string();
    app.cursor_position = app.input.chars().count();
    app.finish_editing();
    // The reject path returns early without writing the formula, so
    // Sheet2!A1 stays empty/uninitialized.
    let cell = app.workbook.sheets[1].get_cell(0, 0);
    assert!(cell.formula.is_none(), "expected cross-sheet cycle rejected, got formula={:?}", cell.formula);
}

#[test]
fn test_cross_sheet_chain_propagates() {
    // Three-link chain: Sheet1!A1 → Sheet2!A1 → Sheet3!A1.
    let mut app = App::default();
    app.workbook.add_sheet("Sheet2".to_string());
    app.workbook.add_sheet("Sheet3".to_string());

    app.workbook.active_sheet = 0;
    app.set_cell_with_undo(0, 0, CellData {
        value: "1".to_string(),
        formula: None,
        format: None,
        comment: None,
    spill_anchor: None,
    });
    app.workbook.active_sheet = 1;
    app.set_cell_with_undo(0, 0, CellData {
        value: "2".to_string(),
        formula: Some("=Sheet1!A1 + 1".to_string()),
        format: None,
        comment: None,
    spill_anchor: None,
    });
    app.workbook.active_sheet = 2;
    app.set_cell_with_undo(0, 0, CellData {
        value: "3".to_string(),
        formula: Some("=Sheet2!A1 + 1".to_string()),
        format: None,
        comment: None,
    spill_anchor: None,
    });
    assert_eq!(app.workbook.sheets[2].get_cell(0, 0).value, "3");

    // Bump the head of the chain.
    app.workbook.active_sheet = 0;
    app.set_cell_with_undo(0, 0, CellData {
        value: "10".to_string(),
        formula: None,
        format: None,
        comment: None,
    spill_anchor: None,
    });
    assert_eq!(app.workbook.sheets[1].get_cell(0, 0).value, "11");
    assert_eq!(app.workbook.sheets[2].get_cell(0, 0).value, "12");
}

#[test]
fn test_smoke_end_to_end_flow() {
    // High-level sanity check exercising the key flows wired by the
    // recent refactors. Does not touch the terminal; only the App API.
    let mut app = App::default();
    assert!(!app.dirty);
    assert!(!app.should_quit);

    // Start an edit and commit a value via the normal Editing flow.
    app.start_editing();
    app.input = "12".to_string();
    app.cursor_position = 2;
    app.finish_editing();
    assert_eq!(app.workbook.current_sheet().get_cell(0, 0).value, "12");
    assert!(app.dirty);

    // A formula with an absolute reference round-trips through autofill.
    app.selected_row = 0;
    app.selected_col = 1;
    app.start_editing();
    app.input = "=A1*$B$5".to_string();
    app.cursor_position = app.input.chars().count();
    app.finish_editing();
    let evaluator = crate::domain::FormulaEvaluator::new(app.workbook.current_sheet());
    // $-anchored part survives an autofill row shift.
    let shifted = evaluator.adjust_formula_references("=A1*$B$5", 1, 0);
    assert_eq!(shifted, "=A2*$B$5");

    // Dirty-aware quit prompts.
    app.dirty = true;
    app.request_quit();
    assert!(matches!(app.mode, AppMode::ConfirmDiscard));
    assert!(!app.should_quit);
    app.confirm_pending_action();
    assert!(app.should_quit);

    // Esc-dismiss clears transient state.
    let mut app2 = App::default();
    app2.search_results.push((1, 1));
    app2.status_message = Some("noise".to_string());
    app2.dismiss_transients();
    assert!(app2.search_results.is_empty());
    assert!(app2.status_message.is_none());

    // recalc_all is callable and idempotent.
    app2.recalc_all();
    assert_eq!(
        app2.status_message.as_deref(),
        Some("Recalculated all formulas")
    );
}

#[test]
fn test_vlookup_basic() {
    use crate::domain::FormulaEvaluator;
    let mut sheet = crate::domain::Spreadsheet::default();
    // Single-column lookup
    for (i, v) in ["a", "b", "c", "d"].iter().enumerate() {
        sheet.set_cell(i, 0, crate::domain::CellData {
            value: v.to_string(), formula: None, format: None, comment: None,
        spill_anchor: None,
        });
    }
    let evaluator = FormulaEvaluator::new(&sheet);
    assert_eq!(
        evaluator.evaluate_formula("=VLOOKUP(\"c\", A1:A4, 1, 0)"),
        "c"
    );
}

#[test]
fn test_app_mode_transitions() {
    let mut app = App::default();
    
    // Normal -> Editing -> Normal
    assert!(matches!(app.mode, AppMode::Normal));
    app.start_editing();
    assert!(matches!(app.mode, AppMode::Editing));
    app.finish_editing();
    assert!(matches!(app.mode, AppMode::Normal));
    
    // Normal -> SaveAs -> Normal
    app.start_save_as();
    assert!(matches!(app.mode, AppMode::SaveAs));
    app.cancel_filename_input();
    assert!(matches!(app.mode, AppMode::Normal));
    
    // Normal -> LoadFile -> Normal
    app.start_load_file();
    assert!(matches!(app.mode, AppMode::LoadFile));
    app.cancel_filename_input();
    assert!(matches!(app.mode, AppMode::Normal));
}

#[test]
fn test_status_message_handling() {
    let mut app = App::default();
    
    // Initially no status message
    assert!(app.status_message.is_none());
    
    // Save success sets status message
    app.set_save_result(Ok("test.tshts".to_string()));
    assert!(app.status_message.is_some());
    
    // Starting save dialog clears status message
    app.start_save_as();
    assert!(app.status_message.is_none());
    
    // Load failure sets status message
    app.set_load_workbook_result(Err("Error".to_string()));
    assert!(app.status_message.is_some());
    
    // Starting load dialog clears status message
    app.start_load_file();
    assert!(app.status_message.is_none());
}

#[test]
fn test_selection_functionality() {
    let mut app = App::default();
    
    // Initially no selection
    assert!(app.get_selection_range().is_none());
    assert!(!app.is_cell_selected(0, 0));
    
    // Start selection
    app.start_selection();
    assert_eq!(app.get_selection_range(), Some(((0, 0), (0, 0))));
    assert!(app.is_cell_selected(0, 0));
    
    // Update selection
    app.update_selection(1, 2);
    assert_eq!(app.get_selection_range(), Some(((0, 0), (1, 2))));
    assert!(app.is_cell_selected(0, 1));
    assert!(app.is_cell_selected(1, 2));
    assert!(!app.is_cell_selected(2, 0));
    
    // Clear selection
    app.clear_selection();
    assert!(app.get_selection_range().is_none());
    assert!(!app.is_cell_selected(0, 0));
}

#[test]
fn test_viewport_and_scrolling() {
    let mut app = App::default();
    
    // Test initial viewport size
    assert_eq!(app.viewport_rows, 20);
    assert_eq!(app.viewport_cols, 8);
    
    // Test updating viewport size
    app.update_viewport_size(15, 10);
    assert_eq!(app.viewport_rows, 15);
    assert_eq!(app.viewport_cols, 10);
    
    // Test ensure_cursor_visible - cursor within viewport
    app.selected_row = 5;
    app.selected_col = 3;
    app.scroll_row = 0;
    app.scroll_col = 0;
    app.ensure_cursor_visible();
    assert_eq!(app.scroll_row, 0);  // No need to scroll
    assert_eq!(app.scroll_col, 0);
    
    // Test ensure_cursor_visible - cursor beyond bottom/right
    app.selected_row = 20;  // Beyond viewport (15 rows)
    app.selected_col = 12;  // Beyond viewport (10 cols)
    app.ensure_cursor_visible();
    assert_eq!(app.scroll_row, 6);  // 20 - 15 + 1 = 6
    assert_eq!(app.scroll_col, 3);  // 12 - 10 + 1 = 3
    
    // Test ensure_cursor_visible - cursor before top/left
    app.selected_row = 2;
    app.selected_col = 1;
    app.ensure_cursor_visible();
    assert_eq!(app.scroll_row, 2);  // Scroll to show cursor
    assert_eq!(app.scroll_col, 1);
}

#[test]
fn test_selection_stats() {
    let mut app = App::default();
    app.set_cell_with_undo(0, 0, CellData { value: "10".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
    app.set_cell_with_undo(1, 0, CellData { value: "20".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
    app.set_cell_with_undo(2, 0, CellData { value: "30".to_string(), formula: None, format: None, comment: None, spill_anchor: None });

    app.selection_start = Some((0, 0));
    app.selection_end = Some((2, 0));

    let stats = app.get_selection_stats();
    assert!(stats.is_some());
    let (sum, avg, count) = stats.unwrap();
    assert_eq!(sum, 60.0);
    assert_eq!(avg, 20.0);
    assert_eq!(count, 3);
}

#[test]
fn test_selection_stats_single_cell() {
    let mut app = App::default();
    app.set_cell_with_undo(0, 0, CellData { value: "10".to_string(), formula: None, format: None, comment: None, spill_anchor: None });

    app.selection_start = Some((0, 0));
    app.selection_end = Some((0, 0));

    // Single-cell selections now also publish stats (SUM=value,
    // AVG=value, COUNT=1) — useful for the user (the value of the
    // cell the cursor sits on) and required by the scenario test
    // framework's status-bar value reads.
    let (sum, avg, count) = app.get_selection_stats()
        .expect("single-cell selection should yield stats");
    assert_eq!(sum, 10.0);
    assert_eq!(avg, 10.0);
    assert_eq!(count, 1);
}

#[test]
fn test_selection_stats_no_numbers() {
    let mut app = App::default();
    app.set_cell_with_undo(0, 0, CellData { value: "hello".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
    app.set_cell_with_undo(1, 0, CellData { value: "world".to_string(), formula: None, format: None, comment: None, spill_anchor: None });

    app.selection_start = Some((0, 0));
    app.selection_end = Some((1, 0));

    // No numeric values should return None
    assert!(app.get_selection_stats().is_none());
}

#[test]
fn test_batch_undo() {
    let mut app = App::default();
    app.set_cell_with_undo(0, 0, CellData { value: "A".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
    app.set_cell_with_undo(1, 0, CellData { value: "B".to_string(), formula: None, format: None, comment: None, spill_anchor: None });

    // Cut = batch undo of clearing cells
    app.selection_start = Some((0, 0));
    app.selection_end = Some((1, 0));
    app.cut_selection();

    assert!(app.workbook.current_sheet().get_cell(0, 0).value.is_empty());
    assert!(app.workbook.current_sheet().get_cell(1, 0).value.is_empty());

    // Single undo should restore both cells
    app.undo();

    assert_eq!(app.workbook.current_sheet().get_cell(0, 0).value, "A");
    assert_eq!(app.workbook.current_sheet().get_cell(1, 0).value, "B");
}

#[test]
fn test_terminal_color_from_name() {
    assert_eq!(TerminalColor::from_name("red"), Some(TerminalColor::Red));
    assert_eq!(TerminalColor::from_name("Blue"), Some(TerminalColor::Blue));
    assert_eq!(TerminalColor::from_name("lightgreen"), Some(TerminalColor::LightGreen));
    assert_eq!(TerminalColor::from_name("CYAN"), Some(TerminalColor::Cyan));
    assert_eq!(TerminalColor::from_name("invalid"), None);
}

#[test]
fn test_parse_column_label() {
    use crate::domain::Spreadsheet;
    assert_eq!(Spreadsheet::parse_column_label("A"), Some(0));
    assert_eq!(Spreadsheet::parse_column_label("B"), Some(1));
    assert_eq!(Spreadsheet::parse_column_label("Z"), Some(25));
    assert_eq!(Spreadsheet::parse_column_label("AA"), Some(26));
    assert_eq!(Spreadsheet::parse_column_label("a"), Some(0));
    assert_eq!(Spreadsheet::parse_column_label(""), None);
    assert_eq!(Spreadsheet::parse_column_label("1"), None);
}

