use super::*;
use crate::domain::CellData;

#[test]
fn test_command_palette_insert_row() {
    let mut app = App::default();
    app.set_cell_with_undo(0, 0, CellData { value: "A1".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
    app.set_cell_with_undo(1, 0, CellData { value: "A2".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
    let orig_rows = app.workbook.current_sheet().rows;

    app.selected_row = 1;
    app.start_command_palette();
    app.command_input = "ir".to_string();
    app.execute_command();

    assert_eq!(app.workbook.current_sheet().rows, orig_rows + 1);
    assert!(matches!(app.mode, AppMode::Normal));
}

#[test]
fn test_command_palette_delete_row() {
    let mut app = App::default();
    app.set_cell_with_undo(0, 0, CellData { value: "A1".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
    app.set_cell_with_undo(1, 0, CellData { value: "A2".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
    let orig_rows = app.workbook.current_sheet().rows;

    app.selected_row = 0;
    app.start_command_palette();
    app.command_input = "dr".to_string();
    app.execute_command();

    assert_eq!(app.workbook.current_sheet().rows, orig_rows - 1);
    assert_eq!(app.workbook.current_sheet().get_cell(0, 0).value, "A2"); // Shifted up
}

#[test]
fn test_command_palette_format_currency() {
    let mut app = App::default();
    app.set_cell_with_undo(0, 0, CellData { value: "1234.5".to_string(), formula: None, format: None, comment: None, spill_anchor: None });

    app.selected_row = 0;
    app.selected_col = 0;
    app.start_command_palette();
    app.command_input = "format currency".to_string();
    app.execute_command();

    let cell = app.workbook.current_sheet().get_cell(0, 0);
    assert!(cell.format.is_some());
    assert!(matches!(cell.format.unwrap().number_format, NumberFormat::Currency { .. }));
}

#[test]
fn test_command_palette_format_percentage() {
    let mut app = App::default();
    app.set_cell_with_undo(0, 0, CellData { value: "0.5".to_string(), formula: None, format: None, comment: None, spill_anchor: None });

    app.selected_row = 0;
    app.selected_col = 0;
    app.start_command_palette();
    app.command_input = "format percent 2".to_string();
    app.execute_command();

    let cell = app.workbook.current_sheet().get_cell(0, 0);
    assert!(cell.format.is_some());
    assert!(matches!(cell.format.unwrap().number_format, NumberFormat::Percentage { decimals: 2 }));
}

#[test]
fn test_command_palette_unknown_command() {
    let mut app = App::default();
    app.start_command_palette();
    app.command_input = "foobar".to_string();
    app.execute_command();

    assert!(app.status_message.as_ref().unwrap().contains("Unknown command"));
    assert!(matches!(app.mode, AppMode::Normal));
}

#[test]
fn test_command_bold() {
    let mut app = App::default();
    app.set_cell_with_undo(0, 0, CellData { value: "test".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
    app.selected_row = 0;
    app.selected_col = 0;

    app.start_command_palette();
    app.command_input = "bold".to_string();
    app.execute_command();

    let cell = app.workbook.current_sheet().get_cell(0, 0);
    assert!(cell.format.as_ref().unwrap().style.bold);
}

#[test]
fn test_command_underline() {
    let mut app = App::default();
    app.set_cell_with_undo(0, 0, CellData { value: "test".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
    app.selected_row = 0;
    app.selected_col = 0;

    app.start_command_palette();
    app.command_input = "underline".to_string();
    app.execute_command();

    let cell = app.workbook.current_sheet().get_cell(0, 0);
    assert!(cell.format.as_ref().unwrap().style.underline);
}

#[test]
fn test_command_color() {
    let mut app = App::default();
    app.set_cell_with_undo(0, 0, CellData { value: "test".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
    app.selected_row = 0;
    app.selected_col = 0;

    app.start_command_palette();
    app.command_input = "color red".to_string();
    app.execute_command();

    let cell = app.workbook.current_sheet().get_cell(0, 0);
    assert_eq!(cell.format.as_ref().unwrap().style.fg_color, Some(TerminalColor::Red));
}

#[test]
fn test_command_bg_color() {
    let mut app = App::default();
    app.set_cell_with_undo(0, 0, CellData { value: "test".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
    app.selected_row = 0;
    app.selected_col = 0;

    app.start_command_palette();
    app.command_input = "bg blue".to_string();
    app.execute_command();

    let cell = app.workbook.current_sheet().get_cell(0, 0);
    assert_eq!(cell.format.as_ref().unwrap().style.bg_color, Some(TerminalColor::Blue));
}

#[test]
fn test_command_color_none_clears() {
    let mut app = App::default();
    app.set_cell_with_undo(0, 0, CellData { value: "test".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
    app.selected_row = 0;
    app.selected_col = 0;

    // Set color first
    app.set_selection_fg_color(Some(TerminalColor::Red));
    assert_eq!(app.workbook.current_sheet().get_cell(0, 0).format.as_ref().unwrap().style.fg_color, Some(TerminalColor::Red));

    // Clear via command
    app.start_command_palette();
    app.command_input = "color none".to_string();
    app.execute_command();

    let cell = app.workbook.current_sheet().get_cell(0, 0);
    assert_eq!(cell.format.as_ref().unwrap().style.fg_color, None);
}

#[test]
fn test_add_sheet_command() {
    let mut app = App::default();
    app.start_command_palette();
    app.command_input = "sheet new".to_string();
    app.execute_command();

    assert_eq!(app.workbook.sheets.len(), 2);
    assert_eq!(app.workbook.active_sheet, 1); // Switched to new sheet
    assert_eq!(app.workbook.sheet_names[1], "Sheet2");
}

#[test]
fn test_delete_sheet_command() {
    let mut app = App::default();
    // Add a second sheet
    app.workbook.add_sheet("Sheet2".to_string());
    app.workbook.active_sheet = 1;

    app.start_command_palette();
    app.command_input = "sheet delete".to_string();
    app.execute_command();

    assert_eq!(app.workbook.sheets.len(), 1);
    assert_eq!(app.workbook.active_sheet, 0);
}

#[test]
fn test_rename_sheet_command() {
    let mut app = App::default();
    app.start_command_palette();
    app.command_input = "rename Revenue".to_string();
    app.execute_command();

    assert_eq!(app.workbook.sheet_names[0], "Revenue");
}

#[test]
fn test_comment_command() {
    let mut app = App::default();
    app.set_cell_with_undo(0, 0, CellData { value: "Data".to_string(), formula: None, format: None, comment: None, spill_anchor: None });

    app.start_command_palette();
    app.command_input = "comment Test note".to_string();
    app.execute_command();

    // Comment text preserves case (it's user-facing prose).
    let cell = app.workbook.current_sheet().get_cell(0, 0);
    assert_eq!(cell.comment, Some("Test note".to_string()));
}

#[test]
fn test_comment_clear_command() {
    let mut app = App::default();
    app.set_cell_with_undo(0, 0, CellData { value: "Data".to_string(), formula: None, format: None, comment: Some("note".to_string()), spill_anchor: None });

    app.start_command_palette();
    app.command_input = "comment clear".to_string();
    app.execute_command();

    let cell = app.workbook.current_sheet().get_cell(0, 0);
    assert_eq!(cell.comment, None);
}

#[test]
fn test_filter_command() {
    let mut app = App::default();
    app.set_cell_with_undo(0, 0, CellData { value: "Yes".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
    app.set_cell_with_undo(1, 0, CellData { value: "No".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
    app.set_cell_with_undo(2, 0, CellData { value: "Yes".to_string(), formula: None, format: None, comment: None, spill_anchor: None });

    app.start_command_palette();
    app.command_input = "filter a yes".to_string();
    app.execute_command();

    assert!(!app.hidden_rows.contains(&0));
    assert!(app.hidden_rows.contains(&1));
    assert!(!app.hidden_rows.contains(&2));
}

#[test]
fn test_unfilter_command() {
    let mut app = App::default();
    app.hidden_rows.insert(1);
    app.filter_column = Some(0);

    app.start_command_palette();
    app.command_input = "unfilter".to_string();
    app.execute_command();

    assert!(app.hidden_rows.is_empty());
    assert_eq!(app.filter_column, None);
}

