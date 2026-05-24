use super::*;
use crate::domain::{CellData};
#[test]
fn test_spreadsheet_default() {
    let sheet = Spreadsheet::default();
    assert_eq!(sheet.rows, 100);
    assert_eq!(sheet.cols, 26);
    assert_eq!(sheet.default_column_width, 8);
    assert!(sheet.cells.is_empty());
    assert!(sheet.column_widths.is_empty());
}

#[test]
fn test_get_cell_empty() {
    let sheet = Spreadsheet::default();
    let cell = sheet.get_cell(0, 0);
    assert!(cell.value.is_empty());
    assert!(cell.formula.is_none());
}

#[test]
fn test_set_and_get_cell() {
    let mut sheet = Spreadsheet::default();
    let cell_data = CellData {
        value: "Hello".to_string(),
        formula: None,
        format: None,
        comment: None,
    spill_anchor: None,
    };
    sheet.set_cell(0, 0, cell_data.clone());
    
    let retrieved = sheet.get_cell(0, 0);
    assert_eq!(retrieved.value, "Hello");
    assert!(retrieved.formula.is_none());
}

#[test]
fn test_set_cell_no_auto_resize() {
    let mut sheet = Spreadsheet::default();
    let initial_width = sheet.get_column_width(0);
    
    let long_cell = CellData {
        value: "This is a very long cell value".to_string(),
        formula: None,
        format: None,
        comment: None,
    spill_anchor: None,
    };
    sheet.set_cell(0, 0, long_cell);
    
    let new_width = sheet.get_column_width(0);
    assert_eq!(new_width, initial_width); // No automatic resizing
}

#[test]
fn test_column_label() {
    assert_eq!(Spreadsheet::column_label(0), "A");
    assert_eq!(Spreadsheet::column_label(1), "B");
    assert_eq!(Spreadsheet::column_label(25), "Z");
    assert_eq!(Spreadsheet::column_label(26), "AA");
    assert_eq!(Spreadsheet::column_label(27), "AB");
    assert_eq!(Spreadsheet::column_label(51), "AZ");
    assert_eq!(Spreadsheet::column_label(52), "BA");
    assert_eq!(Spreadsheet::column_label(701), "ZZ");
    assert_eq!(Spreadsheet::column_label(702), "AAA");
}

#[test]
fn test_parse_cell_reference() {
    // Valid references
    assert_eq!(Spreadsheet::parse_cell_reference("A1"), Some((0, 0)));
    assert_eq!(Spreadsheet::parse_cell_reference("B2"), Some((1, 1)));
    assert_eq!(Spreadsheet::parse_cell_reference("Z26"), Some((25, 25)));
    assert_eq!(Spreadsheet::parse_cell_reference("AA1"), Some((0, 26)));
    assert_eq!(Spreadsheet::parse_cell_reference("AB100"), Some((99, 27)));
    
    // Case insensitive
    assert_eq!(Spreadsheet::parse_cell_reference("a1"), Some((0, 0)));
    assert_eq!(Spreadsheet::parse_cell_reference("b2"), Some((1, 1)));
    
    // Invalid references
    assert_eq!(Spreadsheet::parse_cell_reference(""), None);
    assert_eq!(Spreadsheet::parse_cell_reference("A"), None);
    assert_eq!(Spreadsheet::parse_cell_reference("1"), None);
    assert_eq!(Spreadsheet::parse_cell_reference("A0"), None); // Row 0 doesn't exist in Excel notation
    assert_eq!(Spreadsheet::parse_cell_reference("1A"), None);
    assert_eq!(Spreadsheet::parse_cell_reference("A1B"), None);
    assert_eq!(Spreadsheet::parse_cell_reference("A-1"), None);
}

#[test]
fn test_column_width_management() {
    let mut sheet = Spreadsheet::default();
    
    // Test default width
    assert_eq!(sheet.get_column_width(0), 8);
    
    // Test setting custom width
    sheet.set_column_width(0, 15);
    assert_eq!(sheet.get_column_width(0), 15);
    
    // Test other columns still use default
    assert_eq!(sheet.get_column_width(1), 8);
}

#[test]
fn test_auto_resize_column() {
    let mut sheet = Spreadsheet::default();
    
    // Add cells with varying lengths
    sheet.set_cell(0, 0, CellData { value: "Hi".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
    sheet.set_cell(1, 0, CellData { value: "Medium length".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
    sheet.set_cell(2, 0, CellData { value: "Very long content here".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
    
    sheet.auto_resize_column(0);
    let width = sheet.get_column_width(0);
    
    // Should be at least as wide as the longest content
    assert!(width >= "Very long content here".len());
    // But not more than the maximum of 50
    assert!(width <= 50);
}

#[test]
fn test_auto_resize_all_columns() {
    let mut sheet = Spreadsheet::default();
    
    // Add content to multiple columns
    sheet.set_cell(0, 0, CellData { value: "Short".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
    sheet.set_cell(0, 1, CellData { value: "Much longer content".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
    sheet.set_cell(0, 2, CellData { value: "X".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
    
    sheet.auto_resize_all_columns();
    
    // Each column should be sized appropriately
    assert!(sheet.get_column_width(0) >= 5); // "Short".len()
    assert!(sheet.get_column_width(1) >= 19); // "Much longer content".len()
    assert!(sheet.get_column_width(2) >= 3); // Minimum width
}

#[test]
fn test_formula_cell_no_auto_resize() {
    let mut sheet = Spreadsheet::default();
    let initial_width = sheet.get_column_width(0);
    
    let formula_cell = CellData {
        value: "42".to_string(),
        formula: Some("=SUM(A1:A10)".to_string()),
        format: None,
        comment: None,
    spill_anchor: None,
    };
    
    sheet.set_cell(0, 0, formula_cell);
    let width = sheet.get_column_width(0);
    
    // Width should remain unchanged (no automatic resizing)
    assert_eq!(width, initial_width);
}

#[test]
fn test_serialization_roundtrip() {
    let mut original = Spreadsheet::default();
    original.set_cell(0, 0, CellData {
        value: "test".to_string(),
        formula: Some("=1+1".to_string()),
        format: None,
        comment: None,
    spill_anchor: None,
    });
    original.set_cell(1, 1, CellData {
        value: "42".to_string(),
        formula: None,
        format: None,
        comment: None,
    spill_anchor: None,
    });
    original.set_column_width(0, 15);
    
    // Serialize to JSON
    let json = serde_json::to_string(&original).expect("Serialization failed");
    
    // Deserialize back
    let deserialized: Spreadsheet = serde_json::from_str(&json).expect("Deserialization failed");
    
    // Verify data integrity
    assert_eq!(deserialized.rows, original.rows);
    assert_eq!(deserialized.cols, original.cols);
    assert_eq!(deserialized.default_column_width, original.default_column_width);
    
    let cell_0_0 = deserialized.get_cell(0, 0);
    assert_eq!(cell_0_0.value, "test");
    assert_eq!(cell_0_0.formula.unwrap(), "=1+1");
    
    let cell_1_1 = deserialized.get_cell(1, 1);
    assert_eq!(cell_1_1.value, "42");
    assert!(cell_1_1.formula.is_none());
    
    assert_eq!(deserialized.get_column_width(0), 15);
}

#[test]
// NOTE: Per-sheet cascade tests (test_automatic_recalculation,
// test_dependency_chain_recalculation, test_multiple_dependents,
// test_dependency_removal, test_rebuild_dependencies,
// test_range_dependency_recalculation, test_dependency_tracking_persistence,
// test_diamond_dependency_recalculation) were deleted along with the
// per-sheet cascade itself. The functionality moved to
// Workbook::recalc_via_graph + the executor, and is covered by tests in
// src/domain/services/evaluator.rs and the scenario framework
// (tests/pty_scenarios.rs — DCF, amortization, budgeting all stress
// cumulative dependency chains through real user workflows).
fn _per_sheet_cascade_tests_moved_to_workbook_level() {}

#[test]
fn test_circular_dependency_handling() {
    let mut sheet = Spreadsheet::default();
    
    // Set up a potential circular dependency scenario
    sheet.set_cell(0, 0, CellData { value: "10".to_string(), formula: None, format: None, comment: None, spill_anchor: None }); // A1 = 10
    sheet.set_cell(0, 1, CellData { 
        value: "20".to_string(), 
        formula: Some("=A1*2".to_string()),
        format: None,
        comment: None,
    spill_anchor: None,
    }); // B1 = A1*2 = 20
    
    // Now try to create a circular dependency A1 = B1 + 1
    // This should be prevented by the circular reference check
    use crate::domain::services::FormulaEvaluator;
    let evaluator = FormulaEvaluator::new(&sheet);
    let would_be_circular = evaluator.would_create_circular_reference("=B1+1", (0, 0));
    assert!(would_be_circular); // Should detect the circular reference
    
    // The dependency system should also handle this gracefully
    // Even if somehow a circular dependency got through, recalculation should not hang
}

#[test]
fn test_extract_cell_references_from_formula() {
    use crate::domain::services::FormulaEvaluator;
    
    let sheet = Spreadsheet::default();
    let evaluator = FormulaEvaluator::new(&sheet);
    
    // Test simple cell reference
    let refs = evaluator.extract_cell_references("=A1");
    assert_eq!(refs, vec![(0, 0)]);
    
    // Test multiple cell references
    let refs = evaluator.extract_cell_references("=A1+B2*C3");
    assert_eq!(refs.len(), 3);
    assert!(refs.contains(&(0, 0))); // A1
    assert!(refs.contains(&(1, 1))); // B2
    assert!(refs.contains(&(2, 2))); // C3
    
    // Test range reference
    let refs = evaluator.extract_cell_references("=SUM(A1:A3)");
    assert_eq!(refs.len(), 3);
    assert!(refs.contains(&(0, 0))); // A1
    assert!(refs.contains(&(1, 0))); // A2
    assert!(refs.contains(&(2, 0))); // A3
    
    // Test no references
    let refs = evaluator.extract_cell_references("=5+10");
    assert!(refs.is_empty());
    
    // Test non-formula
    let refs = evaluator.extract_cell_references("Hello World");
    assert!(refs.is_empty());
}


#[test]
fn test_auto_resize_column_shrinks() {
    let mut sheet = Spreadsheet::default();

    // Add wide content and auto-resize
    sheet.set_cell(0, 0, CellData {
        value: "This is very wide content".to_string(),
        formula: None,
        format: None,
        comment: None,
    spill_anchor: None,
    });
    sheet.auto_resize_column(0);
    let wide_width = sheet.get_column_width(0);
    assert!(wide_width >= "This is very wide content".len());

    // Replace with short content
    sheet.set_cell(0, 0, CellData {
        value: "Hi".to_string(),
        formula: None,
        format: None,
        comment: None,
    spill_anchor: None,
    });
    sheet.auto_resize_column(0);
    let narrow_width = sheet.get_column_width(0);

    // Column should have shrunk
    assert!(narrow_width < wide_width);
    assert!(narrow_width >= 3); // minimum width
}

#[test]
fn test_insert_row() {
    let mut sheet = Spreadsheet::default();
    sheet.set_cell(0, 0, CellData { value: "A1".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
    sheet.set_cell(1, 0, CellData { value: "A2".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
    sheet.set_cell(2, 0, CellData { value: "A3".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
    let orig_rows = sheet.rows;

    sheet.insert_row(1); // Insert above row 1 (A2)

    assert_eq!(sheet.rows, orig_rows + 1);
    assert_eq!(sheet.get_cell(0, 0).value, "A1"); // Row 0 unchanged
    assert!(sheet.get_cell(1, 0).value.is_empty()); // New empty row
    assert_eq!(sheet.get_cell(2, 0).value, "A2"); // Shifted down
    assert_eq!(sheet.get_cell(3, 0).value, "A3"); // Shifted down
}

#[test]
fn test_delete_row() {
    let mut sheet = Spreadsheet::default();
    sheet.set_cell(0, 0, CellData { value: "A1".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
    sheet.set_cell(1, 0, CellData { value: "A2".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
    sheet.set_cell(2, 0, CellData { value: "A3".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
    let orig_rows = sheet.rows;

    sheet.delete_row(1); // Delete row 1 (A2)

    assert_eq!(sheet.rows, orig_rows - 1);
    assert_eq!(sheet.get_cell(0, 0).value, "A1"); // Row 0 unchanged
    assert_eq!(sheet.get_cell(1, 0).value, "A3"); // Shifted up
}

#[test]
fn test_insert_col() {
    let mut sheet = Spreadsheet::default();
    sheet.set_cell(0, 0, CellData { value: "A1".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
    sheet.set_cell(0, 1, CellData { value: "B1".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
    sheet.set_cell(0, 2, CellData { value: "C1".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
    let orig_cols = sheet.cols;

    sheet.insert_col(1); // Insert before column B

    assert_eq!(sheet.cols, orig_cols + 1);
    assert_eq!(sheet.get_cell(0, 0).value, "A1");
    assert!(sheet.get_cell(0, 1).value.is_empty()); // New empty column
    assert_eq!(sheet.get_cell(0, 2).value, "B1"); // Shifted right
    assert_eq!(sheet.get_cell(0, 3).value, "C1"); // Shifted right
}

#[test]
fn test_delete_col() {
    let mut sheet = Spreadsheet::default();
    sheet.set_cell(0, 0, CellData { value: "A1".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
    sheet.set_cell(0, 1, CellData { value: "B1".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
    sheet.set_cell(0, 2, CellData { value: "C1".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
    let orig_cols = sheet.cols;

    sheet.delete_col(1); // Delete column B

    assert_eq!(sheet.cols, orig_cols - 1);
    assert_eq!(sheet.get_cell(0, 0).value, "A1");
    assert_eq!(sheet.get_cell(0, 1).value, "C1"); // Shifted left
}

#[test]
fn test_insert_row_adjusts_formulas() {
    let mut sheet = Spreadsheet::default();
    sheet.set_cell(0, 0, CellData { value: "10".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
    sheet.set_cell(1, 0, CellData { value: "20".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
    sheet.set_cell(2, 0, CellData {
        value: "30".to_string(),
        formula: Some("=A1+A2".to_string()),
        format: None,
        comment: None,
    spill_anchor: None,
    });

    sheet.insert_row(1); // Insert before row 1

    // Formula should now reference A1+A3 (A2 shifted to A3)
    let cell = sheet.get_cell(3, 0); // Original row 2 moved to row 3
    assert!(cell.formula.is_some());
    let formula = cell.formula.unwrap();
    assert!(formula.contains("A3"), "Formula should reference A3, got: {}", formula);
}

#[test]
fn test_insert_col_adjusts_formulas() {
    let mut sheet = Spreadsheet::default();
    sheet.set_cell(0, 0, CellData { value: "10".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
    sheet.set_cell(0, 1, CellData { value: "20".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
    sheet.set_cell(0, 2, CellData {
        value: "30".to_string(),
        formula: Some("=A1+B1".to_string()),
        format: None,
        comment: None,
    spill_anchor: None,
    });

    sheet.insert_col(1); // Insert before column B

    // Formula should adjust: B1 -> C1
    let cell = sheet.get_cell(0, 3); // Original col 2 moved to col 3
    assert!(cell.formula.is_some());
    let formula = cell.formula.unwrap();
    assert!(formula.contains("C1"), "Formula should reference C1, got: {}", formula);
}

#[test]
fn test_cell_data_with_format() {
    let cell = CellData {
        value: "42".to_string(),
        formula: None,
        format: Some(CellFormat {
            number_format: NumberFormat::Currency { symbol: "$".to_string(), decimals: 2 },
            ..CellFormat::default()
        }),
        comment: None,
    spill_anchor: None,
    };
    assert!(cell.format.is_some());
    let fmt = cell.format.unwrap();
    assert!(matches!(fmt.number_format, NumberFormat::Currency { .. }));
}

