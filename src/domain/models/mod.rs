//! Domain models for the terminal spreadsheet application.
//!
//! Split into focused submodules:
//!   - style       — NumberFormat, TerminalColor, CellStyle, CellFormat,
//!                   format_cell_value
//!   - refs        — sheet-name rewriting helpers
//!   - cell        — CellData
//!   - spreadsheet — Spreadsheet struct (the workhorse) plus Table and
//!                   ConditionalFormat (tightly coupled types)
//!   - workbook    — Workbook (multi-sheet container) + cross-sheet
//!                   dependency graph

mod style;
mod refs;
mod cell;
mod spreadsheet;
mod workbook;

pub use style::{NumberFormat, TerminalColor, CellStyle, CellFormat, format_cell_value};
pub use refs::{rewrite_sheet_refs, rewrite_sheet_refs_for_name_value};
pub use cell::CellData;
pub use spreadsheet::{Spreadsheet, Table, ConditionalFormat};
pub use workbook::Workbook;
#[allow(unused_imports)]
pub use workbook::CrossSheetKey;

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_conditional_format_fires_on_truthy_predicate() {
        let mut sheet = Spreadsheet::default();
        sheet.set_cell(0, 0, CellData {
            value: "150".to_string(), formula: None, format: None, comment: None,
        spill_anchor: None,
        });
        sheet.set_cell(1, 0, CellData {
            value: "50".to_string(), formula: None, format: None, comment: None,
        spill_anchor: None,
        });
        sheet.conditional_formats.push(ConditionalFormat {
            column: 0,
            predicate: "_ > 100".to_string(),
            style: CellStyle {
                bold: true,
                underline: false,
                fg_color: Some(TerminalColor::Red),
                bg_color: None,
            },
        });
        let s0 = sheet.conditional_style_for(0, 0);
        assert!(s0.is_some());
        assert!(s0.as_ref().unwrap().bold);
        assert_eq!(s0.unwrap().fg_color, Some(TerminalColor::Red));
        // Row 1 doesn't satisfy the predicate.
        assert!(sheet.conditional_style_for(1, 0).is_none());
    }

    #[test]
    fn test_thousands_separator_edge_cases() {
        let fmt = CellFormat {
            number_format: NumberFormat::Number { decimals: 2, thousands_sep: true },
            style: CellStyle::default(),
        };
        assert_eq!(format_cell_value("1234567.89", &fmt), "1,234,567.89");
        assert_eq!(format_cell_value("-1234.5", &fmt), "-1,234.50");
        assert_eq!(format_cell_value("999.99", &fmt), "999.99");
        assert_eq!(format_cell_value("0", &fmt), "0.00");
        assert_eq!(format_cell_value("-0.5", &fmt), "-0.50");

        // Whole-million boundary
        let fmt0 = CellFormat {
            number_format: NumberFormat::Number { decimals: 0, thousands_sep: true },
            style: CellStyle::default(),
        };
        assert_eq!(format_cell_value("1000000", &fmt0), "1,000,000");
        assert_eq!(format_cell_value("-1000000", &fmt0), "-1,000,000");
    }

    #[test]
    fn test_cell_data_default() {
        let cell = CellData::default();
        assert!(cell.value.is_empty());
        assert!(cell.formula.is_none());
    }

    #[test]
    fn test_cell_data_creation() {
        let cell = CellData {
            value: "42".to_string(),
            formula: Some("=6*7".to_string()),
            format: None,
            comment: None,
        spill_anchor: None,
        };
        assert_eq!(cell.value, "42");
        assert_eq!(cell.formula.unwrap(), "=6*7");
    }

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
    fn test_automatic_recalculation() {
        let mut sheet = Spreadsheet::default();
        
        // Set up a simple dependency chain: C1 = A1 + B1
        sheet.set_cell(0, 0, CellData { value: "10".to_string(), formula: None, format: None, comment: None, spill_anchor: None }); // A1 = 10
        sheet.set_cell(0, 1, CellData { value: "20".to_string(), formula: None, format: None, comment: None, spill_anchor: None }); // B1 = 20
        sheet.set_cell(0, 2, CellData { 
            value: "30".to_string(), 
            formula: Some("=A1+B1".to_string()),
            format: None,
            comment: None,
        spill_anchor: None,
        }); // C1 = A1+B1 = 30
        
        // Verify initial state
        assert_eq!(sheet.get_cell(0, 2).value, "30");
        
        // Change A1 and verify C1 updates automatically
        sheet.set_cell(0, 0, CellData { value: "15".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        assert_eq!(sheet.get_cell(0, 2).value, "35"); // Should be 15+20=35
        
        // Change B1 and verify C1 updates automatically
        sheet.set_cell(0, 1, CellData { value: "25".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        assert_eq!(sheet.get_cell(0, 2).value, "40"); // Should be 15+25=40
    }

    #[test]
    fn test_dependency_chain_recalculation() {
        let mut sheet = Spreadsheet::default();
        
        // Set up a dependency chain: A1 -> B1 -> C1
        sheet.set_cell(0, 0, CellData { value: "5".to_string(), formula: None, format: None, comment: None, spill_anchor: None }); // A1 = 5
        sheet.set_cell(0, 1, CellData { 
            value: "10".to_string(), 
            formula: Some("=A1*2".to_string()),
            format: None,
            comment: None,
        spill_anchor: None,
        }); // B1 = A1*2 = 10
        sheet.set_cell(0, 2, CellData { 
            value: "20".to_string(), 
            formula: Some("=B1*2".to_string()),
            format: None,
            comment: None,
        spill_anchor: None,
        }); // C1 = B1*2 = 20
        
        // Verify initial state
        assert_eq!(sheet.get_cell(0, 0).value, "5");
        assert_eq!(sheet.get_cell(0, 1).value, "10");
        assert_eq!(sheet.get_cell(0, 2).value, "20");
        
        // Change A1 and verify the entire chain updates
        sheet.set_cell(0, 0, CellData { value: "10".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        assert_eq!(sheet.get_cell(0, 0).value, "10");
        assert_eq!(sheet.get_cell(0, 1).value, "20"); // 10*2=20
        assert_eq!(sheet.get_cell(0, 2).value, "40"); // 20*2=40
    }

    #[test]
    fn test_multiple_dependents() {
        let mut sheet = Spreadsheet::default();
        
        // Set up multiple cells depending on A1
        sheet.set_cell(0, 0, CellData { value: "10".to_string(), formula: None, format: None, comment: None, spill_anchor: None }); // A1 = 10
        sheet.set_cell(0, 1, CellData { 
            value: "11".to_string(), 
            formula: Some("=A1+1".to_string()),
            format: None,
            comment: None,
        spill_anchor: None,
        }); // B1 = A1+1 = 11
        sheet.set_cell(0, 2, CellData { 
            value: "20".to_string(), 
            formula: Some("=A1*2".to_string()),
            format: None,
            comment: None,
        spill_anchor: None,
        }); // C1 = A1*2 = 20
        sheet.set_cell(0, 3, CellData { 
            value: "100".to_string(), 
            formula: Some("=A1*A1".to_string()),
            format: None,
            comment: None,
        spill_anchor: None,
        }); // D1 = A1*A1 = 100
        
        // Verify initial state
        assert_eq!(sheet.get_cell(0, 1).value, "11");
        assert_eq!(sheet.get_cell(0, 2).value, "20");
        assert_eq!(sheet.get_cell(0, 3).value, "100");
        
        // Change A1 and verify all dependents update
        sheet.set_cell(0, 0, CellData { value: "5".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        assert_eq!(sheet.get_cell(0, 1).value, "6");   // 5+1=6
        assert_eq!(sheet.get_cell(0, 2).value, "10");  // 5*2=10
        assert_eq!(sheet.get_cell(0, 3).value, "25");  // 5*5=25
    }

    #[test]
    fn test_dependency_removal() {
        let mut sheet = Spreadsheet::default();
        
        // Set up a dependency
        sheet.set_cell(0, 0, CellData { value: "10".to_string(), formula: None, format: None, comment: None, spill_anchor: None }); // A1 = 10
        sheet.set_cell(0, 1, CellData { 
            value: "20".to_string(), 
            formula: Some("=A1*2".to_string()),
            format: None,
            comment: None,
        spill_anchor: None,
        }); // B1 = A1*2 = 20
        
        // Verify dependency exists
        assert_eq!(sheet.get_cell(0, 1).value, "20");
        
        // Change A1 and verify B1 updates
        sheet.set_cell(0, 0, CellData { value: "15".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        assert_eq!(sheet.get_cell(0, 1).value, "30");
        
        // Replace B1 with a constant value (remove dependency)
        sheet.set_cell(0, 1, CellData { value: "42".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        assert_eq!(sheet.get_cell(0, 1).value, "42");
        
        // Change A1 again - B1 should NOT update since dependency is removed
        sheet.set_cell(0, 0, CellData { value: "100".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        assert_eq!(sheet.get_cell(0, 1).value, "42"); // Should remain 42, not recalculate
    }

    #[test]
    fn test_rebuild_dependencies() {
        let mut sheet = Spreadsheet::default();
        
        // Manually insert cells with formulas (simulating loading from file)
        sheet.cells.insert((0, 0), CellData { value: "10".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        sheet.cells.insert((0, 1), CellData { 
            value: "20".to_string(), 
            formula: Some("=A1*2".to_string()),
            format: None,
            comment: None,
        spill_anchor: None,
        });
        sheet.cells.insert((0, 2), CellData { 
            value: "40".to_string(), 
            formula: Some("=B1*2".to_string()),
            format: None,
            comment: None,
        spill_anchor: None,
        });
        
        // At this point, dependencies are not tracked
        assert!(sheet.dependencies.is_empty());
        assert!(sheet.dependents.is_empty());
        
        // Rebuild dependencies
        sheet.rebuild_dependencies();
        
        // Verify dependencies are now tracked
        assert!(!sheet.dependencies.is_empty());
        assert!(!sheet.dependents.is_empty());
        
        // Test that recalculation works after rebuilding
        sheet.set_cell(0, 0, CellData { value: "5".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        assert_eq!(sheet.get_cell(0, 1).value, "10"); // 5*2=10
        assert_eq!(sheet.get_cell(0, 2).value, "20"); // 10*2=20
    }

    #[test]
    fn test_range_dependency_recalculation() {
        let mut sheet = Spreadsheet::default();
        
        // Set up cells A1:A3 and a SUM formula
        sheet.set_cell(0, 0, CellData { value: "1".to_string(), formula: None, format: None, comment: None, spill_anchor: None }); // A1 = 1
        sheet.set_cell(1, 0, CellData { value: "2".to_string(), formula: None, format: None, comment: None, spill_anchor: None }); // A2 = 2
        sheet.set_cell(2, 0, CellData { value: "3".to_string(), formula: None, format: None, comment: None, spill_anchor: None }); // A3 = 3
        sheet.set_cell(0, 1, CellData { 
            value: "6".to_string(), 
            formula: Some("=SUM(A1:A3)".to_string()),
            format: None,
            comment: None,
        spill_anchor: None,
        }); // B1 = SUM(A1:A3) = 6
        
        // Verify initial state
        assert_eq!(sheet.get_cell(0, 1).value, "6");
        
        // Change one cell in the range
        sheet.set_cell(1, 0, CellData { value: "5".to_string(), formula: None, format: None, comment: None, spill_anchor: None }); // A2 = 5
        assert_eq!(sheet.get_cell(0, 1).value, "9"); // Should be 1+5+3=9
        
        // Change another cell in the range
        sheet.set_cell(2, 0, CellData { value: "10".to_string(), formula: None, format: None, comment: None, spill_anchor: None }); // A3 = 10
        assert_eq!(sheet.get_cell(0, 1).value, "16"); // Should be 1+5+10=16
    }

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
    fn test_dependency_tracking_persistence() {
        use crate::infrastructure::FileRepository;
        use tempfile::NamedTempFile;
        
        let mut original = Spreadsheet::default();
        
        // Set up dependencies
        original.set_cell(0, 0, CellData { value: "10".to_string(), formula: None, format: None, comment: None, spill_anchor: None }); // A1 = 10
        original.set_cell(0, 1, CellData { 
            value: "20".to_string(), 
            formula: Some("=A1*2".to_string()),
            format: None,
            comment: None,
        spill_anchor: None,
        }); // B1 = A1*2 = 20
        original.set_cell(0, 2, CellData { 
            value: "40".to_string(), 
            formula: Some("=B1*2".to_string()),
            format: None,
            comment: None,
        spill_anchor: None,
        }); // C1 = B1*2 = 40
        
        // Save to file
        let temp_file = NamedTempFile::new().expect("Failed to create temp file");
        let file_path = temp_file.path().to_str().unwrap();
        FileRepository::save_spreadsheet(&original, file_path).expect("Save failed");
        
        // Load from file
        let (mut loaded, _) = FileRepository::load_spreadsheet(file_path).expect("Load failed");
        
        // Dependencies should be rebuilt and functional
        loaded.set_cell(0, 0, CellData { value: "5".to_string(), formula: None, format: None, comment: None, spill_anchor: None }); // Change A1 to 5
        
        // Verify that dependent cells were recalculated
        assert_eq!(loaded.get_cell(0, 1).value, "10"); // B1 = 5*2 = 10
        assert_eq!(loaded.get_cell(0, 2).value, "20"); // C1 = 10*2 = 20
    }

    #[test]
    fn test_diamond_dependency_recalculation() {
        let mut sheet = Spreadsheet::default();

        // Diamond pattern: A1 -> B1, A1 -> C1, B1 -> C1
        // A1 = 10
        sheet.set_cell(0, 0, CellData { value: "10".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        // B1 = A1 * 2
        sheet.set_cell(0, 1, CellData {
            value: "20".to_string(),
            formula: Some("=A1*2".to_string()),
            format: None,
            comment: None,
        spill_anchor: None,
        });
        // C1 = A1 + B1 (depends on both A1 and B1)
        sheet.set_cell(0, 2, CellData {
            value: "30".to_string(),
            formula: Some("=A1+B1".to_string()),
            format: None,
            comment: None,
        spill_anchor: None,
        });

        // Verify initial state
        assert_eq!(sheet.get_cell(0, 0).value, "10");
        assert_eq!(sheet.get_cell(0, 1).value, "20");
        assert_eq!(sheet.get_cell(0, 2).value, "30"); // 10 + 20

        // Change A1 — B1 must update before C1 for correct result
        sheet.set_cell(0, 0, CellData { value: "5".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        assert_eq!(sheet.get_cell(0, 1).value, "10"); // 5*2 = 10
        assert_eq!(sheet.get_cell(0, 2).value, "15"); // 5 + 10 = 15 (not 5 + 20 = 25)
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

    // === Number Formatting Tests ===

    #[test]
    fn test_format_cell_value_general() {
        let fmt = CellFormat { number_format: NumberFormat::General, ..CellFormat::default() };
        assert_eq!(super::format_cell_value("42.5", &fmt), "42.5");
        assert_eq!(super::format_cell_value("hello", &fmt), "hello");
    }

    #[test]
    fn test_format_cell_value_number() {
        let fmt = CellFormat { number_format: NumberFormat::Number { decimals: 2, thousands_sep: false }, ..CellFormat::default() };
        assert_eq!(super::format_cell_value("42", &fmt), "42.00");
        assert_eq!(super::format_cell_value("3.14159", &fmt), "3.14");
        assert_eq!(super::format_cell_value("hello", &fmt), "hello"); // non-numeric passthrough
    }

    #[test]
    fn test_format_cell_value_number_thousands() {
        let fmt = CellFormat { number_format: NumberFormat::Number { decimals: 2, thousands_sep: true }, ..CellFormat::default() };
        assert_eq!(super::format_cell_value("1234567.89", &fmt), "1,234,567.89");
        assert_eq!(super::format_cell_value("42", &fmt), "42.00");
    }

    #[test]
    fn test_format_cell_value_currency() {
        let fmt = CellFormat { number_format: NumberFormat::Currency { symbol: "$".to_string(), decimals: 2 }, ..CellFormat::default() };
        assert_eq!(super::format_cell_value("1234.5", &fmt), "$1,234.50");
        assert_eq!(super::format_cell_value("42", &fmt), "$42.00");
        assert_eq!(super::format_cell_value("hello", &fmt), "hello");
    }

    #[test]
    fn test_format_cell_value_percentage() {
        let fmt = CellFormat { number_format: NumberFormat::Percentage { decimals: 1 }, ..CellFormat::default() };
        assert_eq!(super::format_cell_value("0.75", &fmt), "75.0%");
        assert_eq!(super::format_cell_value("1", &fmt), "100.0%");
        assert_eq!(super::format_cell_value("0.123", &fmt), "12.3%");
    }

    #[test]
    fn test_thousands_separator() {
        assert_eq!(super::style::add_thousands_separator("1234567"), "1,234,567");
        assert_eq!(super::style::add_thousands_separator("123"), "123");
        assert_eq!(super::style::add_thousands_separator("1234.56"), "1,234.56");
        assert_eq!(super::style::add_thousands_separator("-1234567"), "-1,234,567");
    }

    // === Insert/Delete Row/Col Tests ===

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

    #[test]
    fn test_cell_data_format_serialization() {
        let cell = CellData {
            value: "100".to_string(),
            formula: None,
            format: Some(CellFormat {
                number_format: NumberFormat::Percentage { decimals: 1 },
                ..CellFormat::default()
            }),
            comment: None,
        spill_anchor: None,
        };
        let json = serde_json::to_string(&cell).unwrap();
        let deserialized: CellData = serde_json::from_str(&json).unwrap();
        assert_eq!(deserialized.value, "100");
        assert!(deserialized.format.is_some());
        assert!(matches!(deserialized.format.unwrap().number_format, NumberFormat::Percentage { decimals: 1 }));
    }
}