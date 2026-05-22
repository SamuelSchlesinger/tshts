//! Submodule of `state` — see state/mod.rs.

use super::*;

impl App {
    pub fn autofill_selection(&mut self) {
        if let Some(((start_row, start_col), (end_row, end_col))) = self.get_selection_range() {
            let num_rows = end_row - start_row + 1;
            let num_cols = end_col - start_col + 1;

            // Determine fill direction: true = fill down (by rows), false = fill right (by cols)
            let fill_down = num_rows >= num_cols;

            let mut changes = Vec::new();
            let mut pattern_desc = String::new();
            let mut skipped = 0_usize;

            if fill_down {
                for col in start_col..=end_col {
                    let (filled, desc, skip) = self.autofill_column(start_row, end_row, col);
                    changes.extend(filled);
                    skipped += skip;
                    if pattern_desc.is_empty() && !desc.is_empty() {
                        pattern_desc = desc;
                    }
                }
            } else {
                for row in start_row..=end_row {
                    let (filled, desc, skip) = self.autofill_row(row, start_col, end_col);
                    changes.extend(filled);
                    skipped += skip;
                    if pattern_desc.is_empty() && !desc.is_empty() {
                        pattern_desc = desc;
                    }
                }
            }

            let num_changes = changes.len();
            // Single batched undo entry rather than one per cell. A 20-cell
            // autofill used to require 20 separate `u` presses to roll back;
            // now it's one.
            self.set_many_with_undo(changes);

            if num_changes > 0 {
                let suffix = if skipped > 0 {
                    format!(" ({} skipped: would create circular ref)", skipped)
                } else {
                    String::new()
                };
                self.status_message = Some(format!(
                    "Autofilled {} cells using {}{}",
                    num_changes, pattern_desc, suffix
                ));
            } else if skipped > 0 {
                self.status_message = Some(format!(
                    "No cells filled ({} skipped: would create circular ref)",
                    skipped
                ));
            } else {
                self.status_message = Some("No cells to fill".to_string());
            }
        }
    }

    fn autofill_column(&self, start_row: usize, end_row: usize, col: usize) -> (Vec<(usize, usize, CellData)>, String, usize) {
        use crate::domain::services::{FormulaEvaluator, AutofillPattern};

        let mut changes = Vec::new();
        let mut skipped = 0_usize;

        // Collect non-empty cells (pattern cells) and empty cells (target cells)
        let mut pattern_cells: Vec<(usize, CellData)> = Vec::new();
        let mut target_rows: Vec<usize> = Vec::new();

        for row in start_row..=end_row {
            let cell = self.workbook.current_sheet().get_cell(row, col);
            if !cell.value.is_empty() || cell.formula.is_some() {
                pattern_cells.push((row, cell.clone()));
            } else {
                target_rows.push(row);
            }
        }

        if pattern_cells.is_empty() || target_rows.is_empty() {
            return (changes, String::new(), 0);
        }

        let has_formula = pattern_cells.iter().any(|(_, cell)| cell.formula.is_some());

        if has_formula {
            let (source_row, source_cell) = pattern_cells.iter()
                .find(|(_, cell)| cell.formula.is_some())
                .unwrap();

            let evaluator = FormulaEvaluator::new(self.workbook.current_sheet());

            for target_row in &target_rows {
                let row_offset = *target_row as i32 - *source_row as i32;

                if let Some(ref formula) = source_cell.formula {
                    let adjusted_formula = evaluator.adjust_formula_references(formula, row_offset, 0);

                    if evaluator.would_create_circular_reference(&adjusted_formula, (*target_row, col)) {
                        skipped += 1;
                        continue;
                    }

                    let new_value = evaluator.evaluate_formula(&adjusted_formula);
                    changes.push((*target_row, col, CellData {
                        value: new_value,
                        formula: Some(adjusted_formula),
                        format: None,
                        comment: None,
                    spill_anchor: None,
                    }));
                }
            }

            return (changes, "formula".to_string(), skipped);
        }

        // Extract values from pattern cells for pattern detection
        let values: Vec<String> = pattern_cells.iter()
            .map(|(_, cell)| cell.value.clone())
            .collect();

        let pattern = AutofillPattern::detect(&values);
        let pattern_desc = pattern.description();

        // Generate values for target cells
        // The pattern index for targets continues from where pattern cells left off
        let pattern_len = pattern_cells.len();

        for (i, target_row) in target_rows.iter().enumerate() {
            let pattern_index = pattern_len + i;
            let generated_value = pattern.generate(pattern_index);

            changes.push((*target_row, col, CellData {
                value: generated_value,
                formula: None,
                format: None,
                comment: None,
            spill_anchor: None,
            }));
        }

        (changes, pattern_desc, skipped)
    }

    fn autofill_row(&self, row: usize, start_col: usize, end_col: usize) -> (Vec<(usize, usize, CellData)>, String, usize) {
        use crate::domain::services::{FormulaEvaluator, AutofillPattern};

        let mut changes = Vec::new();
        let mut skipped = 0_usize;

        let mut pattern_cells: Vec<(usize, CellData)> = Vec::new();
        let mut target_cols: Vec<usize> = Vec::new();

        for col in start_col..=end_col {
            let cell = self.workbook.current_sheet().get_cell(row, col);
            if !cell.value.is_empty() || cell.formula.is_some() {
                pattern_cells.push((col, cell.clone()));
            } else {
                target_cols.push(col);
            }
        }

        if pattern_cells.is_empty() || target_cols.is_empty() {
            return (changes, String::new(), 0);
        }

        let has_formula = pattern_cells.iter().any(|(_, cell)| cell.formula.is_some());

        if has_formula {
            let (source_col, source_cell) = pattern_cells.iter()
                .find(|(_, cell)| cell.formula.is_some())
                .unwrap();

            let evaluator = FormulaEvaluator::new(self.workbook.current_sheet());

            for target_col in &target_cols {
                let col_offset = *target_col as i32 - *source_col as i32;

                if let Some(ref formula) = source_cell.formula {
                    let adjusted_formula = evaluator.adjust_formula_references(formula, 0, col_offset);

                    if evaluator.would_create_circular_reference(&adjusted_formula, (row, *target_col)) {
                        skipped += 1;
                        continue;
                    }

                    let new_value = evaluator.evaluate_formula(&adjusted_formula);
                    changes.push((row, *target_col, CellData {
                        value: new_value,
                        formula: Some(adjusted_formula),
                        format: None,
                        comment: None,
                    spill_anchor: None,
                    }));
                }
            }

            return (changes, "formula".to_string(), skipped);
        }

        // Extract values from pattern cells for pattern detection
        let values: Vec<String> = pattern_cells.iter()
            .map(|(_, cell)| cell.value.clone())
            .collect();

        let pattern = AutofillPattern::detect(&values);
        let pattern_desc = pattern.description();

        // Generate values for target cells
        // The pattern index for targets continues from where pattern cells left off
        let pattern_len = pattern_cells.len();

        for (i, target_col) in target_cols.iter().enumerate() {
            let pattern_index = pattern_len + i;
            let generated_value = pattern.generate(pattern_index);

            changes.push((row, *target_col, CellData {
                value: generated_value,
                formula: None,
                format: None,
                comment: None,
            spill_anchor: None,
            }));
        }

        (changes, pattern_desc, skipped)
    }

}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::domain::CellData;

    #[test]
    fn test_autofill_simple_values() {
        let mut app = App::default();

        // Set up a simple value in A1 (pattern cell)
        app.set_cell_with_undo(0, 0, CellData {
            value: "Hello".to_string(),
            formula: None,
            format: None,
            comment: None,
        spill_anchor: None,
        });

        // Select A1:A3 (vertical selection, A1 has value, A2-A3 are empty)
        app.selection_start = Some((0, 0));
        app.selection_end = Some((2, 0));

        // Autofill
        app.autofill_selection();

        // Check that the value was copied to empty cells only
        assert_eq!(app.workbook.current_sheet().get_cell(0, 0).value, "Hello"); // Original
        assert_eq!(app.workbook.current_sheet().get_cell(1, 0).value, "Hello"); // Filled
        assert_eq!(app.workbook.current_sheet().get_cell(2, 0).value, "Hello"); // Filled
    }

    #[test]
    fn test_autofill_formula_horizontal() {
        let mut app = App::default();

        // Set up cells with values for reference
        app.set_cell_with_undo(0, 1, CellData { value: "10".to_string(), formula: None, format: None, comment: None, spill_anchor: None }); // B1 = 10
        app.set_cell_with_undo(1, 1, CellData { value: "20".to_string(), formula: None, format: None, comment: None, spill_anchor: None }); // B2 = 20
        app.set_cell_with_undo(0, 2, CellData { value: "30".to_string(), formula: None, format: None, comment: None, spill_anchor: None }); // C1 = 30
        app.set_cell_with_undo(1, 2, CellData { value: "40".to_string(), formula: None, format: None, comment: None, spill_anchor: None }); // C2 = 40

        // Set up a formula in A1 that references B1:B2
        app.set_cell_with_undo(0, 0, CellData {
            value: "30".to_string(),
            formula: Some("=SUM(B1:B2)".to_string()),
            format: None,
            comment: None,
        spill_anchor: None,
        });

        // Select A1:D1 (horizontal autofill, A1 has formula, B1 has value, C1 has value, D1 is empty)
        app.selection_start = Some((0, 0));
        app.selection_end = Some((0, 3));

        // Autofill - should only fill D1 since it's the only empty cell
        app.autofill_selection();

        // Check that only the empty cell D1 got the adjusted formula
        let d1_cell = app.workbook.current_sheet().get_cell(0, 3);
        // The formula from A1 is adjusted by 3 columns: B->E, so =SUM(E1:E2)
        assert_eq!(d1_cell.formula, Some("=SUM(E1:E2)".to_string()));

        // Verify B1 and C1 still have their original values (not overwritten)
        assert_eq!(app.workbook.current_sheet().get_cell(0, 1).value, "10");
        assert_eq!(app.workbook.current_sheet().get_cell(0, 2).value, "30");
    }

    #[test]
    fn test_autofill_formula_vertical() {
        let mut app = App::default();

        // Set up cells with values for reference
        app.set_cell_with_undo(1, 0, CellData { value: "10".to_string(), formula: None, format: None, comment: None, spill_anchor: None }); // A2 = 10
        app.set_cell_with_undo(1, 1, CellData { value: "20".to_string(), formula: None, format: None, comment: None, spill_anchor: None }); // B2 = 20
        app.set_cell_with_undo(2, 0, CellData { value: "30".to_string(), formula: None, format: None, comment: None, spill_anchor: None }); // A3 = 30
        app.set_cell_with_undo(2, 1, CellData { value: "40".to_string(), formula: None, format: None, comment: None, spill_anchor: None }); // B3 = 40

        // Set up a formula in A1 that references A2+B2
        app.set_cell_with_undo(0, 0, CellData {
            value: "30".to_string(),
            formula: Some("=A2+B2".to_string()),
            format: None,
            comment: None,
        spill_anchor: None,
        });

        // Select A1:A4 (vertical autofill, A1 has formula, A2-A3 have values, A4 is empty)
        app.selection_start = Some((0, 0));
        app.selection_end = Some((3, 0));

        // Autofill - should only fill A4 since it's the only empty cell
        app.autofill_selection();

        // Check that only the empty cell A4 got the adjusted formula
        let a4_cell = app.workbook.current_sheet().get_cell(3, 0);
        // The formula from A1 is adjusted by 3 rows: A2->A5, B2->B5, so =A5+B5
        assert_eq!(a4_cell.formula, Some("=A5+B5".to_string()));

        // Verify A2 and A3 still have their original values (not overwritten)
        assert_eq!(app.workbook.current_sheet().get_cell(1, 0).value, "10");
        assert_eq!(app.workbook.current_sheet().get_cell(2, 0).value, "30");
    }

    #[test]
    fn test_autofill_pattern_arithmetic() {
        let mut app = App::default();

        // Set up arithmetic pattern: 1, 2, 3
        app.set_cell_with_undo(0, 0, CellData { value: "1".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        app.set_cell_with_undo(1, 0, CellData { value: "2".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        app.set_cell_with_undo(2, 0, CellData { value: "3".to_string(), formula: None, format: None, comment: None, spill_anchor: None });

        // Select A1:A6 (A1-A3 have values, A4-A6 are empty)
        app.selection_start = Some((0, 0));
        app.selection_end = Some((5, 0));

        // Autofill
        app.autofill_selection();

        // Check pattern continuation
        assert_eq!(app.workbook.current_sheet().get_cell(3, 0).value, "4");
        assert_eq!(app.workbook.current_sheet().get_cell(4, 0).value, "5");
        assert_eq!(app.workbook.current_sheet().get_cell(5, 0).value, "6");
    }

    #[test]
    fn test_autofill_pattern_days() {
        let mut app = App::default();

        // Set up days pattern: Mon, Tue
        app.set_cell_with_undo(0, 0, CellData { value: "Mon".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        app.set_cell_with_undo(1, 0, CellData { value: "Tue".to_string(), formula: None, format: None, comment: None, spill_anchor: None });

        // Select A1:A5 (A1-A2 have values, A3-A5 are empty)
        app.selection_start = Some((0, 0));
        app.selection_end = Some((4, 0));

        // Autofill
        app.autofill_selection();

        // Check pattern continuation
        assert_eq!(app.workbook.current_sheet().get_cell(2, 0).value, "Wed");
        assert_eq!(app.workbook.current_sheet().get_cell(3, 0).value, "Thu");
        assert_eq!(app.workbook.current_sheet().get_cell(4, 0).value, "Fri");
    }

    #[test]
    fn test_autofill_pattern_prefixed() {
        let mut app = App::default();

        // Set up prefixed pattern: Item1, Item2
        app.set_cell_with_undo(0, 0, CellData { value: "Item1".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        app.set_cell_with_undo(1, 0, CellData { value: "Item2".to_string(), formula: None, format: None, comment: None, spill_anchor: None });

        // Select A1:A5 (A1-A2 have values, A3-A5 are empty)
        app.selection_start = Some((0, 0));
        app.selection_end = Some((4, 0));

        // Autofill
        app.autofill_selection();

        // Check pattern continuation
        assert_eq!(app.workbook.current_sheet().get_cell(2, 0).value, "Item3");
        assert_eq!(app.workbook.current_sheet().get_cell(3, 0).value, "Item4");
        assert_eq!(app.workbook.current_sheet().get_cell(4, 0).value, "Item5");
    }

    #[test]
    fn test_autofill_pattern_months_short() {
        let mut app = App::default();

        // Set up months pattern: Jan, Feb, Mar
        app.set_cell_with_undo(0, 0, CellData { value: "Jan".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        app.set_cell_with_undo(1, 0, CellData { value: "Feb".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        app.set_cell_with_undo(2, 0, CellData { value: "Mar".to_string(), formula: None, format: None, comment: None, spill_anchor: None });

        // Select A1:A7 (A1-A3 have values, A4-A7 are empty)
        app.selection_start = Some((0, 0));
        app.selection_end = Some((6, 0));

        // Autofill
        app.autofill_selection();

        // Check pattern continuation
        assert_eq!(app.workbook.current_sheet().get_cell(3, 0).value, "Apr");
        assert_eq!(app.workbook.current_sheet().get_cell(4, 0).value, "May");
        assert_eq!(app.workbook.current_sheet().get_cell(5, 0).value, "Jun");
        assert_eq!(app.workbook.current_sheet().get_cell(6, 0).value, "Jul");
    }

    #[test]
    fn test_autofill_pattern_months_full() {
        let mut app = App::default();

        // Set up full months pattern: January, February
        app.set_cell_with_undo(0, 0, CellData { value: "January".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        app.set_cell_with_undo(1, 0, CellData { value: "February".to_string(), formula: None, format: None, comment: None, spill_anchor: None });

        // Select A1:A5 (A1-A2 have values, A3-A5 are empty)
        app.selection_start = Some((0, 0));
        app.selection_end = Some((4, 0));

        // Autofill
        app.autofill_selection();

        // Check pattern continuation
        assert_eq!(app.workbook.current_sheet().get_cell(2, 0).value, "March");
        assert_eq!(app.workbook.current_sheet().get_cell(3, 0).value, "April");
        assert_eq!(app.workbook.current_sheet().get_cell(4, 0).value, "May");
    }

    #[test]
    fn test_autofill_pattern_quarters() {
        let mut app = App::default();

        // Set up quarters pattern: Q1, Q2
        app.set_cell_with_undo(0, 0, CellData { value: "Q1".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        app.set_cell_with_undo(1, 0, CellData { value: "Q2".to_string(), formula: None, format: None, comment: None, spill_anchor: None });

        // Select A1:A6 (A1-A2 have values, A3-A6 are empty)
        app.selection_start = Some((0, 0));
        app.selection_end = Some((5, 0));

        // Autofill
        app.autofill_selection();

        // Check pattern continuation with wrap-around
        assert_eq!(app.workbook.current_sheet().get_cell(2, 0).value, "Q3");
        assert_eq!(app.workbook.current_sheet().get_cell(3, 0).value, "Q4");
        assert_eq!(app.workbook.current_sheet().get_cell(4, 0).value, "Q1"); // Wraps
        assert_eq!(app.workbook.current_sheet().get_cell(5, 0).value, "Q2"); // Wraps
    }

    #[test]
    fn test_autofill_pattern_months_wrap() {
        let mut app = App::default();

        // Set up months pattern starting near end: Oct, Nov, Dec
        app.set_cell_with_undo(0, 0, CellData { value: "Oct".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        app.set_cell_with_undo(1, 0, CellData { value: "Nov".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        app.set_cell_with_undo(2, 0, CellData { value: "Dec".to_string(), formula: None, format: None, comment: None, spill_anchor: None });

        // Select A1:A6 (A1-A3 have values, A4-A6 are empty)
        app.selection_start = Some((0, 0));
        app.selection_end = Some((5, 0));

        // Autofill
        app.autofill_selection();

        // Check pattern continuation with wrap-around
        assert_eq!(app.workbook.current_sheet().get_cell(3, 0).value, "Jan"); // Wraps
        assert_eq!(app.workbook.current_sheet().get_cell(4, 0).value, "Feb");
        assert_eq!(app.workbook.current_sheet().get_cell(5, 0).value, "Mar");
    }

    #[test]
    fn test_autofill_horizontal_pattern() {
        let mut app = App::default();

        // Set up arithmetic pattern horizontally: 10, 20 in A1, B1
        app.set_cell_with_undo(0, 0, CellData { value: "10".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        app.set_cell_with_undo(0, 1, CellData { value: "20".to_string(), formula: None, format: None, comment: None, spill_anchor: None });

        // Select A1:E1 (A1-B1 have values, C1-E1 are empty) - wide selection = fill right
        app.selection_start = Some((0, 0));
        app.selection_end = Some((0, 4));

        // Autofill
        app.autofill_selection();

        // Check pattern continuation
        assert_eq!(app.workbook.current_sheet().get_cell(0, 2).value, "30");
        assert_eq!(app.workbook.current_sheet().get_cell(0, 3).value, "40");
        assert_eq!(app.workbook.current_sheet().get_cell(0, 4).value, "50");
    }

}
