//! Submodule of `state` — see state/mod.rs.

use super::*;

impl App {
    pub fn autofill_selection(&mut self) {
        if let Some(((start_row, start_col), (end_row, end_col))) = self.get_selection_range() {
            let num_rows = end_row - start_row + 1;
            let num_cols = end_col - start_col + 1;

            // Determine fill direction: true = fill down (by rows), false = fill right (by cols)
            let fill_down = num_rows >= num_cols;

            // Collect cells along the fill direction
            // For fill_down: iterate through rows for each column
            // For fill_right: iterate through columns for each row
            let mut changes = Vec::new();
            let mut pattern_desc = String::new();

            if fill_down {
                // Process each column independently
                for col in start_col..=end_col {
                    let (filled, desc) = self.autofill_column(start_row, end_row, col);
                    changes.extend(filled);
                    if pattern_desc.is_empty() && !desc.is_empty() {
                        pattern_desc = desc;
                    }
                }
            } else {
                // Process each row independently
                for row in start_row..=end_row {
                    let (filled, desc) = self.autofill_row(row, start_col, end_col);
                    changes.extend(filled);
                    if pattern_desc.is_empty() && !desc.is_empty() {
                        pattern_desc = desc;
                    }
                }
            }

            // Apply all changes
            let num_changes = changes.len();
            for (row, col, cell_data) in changes {
                self.set_cell_with_undo(row, col, cell_data);
            }

            if num_changes > 0 {
                self.status_message = Some(format!(
                    "Autofilled {} cells using {}",
                    num_changes,
                    pattern_desc
                ));
            } else {
                self.status_message = Some("No cells to fill".to_string());
            }
        }
    }

    fn autofill_column(&self, start_row: usize, end_row: usize, col: usize) -> (Vec<(usize, usize, CellData)>, String) {
        use crate::domain::services::{FormulaEvaluator, AutofillPattern};

        let mut changes = Vec::new();

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

        // If no pattern cells or no targets, nothing to do
        if pattern_cells.is_empty() || target_rows.is_empty() {
            return (changes, String::new());
        }

        // Check if any pattern cell has a formula - if so, use formula-based fill
        let has_formula = pattern_cells.iter().any(|(_, cell)| cell.formula.is_some());

        if has_formula {
            // Use the first cell with a formula as source, adjust references for targets
            let (source_row, source_cell) = pattern_cells.iter()
                .find(|(_, cell)| cell.formula.is_some())
                .unwrap();

            let evaluator = FormulaEvaluator::new(self.workbook.current_sheet());

            for target_row in &target_rows {
                let row_offset = *target_row as i32 - *source_row as i32;

                if let Some(ref formula) = source_cell.formula {
                    let adjusted_formula = evaluator.adjust_formula_references(formula, row_offset, 0);

                    if evaluator.would_create_circular_reference(&adjusted_formula, (*target_row, col)) {
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

            return (changes, "formula".to_string());
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

        (changes, pattern_desc)
    }

    fn autofill_row(&self, row: usize, start_col: usize, end_col: usize) -> (Vec<(usize, usize, CellData)>, String) {
        use crate::domain::services::{FormulaEvaluator, AutofillPattern};

        let mut changes = Vec::new();

        // Collect non-empty cells (pattern cells) and empty cells (target cells)
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

        // If no pattern cells or no targets, nothing to do
        if pattern_cells.is_empty() || target_cols.is_empty() {
            return (changes, String::new());
        }

        // Check if any pattern cell has a formula - if so, use formula-based fill
        let has_formula = pattern_cells.iter().any(|(_, cell)| cell.formula.is_some());

        if has_formula {
            // Use the first cell with a formula as source, adjust references for targets
            let (source_col, source_cell) = pattern_cells.iter()
                .find(|(_, cell)| cell.formula.is_some())
                .unwrap();

            let evaluator = FormulaEvaluator::new(self.workbook.current_sheet());

            for target_col in &target_cols {
                let col_offset = *target_col as i32 - *source_col as i32;

                if let Some(ref formula) = source_cell.formula {
                    let adjusted_formula = evaluator.adjust_formula_references(formula, 0, col_offset);

                    if evaluator.would_create_circular_reference(&adjusted_formula, (row, *target_col)) {
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

            return (changes, "formula".to_string());
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

        (changes, pattern_desc)
    }

}
