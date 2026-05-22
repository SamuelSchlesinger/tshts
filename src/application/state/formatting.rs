//! Submodule of `state` — see state/mod.rs.

use super::*;

impl App {
    pub fn sort_column_asc(&mut self) {
        self.sort_column(true);
    }

    pub fn sort_column_desc(&mut self) {
        self.sort_column(false);
    }

    fn sort_column(&mut self, ascending: bool) {
        let col = self.selected_col;
        let mut max_row = 0;
        let mut max_col = 0;
        for &(r, c) in self.workbook.current_sheet().cells.keys() {
            max_row = max_row.max(r);
            max_col = max_col.max(c);
        }
        if max_row == 0 {
            self.status_message = Some("Nothing to sort".to_string());
            return;
        }

        // Capture each existing row as (original_row_index, Vec<Option<CellData>>).
        let mut rows: Vec<(usize, Vec<Option<CellData>>)> = Vec::with_capacity(max_row + 1);
        for row in 0..=max_row {
            let mut row_data = Vec::with_capacity(max_col + 1);
            for c in 0..=max_col {
                row_data.push(
                    self.workbook
                        .current_sheet()
                        .cells
                        .contains_key(&(row, c))
                        .then(|| self.workbook.current_sheet().get_cell(row, c)),
                );
            }
            rows.push((row, row_data));
        }

        rows.sort_by(|(_, a), (_, b)| {
            let a_val = a.get(col).and_then(|c| c.as_ref()).map(|c| &c.value);
            let b_val = b.get(col).and_then(|c| c.as_ref()).map(|c| &c.value);
            let cmp = match (a_val, b_val) {
                (None, None) => std::cmp::Ordering::Equal,
                (None, Some(_)) => std::cmp::Ordering::Greater,
                (Some(_), None) => std::cmp::Ordering::Less,
                (Some(a), Some(b)) => match (a.parse::<f64>(), b.parse::<f64>()) {
                    (Ok(an), Ok(bn)) => an.partial_cmp(&bn).unwrap_or(std::cmp::Ordering::Equal),
                    _ => a.cmp(b),
                },
            };
            if ascending { cmp } else { cmp.reverse() }
        });

        // Build old->new row mapping so we can rewrite intra-sort formula refs.
        let mut row_map: std::collections::HashMap<usize, usize> =
            std::collections::HashMap::with_capacity(rows.len());
        for (new_row, (old_row, _)) in rows.iter().enumerate() {
            row_map.insert(*old_row, new_row);
        }

        // Rewrite formula references that fall inside the sort range.
        let evaluator = FormulaEvaluator::new(self.workbook.current_sheet());
        let max_row_bound = max_row;
        for (_, row_cells) in rows.iter_mut() {
            for cell_opt in row_cells.iter_mut() {
                if let Some(cell) = cell_opt.as_mut()
                    && let Some(formula) = cell.formula.clone() {
                        let adjusted = evaluator.remap_row_references(&formula, &row_map, max_row_bound);
                        if adjusted != formula {
                            cell.formula = Some(adjusted);
                        }
                    }
            }
        }

        // Apply, batched for undo + single recalc via set_many.
        let mut batch = Vec::new();
        let mut writes: Vec<(usize, usize, CellData)> = Vec::new();
        let mut clears: Vec<(usize, usize)> = Vec::new();
        for (new_row, (_old, row_data)) in rows.iter().enumerate() {
            for (col_idx, cell_opt) in row_data.iter().enumerate() {
                let old = if self
                    .workbook
                    .current_sheet()
                    .cells
                    .contains_key(&(new_row, col_idx))
                {
                    Some(self.workbook.current_sheet().get_cell(new_row, col_idx))
                } else {
                    None
                };
                let new = cell_opt.clone();
                if old != new {
                    batch.push(UndoAction::CellModified {
                        row: new_row,
                        col: col_idx,
                        old_cell: old,
                        new_cell: new.clone(),
                    });
                    match new {
                        Some(cell) => writes.push((new_row, col_idx, cell)),
                        None => clears.push((new_row, col_idx)),
                    }
                }
            }
        }
        // Clears first (so dependent invalidation happens), then bulk writes.
        for (r, c) in clears {
            self.workbook.current_sheet_mut().clear_cell(r, c);
        }
        if !writes.is_empty() {
            self.workbook.current_sheet_mut().set_many(writes);
        }
        if !batch.is_empty() {
            self.record_action(UndoAction::Batch(batch));
        }
        let dir = if ascending { "ascending" } else { "descending" };
        self.status_message = Some(format!(
            "Sorted by column {} {}",
            crate::domain::Spreadsheet::column_label(col),
            dir
        ));
    }

    pub fn set_selection_format(&mut self, number_format: NumberFormat) {
        let range = if let Some(range) = self.get_selection_range() {
            range
        } else {
            ((self.selected_row, self.selected_col), (self.selected_row, self.selected_col))
        };
        let ((start_row, start_col), (end_row, end_col)) = range;
        let fmt_name = match &number_format {
            NumberFormat::General => "General",
            NumberFormat::Number { .. } => "Number",
            NumberFormat::Currency { .. } => "Currency",
            NumberFormat::Percentage { .. } => "Percentage",
        };
        let mut count = 0;
        for row in start_row..=end_row {
            for col in start_col..=end_col {
                let mut cell = self.workbook.current_sheet().get_cell(row, col);
                let format = match &number_format {
                    NumberFormat::General => None,
                    _ => {
                        let existing_style = cell.format.as_ref().map(|f| f.style.clone()).unwrap_or_default();
                        Some(CellFormat { number_format: number_format.clone(), style: existing_style })
                    }
                };
                cell.format = format;
                self.workbook.current_sheet_mut().set_cell(row, col, cell);
                count += 1;
            }
        }
        self.status_message = Some(format!("Applied {} format to {} cell(s)", fmt_name, count));
    }

    pub fn toggle_bold(&mut self) {
        let range = if let Some(range) = self.get_selection_range() {
            range
        } else {
            ((self.selected_row, self.selected_col), (self.selected_row, self.selected_col))
        };
        let ((start_row, start_col), (end_row, end_col)) = range;
        // Check current state from first cell to determine toggle direction
        let first_cell = self.workbook.current_sheet().get_cell(start_row, start_col);
        let currently_bold = first_cell.format.as_ref().map(|f| f.style.bold).unwrap_or(false);
        let new_bold = !currently_bold;

        for row in start_row..=end_row {
            for col in start_col..=end_col {
                let mut cell = self.workbook.current_sheet().get_cell(row, col);
                let mut fmt = cell.format.unwrap_or_default();
                fmt.style.bold = new_bold;
                cell.format = Some(fmt);
                self.workbook.current_sheet_mut().set_cell(row, col, cell);
            }
        }
        self.status_message = Some(format!("Bold {}", if new_bold { "on" } else { "off" }));
    }

    pub fn toggle_underline(&mut self) {
        let range = if let Some(range) = self.get_selection_range() {
            range
        } else {
            ((self.selected_row, self.selected_col), (self.selected_row, self.selected_col))
        };
        let ((start_row, start_col), (end_row, end_col)) = range;
        let first_cell = self.workbook.current_sheet().get_cell(start_row, start_col);
        let currently_underline = first_cell.format.as_ref().map(|f| f.style.underline).unwrap_or(false);
        let new_underline = !currently_underline;

        for row in start_row..=end_row {
            for col in start_col..=end_col {
                let mut cell = self.workbook.current_sheet().get_cell(row, col);
                let mut fmt = cell.format.unwrap_or_default();
                fmt.style.underline = new_underline;
                cell.format = Some(fmt);
                self.workbook.current_sheet_mut().set_cell(row, col, cell);
            }
        }
        self.status_message = Some(format!("Underline {}", if new_underline { "on" } else { "off" }));
    }

    pub fn set_selection_fg_color(&mut self, color: Option<TerminalColor>) {
        let range = if let Some(range) = self.get_selection_range() {
            range
        } else {
            ((self.selected_row, self.selected_col), (self.selected_row, self.selected_col))
        };
        let ((start_row, start_col), (end_row, end_col)) = range;
        for row in start_row..=end_row {
            for col in start_col..=end_col {
                let mut cell = self.workbook.current_sheet().get_cell(row, col);
                let mut fmt = cell.format.unwrap_or_default();
                fmt.style.fg_color = color.clone();
                cell.format = Some(fmt);
                self.workbook.current_sheet_mut().set_cell(row, col, cell);
            }
        }
        let color_name = color.as_ref().map(|c| format!("{:?}", c)).unwrap_or("default".to_string());
        self.status_message = Some(format!("Set foreground color to {}", color_name));
    }

    pub fn set_selection_bg_color(&mut self, color: Option<TerminalColor>) {
        let range = if let Some(range) = self.get_selection_range() {
            range
        } else {
            ((self.selected_row, self.selected_col), (self.selected_row, self.selected_col))
        };
        let ((start_row, start_col), (end_row, end_col)) = range;
        for row in start_row..=end_row {
            for col in start_col..=end_col {
                let mut cell = self.workbook.current_sheet().get_cell(row, col);
                let mut fmt = cell.format.unwrap_or_default();
                fmt.style.bg_color = color.clone();
                cell.format = Some(fmt);
                self.workbook.current_sheet_mut().set_cell(row, col, cell);
            }
        }
        let color_name = color.as_ref().map(|c| format!("{:?}", c)).unwrap_or("default".to_string());
        self.status_message = Some(format!("Set background color to {}", color_name));
    }

    pub fn set_cell_comment(&mut self, comment: Option<String>) {
        let row = self.selected_row;
        let col = self.selected_col;
        let exists = self.workbook.current_sheet().cells.contains_key(&(row, col));
        let mut cell = self.workbook.current_sheet().get_cell(row, col);
        let old_cell = if exists { Some(cell.clone()) } else { None };
        if cell.comment == comment {
            // No change, nothing to record.
            self.status_message = Some("Comment unchanged".to_string());
            return;
        }
        cell.comment = comment.clone();
        self.workbook.current_sheet_mut().set_cell(row, col, cell.clone());
        self.record_action(UndoAction::CellModified {
            row, col,
            old_cell,
            new_cell: Some(cell),
        });
        if let Some(ref text) = comment {
            self.status_message = Some(format!("Comment set: {}", text));
        } else {
            self.status_message = Some("Comment cleared".to_string());
        }
    }

    pub fn apply_filter(&mut self, col: usize, criteria: Option<String>) {
        self.hidden_rows.clear();
        self.filter_column = Some(col);
        if let Some(ref criteria) = criteria {
            self.filter_value = Some(criteria.clone());
            let criteria_lower = criteria.to_lowercase();
            // Find data extent
            let max_row = self.workbook.current_sheet().cells.keys()
                .map(|&(r, _)| r)
                .max()
                .unwrap_or(0);
            for row in 0..=max_row {
                let cell = self.workbook.current_sheet().get_cell(row, col);
                if !cell.value.to_lowercase().contains(&criteria_lower) {
                    self.hidden_rows.insert(row);
                }
            }
            let hidden_count = self.hidden_rows.len();
            self.status_message = Some(format!("Filter applied: {} rows hidden", hidden_count));
        } else {
            self.filter_value = None;
            self.status_message = Some(format!("Filter set on column {}", Spreadsheet::column_label(col)));
        }
    }

    pub fn clear_filter(&mut self) {
        self.hidden_rows.clear();
        self.filter_column = None;
        self.filter_value = None;
        self.status_message = Some("Filter cleared".to_string());
    }

}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::domain::CellData;

    #[test]
    fn test_sort_column_ascending() {
        let mut app = App::default();
        app.set_cell_with_undo(0, 0, CellData { value: "30".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        app.set_cell_with_undo(1, 0, CellData { value: "10".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        app.set_cell_with_undo(2, 0, CellData { value: "20".to_string(), formula: None, format: None, comment: None, spill_anchor: None });

        app.selected_col = 0;
        app.sort_column_asc();

        assert_eq!(app.workbook.current_sheet().get_cell(0, 0).value, "10");
        assert_eq!(app.workbook.current_sheet().get_cell(1, 0).value, "20");
        assert_eq!(app.workbook.current_sheet().get_cell(2, 0).value, "30");
    }

    #[test]
    fn test_sort_column_descending() {
        let mut app = App::default();
        app.set_cell_with_undo(0, 0, CellData { value: "10".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        app.set_cell_with_undo(1, 0, CellData { value: "30".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        app.set_cell_with_undo(2, 0, CellData { value: "20".to_string(), formula: None, format: None, comment: None, spill_anchor: None });

        app.selected_col = 0;
        app.sort_column_desc();

        assert_eq!(app.workbook.current_sheet().get_cell(0, 0).value, "30");
        assert_eq!(app.workbook.current_sheet().get_cell(1, 0).value, "20");
        assert_eq!(app.workbook.current_sheet().get_cell(2, 0).value, "10");
    }

    #[test]
    fn test_sort_preserves_other_columns() {
        let mut app = App::default();
        app.set_cell_with_undo(0, 0, CellData { value: "30".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        app.set_cell_with_undo(0, 1, CellData { value: "C".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        app.set_cell_with_undo(1, 0, CellData { value: "10".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        app.set_cell_with_undo(1, 1, CellData { value: "A".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        app.set_cell_with_undo(2, 0, CellData { value: "20".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        app.set_cell_with_undo(2, 1, CellData { value: "B".to_string(), formula: None, format: None, comment: None, spill_anchor: None });

        app.selected_col = 0;
        app.sort_column_asc();

        // Column B should follow the sort
        assert_eq!(app.workbook.current_sheet().get_cell(0, 1).value, "A");
        assert_eq!(app.workbook.current_sheet().get_cell(1, 1).value, "B");
        assert_eq!(app.workbook.current_sheet().get_cell(2, 1).value, "C");
    }

    #[test]
    fn test_sort_undo() {
        let mut app = App::default();
        app.set_cell_with_undo(0, 0, CellData { value: "30".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        app.set_cell_with_undo(1, 0, CellData { value: "10".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        app.set_cell_with_undo(2, 0, CellData { value: "20".to_string(), formula: None, format: None, comment: None, spill_anchor: None });

        app.selected_col = 0;
        app.sort_column_asc();

        assert_eq!(app.workbook.current_sheet().get_cell(0, 0).value, "10");

        app.undo();

        assert_eq!(app.workbook.current_sheet().get_cell(0, 0).value, "30");
        assert_eq!(app.workbook.current_sheet().get_cell(1, 0).value, "10");
        assert_eq!(app.workbook.current_sheet().get_cell(2, 0).value, "20");
    }

    #[test]
    fn test_freeze_panes() {
        let mut app = App::default();
        app.selected_row = 2;
        app.selected_col = 1;

        app.start_command_palette();
        app.command_input = "freeze".to_string();
        app.execute_command();

        assert_eq!(app.frozen_rows, 2);
        assert_eq!(app.frozen_cols, 1);
    }

    #[test]
    fn test_unfreeze_panes() {
        let mut app = App::default();
        app.frozen_rows = 3;
        app.frozen_cols = 2;

        app.start_command_palette();
        app.command_input = "unfreeze".to_string();
        app.execute_command();

        assert_eq!(app.frozen_rows, 0);
        assert_eq!(app.frozen_cols, 0);
    }

    #[test]
    fn test_set_format_on_selection() {
        let mut app = App::default();
        app.set_cell_with_undo(0, 0, CellData { value: "100".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        app.set_cell_with_undo(0, 1, CellData { value: "200".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        app.set_cell_with_undo(1, 0, CellData { value: "300".to_string(), formula: None, format: None, comment: None, spill_anchor: None });

        app.selection_start = Some((0, 0));
        app.selection_end = Some((1, 1));

        app.set_selection_format(NumberFormat::Currency { symbol: "$".to_string(), decimals: 2 });

        for row in 0..=1 {
            for col in 0..=1 {
                let cell = app.workbook.current_sheet().get_cell(row, col);
                assert!(cell.format.is_some(), "Cell ({},{}) should have format", row, col);
            }
        }
    }

    #[test]
    fn test_set_format_general_clears_format() {
        let mut app = App::default();
        app.set_cell_with_undo(0, 0, CellData {
            value: "100".to_string(),
            formula: None,
            format: Some(CellFormat { number_format: NumberFormat::Currency { symbol: "$".to_string(), decimals: 2 }, ..CellFormat::default() }),
            comment: None,
        spill_anchor: None,
        });

        app.selected_row = 0;
        app.selected_col = 0;
        app.set_selection_format(NumberFormat::General);

        let cell = app.workbook.current_sheet().get_cell(0, 0);
        assert!(cell.format.is_none()); // General clears format
    }

    #[test]
    fn test_toggle_bold() {
        let mut app = App::default();
        app.set_cell_with_undo(0, 0, CellData { value: "test".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        app.selected_row = 0;
        app.selected_col = 0;

        // Toggle bold on
        app.toggle_bold();
        let cell = app.workbook.current_sheet().get_cell(0, 0);
        assert!(cell.format.as_ref().unwrap().style.bold);

        // Toggle bold off
        app.toggle_bold();
        let cell = app.workbook.current_sheet().get_cell(0, 0);
        assert!(!cell.format.as_ref().unwrap().style.bold);
    }

    #[test]
    fn test_toggle_underline() {
        let mut app = App::default();
        app.set_cell_with_undo(0, 0, CellData { value: "test".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        app.selected_row = 0;
        app.selected_col = 0;

        // Toggle underline on
        app.toggle_underline();
        let cell = app.workbook.current_sheet().get_cell(0, 0);
        assert!(cell.format.as_ref().unwrap().style.underline);

        // Toggle underline off
        app.toggle_underline();
        let cell = app.workbook.current_sheet().get_cell(0, 0);
        assert!(!cell.format.as_ref().unwrap().style.underline);
    }

    #[test]
    fn test_toggle_bold_on_selection() {
        let mut app = App::default();
        app.set_cell_with_undo(0, 0, CellData { value: "A".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        app.set_cell_with_undo(0, 1, CellData { value: "B".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        app.set_cell_with_undo(1, 0, CellData { value: "C".to_string(), formula: None, format: None, comment: None, spill_anchor: None });

        app.selection_start = Some((0, 0));
        app.selection_end = Some((1, 1));
        app.selecting = true;

        app.toggle_bold();

        for row in 0..=1 {
            for col in 0..=1 {
                let cell = app.workbook.current_sheet().get_cell(row, col);
                assert!(cell.format.as_ref().unwrap().style.bold, "Cell ({},{}) should be bold", row, col);
            }
        }
    }

    #[test]
    fn test_set_fg_color() {
        let mut app = App::default();
        app.set_cell_with_undo(0, 0, CellData { value: "test".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        app.selected_row = 0;
        app.selected_col = 0;

        app.set_selection_fg_color(Some(TerminalColor::Red));
        let cell = app.workbook.current_sheet().get_cell(0, 0);
        assert_eq!(cell.format.as_ref().unwrap().style.fg_color, Some(TerminalColor::Red));

        // Clear color
        app.set_selection_fg_color(None);
        let cell = app.workbook.current_sheet().get_cell(0, 0);
        assert_eq!(cell.format.as_ref().unwrap().style.fg_color, None);
    }

    #[test]
    fn test_set_bg_color() {
        let mut app = App::default();
        app.set_cell_with_undo(0, 0, CellData { value: "test".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        app.selected_row = 0;
        app.selected_col = 0;

        app.set_selection_bg_color(Some(TerminalColor::Blue));
        let cell = app.workbook.current_sheet().get_cell(0, 0);
        assert_eq!(cell.format.as_ref().unwrap().style.bg_color, Some(TerminalColor::Blue));
    }

    #[test]
    fn test_set_cell_comment() {
        let mut app = App::default();
        app.set_cell_with_undo(0, 0, CellData { value: "Hello".to_string(), formula: None, format: None, comment: None, spill_anchor: None });

        app.set_cell_comment(Some("This is a comment".to_string()));
        let cell = app.workbook.current_sheet().get_cell(0, 0);
        assert_eq!(cell.comment, Some("This is a comment".to_string()));
        assert_eq!(cell.value, "Hello"); // Value preserved
    }

    #[test]
    fn test_clear_cell_comment() {
        let mut app = App::default();
        app.set_cell_with_undo(0, 0, CellData { value: "Hello".to_string(), formula: None, format: None, comment: Some("old".to_string()), spill_anchor: None });

        app.set_cell_comment(None);
        let cell = app.workbook.current_sheet().get_cell(0, 0);
        assert_eq!(cell.comment, None);
    }

    #[test]
    fn test_apply_filter() {
        let mut app = App::default();
        app.set_cell_with_undo(0, 0, CellData { value: "Apple".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        app.set_cell_with_undo(1, 0, CellData { value: "Banana".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        app.set_cell_with_undo(2, 0, CellData { value: "Apple".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        app.set_cell_with_undo(3, 0, CellData { value: "Cherry".to_string(), formula: None, format: None, comment: None, spill_anchor: None });

        app.apply_filter(0, Some("Apple".to_string()));

        // Rows 1 and 3 should be hidden (Banana and Cherry)
        assert!(!app.hidden_rows.contains(&0));
        assert!(app.hidden_rows.contains(&1));
        assert!(!app.hidden_rows.contains(&2));
        assert!(app.hidden_rows.contains(&3));
    }

    #[test]
    fn test_clear_filter() {
        let mut app = App::default();
        app.set_cell_with_undo(0, 0, CellData { value: "Apple".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        app.set_cell_with_undo(1, 0, CellData { value: "Banana".to_string(), formula: None, format: None, comment: None, spill_anchor: None });

        app.apply_filter(0, Some("Apple".to_string()));
        assert!(!app.hidden_rows.is_empty());

        app.clear_filter();
        assert!(app.hidden_rows.is_empty());
        assert_eq!(app.filter_column, None);
        assert_eq!(app.filter_value, None);
    }

    #[test]
    fn test_comment_undo() {
        let mut app = App::default();
        app.set_cell_with_undo(0, 0, CellData { value: "Hello".to_string(), formula: None, format: None, comment: None, spill_anchor: None });

        app.set_cell_comment(Some("My comment".to_string()));
        assert_eq!(app.workbook.current_sheet().get_cell(0, 0).comment, Some("My comment".to_string()));

        app.undo();
        assert_eq!(app.workbook.current_sheet().get_cell(0, 0).comment, None);
    }

}
