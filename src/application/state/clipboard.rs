//! Submodule of `state` — see state/mod.rs.

use super::*;

impl App {
    pub fn vim_paste_below(&mut self) {
        let Some(cb) = self.clipboard.as_ref() else {
            self.status_message = Some("Nothing to paste".to_string());
            return;
        };
        if cb.is_row_op {
            self.selected_row = (self.selected_row + 1)
                .min(self.workbook.current_sheet().rows.saturating_sub(1));
        } else {
            self.selected_col = (self.selected_col + 1)
                .min(self.workbook.current_sheet().cols.saturating_sub(1));
        }
        self.paste();
    }

    pub fn vim_paste_above(&mut self) {
        let Some(cb) = self.clipboard.as_ref() else {
            self.status_message = Some("Nothing to paste".to_string());
            return;
        };
        if cb.is_row_op {
            self.selected_row = self.selected_row.saturating_sub(1);
        } else {
            self.selected_col = self.selected_col.saturating_sub(1);
        }
        self.paste();
    }

    pub fn copy_selection(&mut self) {
        let range = if let Some(range) = self.get_selection_range() {
            range
        } else {
            ((self.selected_row, self.selected_col), (self.selected_row, self.selected_col))
        };
        let ((start_row, start_col), (end_row, end_col)) = range;
        let mut cells = Vec::new();
        for row in start_row..=end_row {
            for col in start_col..=end_col {
                let cell = self.workbook.current_sheet().get_cell(row, col);
                // Skip spill ghosts: their value is derived from an anchor
                // formula and they have no formula of their own, so pasting
                // them as-is would produce inert duplicates of the anchor's
                // top-left value. The anchor (when included in the range)
                // carries the formula and will re-spill at the destination.
                if cell.spill_anchor.is_some() {
                    continue;
                }
                if !cell.value.is_empty() || cell.formula.is_some() {
                    cells.push((row - start_row, col - start_col, cell));
                }
            }
        }
        let count = (end_row - start_row + 1) * (end_col - start_col + 1);

        // Sentinel-prefixed TSV for system clipboard.
        let mut tsv = String::from(crate::infrastructure::sidecar::SENTINEL);
        for row in start_row..=end_row {
            for col in start_col..=end_col {
                if col > start_col {
                    tsv.push('\t');
                }
                let cell = self.workbook.current_sheet().get_cell(row, col);
                tsv.push_str(&cell.value);
            }
            tsv.push('\n');
        }
        if let Ok(mut board) = arboard::Clipboard::new() {
            let _ = board.set_text(tsv);
        }
        // Sidecar JSON for formula round-trip. Skipped in tests for the same
        // state-isolation reason as the read path.
        if !cfg!(test) {
            crate::infrastructure::sidecar::write(cells.clone(), start_row, start_col);
        }

        self.clipboard = Some(ClipboardData {
            cells,
            source_row: start_row,
            source_col: start_col,
            is_row_op: false,
        });
        self.status_message = Some(format!("Copied {} cell(s)", count));
    }

    pub fn cut_selection(&mut self) {
        let range = if let Some(range) = self.get_selection_range() {
            range
        } else {
            ((self.selected_row, self.selected_col), (self.selected_row, self.selected_col))
        };
        let ((start_row, start_col), (end_row, end_col)) = range;
        let mut cells = Vec::new();
        for row in start_row..=end_row {
            for col in start_col..=end_col {
                let cell = self.workbook.current_sheet().get_cell(row, col);
                if cell.spill_anchor.is_some() {
                    continue;
                }
                if !cell.value.is_empty() || cell.formula.is_some() {
                    cells.push((row - start_row, col - start_col, cell));
                }
            }
        }
        let count = cells.len();

        // Symmetric with copy_selection: also push the cut region to the
        // system clipboard so an external paste after `dd`/`x` gets the
        // cut contents instead of stale data from the previous copy.
        let mut tsv = String::from(crate::infrastructure::sidecar::SENTINEL);
        for row in start_row..=end_row {
            for col in start_col..=end_col {
                if col > start_col {
                    tsv.push('\t');
                }
                let cell = self.workbook.current_sheet().get_cell(row, col);
                tsv.push_str(&cell.value);
            }
            tsv.push('\n');
        }
        if let Ok(mut board) = arboard::Clipboard::new() {
            let _ = board.set_text(tsv);
        }
        if !cfg!(test) {
            crate::infrastructure::sidecar::write(cells.clone(), start_row, start_col);
        }

        self.clipboard = Some(ClipboardData {
            cells,
            source_row: start_row,
            source_col: start_col,
            is_row_op: false,
        });
        // Clear the cut cells and notify cross-sheet listeners so any
        // formula on another sheet that referenced these now goes stale.
        // Route the clears through `clear_cells_on_active` so the dirty
        // set is populated (single workbook call, one cross-sheet pass).
        let mut batch = Vec::new();
        let mut positions: Vec<(usize, usize)> = Vec::new();
        for row in start_row..=end_row {
            for col in start_col..=end_col {
                let old = if self.workbook.current_sheet().cells.contains_key(&(row, col)) {
                    Some(self.workbook.current_sheet().get_cell(row, col))
                } else {
                    None
                };
                if old.is_some() {
                    batch.push(UndoAction::CellModified { row, col, old_cell: old, new_cell: None });
                    positions.push((row, col));
                }
            }
        }
        if !positions.is_empty() {
            self.workbook.clear_cells_on_active(positions);
        }
        if !batch.is_empty() {
            self.record_action(UndoAction::Batch(batch));
        }
        self.status_message = Some(format!("Cut {} cell(s)", count));
    }

    pub fn paste(&mut self) {
        // If the system clipboard has our sentinel header, load the matching
        // sidecar JSON (preserves formulas/formats/comments). Otherwise fall
        // back to internal clipboard, then plain-text TSV.
        // Skipped in tests because $HOME and ~/.cache/tshts persist between
        // test runs and would otherwise leak state between cases.
        if !cfg!(test)
            && let Ok(mut board) = arboard::Clipboard::new()
                && let Ok(text) = board.get_text()
                    && crate::infrastructure::sidecar::strip_sentinel(&text).is_some()
                        && let Some(payload) = crate::infrastructure::sidecar::read() {
                            let cb = ClipboardData {
                                cells: payload.cells,
                                source_row: payload.source_row,
                                source_col: payload.source_col,
                                is_row_op: false,
                            };
                            self.clipboard = Some(cb);
                        }
        let clipboard = if let Some(ref cb) = self.clipboard {
            cb.clone()
        } else {
            // Tests don't touch the system clipboard (cross-test contamination).
            if !cfg!(test)
                && let Ok(mut board) = arboard::Clipboard::new()
                    && let Ok(text) = board.get_text()
                        && !text.is_empty() {
                            let body = crate::infrastructure::sidecar::strip_sentinel(&text)
                                .unwrap_or(&text);
                            self.paste_tsv(body);
                            return;
                        }
            self.status_message = Some("Nothing to paste".to_string());
            return;
        };

        let dest_row = self.selected_row;
        let dest_col = self.selected_col;

        // Compute all new cells first (evaluator borrows spreadsheet immutably)
        let new_cells: Vec<_> = {
            let evaluator = crate::domain::FormulaEvaluator::new(self.workbook.current_sheet());
            clipboard.cells.iter().filter_map(|(row_off, col_off, cell)| {
                let target_row = dest_row + row_off;
                let target_col = dest_col + col_off;
                if target_row >= self.workbook.current_sheet().rows || target_col >= self.workbook.current_sheet().cols {
                    return None;
                }
                let new_cell = if let Some(ref formula) = cell.formula {
                    let row_offset = target_row as i32 - (clipboard.source_row + row_off) as i32;
                    let col_offset = target_col as i32 - (clipboard.source_col + col_off) as i32;
                    let adjusted = evaluator.adjust_formula_references(formula, row_offset, col_offset);
                    let value = evaluator.evaluate_formula(&adjusted);
                    CellData { value, formula: Some(adjusted), format: cell.format.clone(), comment: cell.comment.clone(), spill_anchor: None }
                } else {
                    cell.clone()
                };
                Some((target_row, target_col, new_cell))
            }).collect()
        };

        // Now apply changes (mutably borrows spreadsheet) — collect the
        // undo batch, then push all writes through `set_many` in one shot
        // so dependents recalc just once.
        let mut batch = Vec::new();
        let mut writes: Vec<(usize, usize, CellData)> = Vec::with_capacity(new_cells.len());
        for (target_row, target_col, new_cell) in &new_cells {
            let old = if self
                .workbook
                .current_sheet()
                .cells
                .contains_key(&(*target_row, *target_col))
            {
                Some(self.workbook.current_sheet().get_cell(*target_row, *target_col))
            } else {
                None
            };
            batch.push(UndoAction::CellModified {
                row: *target_row,
                col: *target_col,
                old_cell: old,
                new_cell: Some(new_cell.clone()),
            });
            writes.push((*target_row, *target_col, new_cell.clone()));
        }
        // Single workbook API call handles same-sheet recalc + cross-sheet
        // propagation for the whole paste batch.
        self.workbook.write_cells_on_active(writes);
        if !batch.is_empty() {
            self.record_action(UndoAction::Batch(batch));
        }
        self.status_message = Some(format!("Pasted {} cell(s)", clipboard.cells.len()));
    }

    fn paste_tsv(&mut self, text: &str) {
        let dest_row = self.selected_row;
        let dest_col = self.selected_col;
        let mut batch = Vec::new();
        let mut writes: Vec<(usize, usize, CellData)> = Vec::new();

        for (row_offset, line) in text.lines().enumerate() {
            if line.is_empty() { continue; }
            for (col_offset, value) in line.split('\t').enumerate() {
                let target_row = dest_row + row_offset;
                let target_col = dest_col + col_offset;
                if target_row >= self.workbook.current_sheet().rows || target_col >= self.workbook.current_sheet().cols {
                    continue;
                }
                let old = if self.workbook.current_sheet().cells.contains_key(&(target_row, target_col)) {
                    Some(self.workbook.current_sheet().get_cell(target_row, target_col))
                } else {
                    None
                };
                // If the pasted cell starts with `=`, treat it as a formula
                // and evaluate it at the destination. References are NOT
                // adjusted: this path is only reached when the source was
                // not tshts (no sentinel/sidecar), so we have no original
                // anchor to shift against. Excel and Sheets behave the same
                // way for plain-text formula paste. The tshts→tshts path in
                // `paste()` handles relative adjustment properly.
                let new_cell = if value.starts_with('=') {
                    // Cycle check: pasted text bypasses finish_editing_in_direction,
                    // so a `=A1` pasted into A1 would otherwise be accepted.
                    if !self.iterative_calc {
                        let names = self.workbook.named_ranges.clone();
                        let evaluator = FormulaEvaluator::for_workbook(
                            &self.workbook,
                            self.workbook.current_sheet(),
                            &names,
                        );
                        let same = evaluator.would_create_circular_reference(
                            value,
                            (target_row, target_col),
                        );
                        let precedents = evaluator.extract_qualified_refs(value);
                        let sheet_name = self
                            .workbook
                            .sheet_names[self.workbook.active_sheet]
                            .clone();
                        let cross = self.workbook.would_create_cross_sheet_cycle(
                            &sheet_name,
                            target_row,
                            target_col,
                            &precedents,
                        );
                        if same || cross {
                            // Skip this one cell; keep going. Surfacing via
                            // status_message at end of paste.
                            continue;
                        }
                    }
                    let evaluator = FormulaEvaluator::new(self.workbook.current_sheet());
                    let evaluated = evaluator.evaluate_formula(value);
                    CellData {
                        value: evaluated,
                        formula: Some(value.to_string()),
                        format: None,
                        comment: None,
                        spill_anchor: None,
                    }
                } else {
                    CellData {
                        value: value.to_string(),
                        formula: None,
                        format: None,
                        comment: None,
                        spill_anchor: None,
                    }
                };
                batch.push(UndoAction::CellModified {
                    row: target_row,
                    col: target_col,
                    old_cell: old,
                    new_cell: Some(new_cell.clone()),
                });
                writes.push((target_row, target_col, new_cell));
            }
        }
        let count = writes.len();
        // Single workbook API call handles same-sheet recalc + cross-sheet
        // propagation for the whole paste batch.
        self.workbook.write_cells_on_active(writes);
        if !batch.is_empty() {
            self.record_action(UndoAction::Batch(batch));
        }
        self.status_message = Some(format!("Pasted {} cell(s) from system clipboard", count));
    }

    pub fn insert_row(&mut self) {
        let insert_at = self.selected_row;
        let sheet_idx = self.workbook.active_sheet;
        // Routes through Workbook so cross-sheet refs to this sheet adjust
        // too (e.g. `Sheet2!A5` shifts to `Sheet2!A6` on insertion above A5).
        self.workbook.insert_row_on_active(insert_at);
        self.record_action(UndoAction::RowInserted { sheet_idx, at: insert_at });
        self.status_message = Some(format!("Inserted row at {}", insert_at + 1));
    }

    pub fn delete_row(&mut self) {
        let delete_at = self.selected_row;
        let sheet_idx = self.workbook.active_sheet;
        // Snapshot pre-delete: structural cross-sheet shifts can't be
        // reconstructed from the deleted row alone.
        let pre = Box::new(self.workbook.clone());
        self.workbook.delete_row_on_active(delete_at);
        if self.selected_row >= self.workbook.current_sheet().rows {
            self.selected_row = self.workbook.current_sheet().rows.saturating_sub(1);
        }
        self.record_action(UndoAction::RowDeleted { sheet_idx, at: delete_at, pre });
        self.status_message = Some(format!("Deleted row {}", delete_at + 1));
    }

    pub fn insert_col(&mut self) {
        let insert_at = self.selected_col;
        let sheet_idx = self.workbook.active_sheet;
        self.workbook.insert_col_on_active(insert_at);
        self.record_action(UndoAction::ColInserted { sheet_idx, at: insert_at });
        self.status_message = Some(format!("Inserted column at {}", crate::domain::Spreadsheet::column_label(insert_at)));
    }

    pub fn delete_col(&mut self) {
        let delete_at = self.selected_col;
        let sheet_idx = self.workbook.active_sheet;
        let pre = Box::new(self.workbook.clone());
        self.workbook.delete_col_on_active(delete_at);
        if self.selected_col >= self.workbook.current_sheet().cols {
            self.selected_col = self.workbook.current_sheet().cols.saturating_sub(1);
        }
        self.record_action(UndoAction::ColDeleted { sheet_idx, at: delete_at, pre });
        self.status_message = Some(format!("Deleted column {}", crate::domain::Spreadsheet::column_label(delete_at)));
    }

}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::domain::CellData;

    #[test]
    fn test_copy_paste_single_cell() {
        let mut app = App::default();
        app.set_cell_with_undo(0, 0, CellData { value: "Hello".to_string(), formula: None, format: None, comment: None, spill_anchor: None });

        // Copy A1
        app.selected_row = 0;
        app.selected_col = 0;
        app.copy_selection();
        assert!(app.clipboard.is_some());

        // Paste to B2
        app.selected_row = 1;
        app.selected_col = 1;
        app.paste();

        assert_eq!(app.workbook.current_sheet().get_cell(1, 1).value, "Hello");
    }

    #[test]
    fn test_copy_paste_range() {
        let mut app = App::default();
        app.set_cell_with_undo(0, 0, CellData { value: "A".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        app.set_cell_with_undo(0, 1, CellData { value: "B".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        app.set_cell_with_undo(1, 0, CellData { value: "C".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        app.set_cell_with_undo(1, 1, CellData { value: "D".to_string(), formula: None, format: None, comment: None, spill_anchor: None });

        // Select A1:B2
        app.selection_start = Some((0, 0));
        app.selection_end = Some((1, 1));
        app.copy_selection();

        // Paste to C3
        app.selected_row = 2;
        app.selected_col = 2;
        app.paste();

        assert_eq!(app.workbook.current_sheet().get_cell(2, 2).value, "A");
        assert_eq!(app.workbook.current_sheet().get_cell(2, 3).value, "B");
        assert_eq!(app.workbook.current_sheet().get_cell(3, 2).value, "C");
        assert_eq!(app.workbook.current_sheet().get_cell(3, 3).value, "D");
    }

    #[test]
    fn test_cut_paste() {
        let mut app = App::default();
        app.set_cell_with_undo(0, 0, CellData { value: "Move me".to_string(), formula: None, format: None, comment: None, spill_anchor: None });

        app.selected_row = 0;
        app.selected_col = 0;
        app.cut_selection();

        // Original cell should be cleared
        assert!(app.workbook.current_sheet().get_cell(0, 0).value.is_empty());

        // Paste to new location
        app.selected_row = 2;
        app.selected_col = 2;
        app.paste();

        assert_eq!(app.workbook.current_sheet().get_cell(2, 2).value, "Move me");
    }

    #[test]
    fn test_paste_nothing() {
        let mut app = App::default();
        app.paste(); // Should not crash
        assert!(app.status_message.as_ref().unwrap().contains("Nothing to paste"));
    }

    #[test]
    fn test_copy_paste_formula_adjusts_refs() {
        let mut app = App::default();
        app.set_cell_with_undo(0, 0, CellData { value: "10".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        app.set_cell_with_undo(0, 1, CellData {
            value: "20".to_string(),
            formula: Some("=A1*2".to_string()),
            format: None,
            comment: None,
        spill_anchor: None,
        });

        // Copy B1 (has formula =A1*2)
        app.selected_row = 0;
        app.selected_col = 1;
        app.copy_selection();

        // Paste to B2 (should adjust to =A2*2)
        app.selected_row = 1;
        app.selected_col = 1;
        app.paste();

        let pasted = app.workbook.current_sheet().get_cell(1, 1);
        assert!(pasted.formula.is_some());
        assert_eq!(pasted.formula.unwrap(), "=A2*2");
    }

}
