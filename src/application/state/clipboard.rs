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
                if !cell.value.is_empty() || cell.formula.is_some() {
                    cells.push((row - start_row, col - start_col, cell));
                }
            }
        }
        let count = (end_row - start_row + 1) * (end_col - start_col + 1);
        self.clipboard = Some(ClipboardData {
            cells,
            source_row: start_row,
            source_col: start_col,
            is_row_op: false,
        });
        // Clear the cut cells
        let mut batch = Vec::new();
        for row in start_row..=end_row {
            for col in start_col..=end_col {
                let old = if self.workbook.current_sheet().cells.contains_key(&(row, col)) {
                    Some(self.workbook.current_sheet().get_cell(row, col))
                } else {
                    None
                };
                if old.is_some() {
                    batch.push(UndoAction::CellModified { row, col, old_cell: old, new_cell: None });
                    self.workbook.current_sheet_mut().clear_cell(row, col);
                }
            }
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
        if !cfg!(test) {
            if let Ok(mut board) = arboard::Clipboard::new() {
                if let Ok(text) = board.get_text() {
                    if crate::infrastructure::sidecar::strip_sentinel(&text).is_some() {
                        if let Some(payload) = crate::infrastructure::sidecar::read() {
                            let cb = ClipboardData {
                                cells: payload.cells,
                                source_row: payload.source_row,
                                source_col: payload.source_col,
                                is_row_op: false,
                            };
                            self.clipboard = Some(cb);
                        }
                    }
                }
            }
        }
        let clipboard = if let Some(ref cb) = self.clipboard {
            cb.clone()
        } else {
            // Tests don't touch the system clipboard (cross-test contamination).
            if !cfg!(test) {
                if let Ok(mut board) = arboard::Clipboard::new() {
                    if let Ok(text) = board.get_text() {
                        if !text.is_empty() {
                            let body = crate::infrastructure::sidecar::strip_sentinel(&text)
                                .unwrap_or(&text);
                            self.paste_tsv(body);
                            return;
                        }
                    }
                }
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
        if !writes.is_empty() {
            self.workbook.current_sheet_mut().set_many(writes);
        }
        if !batch.is_empty() {
            self.record_action(UndoAction::Batch(batch));
        }
        self.status_message = Some(format!("Pasted {} cell(s)", clipboard.cells.len()));
    }

    fn paste_tsv(&mut self, text: &str) {
        let dest_row = self.selected_row;
        let dest_col = self.selected_col;
        let mut batch = Vec::new();
        let mut count = 0;

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
                // and evaluate it relative to the destination cell.
                let new_cell = if value.starts_with('=') {
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
                self.workbook.current_sheet_mut().set_cell(target_row, target_col, new_cell);
                count += 1;
            }
        }
        if !batch.is_empty() {
            self.record_action(UndoAction::Batch(batch));
        }
        self.status_message = Some(format!("Pasted {} cell(s) from system clipboard", count));
    }

    pub fn insert_row(&mut self) {
        let insert_at = self.selected_row;
        self.workbook.current_sheet_mut().insert_row(insert_at);
        self.dirty = true;
        self.status_message = Some(format!("Inserted row at {}", insert_at + 1));
    }

    pub fn delete_row(&mut self) {
        let delete_at = self.selected_row;
        self.workbook.current_sheet_mut().delete_row(delete_at);
        if self.selected_row >= self.workbook.current_sheet().rows {
            self.selected_row = self.workbook.current_sheet().rows.saturating_sub(1);
        }
        self.dirty = true;
        self.status_message = Some(format!("Deleted row {}", delete_at + 1));
    }

    pub fn insert_col(&mut self) {
        let insert_at = self.selected_col;
        self.workbook.current_sheet_mut().insert_col(insert_at);
        self.dirty = true;
        self.status_message = Some(format!("Inserted column at {}", crate::domain::Spreadsheet::column_label(insert_at)));
    }

    pub fn delete_col(&mut self) {
        let delete_at = self.selected_col;
        self.workbook.current_sheet_mut().delete_col(delete_at);
        if self.selected_col >= self.workbook.current_sheet().cols {
            self.selected_col = self.workbook.current_sheet().cols.saturating_sub(1);
        }
        self.dirty = true;
        self.status_message = Some(format!("Deleted column {}", crate::domain::Spreadsheet::column_label(delete_at)));
    }

}
