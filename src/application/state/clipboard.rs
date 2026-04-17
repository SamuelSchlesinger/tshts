//! Copy / cut / paste and row/column insert/delete operations.

use crate::domain::{CellData, Spreadsheet};
use super::{App, ClipboardData, UndoAction};

impl App {
    /// Copies the current selection to both the internal and system clipboards.
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

        let mut tsv = String::new();
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

        self.clipboard = Some(ClipboardData {
            cells,
            source_row: start_row,
            source_col: start_col,
        });
        self.status_message = Some(format!("Copied {} cell(s)", count));
    }

    /// Cuts the current selection to the internal clipboard.
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
        });
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

    /// Pastes clipboard contents at the current cursor position.
    /// Falls back to system clipboard if internal clipboard is empty.
    pub fn paste(&mut self) {
        let clipboard = if let Some(ref cb) = self.clipboard {
            cb.clone()
        } else {
            if let Ok(mut board) = arboard::Clipboard::new() {
                if let Ok(text) = board.get_text() {
                    if !text.is_empty() {
                        self.paste_tsv(&text);
                        return;
                    }
                }
            }
            self.status_message = Some("Nothing to paste".to_string());
            return;
        };

        let dest_row = self.selected_row;
        let dest_col = self.selected_col;

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
                    CellData { value, formula: Some(adjusted), format: cell.format.clone(), comment: cell.comment.clone() }
                } else {
                    cell.clone()
                };
                Some((target_row, target_col, new_cell))
            }).collect()
        };

        let mut batch = Vec::new();
        for (target_row, target_col, new_cell) in &new_cells {
            let old = if self.workbook.current_sheet().cells.contains_key(&(*target_row, *target_col)) {
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
            self.workbook.current_sheet_mut().set_cell(*target_row, *target_col, new_cell.clone());
        }

        if !batch.is_empty() {
            self.record_action(UndoAction::Batch(batch));
        }
        self.status_message = Some(format!("Pasted {} cell(s)", clipboard.cells.len()));
    }

    /// Pastes TSV text from the system clipboard at the current cursor position.
    pub(crate) fn paste_tsv(&mut self, text: &str) {
        let dest_row = self.selected_row;
        let dest_col = self.selected_col;
        let mut batch = Vec::new();
        let mut count = 0;

        for (row_offset, line) in text.lines().enumerate() {
            for (col_offset, value) in line.split('\t').enumerate() {
                if value.is_empty() {
                    continue;
                }
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
                let new_cell = CellData { value: value.to_string(), formula: None, format: None, comment: None };
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

    /// Inserts a row above the current cursor position.
    pub fn insert_row(&mut self) {
        let insert_at = self.selected_row;
        self.workbook.current_sheet_mut().insert_row(insert_at);
        self.status_message = Some(format!("Inserted row at {}", insert_at + 1));
    }

    /// Deletes the current row.
    pub fn delete_row(&mut self) {
        let delete_at = self.selected_row;
        self.workbook.current_sheet_mut().delete_row(delete_at);
        if self.selected_row >= self.workbook.current_sheet().rows {
            self.selected_row = self.workbook.current_sheet().rows.saturating_sub(1);
        }
        self.status_message = Some(format!("Deleted row {}", delete_at + 1));
    }

    /// Inserts a column to the left of the current cursor position.
    pub fn insert_col(&mut self) {
        let insert_at = self.selected_col;
        self.workbook.current_sheet_mut().insert_col(insert_at);
        self.status_message = Some(format!("Inserted column at {}", Spreadsheet::column_label(insert_at)));
    }

    /// Deletes the current column.
    pub fn delete_col(&mut self) {
        let delete_at = self.selected_col;
        self.workbook.current_sheet_mut().delete_col(delete_at);
        if self.selected_col >= self.workbook.current_sheet().cols {
            self.selected_col = self.workbook.current_sheet().cols.saturating_sub(1);
        }
        self.status_message = Some(format!("Deleted column {}", Spreadsheet::column_label(delete_at)));
    }
}
