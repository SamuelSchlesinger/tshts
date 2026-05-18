//! Submodule of `state` — see state/mod.rs.

use super::*;

impl App {
    pub fn vim_reset_pending(&mut self) {
        self.vim_count = None;
        self.vim_pending_op = None;
        self.vim_awaiting_g = false;
    }

    pub fn vim_enter_visual(&mut self, kind: VisualKind) {
        self.start_selection();
        match kind {
            VisualKind::Row => {
                let last_col = self.workbook.current_sheet().cols.saturating_sub(1);
                self.selection_start = Some((self.selected_row, 0));
                self.selection_end = Some((self.selected_row, last_col));
            }
            VisualKind::Cell | VisualKind::Block => {
                self.selection_start = Some((self.selected_row, self.selected_col));
                self.selection_end = Some((self.selected_row, self.selected_col));
            }
        }
        self.mode = AppMode::Visual { kind };
    }

    pub fn vim_exit_visual(&mut self) {
        self.clear_selection();
        self.mode = AppMode::Normal;
        self.vim_reset_pending();
    }

    pub fn vim_apply_operator(
        &mut self,
        op: VimOperator,
        r0: usize,
        c0: usize,
        r1: usize,
        c1: usize,
    ) {
        let (r0, r1) = (r0.min(r1), r0.max(r1));
        let (c0, c1) = (c0.min(c1), c0.max(c1));
        // Yank/Delete both copy first; we temporarily set the selection so
        // copy_selection picks up the right cells without us re-implementing it.
        let prev_start = self.selection_start;
        let prev_end = self.selection_end;
        self.selection_start = Some((r0, c0));
        self.selection_end = Some((r1, c1));
        self.copy_selection();
        self.selection_start = prev_start;
        self.selection_end = prev_end;

        match op {
            VimOperator::Yank => {
                self.status_message =
                    Some(format!("Yanked {} cell(s)", (r1 - r0 + 1) * (c1 - c0 + 1)));
            }
            VimOperator::Delete | VimOperator::Change => {
                let mut batch = Vec::new();
                for row in r0..=r1 {
                    for col in c0..=c1 {
                        let old = self
                            .workbook
                            .current_sheet()
                            .cells
                            .get(&(row, col))
                            .cloned();
                        if old.is_some() {
                            self.workbook.current_sheet_mut().clear_cell(row, col);
                            batch.push(UndoAction::CellModified {
                                row,
                                col,
                                old_cell: old,
                                new_cell: None,
                            });
                        }
                    }
                }
                if !batch.is_empty() {
                    self.record_action(UndoAction::Batch(batch));
                    self.dirty = true;
                }
                self.status_message =
                    Some(format!("Deleted {} cell(s)", (r1 - r0 + 1) * (c1 - c0 + 1)));
                if matches!(op, VimOperator::Change) {
                    self.selected_row = r0;
                    self.selected_col = c0;
                    self.start_editing();
                }
            }
        }
        self.vim_reset_pending();
    }

    pub fn vim_apply_line_op(&mut self, op: VimOperator, count: usize) {
        let count = count.max(1);
        let last_col = self.workbook.current_sheet().cols.saturating_sub(1);
        let r0 = self.selected_row;
        let r1 = (r0 + count - 1).min(self.workbook.current_sheet().rows.saturating_sub(1));
        self.vim_apply_operator(op, r0, 0, r1, last_col);
        if let Some(cb) = self.clipboard.as_mut() {
            cb.is_row_op = true;
        }
    }

    pub fn vim_motion_row_start(&mut self) {
        self.selected_col = 0;
        self.ensure_cursor_visible();
    }

    pub fn vim_motion_row_end(&mut self) {
        let row = self.selected_row;
        let last_data = self
            .workbook
            .current_sheet()
            .cells
            .iter()
            .filter(|((r, _), c)| *r == row && !c.value.is_empty())
            .map(|((_, c), _)| *c)
            .max();
        self.selected_col = last_data
            .unwrap_or_else(|| self.workbook.current_sheet().cols.saturating_sub(1));
        self.ensure_cursor_visible();
    }

    pub fn vim_motion_row_first_data(&mut self) {
        let row = self.selected_row;
        let first_data = self
            .workbook
            .current_sheet()
            .cells
            .iter()
            .filter(|((r, _), c)| *r == row && !c.value.is_empty())
            .map(|((_, c), _)| *c)
            .min();
        self.selected_col = first_data.unwrap_or(0);
        self.ensure_cursor_visible();
    }

    pub fn vim_motion_top(&mut self) {
        self.selected_row = 0;
        self.ensure_cursor_visible();
    }

    pub fn vim_motion_bottom(&mut self) {
        let max_data_row = self
            .workbook
            .current_sheet()
            .cells
            .iter()
            .filter(|(_, c)| !c.value.is_empty())
            .map(|((r, _), _)| *r)
            .max();
        self.selected_row = max_data_row
            .unwrap_or_else(|| self.workbook.current_sheet().rows.saturating_sub(1));
        self.ensure_cursor_visible();
    }

    pub fn vim_motion_goto_row(&mut self, n: usize) {
        if n == 0 {
            return;
        }
        let target = (n - 1).min(self.workbook.current_sheet().rows.saturating_sub(1));
        self.selected_row = target;
        self.ensure_cursor_visible();
    }

}
