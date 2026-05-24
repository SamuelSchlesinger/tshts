//! Undo/redo action enum plus the small set of revert/apply helpers that
//! mutate the live workbook for each variant.
//!
//! Each `UndoAction` variant co-locates its data with its `apply` / `revert`
//! logic via the impl below — adding a new mutation type means adding a
//! variant AND its two methods in one place, rather than touching three
//! sites (variant + apply_undo arm + apply_redo arm in App).

use crate::domain::{CellData, Workbook};

/// Represents an action that can be undone/redone.
//
// `large_enum_variant` allowed: `CellModified` holds two `Option<CellData>`
// (~288 bytes); other variants are ~56 bytes. Boxing `CellData` would touch
// ~25 construction/match sites for negligible gain — undo stacks are
// bounded (~hundreds of entries), so total overhead stays under 300 KB.
#[allow(clippy::large_enum_variant)]
#[derive(Debug, Clone)]
pub enum UndoAction {
    /// Cell was modified (row, col, old_value, new_value)
    CellModified {
        row: usize,
        col: usize,
        old_cell: Option<CellData>,
        new_cell: Option<CellData>,
    },
    /// Multiple actions that should be undone/redone atomically
    Batch(Vec<UndoAction>),
    /// Sheet-level conditional-format list replaced (sheet_idx, old, new).
    /// Lets `:cf <col> ...` and `:cf clear` participate in undo/redo.
    ConditionalFormatsReplaced {
        sheet_idx: usize,
        old: Vec<crate::domain::ConditionalFormat>,
        new: Vec<crate::domain::ConditionalFormat>,
    },
    /// Row was inserted at `at` on `sheet_idx`. Undo = delete the row.
    /// Used by vim `o`/`O` so an unwanted opened row can be rolled back. We
    /// don't carry cell data because the row is always inserted empty —
    /// any content typed in after gets its own CellModified entry.
    RowInserted {
        sheet_idx: usize,
        at: usize,
    },
    /// Row was deleted at `at` on `sheet_idx`. Undo restores `pre`;
    /// redo re-runs the delete. We snapshot the entire workbook because
    /// cross-sheet refs to the deleted row become `#REF!` and per-sheet
    /// formula shifts can't be undone purely from the deleted row's contents.
    RowDeleted {
        sheet_idx: usize,
        at: usize,
        pre: Box<Workbook>,
    },
    /// Column was inserted at `at` on `sheet_idx`. Undo = delete the column.
    ColInserted {
        sheet_idx: usize,
        at: usize,
    },
    /// Column was deleted. Same reasoning as RowDeleted.
    ColDeleted {
        sheet_idx: usize,
        at: usize,
        pre: Box<Workbook>,
    },
    /// Coarse workbook-level snapshot. Used as an escape hatch for
    /// structural operations (sheet add/delete/rename, freeze, filter,
    /// table create, iterative-calc toggles, named-range edits) where
    /// fine-grained reversal would require its own variant and the
    /// command is rare enough that round-tripping a whole workbook is
    /// acceptable. `pre`/`post` are the workbook state before/after the
    /// command; revert and apply just swap them in.
    WorkbookSnapshot {
        description: String,
        pre: Box<Workbook>,
        post: Box<Workbook>,
    },
}

impl UndoAction {
    /// Human-readable label for use in undo/redo status messages.
    /// `WorkbookSnapshot` carries its own per-command description (passed
    /// to `with_snapshot_undo`); everything else gets a static name.
    pub fn description(&self) -> String {
        match self {
            UndoAction::CellModified { .. } => "cell edit".to_string(),
            UndoAction::Batch(_) => "batch edit".to_string(),
            UndoAction::ConditionalFormatsReplaced { .. } => "CF rules".to_string(),
            UndoAction::RowInserted { .. } => "insert row".to_string(),
            UndoAction::RowDeleted { .. } => "delete row".to_string(),
            UndoAction::ColInserted { .. } => "insert column".to_string(),
            UndoAction::ColDeleted { .. } => "delete column".to_string(),
            UndoAction::WorkbookSnapshot { description, .. } => description.clone(),
        }
    }

    /// Roll the workbook back to the state before this action was applied.
    /// Used by `App::undo`. Returns any non-fatal recalc error so the
    /// caller can surface it (e.g. iterative-calc non-convergence on a
    /// cyclic workbook) instead of corrupting the TUI via eprintln.
    pub fn revert(
        &self,
        workbook: &mut Workbook,
    ) -> Result<(), crate::domain::services::CalcError> {
        match self {
            UndoAction::CellModified { row, col, old_cell, new_cell: _ } => {
                restore_cell(workbook, *row, *col, old_cell.as_ref());
                Ok(())
            }
            UndoAction::Batch(actions) => {
                // O(N) batch revert: write every cell directly (no
                // per-cell cascade), mark dirty, then run a single
                // graph-driven recalc. The naive loop-and-cascade
                // approach would be O(N²) for a chain — undo of a
                // 10k-cell replace_all would take ~15s rather than
                // ~150ms.
                let batch_res = batch_apply_cells_then_recalc(
                    workbook,
                    actions.iter().rev(),
                    |a| {
                        if let UndoAction::CellModified { row, col, old_cell, .. } = a {
                            Some((*row, *col, old_cell.clone()))
                        } else {
                            None
                        }
                    },
                );
                // Non-CellModified actions inside the batch (rare today
                // — most batches are pure CellModified) get the usual
                // per-action revert. This is still O(actions) in the
                // count of non-CellModified entries, which is bounded.
                // Surface the first error encountered.
                let mut first_err = batch_res.err();
                for a in actions.iter().rev() {
                    if !matches!(a, UndoAction::CellModified { .. })
                        && let Err(e) = a.revert(workbook) {
                            first_err.get_or_insert(e);
                        }
                }
                match first_err {
                    Some(e) => Err(e),
                    None => Ok(()),
                }
            }
            UndoAction::ConditionalFormatsReplaced { sheet_idx, old, new: _ } => {
                restore_cf(workbook, *sheet_idx, old);
                Ok(())
            }
            UndoAction::RowInserted { sheet_idx, at } => {
                if *sheet_idx < workbook.sheets.len() {
                    with_active_sheet(workbook, *sheet_idx, |wb| wb.delete_row_on_active(*at));
                }
                Ok(())
            }
            UndoAction::RowDeleted { pre, .. } => {
                restore_workbook(workbook, pre);
                Ok(())
            }
            UndoAction::ColInserted { sheet_idx, at } => {
                if *sheet_idx < workbook.sheets.len() {
                    with_active_sheet(workbook, *sheet_idx, |wb| wb.delete_col_on_active(*at));
                }
                Ok(())
            }
            UndoAction::ColDeleted { pre, .. } => {
                restore_workbook(workbook, pre);
                Ok(())
            }
            UndoAction::WorkbookSnapshot { pre, .. } => {
                restore_workbook(workbook, pre);
                Ok(())
            }
        }
    }

    /// Re-apply this action to the workbook. Used by `App::redo`. Returns
    /// any non-fatal recalc error so the caller can surface it.
    pub fn apply(
        &self,
        workbook: &mut Workbook,
    ) -> Result<(), crate::domain::services::CalcError> {
        match self {
            UndoAction::CellModified { row, col, old_cell: _, new_cell } => {
                restore_cell(workbook, *row, *col, new_cell.as_ref());
                Ok(())
            }
            UndoAction::Batch(actions) => {
                // O(N) batch apply — same shape as revert above. See
                // `batch_apply_cells_then_recalc` for the rationale.
                let batch_res = batch_apply_cells_then_recalc(workbook, actions.iter(), |a| {
                    if let UndoAction::CellModified { row, col, new_cell, .. } = a {
                        Some((*row, *col, new_cell.clone()))
                    } else {
                        None
                    }
                });
                let mut first_err = batch_res.err();
                for a in actions {
                    if !matches!(a, UndoAction::CellModified { .. })
                        && let Err(e) = a.apply(workbook) {
                            first_err.get_or_insert(e);
                        }
                }
                match first_err {
                    Some(e) => Err(e),
                    None => Ok(()),
                }
            }
            UndoAction::ConditionalFormatsReplaced { sheet_idx, old: _, new } => {
                restore_cf(workbook, *sheet_idx, new);
                Ok(())
            }
            UndoAction::RowInserted { sheet_idx, at } => {
                if *sheet_idx < workbook.sheets.len() {
                    with_active_sheet(workbook, *sheet_idx, |wb| wb.insert_row_on_active(*at));
                }
                Ok(())
            }
            UndoAction::RowDeleted { sheet_idx, at, .. } => {
                if *sheet_idx < workbook.sheets.len() {
                    with_active_sheet(workbook, *sheet_idx, |wb| wb.delete_row_on_active(*at));
                }
                Ok(())
            }
            UndoAction::ColInserted { sheet_idx, at } => {
                if *sheet_idx < workbook.sheets.len() {
                    with_active_sheet(workbook, *sheet_idx, |wb| wb.insert_col_on_active(*at));
                }
                Ok(())
            }
            UndoAction::ColDeleted { sheet_idx, at, .. } => {
                if *sheet_idx < workbook.sheets.len() {
                    with_active_sheet(workbook, *sheet_idx, |wb| wb.delete_col_on_active(*at));
                }
                Ok(())
            }
            UndoAction::WorkbookSnapshot { post, .. } => {
                restore_workbook(workbook, post);
                Ok(())
            }
        }
    }
}

/// Replace `workbook` with `pre`'s contents (deep clone). Used by
/// RowDeleted/ColDeleted undo where structural snapshots are the most
/// reliable way to roll back.
fn restore_workbook(workbook: &mut Workbook, pre: &Workbook) {
    *workbook = (*pre).clone();
    workbook.rebuild_cross_sheet_deps();
    for sheet in &mut workbook.sheets {
        sheet.resweep_all_spills();
    }
    // Conservatively re-dirty every formula cell. The snapshot's `dirty`
    // is unrelated to what's stale now (cells dirtied between the
    // snapshot and this undo would otherwise be lost), and the
    // wholesale-replace makes per-cell tracking unreliable.
    workbook.mark_all_formula_cells_dirty();
}

/// Restore `(row, col)` to `data` (or clear if `None`), then propagate.
/// Shared by `apply` and `revert` for `CellModified`.
fn restore_cell(workbook: &mut Workbook, row: usize, col: usize, data: Option<&CellData>) {
    // Route through the workbook chokepoints so the dirty set is
    // populated; undo/redo were previously skipping dirty entirely.
    // `set_cell_on_active` / `clear_cell_on_active` both run a single
    // graph-driven recalc internally, so no separate propagate call.
    match data {
        Some(d) => workbook.set_cell_on_active(row, col, d.clone()),
        None => workbook.clear_cell_on_active(row, col),
    }
}

/// Bulk-apply CellModified-style entries from a batch in O(N) total —
/// write every cell directly to the live workbook, mark the union
/// dirty, then run a single graph-driven recalc. The naive
/// `for a in batch { a.revert(wb) }` approach is O(N²) for a chain
/// because each per-action revert triggers its own downstream cascade.
///
/// Generic over the iterator + projection so it can serve both apply
/// (reads `new_cell`) and revert (reads `old_cell` in reverse order)
/// without duplicating the body.
fn batch_apply_cells_then_recalc<'a, I, F>(
    workbook: &mut Workbook,
    actions: I,
    project: F,
) -> Result<(), crate::domain::services::CalcError>
where
    I: Iterator<Item = &'a UndoAction>,
    F: Fn(&'a UndoAction) -> Option<(usize, usize, Option<CellData>)>,
{
    let sheet_name = workbook.sheet_names[workbook.active_sheet].clone();
    let active_idx = workbook.active_sheet;
    let mut count = 0usize;
    for a in actions {
        if let Some((row, col, data)) = project(a) {
            // Write the cell value directly into the active sheet's
            // cells map. Skip set_cell so the per-sheet cascade
            // doesn't fire; we'll do ONE recalc at the end.
            let sheet = &mut workbook.sheets[active_idx];
            match data {
                Some(d) => {
                    // sweep prior spill ghosts if this cell was an
                    // anchor; the new cell may not be.
                    sheet.sweep_spill_ghosts_for(row, col);
                    sheet.cells.insert((row, col), d);
                }
                None => {
                    sheet.sweep_spill_ghosts_for(row, col);
                    sheet.cells.remove(&(row, col));
                }
            }
            // Clear the CF cache once per write (cheap; render-only).
            sheet.cf_cache.lock().unwrap().clear();
            workbook.dirty.insert((sheet_name.clone(), row, col));
            count += 1;
        }
    }
    if count == 0 {
        return Ok(());
    }
    // The legacy per-sheet dep graph needs rebuilding because we
    // bypassed `add_cell_dependencies`. The unified graph gets
    // rebuilt lazily inside recalc_via_graph_result.
    workbook.rebuild_cross_sheet_deps();
    workbook.build_dep_graph_from_scratch();
    // Use the Result variant — the eprintln-swallowing wrapper would
    // corrupt the TUI alt-screen during an undo/redo on a non-converging
    // cyclic workbook.
    workbook.recalc_via_graph_result()
}

fn restore_cf(workbook: &mut Workbook, sheet_idx: usize, rules: &[crate::domain::ConditionalFormat]) {
    if let Some(sheet) = workbook.sheets.get_mut(sheet_idx) {
        sheet.conditional_formats = rules.to_vec();
        sheet.cf_cache.lock().unwrap().clear();
    }
}

/// Run `f` with `sheet_idx` as the active sheet, restoring the prior active
/// sheet afterward. Lets undo/redo operations that take an explicit
/// sheet_idx reuse the workbook's `*_on_active` family without a permanent
/// active-sheet switch surfacing to the UI.
fn with_active_sheet<F: FnOnce(&mut Workbook)>(workbook: &mut Workbook, sheet_idx: usize, f: F) {
    let prior = workbook.active_sheet;
    workbook.active_sheet = sheet_idx;
    f(workbook);
    workbook.active_sheet = prior;
}
