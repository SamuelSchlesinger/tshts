//! Submodule of `models` — see models/mod.rs.

#![allow(unused_imports)]
use serde::{Deserialize, Serialize};
use std::collections::{HashMap, HashSet, VecDeque};
use super::*;

/// A cell address that includes the sheet name. Used as the key type for
/// the workbook-level cross-sheet dependency graph.
pub type CrossSheetKey = (String, usize, usize);

/// A workbook containing multiple spreadsheets (tabs).
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Workbook {
    /// The sheets in this workbook
    pub sheets: Vec<Spreadsheet>,
    /// Names for each sheet
    pub sheet_names: Vec<String>,
    /// Index of the currently active sheet
    pub active_sheet: usize,
    /// Named ranges: name -> cell reference string (e.g., "Revenue" -> "B2:B50")
    pub named_ranges: HashMap<String, String>,
    /// Workbook-level dep graph for CROSS-sheet references only. Same-sheet
    /// deps remain on each Spreadsheet. `dependents[P]` is the set of
    /// cells that reference `P` from a different sheet. Not serialized —
    /// rebuilt on load.
    #[serde(skip)]
    pub cross_sheet_dependents: HashMap<CrossSheetKey, HashSet<CrossSheetKey>>,
    /// Inverse of `cross_sheet_dependents`: `dependencies[X]` is the set of
    /// cells `X` references from other sheets. Used to clear stale deps when
    /// a formula changes.
    #[serde(skip)]
    pub cross_sheet_dependencies: HashMap<CrossSheetKey, HashSet<CrossSheetKey>>,
}

impl Default for Workbook {
    fn default() -> Self {
        Self {
            sheets: vec![Spreadsheet::default()],
            sheet_names: vec!["Sheet1".to_string()],
            active_sheet: 0,
            named_ranges: HashMap::new(),
            cross_sheet_dependents: HashMap::new(),
            cross_sheet_dependencies: HashMap::new(),
        }
    }
}

impl Workbook {
    /// Gets a reference to the active sheet.
    pub fn current_sheet(&self) -> &Spreadsheet {
        &self.sheets[self.active_sheet]
    }

    /// Gets a mutable reference to the active sheet.
    pub fn current_sheet_mut(&mut self) -> &mut Spreadsheet {
        &mut self.sheets[self.active_sheet]
    }

    /// Adds a new empty sheet with the given name.
    pub fn add_sheet(&mut self, name: String) {
        self.sheets.push(Spreadsheet::default());
        self.sheet_names.push(name);
    }

    /// Removes a sheet by index. Adjusts active_sheet if needed.
    /// Won't remove the last sheet.
    pub fn remove_sheet(&mut self, index: usize) -> bool {
        if self.sheets.len() <= 1 || index >= self.sheets.len() {
            return false;
        }
        let removed_name = self.sheet_names[index].clone();
        self.sheets.remove(index);
        self.sheet_names.remove(index);
        if self.active_sheet >= self.sheets.len() {
            self.active_sheet = self.sheets.len() - 1;
        } else if self.active_sheet > index {
            self.active_sheet -= 1;
        }
        // Purge any cross-sheet dep entries that touched the removed sheet.
        self.cross_sheet_dependents
            .retain(|k, _| !k.0.eq_ignore_ascii_case(&removed_name));
        for set in self.cross_sheet_dependents.values_mut() {
            set.retain(|k| !k.0.eq_ignore_ascii_case(&removed_name));
        }
        self.cross_sheet_dependencies
            .retain(|k, _| !k.0.eq_ignore_ascii_case(&removed_name));
        for set in self.cross_sheet_dependencies.values_mut() {
            set.retain(|k| !k.0.eq_ignore_ascii_case(&removed_name));
        }
        true
    }

    /// Renames the active sheet.
    /// Rename the active sheet. Returns false (no-op) if `new_name` is empty
    /// or duplicates another sheet's name (case-insensitive). On success,
    /// rewrites formulas in every sheet AND named-range values that
    /// referenced the old name.
    pub fn rename_sheet(&mut self, new_name: String) -> bool {
        let old_name = self.sheet_names[self.active_sheet].clone();
        if old_name == new_name {
            return true;
        }
        // Reject empty names (unreferenceable in formulas).
        if new_name.trim().is_empty() {
            return false;
        }
        // Reject duplicates against any OTHER sheet (case-insensitive).
        if self
            .sheet_names
            .iter()
            .enumerate()
            .any(|(i, n)| i != self.active_sheet && n.eq_ignore_ascii_case(&new_name))
        {
            return false;
        }
        self.sheet_names[self.active_sheet] = new_name.clone();
        // Rewrite formulas in every sheet.
        for sheet in &mut self.sheets {
            let updates: Vec<(usize, usize, String)> = sheet
                .cells
                .iter()
                .filter_map(|(&(r, c), cd)| {
                    cd.formula
                        .as_ref()
                        .map(|f| (r, c, rewrite_sheet_refs(f, &old_name, &new_name)))
                })
                .filter(|(r, c, new_formula)| {
                    sheet
                        .cells
                        .get(&(*r, *c))
                        .and_then(|cd| cd.formula.as_ref())
                        .map(|old| old != new_formula)
                        .unwrap_or(false)
                })
                .collect();
            for (r, c, formula) in updates {
                if let Some(cd) = sheet.cells.get_mut(&(r, c)) {
                    cd.formula = Some(formula);
                }
            }
            sheet.rebuild_dependencies();
        }
        // Update named-range values too — they often contain sheet-qualified
        // ranges (`Sheet1!A1:B10`) that need the rename.
        let updated: Vec<(String, String)> = self
            .named_ranges
            .iter()
            .map(|(k, v)| (k.clone(), rewrite_sheet_refs_for_name_value(v, &old_name, &new_name)))
            .filter(|(k, v)| self.named_ranges.get(k).map(|orig| orig != v).unwrap_or(false))
            .collect();
        for (k, v) in updated {
            self.named_ranges.insert(k.clone(), v.clone());
            for sheet in &mut self.sheets {
                sheet.named_ranges.insert(k.clone(), v.clone());
            }
        }
        // Cross-sheet dep keys reference sheet names by string; rebuild.
        self.rebuild_cross_sheet_deps();
        true
    }

    /// Define or replace a named range. Keys are normalized to uppercase so
    /// formulas can reference them case-insensitively.
    pub fn set_name(&mut self, name: &str, value: &str) {
        let key = name.to_uppercase();
        self.named_ranges.insert(key.clone(), value.to_string());
        for sheet in &mut self.sheets {
            sheet.named_ranges.insert(key.clone(), value.to_string());
            sheet.rebuild_dependencies();
        }
    }

    /// Remove a named range. Returns true if it existed.
    pub fn remove_name(&mut self, name: &str) -> bool {
        let key = name.to_uppercase();
        let existed = self.named_ranges.remove(&key).is_some();
        for sheet in &mut self.sheets {
            sheet.named_ranges.remove(&key);
            sheet.rebuild_dependencies();
        }
        existed
    }

    /// Creates a Workbook from a single Spreadsheet (for backward compatibility).
    pub fn from_spreadsheet(sheet: Spreadsheet) -> Self {
        Self {
            sheets: vec![sheet],
            sheet_names: vec!["Sheet1".to_string()],
            active_sheet: 0,
            named_ranges: HashMap::new(),
            cross_sheet_dependents: HashMap::new(),
            cross_sheet_dependencies: HashMap::new(),
        }
    }

    /// Re-register the cross-sheet dependencies for the cell at
    /// `(sheet_name, row, col)`. Called after every cell write. Removes
    /// stale entries from the old formula and inserts new ones from the
    /// current one.
    pub fn register_cross_sheet_deps(&mut self, sheet_name: &str, row: usize, col: usize) {
        use crate::domain::services::FormulaEvaluator;
        let key: CrossSheetKey = (sheet_name.to_string(), row, col);

        // Step 1: clear old reverse links.
        if let Some(old_precs) = self.cross_sheet_dependencies.remove(&key) {
            for p in old_precs {
                if let Some(set) = self.cross_sheet_dependents.get_mut(&p) {
                    set.remove(&key);
                    if set.is_empty() {
                        self.cross_sheet_dependents.remove(&p);
                    }
                }
            }
        }

        // Step 2: pull the current formula.
        let sheet_idx = match self.sheet_names.iter().position(|n| n == sheet_name) {
            Some(i) => i,
            None => return,
        };
        let formula = match self.sheets[sheet_idx]
            .cells
            .get(&(row, col))
            .and_then(|cd| cd.formula.clone())
        {
            Some(f) => f,
            None => return,
        };

        // Step 3: extract qualified refs from the formula. Use a snapshot
        // so the evaluator can borrow the workbook immutably.
        let qualified_refs: Vec<(Option<String>, usize, usize)> = {
            let names = self.named_ranges.clone();
            let evaluator = FormulaEvaluator::with_workbook(
                self,
                &self.sheets[sheet_idx],
                &names,
            );
            evaluator.extract_qualified_refs(&formula)
        };

        // Step 4: register the cross-sheet ones (skip refs back to the same
        // sheet — those already live in the per-sheet dep graph).
        for (ref_sheet, ref_row, ref_col) in qualified_refs {
            // Skip if no explicit sheet or if it points to the same sheet.
            let resolved_sheet = match ref_sheet {
                Some(s) if !s.eq_ignore_ascii_case(sheet_name) => s,
                _ => continue,
            };
            // Normalize to the canonical sheet-name casing in `sheet_names`.
            let canon = self
                .sheet_names
                .iter()
                .find(|n| n.eq_ignore_ascii_case(&resolved_sheet))
                .cloned()
                .unwrap_or(resolved_sheet);
            let prec: CrossSheetKey = (canon, ref_row, ref_col);
            self.cross_sheet_dependencies
                .entry(key.clone())
                .or_insert_with(HashSet::new)
                .insert(prec.clone());
            self.cross_sheet_dependents
                .entry(prec)
                .or_insert_with(HashSet::new)
                .insert(key.clone());
        }
    }

    /// Recalculate every cell that depends on `(sheet_name, row, col)` via
    /// the cross-sheet graph. Walks transitively (BFS) so chains like
    /// Sheet1!A1 → Sheet2!A1 → Sheet3!A1 all update.
    pub fn propagate_cross_sheet_changes(
        &mut self,
        sheet_name: &str,
        row: usize,
        col: usize,
    ) {
        use crate::domain::services::FormulaEvaluator;
        let mut queue: std::collections::VecDeque<CrossSheetKey> =
            std::collections::VecDeque::new();
        queue.push_back((sheet_name.to_string(), row, col));
        let mut visited: HashSet<CrossSheetKey> = HashSet::new();

        while let Some(key) = queue.pop_front() {
            if !visited.insert(key.clone()) {
                continue;
            }
            let deps: Vec<CrossSheetKey> = self
                .cross_sheet_dependents
                .get(&key)
                .cloned()
                .unwrap_or_default()
                .into_iter()
                .collect();
            if deps.is_empty() {
                continue;
            }
            // Snapshot the workbook so the evaluator can borrow it while
            // we mutate `self.sheets`. Cloning is heavy but only happens
            // when something cross-sheet actually fires; for the common
            // case of no cross-sheet deps, this is never reached.
            let snapshot = self.clone();
            let names = snapshot.named_ranges.clone();
            for dep in deps {
                let (dep_sheet, dep_row, dep_col) = dep.clone();
                let Some(dep_idx) = snapshot
                    .sheet_names
                    .iter()
                    .position(|n| n.eq_ignore_ascii_case(&dep_sheet))
                else {
                    continue;
                };
                let snap_sheet = &snapshot.sheets[dep_idx];
                let Some(cd) = snap_sheet.cells.get(&(dep_row, dep_col)) else {
                    continue;
                };
                let Some(formula) = cd.formula.clone() else { continue };
                let evaluator =
                    FormulaEvaluator::with_workbook(&snapshot, snap_sheet, &names);
                let new_value = evaluator.evaluate_formula(&formula);
                // Write to the real workbook.
                let dep_real_idx = self
                    .sheet_names
                    .iter()
                    .position(|n| n.eq_ignore_ascii_case(&dep_sheet));
                if let Some(idx) = dep_real_idx {
                    if let Some(real_cd) = self.sheets[idx].cells.get_mut(&(dep_row, dep_col))
                    {
                        if real_cd.value != new_value {
                            real_cd.value = new_value;
                            queue.push_back(dep);
                        }
                    }
                }
            }
        }
    }

    /// Check whether adding a new formula at `(sheet_name, row, col)` with
    /// the given precedents would create a cross-sheet cycle. Walks the
    /// existing cross-sheet graph from each precedent; if any path reaches
    /// `(sheet_name, row, col)`, we'd loop. The same-sheet check still runs
    /// separately via `FormulaEvaluator::would_create_circular_reference`.
    pub fn would_create_cross_sheet_cycle(
        &self,
        sheet_name: &str,
        row: usize,
        col: usize,
        candidate_precedents: &[(Option<String>, usize, usize)],
    ) -> bool {
        let target: CrossSheetKey = (sheet_name.to_string(), row, col);
        let mut stack: Vec<CrossSheetKey> = Vec::new();
        for (prec_sheet, prec_row, prec_col) in candidate_precedents {
            // Only consider cross-sheet precedents (same-sheet cycles are
            // caught by the existing AST walker).
            let Some(ps) = prec_sheet else { continue };
            if ps.eq_ignore_ascii_case(sheet_name) {
                continue;
            }
            let canon = self
                .sheet_names
                .iter()
                .find(|n| n.eq_ignore_ascii_case(ps))
                .cloned()
                .unwrap_or_else(|| ps.clone());
            stack.push((canon, *prec_row, *prec_col));
        }
        let mut visited: HashSet<CrossSheetKey> = HashSet::new();
        while let Some(node) = stack.pop() {
            if node == target {
                return true;
            }
            if !visited.insert(node.clone()) {
                continue;
            }
            // Also walk down: the cells that *node* in turn depends on.
            if let Some(deps) = self.cross_sheet_dependencies.get(&node) {
                for d in deps {
                    stack.push(d.clone());
                }
            }
        }
        false
    }

    /// Rebuild the cross-sheet dep graph from scratch by scanning every
    /// formula in every sheet. Called after load (since the graph isn't
    /// serialized) and as a fallback when state drifts.
    pub fn rebuild_cross_sheet_deps(&mut self) {
        self.cross_sheet_dependents.clear();
        self.cross_sheet_dependencies.clear();
        let cells: Vec<(String, usize, usize)> = self
            .sheet_names
            .iter()
            .enumerate()
            .flat_map(|(idx, name)| {
                self.sheets[idx]
                    .cells
                    .iter()
                    .filter(|(_, cd)| cd.formula.is_some())
                    .map(move |(&(r, c), _)| (name.clone(), r, c))
            })
            .collect();
        for (sheet, r, c) in cells {
            self.register_cross_sheet_deps(&sheet, r, c);
        }
    }
}
