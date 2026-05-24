//! Cross-sheet ref adjustment. When a row/col is inserted or deleted on
//! Sheet1, every OTHER sheet's `=Sheet1!A5` refs need shifting in-place —
//! the same-sheet `map_ast_refs` doesn't touch sheet-qualified refs, so we
//! need a parallel `map_qualified_ast_refs` that filters on the sheet name.
//!
//! `ast_to_string` is shared with `adjust.rs`.

use super::FormulaEvaluator;
use crate::domain::models::Spreadsheet;
use crate::domain::parser::{ErrorKind, Expr, Parser};

impl<'a> FormulaEvaluator<'a> {
    /// Adjusts SHEET-QUALIFIED references in a formula whose sheet matches
    /// `target_sheet` (case-insensitive). Used by `Workbook` when a
    /// structural mutation on one sheet needs to update cross-sheet refs on
    /// every OTHER sheet — e.g. inserting a row into Sheet1 must shift
    /// `=Sheet1!A5` on Sheet2 to `=Sheet1!A6`. Unqualified refs and refs to
    /// other sheets are left untouched.
    pub fn adjust_formula_for_sheet_row_insert(
        &self,
        formula: &str,
        target_sheet: &str,
        at: usize,
    ) -> String {
        self.adjust_formula_for_sheet_structural(formula, target_sheet, |row, col| {
            if row >= at { Some((row + 1, col)) } else { Some((row, col)) }
        })
    }

    pub fn adjust_formula_for_sheet_row_delete(
        &self,
        formula: &str,
        target_sheet: &str,
        at: usize,
    ) -> String {
        self.adjust_formula_for_sheet_structural(formula, target_sheet, |row, col| {
            if row == at { None }
            else if row > at { Some((row - 1, col)) }
            else { Some((row, col)) }
        })
    }

    pub fn adjust_formula_for_sheet_col_insert(
        &self,
        formula: &str,
        target_sheet: &str,
        at: usize,
    ) -> String {
        self.adjust_formula_for_sheet_structural(formula, target_sheet, |row, col| {
            if col >= at { Some((row, col + 1)) } else { Some((row, col)) }
        })
    }

    pub fn adjust_formula_for_sheet_col_delete(
        &self,
        formula: &str,
        target_sheet: &str,
        at: usize,
    ) -> String {
        self.adjust_formula_for_sheet_structural(formula, target_sheet, |row, col| {
            if col == at { None }
            else if col > at { Some((row, col - 1)) }
            else { Some((row, col)) }
        })
    }

    /// Generic structural-mutation adjustment scoped to one sheet's
    /// qualified refs. Mirrors `adjust_formula_structural` but operates on
    /// sheet-qualified refs whose sheet name matches `target_sheet`.
    fn adjust_formula_for_sheet_structural<F>(
        &self,
        formula: &str,
        target_sheet: &str,
        map_ref: F,
    ) -> String
    where
        F: Fn(usize, usize) -> Option<(usize, usize)> + Copy,
    {
        if !formula.starts_with('=') {
            return formula.to_string();
        }
        let expr = &formula[1..];
        match Parser::new(expr) {
            Ok(mut parser) => match parser.parse() {
                Ok(ast) => {
                    let adjusted = self.map_qualified_ast_refs(&ast, target_sheet, map_ref);
                    format!("={}", self.ast_to_string(&adjusted))
                }
                Err(_) => formula.to_string(),
            },
            Err(_) => formula.to_string(),
        }
    }

    /// AST walker that applies `map_ref` only to cell references whose sheet
    /// qualifier matches `target_sheet` (case-insensitive). Refs without a
    /// sheet qualifier, and refs to other sheets, pass through unchanged.
    fn map_qualified_ast_refs<F>(&self, expr: &Expr, target_sheet: &str, map_ref: F) -> Expr
    where
        F: Fn(usize, usize) -> Option<(usize, usize)> + Copy,
    {
        // Returns Ok(new_ref_string) if the ref was on the target sheet and
        // successfully shifted; Ok(unchanged) if it's a different sheet or
        // an unqualified ref; Err(()) if map_ref refused the shift (caller
        // emits Expr::ErrorLit(Ref)).
        let remap = |cell_ref: &str| -> Result<String, ()> {
            let Some((Some(sheet_name), row, col, abs_row, abs_col)) =
                Spreadsheet::parse_qualified_reference(cell_ref)
            else {
                return Ok(cell_ref.to_string());
            };
            if !sheet_name.eq_ignore_ascii_case(target_sheet) {
                return Ok(cell_ref.to_string());
            }
            match map_ref(row, col) {
                Some((nr, nc)) => {
                    let new_cell = Spreadsheet::format_cell_reference(nr, nc, abs_row, abs_col);
                    // Re-quote sheet names that aren't bare identifiers
                    // (anything beyond [A-Za-z0-9_]); `parse_qualified_reference`
                    // strips quotes, so naively concatenating with `!` would
                    // emit a syntactically broken ref for names with spaces.
                    Ok(format!("{}!{}", super::format_sheet_name(&sheet_name), new_cell))
                }
                None => Err(()),
            }
        };
        match expr {
            Expr::String(s) => Expr::String(s.clone()),
            Expr::Number(n) => Expr::Number(*n),
            Expr::ErrorLit(k) => Expr::ErrorLit(*k),
            Expr::NamedRef(n) => Expr::NamedRef(n.clone()),
            Expr::CellRef(cell_ref) => match remap(cell_ref) {
                Ok(s) => Expr::CellRef(s),
                Err(()) => Expr::ErrorLit(ErrorKind::Ref),
            },
            Expr::Range(start, end) => match (remap(start), remap(end)) {
                (Ok(s), Ok(e)) => Expr::Range(s, e),
                _ => Expr::ErrorLit(ErrorKind::Ref),
            },
            Expr::Binary { left, operator, right } => Expr::Binary {
                left: Box::new(self.map_qualified_ast_refs(left, target_sheet, map_ref)),
                operator: operator.clone(),
                right: Box::new(self.map_qualified_ast_refs(right, target_sheet, map_ref)),
            },
            Expr::Unary { operator, operand } => Expr::Unary {
                operator: operator.clone(),
                operand: Box::new(self.map_qualified_ast_refs(operand, target_sheet, map_ref)),
            },
            Expr::FunctionCall { name, args } => Expr::FunctionCall {
                name: name.clone(),
                args: args
                    .iter()
                    .map(|a| self.map_qualified_ast_refs(a, target_sheet, map_ref))
                    .collect(),
            },
            Expr::Let { bindings, body } => Expr::Let {
                bindings: bindings
                    .iter()
                    .map(|(n, v)| (n.clone(), Box::new(self.map_qualified_ast_refs(v, target_sheet, map_ref))))
                    .collect(),
                body: Box::new(self.map_qualified_ast_refs(body, target_sheet, map_ref)),
            },
            Expr::Lambda { params, body } => Expr::Lambda {
                params: params.clone(),
                body: Box::new(self.map_qualified_ast_refs(body, target_sheet, map_ref)),
            },
            Expr::ArrayLiteral { rows } => Expr::ArrayLiteral {
                rows: rows
                    .iter()
                    .map(|r| {
                        r.iter()
                            .map(|c| self.map_qualified_ast_refs(c, target_sheet, map_ref))
                            .collect()
                    })
                    .collect(),
            },
        }
    }
}

