//! Cell-reference extraction. Drives the workbook dep graph: each formula's
//! refs become incoming edges for the cell that contains it.

use super::FormulaEvaluator;
use crate::domain::models::Spreadsheet;
use crate::domain::parser::{Expr, Parser};
use super::super::with_parse_cache;

impl<'a> FormulaEvaluator<'a> {
    /// Returns the cells (row, col) referenced by `formula`. Used for the
    /// dependency graph that drives automatic recalc.
    pub fn extract_cell_references(&self, formula: &str) -> Vec<(usize, usize)> {
        if !formula.starts_with('=') {
            return Vec::new();
        }
        let expr = &formula[1..];
        match with_parse_cache(expr, |s| {
            let mut p = Parser::new(s)?;
            p.parse()
        }) {
            Ok(ast) => self.extract_cell_references_from_ast(&ast),
            Err(_) => Vec::new(),
        }
    }

    /// Like `extract_cell_references` but preserves the sheet name of each
    /// reference. Returns `(sheet_name, row, col)` triples; `sheet_name` is
    /// `None` for unqualified refs (which live in the current sheet) and
    /// `Some(name)` for cross-sheet refs. Used by the workbook to build
    /// its cross-sheet dependency graph.
    pub fn extract_qualified_refs(&self, formula: &str) -> Vec<(Option<String>, usize, usize)> {
        if !formula.starts_with('=') {
            return Vec::new();
        }
        let expr = &formula[1..];
        match with_parse_cache(expr, |s| {
            let mut p = Parser::new(s)?;
            p.parse()
        }) {
            Ok(ast) => {
                let mut out = Vec::new();
                self.extract_qualified_refs_from_ast(&ast, &mut out);
                out
            }
            Err(_) => Vec::new(),
        }
    }

    fn extract_qualified_refs_from_ast(
        &self,
        expr: &Expr,
        out: &mut Vec<(Option<String>, usize, usize)>,
    ) {
        match expr {
            Expr::String(_) | Expr::Number(_) | Expr::ErrorLit(_) => {}
            Expr::CellRef(cell_ref) => {
                if let Some((sheet, r, c, _, _)) =
                    Spreadsheet::parse_qualified_reference(cell_ref)
                {
                    out.push((sheet, r, c));
                }
            }
            Expr::Range(start_cell, end_cell) => {
                // 3-D markers (`<S1>..<S3>!<cell>`): expand the cell across
                // each sheet in the span [S1..=S3], inclusive of intermediates.
                // The previous version only emitted S1 and S3, so a change
                // on an intermediate sheet (S2!A1) failed to trigger recalc
                // of a cell that referenced `S1:S3!A1`.
                if let Some((s1, s2, cell)) = Spreadsheet::parse_three_d_marker(start_cell)
                {
                    if let Some((row, col)) = Spreadsheet::parse_cell_reference(&cell) {
                        let pushed = if let Some(wb) = self.workbook {
                            let lo = wb.sheet_names.iter()
                                .position(|n| n.eq_ignore_ascii_case(&s1));
                            let hi = wb.sheet_names.iter()
                                .position(|n| n.eq_ignore_ascii_case(&s2));
                            if let (Some(lo), Some(hi)) = (lo, hi) {
                                let (a, b) = if lo <= hi { (lo, hi) } else { (hi, lo) };
                                for i in a..=b {
                                    out.push((Some(wb.sheet_names[i].clone()), row, col));
                                }
                                true
                            } else {
                                false
                            }
                        } else {
                            false
                        };
                        if !pushed {
                            // Workbook context absent (or sheet name not found):
                            // fall back to the boundaries so dep registration
                            // still covers something. Single-sheet evaluator
                            // path; intermediate sheets are unknowable here.
                            out.push((Some(s1.clone()), row, col));
                            if !s1.eq_ignore_ascii_case(&s2) {
                                out.push((Some(s2), row, col));
                            }
                        }
                    }
                    return;
                }
                let sp = Spreadsheet::parse_qualified_reference(start_cell);
                let ep = Spreadsheet::parse_qualified_reference(end_cell);
                if let (
                    Some((sheet, sr, sc, _, _)),
                    Some((_, er, ec, _, _)),
                ) = (sp, ep)
                {
                    for row in sr..=er {
                        for col in sc..=ec {
                            out.push((sheet.clone(), row, col));
                        }
                    }
                }
            }
            Expr::Binary { left, right, .. } => {
                self.extract_qualified_refs_from_ast(left, out);
                self.extract_qualified_refs_from_ast(right, out);
            }
            Expr::Unary { operand, .. } => {
                self.extract_qualified_refs_from_ast(operand, out);
            }
            Expr::FunctionCall { args, .. } => {
                for arg in args {
                    self.extract_qualified_refs_from_ast(arg, out);
                }
            }
            Expr::NamedRef(name) => {
                if let Some(names) = self.names
                    && let Some(value) =
                        names.get(&name.to_uppercase()).or_else(|| names.get(name))
                        && let Ok(mut p) = Parser::new(value)
                            && let Ok(ast) = p.parse() {
                                self.extract_qualified_refs_from_ast(&ast, out);
                            }
            }
            Expr::Let { bindings, body } => {
                for (_, v) in bindings {
                    self.extract_qualified_refs_from_ast(v, out);
                }
                self.extract_qualified_refs_from_ast(body, out);
            }
            Expr::Lambda { body, .. } => {
                self.extract_qualified_refs_from_ast(body, out);
            }
            Expr::ArrayLiteral { rows } => {
                for row in rows {
                    for c in row {
                        self.extract_qualified_refs_from_ast(c, out);
                    }
                }
            }
        }
    }

    pub(super) fn extract_cell_references_from_ast(&self, expr: &Expr) -> Vec<(usize, usize)> {
        let mut references = Vec::new();

        match expr {
            Expr::String(_) | Expr::ErrorLit(_) => {},
            Expr::CellRef(cell_ref) => {
                if let Some((row, col)) = Spreadsheet::parse_cell_reference(cell_ref) {
                    references.push((row, col));
                }
            }
            Expr::Range(start_cell, end_cell) => {
                if let (Some((start_row, start_col)), Some((end_row, end_col))) =
                    (Spreadsheet::parse_cell_reference(start_cell), Spreadsheet::parse_cell_reference(end_cell)) {
                    for row in start_row..=end_row {
                        for col in start_col..=end_col {
                            references.push((row, col));
                        }
                    }
                }
            }
            Expr::Binary { left, right, .. } => {
                references.extend(self.extract_cell_references_from_ast(left));
                references.extend(self.extract_cell_references_from_ast(right));
            }
            Expr::Unary { operand, .. } => {
                references.extend(self.extract_cell_references_from_ast(operand));
            }
            Expr::FunctionCall { args, .. } => {
                for arg in args {
                    references.extend(self.extract_cell_references_from_ast(arg));
                }
            }
            Expr::NamedRef(name) => {
                if let Some(names) = self.names
                    && let Some(value) = names.get(&name.to_uppercase()).or_else(|| names.get(name))
                        && let Ok(mut p) = Parser::new(value)
                            && let Ok(ast) = p.parse() {
                                references.extend(self.extract_cell_references_from_ast(&ast));
                            }
            }
            Expr::Number(_) => {}
            Expr::Let { bindings, body } => {
                for (_, v) in bindings {
                    references.extend(self.extract_cell_references_from_ast(v));
                }
                references.extend(self.extract_cell_references_from_ast(body));
            }
            Expr::Lambda { body, .. } => {
                references.extend(self.extract_cell_references_from_ast(body));
            }
            Expr::ArrayLiteral { rows } => {
                for row in rows {
                    for c in row {
                        references.extend(self.extract_cell_references_from_ast(c));
                    }
                }
            }
        }

        references
    }
}
