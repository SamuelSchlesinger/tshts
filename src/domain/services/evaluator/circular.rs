//! Circular-reference detection. `would_create_circular_reference` in
//! `mod.rs` is the public entry; this file holds the AST walker.

use super::FormulaEvaluator;
use crate::domain::models::Spreadsheet;
use crate::domain::parser::{Expr, Parser};
use std::collections::HashSet;

impl<'a> FormulaEvaluator<'a> {
    /// Checks for circular references in a formula string, reusing the visited set.
    fn check_circular_in_formula(&self, formula: &str, target_cell: (usize, usize), visited: &mut HashSet<(usize, usize)>) -> bool {
        if !formula.starts_with('=') {
            return false;
        }
        let expr = &formula[1..];
        match Parser::new(expr) {
            Ok(mut parser) => match parser.parse() {
                Ok(ast) => self.check_circular_reference_in_ast(&ast, target_cell, visited),
                Err(_) => false,
            },
            Err(_) => false,
        }
    }

    /// Checks for circular references in an AST.
    pub(super) fn check_circular_reference_in_ast(&self, expr: &Expr, target_cell: (usize, usize), visited: &mut HashSet<(usize, usize)>) -> bool {
        match expr {
            Expr::String(_) | Expr::ErrorLit(_) => false,
            Expr::CellRef(cell_ref) => {
                if let Some((row, col)) = Spreadsheet::parse_cell_reference(cell_ref) {
                    if (row, col) == target_cell {
                        return true;
                    }

                    if visited.contains(&(row, col)) {
                        return false;
                    }

                    visited.insert((row, col));

                    let cell = self.spreadsheet.get_cell(row, col);
                    if let Some(ref cell_formula) = cell.formula
                        && self.check_circular_in_formula(cell_formula, target_cell, visited) {
                            return true;
                        }
                }
                false
            }
            Expr::Range(start_cell, end_cell) => {
                if let (Some((start_row, start_col)), Some((end_row, end_col))) =
                    (Spreadsheet::parse_cell_reference(start_cell), Spreadsheet::parse_cell_reference(end_cell)) {
                    for row in start_row..=end_row {
                        for col in start_col..=end_col {
                            if (row, col) == target_cell {
                                return true;
                            }
                            if !visited.contains(&(row, col)) {
                                visited.insert((row, col));
                                let cell = self.spreadsheet.get_cell(row, col);
                                if let Some(ref cell_formula) = cell.formula
                                    && self.check_circular_in_formula(cell_formula, target_cell, visited) {
                                        return true;
                                    }
                            }
                        }
                    }
                }
                false
            }
            Expr::Binary { left, right, .. } => {
                self.check_circular_reference_in_ast(left, target_cell, visited) ||
                self.check_circular_reference_in_ast(right, target_cell, visited)
            }
            Expr::Unary { operand, .. } => {
                self.check_circular_reference_in_ast(operand, target_cell, visited)
            }
            Expr::FunctionCall { args, .. } => {
                args.iter().any(|arg| self.check_circular_reference_in_ast(arg, target_cell, visited))
            }
            // NamedRef: resolve via names and recursively check.
            Expr::NamedRef(name) => {
                if let Some(names) = self.names
                    && let Some(value) = names.get(&name.to_uppercase()).or_else(|| names.get(name))
                        && let Ok(mut p) = Parser::new(value)
                            && let Ok(ast) = p.parse() {
                                return self.check_circular_reference_in_ast(&ast, target_cell, visited);
                            }
                false
            }
            Expr::Number(_) => false,
            // LET/LAMBDA: walk bindings/body, ignoring scoping issues for
            // circular-ref detection (over-approximation is acceptable).
            Expr::Let { bindings, body } => {
                bindings.iter().any(|(_, v)| {
                    self.check_circular_reference_in_ast(v, target_cell, visited)
                }) || self.check_circular_reference_in_ast(body, target_cell, visited)
            }
            Expr::Lambda { body, .. } => {
                self.check_circular_reference_in_ast(body, target_cell, visited)
            }
            Expr::ArrayLiteral { rows } => rows.iter().flatten().any(|c| {
                self.check_circular_reference_in_ast(c, target_cell, visited)
            }),
        }
    }
}
