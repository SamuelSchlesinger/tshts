//! Same-sheet ref adjustment plus AST↔string serialization.
//!
//! Two flavors of adjustment live here:
//! 1. **Relative shift** (`adjust_formula_references`) — paste/autofill;
//!    relative parts move by `(row_off, col_off)`, absolute parts stay.
//! 2. **Structural row/col insert/delete** (`adjust_formula_for_row_insert`
//!    et al) — driven by `adjust_formula_structural` / `map_ast_refs`.
//!
//! `ast_to_string` is shared with `cross_sheet.rs`; `binary_op_to_string`
//! and `unary_op_to_string` are only used by `ast_to_string`.

use super::FormulaEvaluator;
use crate::domain::models::Spreadsheet;
use crate::domain::parser::{ErrorKind, Expr, Parser};
use std::collections::HashMap;

impl<'a> FormulaEvaluator<'a> {
    /// Shifts relative cell references by the given offsets; absolute parts
    /// (`$col` / `$row`) are preserved. Used for fill and paste.
    ///
    /// ```
    /// use tshts::domain::{Spreadsheet, FormulaEvaluator};
    ///
    /// let sheet = Spreadsheet::default();
    /// let evaluator = FormulaEvaluator::new(&sheet);
    /// assert_eq!(evaluator.adjust_formula_references("=SUM(B4:B6)", 0, 1), "=SUM(C4:C6)");
    /// assert_eq!(evaluator.adjust_formula_references("=A1+B1", 1, 0), "=A2+B2");
    /// ```
    pub fn adjust_formula_references(&self, formula: &str, row_offset: i32, col_offset: i32) -> String {
        if !formula.starts_with('=') {
            return formula.to_string();
        }

        let expr = &formula[1..];
        match Parser::new(expr) {
            Ok(mut parser) => match parser.parse() {
                Ok(ast) => {
                    let adjusted_ast = self.adjust_ast_references(&ast, row_offset, col_offset);
                    format!("={}", self.ast_to_string(&adjusted_ast))
                }
                Err(_) => formula.to_string(),
            },
            Err(_) => formula.to_string(),
        }
    }

    /// Adjusts cell references in an AST. Absolute parts (`$col` / `$row`) are
    /// preserved verbatim — only the relative parts shift by the offsets.
    /// If a relative shift would take a reference past the origin (row or
    /// col < 0), the ref becomes `Expr::ErrorLit(ErrorKind::Ref)`. Error
    /// propagation through Binary/FunctionCall/etc. handles the cascade —
    /// no whole-formula collapse needed.
    fn adjust_ast_references(&self, expr: &Expr, row_offset: i32, col_offset: i32) -> Expr {
        // Returns Ok(new_ref) on success, Err(()) when the shift would take
        // the reference past the origin (caller emits Expr::ErrorLit(Ref)).
        // Unparseable refs (e.g. sheet-qualified like "Sheet2!A1") pass
        // through verbatim — they're handled by a separate adjustment pass
        // outside this function.
        let shift = |cell_ref: &str| -> Result<String, ()> {
            let Some((row, col, abs_row, abs_col)) =
                Spreadsheet::parse_cell_reference_with_flags(cell_ref)
            else {
                return Ok(cell_ref.to_string());
            };
            let new_row = if abs_row {
                Some(row)
            } else {
                let s = row as i32 + row_offset;
                if s >= 0 { Some(s as usize) } else { None }
            };
            let new_col = if abs_col {
                Some(col)
            } else {
                let s = col as i32 + col_offset;
                if s >= 0 { Some(s as usize) } else { None }
            };
            match (new_row, new_col) {
                (Some(r), Some(c)) => {
                    Ok(Spreadsheet::format_cell_reference(r, c, abs_row, abs_col))
                }
                _ => Err(()),
            }
        };
        match expr {
            Expr::String(s) => Expr::String(s.clone()),
            Expr::ErrorLit(k) => Expr::ErrorLit(*k),
            Expr::CellRef(cell_ref) => match shift(cell_ref) {
                Ok(s) => Expr::CellRef(s),
                Err(()) => Expr::ErrorLit(ErrorKind::Ref),
            },
            // If either endpoint shifts past the origin, the whole range is
            // invalid → ErrorLit(Ref). Error propagation through the
            // containing expression (SUM, Binary, etc.) does the rest.
            Expr::Range(start, end) => match (shift(start), shift(end)) {
                (Ok(s), Ok(e)) => Expr::Range(s, e),
                _ => Expr::ErrorLit(ErrorKind::Ref),
            },
            Expr::Binary { left, operator, right } => Expr::Binary {
                left: Box::new(self.adjust_ast_references(left, row_offset, col_offset)),
                operator: operator.clone(),
                right: Box::new(self.adjust_ast_references(right, row_offset, col_offset)),
            },
            Expr::Unary { operator, operand } => Expr::Unary {
                operator: operator.clone(),
                operand: Box::new(self.adjust_ast_references(operand, row_offset, col_offset)),
            },
            Expr::FunctionCall { name, args } => Expr::FunctionCall {
                name: name.clone(),
                args: args
                    .iter()
                    .map(|arg| self.adjust_ast_references(arg, row_offset, col_offset))
                    .collect(),
            },
            Expr::NamedRef(name) => Expr::NamedRef(name.clone()),
            Expr::Number(n) => Expr::Number(*n),
            // LET/LAMBDA: rewrite their inner refs but don't shift the names.
            Expr::Let { bindings, body } => Expr::Let {
                bindings: bindings
                    .iter()
                    .map(|(n, v)| {
                        (
                            n.clone(),
                            Box::new(self.adjust_ast_references(v, row_offset, col_offset)),
                        )
                    })
                    .collect(),
                body: Box::new(self.adjust_ast_references(body, row_offset, col_offset)),
            },
            Expr::Lambda { params, body } => Expr::Lambda {
                params: params.clone(),
                body: Box::new(self.adjust_ast_references(body, row_offset, col_offset)),
            },
            Expr::ArrayLiteral { rows } => Expr::ArrayLiteral {
                rows: rows
                    .iter()
                    .map(|r| {
                        r.iter()
                            .map(|c| self.adjust_ast_references(c, row_offset, col_offset))
                            .collect()
                    })
                    .collect(),
            },
        }
    }

    /// Converts an AST back to a formula string. Used by both the same-sheet
    /// and cross-sheet adjustment paths after they rewrite refs.
    pub(super) fn ast_to_string(&self, expr: &Expr) -> String {
        match expr {
            Expr::String(s) => format!("\"{}\"", s.replace("\"", "\"\"")),
            Expr::CellRef(cell_ref) => cell_ref.clone(),
            // Error literals serialize as their Excel string (`#REF!`, etc.).
            // The lexer round-trips them back to Token::ErrorLit on re-parse.
            Expr::ErrorLit(kind) => kind.as_str().to_string(),
            Expr::Range(start_ref, end_ref) => format!("{}:{}", start_ref, end_ref),
            Expr::Binary { left, operator, right } => {
                format!("{}{}{}",
                    self.ast_to_string(left),
                    self.binary_op_to_string(operator),
                    self.ast_to_string(right))
            }
            Expr::Unary { operator, operand } => {
                format!("{}{}", self.unary_op_to_string(operator), self.ast_to_string(operand))
            }
            Expr::FunctionCall { name, args } => {
                let arg_strs: Vec<String> = args.iter()
                    .map(|arg| self.ast_to_string(arg))
                    .collect();
                format!("{}({})", name, arg_strs.join(","))
            }
            Expr::NamedRef(name) => name.clone(),
            Expr::Number(n) => n.to_string(),
            Expr::Let { bindings, body } => {
                let mut parts = Vec::with_capacity(bindings.len() * 2 + 1);
                for (n, v) in bindings {
                    parts.push(n.clone());
                    parts.push(self.ast_to_string(v));
                }
                parts.push(self.ast_to_string(body));
                format!("LET({})", parts.join(","))
            }
            Expr::Lambda { params, body } => {
                let mut parts = params.clone();
                parts.push(self.ast_to_string(body));
                format!("LAMBDA({})", parts.join(","))
            }
            Expr::ArrayLiteral { rows } => {
                let row_strs: Vec<String> = rows
                    .iter()
                    .map(|r| {
                        r.iter()
                            .map(|c| self.ast_to_string(c))
                            .collect::<Vec<_>>()
                            .join(",")
                    })
                    .collect();
                format!("{{{}}}", row_strs.join(";"))
            }
        }
    }

    /// Converts a binary operator enum back to a string.
    fn binary_op_to_string(&self, op: &crate::domain::parser::BinaryOp) -> &'static str {
        use crate::domain::parser::BinaryOp;
        match op {
            BinaryOp::Add => "+",
            BinaryOp::Subtract => "-",
            BinaryOp::Multiply => "*",
            BinaryOp::Divide => "/",
            BinaryOp::Power => "**",
            BinaryOp::Modulo => "%",
            BinaryOp::Equal => "=",
            BinaryOp::NotEqual => "<>",
            BinaryOp::Less => "<",
            BinaryOp::LessEqual => "<=",
            BinaryOp::Greater => ">",
            BinaryOp::GreaterEqual => ">=",
            BinaryOp::Concatenate => "&",
        }
    }

    /// Converts a unary operator enum back to a string.
    fn unary_op_to_string(&self, op: &crate::domain::parser::UnaryOp) -> &'static str {
        use crate::domain::parser::UnaryOp;
        match op {
            UnaryOp::Minus => "-",
            UnaryOp::Plus => "+",
        }
    }

    /// Adjusts formula references when a row is inserted at `at`.
    /// References to rows >= at are incremented by 1.
    pub fn adjust_formula_for_row_insert(&self, formula: &str, at: usize) -> String {
        self.adjust_formula_structural(formula, |row, col| {
            if row >= at { Some((row + 1, col)) } else { Some((row, col)) }
        })
    }

    /// Adjusts formula references when a row is deleted at `at`.
    /// References to rows > at are decremented by 1. References to the
    /// deleted row itself become `#REF!`.
    pub fn adjust_formula_for_row_delete(&self, formula: &str, at: usize) -> String {
        self.adjust_formula_structural(formula, |row, col| {
            if row == at { None }
            else if row > at { Some((row - 1, col)) }
            else { Some((row, col)) }
        })
    }

    /// Adjusts formula references when a column is inserted at `at`.
    /// References to cols >= at are incremented by 1.
    pub fn adjust_formula_for_col_insert(&self, formula: &str, at: usize) -> String {
        self.adjust_formula_structural(formula, |row, col| {
            if col >= at { Some((row, col + 1)) } else { Some((row, col)) }
        })
    }

    /// Adjusts formula references when a column is deleted at `at`.
    /// References to cols > at are decremented by 1. References to the
    /// deleted column itself become `#REF!`.
    pub fn adjust_formula_for_col_delete(&self, formula: &str, at: usize) -> String {
        self.adjust_formula_structural(formula, |row, col| {
            if col == at { None }
            else if col > at { Some((row, col - 1)) }
            else { Some((row, col)) }
        })
    }

    /// Remaps row references in a formula using the given old→new row map.
    /// Cells whose row is in the map are remapped; references outside the map
    /// are left alone (so formulas pointing into unsorted regions still work).
    /// Range endpoints are remapped independently.
    pub fn remap_row_references(
        &self,
        formula: &str,
        row_map: &HashMap<usize, usize>,
        _max_row: usize,
    ) -> String {
        self.adjust_formula_structural(formula, |row, col| {
            if let Some(&new_row) = row_map.get(&row) {
                Some((new_row, col))
            } else {
                Some((row, col))
            }
        })
    }

    /// Generic structural formula adjustment using a mapping function on (row, col).
    /// Returning `None` from `map_ref` marks the reference as `#REF!`. Any
    /// formula containing a `#REF!` reference collapses to `=#REF!` so the
    /// re-serialized form round-trips through the lexer cleanly (the parser
    /// cannot tokenize `#`).
    fn adjust_formula_structural<F>(&self, formula: &str, map_ref: F) -> String
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
                    let adjusted = self.map_ast_refs(&ast, map_ref);
                    format!("={}", self.ast_to_string(&adjusted))
                }
                Err(_) => formula.to_string(),
            },
            Err(_) => formula.to_string(),
        }
    }

    /// Maps all cell references in an AST using the given function.
    /// Absolute-row/col markers are preserved on output even when remapped.
    /// `map_ref` returning `None` produces an `Expr::ErrorLit(Ref)`, which
    /// the parser/evaluator handles via standard error propagation.
    fn map_ast_refs<F>(&self, expr: &Expr, map_ref: F) -> Expr
    where
        F: Fn(usize, usize) -> Option<(usize, usize)> + Copy,
    {
        // Returns Ok(new_ref) on success, Err(()) when map_ref refused the
        // shift (caller emits Expr::ErrorLit(Ref)). Unparseable refs (e.g.
        // sheet-qualified) pass through verbatim.
        let remap = |cell_ref: &str| -> Result<String, ()> {
            let Some((row, col, abs_row, abs_col)) =
                Spreadsheet::parse_cell_reference_with_flags(cell_ref)
            else {
                return Ok(cell_ref.to_string());
            };
            match map_ref(row, col) {
                Some((nr, nc)) => Ok(Spreadsheet::format_cell_reference(nr, nc, abs_row, abs_col)),
                None => Err(()),
            }
        };
        match expr {
            Expr::String(s) => Expr::String(s.clone()),
            Expr::Number(n) => Expr::Number(*n),
            Expr::ErrorLit(k) => Expr::ErrorLit(*k),
            Expr::CellRef(cell_ref) => match remap(cell_ref) {
                Ok(s) => Expr::CellRef(s),
                Err(()) => Expr::ErrorLit(ErrorKind::Ref),
            },
            Expr::Range(start, end) => match (remap(start), remap(end)) {
                (Ok(s), Ok(e)) => Expr::Range(s, e),
                _ => Expr::ErrorLit(ErrorKind::Ref),
            },
            Expr::Binary { left, operator, right } => {
                Expr::Binary {
                    left: Box::new(self.map_ast_refs(left, map_ref)),
                    operator: operator.clone(),
                    right: Box::new(self.map_ast_refs(right, map_ref)),
                }
            }
            Expr::Unary { operator, operand } => {
                Expr::Unary {
                    operator: operator.clone(),
                    operand: Box::new(self.map_ast_refs(operand, map_ref)),
                }
            }
            Expr::FunctionCall { name, args } => {
                Expr::FunctionCall {
                    name: name.clone(),
                    args: args.iter().map(|a| self.map_ast_refs(a, map_ref)).collect(),
                }
            }
            Expr::NamedRef(name) => Expr::NamedRef(name.clone()),
            Expr::Let { bindings, body } => Expr::Let {
                bindings: bindings
                    .iter()
                    .map(|(n, v)| (n.clone(), Box::new(self.map_ast_refs(v, map_ref))))
                    .collect(),
                body: Box::new(self.map_ast_refs(body, map_ref)),
            },
            Expr::Lambda { params, body } => Expr::Lambda {
                params: params.clone(),
                body: Box::new(self.map_ast_refs(body, map_ref)),
            },
            Expr::ArrayLiteral { rows } => Expr::ArrayLiteral {
                rows: rows
                    .iter()
                    .map(|r| r.iter().map(|c| self.map_ast_refs(c, map_ref)).collect())
                    .collect(),
            },
        }
    }
}
