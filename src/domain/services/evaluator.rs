//! Submodule of `services` — see services/mod.rs.

#![allow(unused_imports)]
use super::*;
use crate::domain::models::Spreadsheet;

pub struct FormulaEvaluator<'a> {
    spreadsheet: &'a Spreadsheet,
    /// Optional named-ranges context. Resolution of bare identifiers in
    /// formulas (e.g. `=Revenue + 10`) uses this map; absent → unknown.
    names: Option<&'a HashMap<String, String>>,
    /// Optional workbook for cross-sheet refs (`Sheet2!A1`). When absent,
    /// sheet-qualified refs fail with `#REF!`.
    workbook: Option<&'a Workbook>,
}

impl<'a> FormulaEvaluator<'a> {
    pub fn new(spreadsheet: &'a Spreadsheet) -> Self {
        Self { spreadsheet, names: None, workbook: None }
    }

    /// Variant that resolves bare identifiers via `names`.
    pub fn with_names(
        spreadsheet: &'a Spreadsheet,
        names: &'a HashMap<String, String>,
    ) -> Self {
        Self { spreadsheet, names: Some(names), workbook: None }
    }

    /// Full-context variant that resolves both named ranges and cross-sheet
    /// references. The `spreadsheet` argument is the "current" sheet for
    /// unqualified refs; the workbook is consulted for `Sheet2!A1`.
    pub fn with_workbook(
        workbook: &'a Workbook,
        spreadsheet: &'a Spreadsheet,
        names: &'a HashMap<String, String>,
    ) -> Self {
        Self {
            spreadsheet,
            names: Some(names),
            workbook: Some(workbook),
        }
    }

    /// Returns the evaluated result, "#ERROR" on failure, or `formula`
    /// unchanged if it doesn't start with `=`.
    ///
    /// ```
    /// use tshts::domain::{Spreadsheet, FormulaEvaluator};
    ///
    /// let sheet = Spreadsheet::default();
    /// let evaluator = FormulaEvaluator::new(&sheet);
    /// assert_eq!(evaluator.evaluate_formula("=2+3"), "5");
    /// assert_eq!(evaluator.evaluate_formula("hello"), "hello");
    /// ```
    pub fn evaluate_formula(&self, formula: &str) -> String {
        if formula.starts_with('=') {
            let expr = &formula[1..];
            match self.parse_and_evaluate(expr) {
                Ok(result) => result.to_string(),
                // Unhandled-error fallback: Excel's default is #VALUE! for
                // generic argument issues, but we keep "#ERROR" for the
                // truly-untyped failure mode (parse errors, unknown
                // functions). Heuristic: if the error message implies a
                // recognized error class, surface that; else "#ERROR".
                Err(msg) => classify_err(&msg).to_string(),
            }
        } else {
            formula.to_string()
        }
    }

    /// Walks the AST checking whether `formula` directly or transitively
    /// references `current_cell`.
    ///
    /// ```
    /// use tshts::domain::{Spreadsheet, FormulaEvaluator};
    ///
    /// let sheet = Spreadsheet::default();
    /// let evaluator = FormulaEvaluator::new(&sheet);
    /// assert!(evaluator.would_create_circular_reference("=A1+1", (0, 0)));
    /// assert!(!evaluator.would_create_circular_reference("=AND(B1>0,C1<10)", (0, 0)));
    /// ```
    pub fn would_create_circular_reference(&self, formula: &str, current_cell: (usize, usize)) -> bool {
        if !formula.starts_with('=') {
            return false;
        }
        
        let expr = &formula[1..];
        match Parser::new(expr) {
            Ok(mut parser) => {
                match parser.parse() {
                    Ok(ast) => self.check_circular_reference_in_ast(&ast, current_cell, &mut HashSet::new()),
                    Err(_) => false, // If we can't parse it, assume it's not circular
                }
            }
            Err(_) => false,
        }
    }

    /// Parses and evaluates an expression using the new parser. Uses a
    /// thread-local LRU cache so repeated formulas (common during paste/
    /// autofill) don't re-tokenize.
    fn parse_and_evaluate(&self, expr: &str) -> Result<Value, String> {
        let ast = with_parse_cache(expr, |s| {
            let mut parser = Parser::new(s)?;
            parser.parse()
        })?;
        let function_registry = FunctionRegistry::new();
        let evaluator = match (self.workbook, self.names) {
            (Some(wb), Some(names)) => ExpressionEvaluator::with_workbook(
                wb,
                self.spreadsheet,
                &function_registry,
                names,
            ),
            (None, Some(names)) => {
                ExpressionEvaluator::with_names(self.spreadsheet, &function_registry, names)
            }
            _ => ExpressionEvaluator::new(self.spreadsheet, &function_registry),
        };
        evaluator.evaluate(&ast)
    }
    
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
    fn check_circular_reference_in_ast(&self, expr: &Expr, target_cell: (usize, usize), visited: &mut HashSet<(usize, usize)>) -> bool {
        match expr {
            Expr::String(_) => false,
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
                    if let Some(ref cell_formula) = cell.formula {
                        if self.check_circular_in_formula(cell_formula, target_cell, visited) {
                            return true;
                        }
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
                                if let Some(ref cell_formula) = cell.formula {
                                    if self.check_circular_in_formula(cell_formula, target_cell, visited) {
                                        return true;
                                    }
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
                if let Some(names) = self.names {
                    if let Some(value) = names.get(&name.to_uppercase()).or_else(|| names.get(name)) {
                        if let Ok(mut p) = Parser::new(value) {
                            if let Ok(ast) = p.parse() {
                                return self.check_circular_reference_in_ast(&ast, target_cell, visited);
                            }
                        }
                    }
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
            Expr::String(_) | Expr::Number(_) => {}
            Expr::CellRef(cell_ref) => {
                if let Some((sheet, r, c, _, _)) =
                    Spreadsheet::parse_qualified_reference(cell_ref)
                {
                    out.push((sheet, r, c));
                }
            }
            Expr::Range(start_cell, end_cell) => {
                // 3-D markers (`<S1>..<S3>!<cell>`): expand the cell across
                // each named sheet in the span.
                if let Some((s1, s2, cell)) = Spreadsheet::parse_three_d_marker(start_cell)
                {
                    if let Some((row, col)) = Spreadsheet::parse_cell_reference(&cell) {
                        // Walk the sheet-name list; if the names aren't in
                        // the workbook, we silently skip.
                        // (We don't have workbook ordering info in this
                        // method directly — caller can resolve as needed.)
                        out.push((Some(s1.clone()), row, col));
                        if !s1.eq_ignore_ascii_case(&s2) {
                            out.push((Some(s2), row, col));
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
                if let Some(names) = self.names {
                    if let Some(value) =
                        names.get(&name.to_uppercase()).or_else(|| names.get(name))
                    {
                        if let Ok(mut p) = Parser::new(value) {
                            if let Ok(ast) = p.parse() {
                                self.extract_qualified_refs_from_ast(&ast, out);
                            }
                        }
                    }
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
            Expr::String(_) => {},
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
                if let Some(names) = self.names {
                    if let Some(value) = names.get(&name.to_uppercase()).or_else(|| names.get(name)) {
                        if let Ok(mut p) = Parser::new(value) {
                            if let Ok(ast) = p.parse() {
                                references.extend(self.extract_cell_references_from_ast(&ast));
                            }
                        }
                    }
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
            Ok(mut parser) => {
                match parser.parse() {
                    Ok(ast) => {
                        let adjusted_ast = self.adjust_ast_references(&ast, row_offset, col_offset);
                        format!("={}", self.ast_to_string(&adjusted_ast))
                    }
                    Err(_) => formula.to_string(),
                }
            }
            Err(_) => formula.to_string(),
        }
    }

    /// Adjusts cell references in an AST. Absolute parts (`$col` / `$row`) are
    /// preserved verbatim — only the relative parts shift by the offsets.
    fn adjust_ast_references(&self, expr: &Expr, row_offset: i32, col_offset: i32) -> Expr {
        let adjust = |cell_ref: &str| -> String {
            if let Some((row, col, abs_row, abs_col)) =
                Spreadsheet::parse_cell_reference_with_flags(cell_ref)
            {
                let new_row = if abs_row {
                    row
                } else {
                    (row as i32 + row_offset).max(0) as usize
                };
                let new_col = if abs_col {
                    col
                } else {
                    (col as i32 + col_offset).max(0) as usize
                };
                Spreadsheet::format_cell_reference(new_row, new_col, abs_row, abs_col)
            } else {
                cell_ref.to_string()
            }
        };
        match expr {
            Expr::String(s) => Expr::String(s.clone()),
            Expr::CellRef(cell_ref) => Expr::CellRef(adjust(cell_ref)),
            Expr::Range(start_ref, end_ref) => Expr::Range(adjust(start_ref), adjust(end_ref)),
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

    /// Converts an AST back to a formula string.
    fn ast_to_string(&self, expr: &Expr) -> String {
        match expr {
            Expr::String(s) => format!("\"{}\"", s.replace("\"", "\"\"")),
            Expr::CellRef(cell_ref) => cell_ref.clone(),
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
            if row >= at { (row + 1, col) } else { (row, col) }
        })
    }

    /// Adjusts formula references when a row is deleted at `at`.
    /// References to rows > at are decremented by 1. References to row `at` become #REF.
    pub fn adjust_formula_for_row_delete(&self, formula: &str, at: usize) -> String {
        self.adjust_formula_structural(formula, |row, col| {
            if row == at { (row, col) } // Keep as-is; ideally would be #REF but simplify
            else if row > at { (row - 1, col) }
            else { (row, col) }
        })
    }

    /// Adjusts formula references when a column is inserted at `at`.
    /// References to cols >= at are incremented by 1.
    pub fn adjust_formula_for_col_insert(&self, formula: &str, at: usize) -> String {
        self.adjust_formula_structural(formula, |row, col| {
            if col >= at { (row, col + 1) } else { (row, col) }
        })
    }

    /// Adjusts formula references when a column is deleted at `at`.
    /// References to cols > at are decremented by 1.
    pub fn adjust_formula_for_col_delete(&self, formula: &str, at: usize) -> String {
        self.adjust_formula_structural(formula, |row, col| {
            if col == at { (row, col) }
            else if col > at { (row, col - 1) }
            else { (row, col) }
        })
    }

    /// Remaps row references in a formula using the given old→new row map.
    /// Cells whose row is in the map are remapped; references outside the map
    /// are left alone (so formulas pointing into unsorted regions still work).
    /// Range endpoints are remapped independently.
    pub fn remap_row_references(
        &self,
        formula: &str,
        row_map: &std::collections::HashMap<usize, usize>,
        _max_row: usize,
    ) -> String {
        self.adjust_formula_structural(formula, |row, col| {
            if let Some(&new_row) = row_map.get(&row) {
                (new_row, col)
            } else {
                (row, col)
            }
        })
    }

    /// Generic structural formula adjustment using a mapping function on (row, col).
    fn adjust_formula_structural<F>(&self, formula: &str, map_ref: F) -> String
    where
        F: Fn(usize, usize) -> (usize, usize) + Copy,
    {
        if !formula.starts_with('=') {
            return formula.to_string();
        }
        let expr = &formula[1..];
        match Parser::new(expr) {
            Ok(mut parser) => {
                match parser.parse() {
                    Ok(ast) => {
                        let adjusted = self.map_ast_refs(&ast, map_ref);
                        format!("={}", self.ast_to_string(&adjusted))
                    }
                    Err(_) => formula.to_string(),
                }
            }
            Err(_) => formula.to_string(),
        }
    }

    /// Maps all cell references in an AST using the given function.
    /// Absolute-row/col markers are preserved on output even when remapped.
    fn map_ast_refs<F>(&self, expr: &Expr, map_ref: F) -> Expr
    where
        F: Fn(usize, usize) -> (usize, usize) + Copy,
    {
        let remap = |cell_ref: &str| -> String {
            if let Some((row, col, abs_row, abs_col)) =
                Spreadsheet::parse_cell_reference_with_flags(cell_ref)
            {
                let (nr, nc) = map_ref(row, col);
                Spreadsheet::format_cell_reference(nr, nc, abs_row, abs_col)
            } else {
                cell_ref.to_string()
            }
        };
        match expr {
            Expr::String(s) => Expr::String(s.clone()),
            Expr::Number(n) => Expr::Number(*n),
            Expr::CellRef(cell_ref) => Expr::CellRef(remap(cell_ref)),
            Expr::Range(start_ref, end_ref) => Expr::Range(remap(start_ref), remap(end_ref)),
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
