//! Submodule of `parser` — see parser/mod.rs.

#![allow(unused_imports)]
use super::*;

pub struct ExpressionEvaluator<'a> {
    spreadsheet: &'a Spreadsheet,
    function_registry: &'a FunctionRegistry,
    named_ranges: Option<&'a HashMap<String, String>>,
    /// Workbook for cross-sheet refs. When absent, sheet-qualified refs
    /// return `#REF!`.
    workbook: Option<&'a crate::domain::models::Workbook>,
    /// Stack of local-binding scopes for LET / LAMBDA. Innermost is last.
    /// Uses interior mutability so `evaluate(&self)` can push/pop.
    local_scope: std::cell::RefCell<Vec<HashMap<String, Value>>>,
}

impl<'a> ExpressionEvaluator<'a> {
    pub fn new(spreadsheet: &'a Spreadsheet, function_registry: &'a FunctionRegistry) -> Self {
        Self {
            spreadsheet,
            function_registry,
            named_ranges: None,
            workbook: None,
            local_scope: std::cell::RefCell::new(Vec::new()),
        }
    }

    pub fn with_names(
        spreadsheet: &'a Spreadsheet,
        function_registry: &'a FunctionRegistry,
        names: &'a HashMap<String, String>,
    ) -> Self {
        Self {
            spreadsheet,
            function_registry,
            named_ranges: Some(names),
            workbook: None,
            local_scope: std::cell::RefCell::new(Vec::new()),
        }
    }

    pub fn with_workbook(
        workbook: &'a crate::domain::models::Workbook,
        spreadsheet: &'a Spreadsheet,
        function_registry: &'a FunctionRegistry,
        names: &'a HashMap<String, String>,
    ) -> Self {
        Self {
            spreadsheet,
            function_registry,
            named_ranges: Some(names),
            workbook: Some(workbook),
            local_scope: std::cell::RefCell::new(Vec::new()),
        }
    }

    /// Look up `name` in the local scope (innermost first).
    fn lookup_local(&self, name: &str) -> Option<Value> {
        let scope = self.local_scope.borrow();
        for frame in scope.iter().rev() {
            if let Some(v) = frame.get(name).or_else(|| frame.get(&name.to_uppercase())) {
                return Some(v.clone());
            }
        }
        None
    }

    /// Resolve a possibly sheet-qualified cell ref to the addressed sheet.
    /// Returns the "current" sheet for unqualified refs. Currently used only
    /// for cross-sheet INDIRECT/OFFSET work; kept here for symmetry.
    #[allow(dead_code)]
    fn sheet_for_ref(&self, cell_ref: &str) -> &Spreadsheet {
        if let (Some(wb), Some((Some(sheet_name), _, _, _, _))) = (
            self.workbook,
            crate::domain::models::Spreadsheet::parse_qualified_reference(cell_ref),
        ) {
            if let Some(idx) = wb.sheet_names.iter().position(|n| n == &sheet_name) {
                return &wb.sheets[idx];
            }
        }
        self.spreadsheet
    }

    /// Look up a name and parse its value into an `Expr`. The value may be
    /// a single cell ref or a range like `A1:B10`.
    fn resolve_name(&self, name: &str) -> Result<Expr, String> {
        let names = self
            .named_ranges
            .ok_or_else(|| format!("Unknown identifier: {}", name))?;
        let value = names
            .get(&name.to_uppercase())
            .or_else(|| names.get(name))
            .ok_or_else(|| format!("Unknown name: {}", name))?;
        let mut parser = Parser::new(value)?;
        parser.parse()
    }
    
    /// Evaluates an expression AST to a value result.
    pub fn evaluate(&self, expr: &Expr) -> Result<Value, String> {
        match expr {
            Expr::Number(value) => Ok(Value::Number(*value)),

            Expr::String(text) => Ok(Value::String(text.clone())),

            Expr::NamedRef(name) => {
                // Local scope (LET / LAMBDA params) takes priority over
                // workbook-level named ranges.
                if let Some(v) = self.lookup_local(name) {
                    return Ok(v);
                }
                let resolved = self.resolve_name(name)?;
                self.evaluate(&resolved)
            }

            Expr::Let { bindings, body } => {
                let mut frame: HashMap<String, Value> = HashMap::new();
                for (name, value_expr) in bindings {
                    // Bindings see earlier siblings — push the frame
                    // incrementally so later RHS expressions can refer to
                    // earlier ones (Excel LET semantics).
                    self.local_scope.borrow_mut().push(frame.clone());
                    let v = self.evaluate(value_expr);
                    self.local_scope.borrow_mut().pop();
                    let v = v?;
                    frame.insert(name.to_uppercase(), v);
                }
                self.local_scope.borrow_mut().push(frame);
                let result = self.evaluate(body);
                self.local_scope.borrow_mut().pop();
                result
            }

            // A bare LAMBDA outside an invocation evaluates to itself as a
            // single-cell value. The string form is its formula source so
            // round-trip through named ranges works.
            Expr::Lambda { .. } => Ok(Value::String(format!("[LAMBDA]"))),

            Expr::ArrayLiteral { rows } => {
                let r = rows.len();
                let c = rows.first().map(|r| r.len()).unwrap_or(0);
                if r == 0 || c == 0 {
                    return Ok(Value::Array { rows: 0, cols: 0, data: Vec::new() });
                }
                let mut data = Vec::with_capacity(r * c);
                for row in rows {
                    for cell in row {
                        data.push(self.evaluate(cell)?);
                    }
                }
                Ok(Value::Array { rows: r, cols: c, data })
            }

            Expr::CellRef(cell_ref) => {
                let parsed = crate::domain::models::Spreadsheet::parse_qualified_reference(cell_ref)
                    .ok_or_else(|| format!("Invalid cell reference: {}", cell_ref))?;
                let (sheet_opt, row, col, _, _) = parsed;
                let sheet = if let Some(name) = sheet_opt {
                    let wb = self.workbook.ok_or_else(|| {
                        format!("Cross-sheet ref {} but no workbook context", cell_ref)
                    })?;
                    // Sheet names are case-insensitive (Excel convention),
                    // and the lexer uppercases unquoted identifiers.
                    let idx = wb
                        .sheet_names
                        .iter()
                        .position(|n| n.eq_ignore_ascii_case(&name))
                        .ok_or_else(|| format!("Unknown sheet: {}", name))?;
                    &wb.sheets[idx]
                } else {
                    self.spreadsheet
                };
                let cell = sheet.get_cell(row, col);
                if let Ok(num) = cell.value.parse::<f64>() {
                    Ok(Value::Number(num))
                } else {
                    Ok(Value::String(cell.value))
                }
            }

            Expr::Range(start_cell, end_cell) => {
                // Expand the range to a Value::Array so it can participate in
                // arithmetic broadcasting (e.g. `=A1:A10 * 2`). Aggregate
                // functions flatten it transparently.
                let arr = self.evaluate_function_args(&[Expr::Range(
                    start_cell.clone(),
                    end_cell.clone(),
                )])?;
                Ok(arr.into_iter().next().unwrap_or(Value::List(Vec::new())))
            }
            
            Expr::Binary { left, operator, right } => {
                let left_val = self.evaluate(left)?;
                let right_val = self.evaluate(right)?;

                // Error propagation: any operand carrying an error returns
                // that error directly. This matches Excel semantics.
                if let Some(e) = left_val.first_error() {
                    return Ok(Value::Error(e));
                }
                if let Some(e) = right_val.first_error() {
                    return Ok(Value::Error(e));
                }
                match operator {
                    BinaryOp::Add => broadcast_binary(&left_val, &right_val, |a, b| {
                        Ok(Value::Number(a.to_number() + b.to_number()))
                    }),
                    BinaryOp::Subtract => broadcast_binary(&left_val, &right_val, |a, b| {
                        Ok(Value::Number(a.to_number() - b.to_number()))
                    }),
                    BinaryOp::Multiply => broadcast_binary(&left_val, &right_val, |a, b| {
                        Ok(Value::Number(a.to_number() * b.to_number()))
                    }),
                    BinaryOp::Divide => broadcast_binary(&left_val, &right_val, |a, b| {
                        let r = b.to_number();
                        if r == 0.0 {
                            Ok(Value::Error(ErrorKind::Div0))
                        } else {
                            Ok(Value::Number(a.to_number() / r))
                        }
                    }),
                    BinaryOp::Modulo => broadcast_binary(&left_val, &right_val, |a, b| {
                        let r = b.to_number();
                        if r == 0.0 {
                            Ok(Value::Error(ErrorKind::Div0))
                        } else {
                            Ok(Value::Number(a.to_number() % r))
                        }
                    }),
                    BinaryOp::Power => broadcast_binary(&left_val, &right_val, |a, b| {
                        Ok(Value::Number(a.to_number().powf(b.to_number())))
                    }),
                    BinaryOp::Concatenate => {
                        let left_str = left_val.to_string();
                        let right_str = right_val.to_string();
                        Ok(Value::String(format!("{}{}", left_str, right_str)))
                    }
                    BinaryOp::Less => {
                        let left_num = left_val.to_number();
                        let right_num = right_val.to_number();
                        Ok(Value::Number(if left_num < right_num { 1.0 } else { 0.0 }))
                    }
                    BinaryOp::LessEqual => {
                        let left_num = left_val.to_number();
                        let right_num = right_val.to_number();
                        Ok(Value::Number(if left_num <= right_num { 1.0 } else { 0.0 }))
                    }
                    BinaryOp::Greater => {
                        let left_num = left_val.to_number();
                        let right_num = right_val.to_number();
                        Ok(Value::Number(if left_num > right_num { 1.0 } else { 0.0 }))
                    }
                    BinaryOp::GreaterEqual => {
                        let left_num = left_val.to_number();
                        let right_num = right_val.to_number();
                        Ok(Value::Number(if left_num >= right_num { 1.0 } else { 0.0 }))
                    }
                    BinaryOp::Equal => {
                        let result = match (&left_val, &right_val) {
                            (Value::Number(l), Value::Number(r)) => numbers_equal(*l, *r),
                            (Value::String(l), Value::String(r)) => l == r,
                            (Value::Bool(l), Value::Bool(r)) => l == r,
                            _ => left_val.to_string() == right_val.to_string(),
                        };
                        Ok(Value::Number(if result { 1.0 } else { 0.0 }))
                    }
                    BinaryOp::NotEqual => {
                        let result = match (&left_val, &right_val) {
                            (Value::Number(l), Value::Number(r)) => !numbers_equal(*l, *r),
                            (Value::String(l), Value::String(r)) => l != r,
                            (Value::Bool(l), Value::Bool(r)) => l != r,
                            _ => left_val.to_string() != right_val.to_string(),
                        };
                        Ok(Value::Number(if result { 1.0 } else { 0.0 }))
                    }
                }
            }
            
            Expr::Unary { operator, operand } => {
                let operand_val = self.evaluate(operand)?;
                
                match operator {
                    UnaryOp::Plus => Ok(Value::Number(operand_val.to_number())),
                    UnaryOp::Minus => Ok(Value::Number(-operand_val.to_number())),
                }
            }
            
            Expr::FunctionCall { name, args } => {
                let upper = name.to_uppercase();
                if upper == "INDIRECT" {
                    return self.eval_indirect(args);
                }
                if upper == "OFFSET" {
                    return self.eval_offset(args);
                }
                // Higher-order lambda helpers: last arg must be a LAMBDA.
                if matches!(upper.as_str(), "MAP" | "REDUCE" | "BYROW" | "BYCOL" | "SCAN" | "MAKEARRAY") {
                    return self.eval_lambda_helper(&upper, args);
                }
                // User-defined LAMBDA stored as a named range:
                // `:name DOUBLE LAMBDA(x, x*2)` then `=DOUBLE(7)` looks up
                // DOUBLE, parses it as a lambda, and invokes with the args.
                if let Some(names) = self.named_ranges {
                    if let Some(value) = names
                        .get(&upper)
                        .or_else(|| names.get(name))
                    {
                        if let Ok(mut p) = Parser::new(value) {
                            if let Ok(Expr::Lambda { params, body }) = p.parse() {
                                if params.len() != args.len() {
                                    return Err(format!(
                                        "{}: expected {} arg(s), got {}",
                                        name,
                                        params.len(),
                                        args.len()
                                    ));
                                }
                                let mut frame: HashMap<String, Value> = HashMap::new();
                                for (p_name, a) in params.iter().zip(args.iter()) {
                                    let v = self.evaluate(a)?;
                                    frame.insert(p_name.to_uppercase(), v);
                                }
                                self.local_scope.borrow_mut().push(frame);
                                let result = self.evaluate(&body);
                                self.local_scope.borrow_mut().pop();
                                return result;
                            }
                        }
                    }
                }
                let func = self.function_registry.get_function(name)
                    .ok_or_else(|| format!("Unknown function: {}", name))?;
                let arg_values = self.evaluate_function_args(args)?;
                func(&arg_values)
            }
        }
    }

    /// Dispatch table for higher-order lambda helpers: MAP, REDUCE, BYROW,
    /// BYCOL, SCAN, MAKEARRAY. Each expects a LAMBDA as the last argument.
    fn eval_lambda_helper(&self, name: &str, args: &[Expr]) -> Result<Value, String> {
        let (params, body) = match args.last() {
            Some(Expr::Lambda { params, body }) => (params.clone(), body.clone()),
            _ => return Err(format!("{}: last argument must be a LAMBDA", name)),
        };
        let invoke = |inputs: Vec<Value>| -> Result<Value, String> {
            if inputs.len() != params.len() {
                return Err(format!("{}: lambda arity mismatch", name));
            }
            let mut frame: HashMap<String, Value> = HashMap::new();
            for (p, v) in params.iter().zip(inputs.iter()) {
                frame.insert(p.to_uppercase(), v.clone());
            }
            self.local_scope.borrow_mut().push(frame);
            let result = self.evaluate(&body);
            self.local_scope.borrow_mut().pop();
            result
        };
        match name {
            "MAP" => {
                // MAP(array1, [array2, ...], lambda) — element-wise.
                let arrays: Vec<_> = args[..args.len() - 1]
                    .iter()
                    .map(|a| self.evaluate(a))
                    .collect::<Result<Vec<_>, _>>()?;
                if arrays.is_empty() {
                    return Err("MAP requires at least one array".to_string());
                }
                let first_shape = shape_of(&arrays[0]);
                let rows = first_shape.0;
                let cols = first_shape.1;
                let mut out = Vec::with_capacity(rows * cols);
                for i in 0..(rows * cols) {
                    let inputs: Vec<Value> = arrays
                        .iter()
                        .map(|a| {
                            let (_, _, d) = shape_of(a);
                            d.get(i).cloned().unwrap_or(Value::String(String::new()))
                        })
                        .collect();
                    out.push(invoke(inputs)?);
                }
                Ok(Value::Array { rows, cols, data: out })
            }
            "REDUCE" => {
                // REDUCE(initial, array, lambda(acc, val))
                if args.len() != 3 {
                    return Err("REDUCE requires 3 arguments".to_string());
                }
                let mut acc = self.evaluate(&args[0])?;
                let arr = self.evaluate(&args[1])?;
                for v in arr.flatten() {
                    acc = invoke(vec![acc, v])?;
                }
                Ok(acc)
            }
            "BYROW" | "BYCOL" => {
                if args.len() != 2 {
                    return Err(format!("{} requires 2 arguments", name));
                }
                let arr = self.evaluate(&args[0])?;
                let (rows, cols, data) = shape_of(&arr);
                if name == "BYROW" {
                    let mut out = Vec::with_capacity(rows);
                    for r in 0..rows {
                        let row: Vec<Value> = (0..cols)
                            .map(|c| data[r * cols + c].clone())
                            .collect();
                        let v = invoke(vec![Value::Array {
                            rows: 1,
                            cols,
                            data: row,
                        }])?;
                        out.push(v);
                    }
                    Ok(Value::Array { rows, cols: 1, data: out })
                } else {
                    let mut out = Vec::with_capacity(cols);
                    for c in 0..cols {
                        let col: Vec<Value> = (0..rows)
                            .map(|r| data[r * cols + c].clone())
                            .collect();
                        let v = invoke(vec![Value::Array {
                            rows,
                            cols: 1,
                            data: col,
                        }])?;
                        out.push(v);
                    }
                    Ok(Value::Array { rows: 1, cols, data: out })
                }
            }
            "SCAN" => {
                // SCAN(initial, array, lambda(acc, val))
                if args.len() != 3 {
                    return Err("SCAN requires 3 arguments".to_string());
                }
                let mut acc = self.evaluate(&args[0])?;
                let arr = self.evaluate(&args[1])?;
                let flat = arr.flatten();
                let mut out = Vec::with_capacity(flat.len());
                for v in flat {
                    acc = invoke(vec![acc.clone(), v])?;
                    out.push(acc.clone());
                }
                let cols = out.len();
                Ok(Value::Array { rows: 1, cols, data: out })
            }
            "MAKEARRAY" => {
                // MAKEARRAY(rows, cols, lambda(r, c))
                if args.len() != 3 {
                    return Err("MAKEARRAY requires 3 arguments".to_string());
                }
                let rows = self.evaluate(&args[0])?.to_number() as usize;
                let cols = self.evaluate(&args[1])?.to_number() as usize;
                let mut out = Vec::with_capacity(rows * cols);
                for r in 0..rows {
                    for c in 0..cols {
                        out.push(invoke(vec![
                            Value::Number((r + 1) as f64),
                            Value::Number((c + 1) as f64),
                        ])?);
                    }
                }
                Ok(Value::Array { rows, cols, data: out })
            }
            _ => unreachable!(),
        }
    }

    /// INDIRECT(ref_text) → value at the cell whose A1-style address is `ref_text`.
    /// Errors if the string can't be parsed as a cell reference.
    fn eval_indirect(&self, args: &[Expr]) -> Result<Value, String> {
        if args.len() != 1 {
            return Err("INDIRECT requires exactly 1 argument".to_string());
        }
        let ref_text = self.evaluate(&args[0])?.to_string();
        // Accept either a single cell ref or a range. Ranges in scalar
        // context degrade to first cell.
        if let Some(colon) = ref_text.find(':') {
            let start = &ref_text[..colon];
            let end = &ref_text[colon + 1..];
            let range_expr = Expr::Range(start.to_string(), end.to_string());
            // Implicit intersection: first cell of range.
            let values = self.evaluate_function_args(&[range_expr])?;
            return Ok(values.into_iter().next().unwrap_or(Value::String(String::new())));
        }
        let (row, col) = Spreadsheet::parse_cell_reference(&ref_text)
            .ok_or_else(|| format!("INDIRECT: invalid reference: {}", ref_text))?;
        let cell = self.spreadsheet.get_cell(row, col);
        if let Ok(n) = cell.value.parse::<f64>() {
            Ok(Value::Number(n))
        } else {
            Ok(Value::String(cell.value))
        }
    }

    /// OFFSET(base, rows, cols, [height], [width]) → value or range starting
    /// at (base.row + rows, base.col + cols) of the given size. Height and
    /// width default to 1. Returns implicit-intersection value when size is 1×1.
    fn eval_offset(&self, args: &[Expr]) -> Result<Value, String> {
        if args.len() < 3 || args.len() > 5 {
            return Err("OFFSET requires 3-5 arguments".to_string());
        }
        // Base must be a CellRef (extract row/col without evaluating).
        let (base_row, base_col) = match &args[0] {
            Expr::CellRef(r) => Spreadsheet::parse_cell_reference(r)
                .ok_or_else(|| format!("OFFSET: invalid base: {}", r))?,
            Expr::NamedRef(name) => {
                let resolved = self.resolve_name(name)?;
                match resolved {
                    Expr::CellRef(r) => Spreadsheet::parse_cell_reference(&r)
                        .ok_or_else(|| format!("OFFSET: invalid base: {}", r))?,
                    Expr::Range(start, _) => Spreadsheet::parse_cell_reference(&start)
                        .ok_or_else(|| format!("OFFSET: invalid base: {}", start))?,
                    _ => return Err("OFFSET: base must be a cell ref".to_string()),
                }
            }
            _ => return Err("OFFSET: base must be a cell ref".to_string()),
        };
        let row_off = self.evaluate(&args[1])?.to_number() as i64;
        let col_off = self.evaluate(&args[2])?.to_number() as i64;
        let height = args
            .get(3)
            .map(|e| self.evaluate(e).map(|v| v.to_number() as i64))
            .transpose()?
            .unwrap_or(1)
            .max(1);
        let width = args
            .get(4)
            .map(|e| self.evaluate(e).map(|v| v.to_number() as i64))
            .transpose()?
            .unwrap_or(1)
            .max(1);
        let new_row_i = base_row as i64 + row_off;
        let new_col_i = base_col as i64 + col_off;
        // Negative offsets and rows/cols beyond the sheet are `#REF!`.
        if new_row_i < 0
            || new_col_i < 0
            || new_row_i as usize >= self.spreadsheet.rows
            || new_col_i as usize >= self.spreadsheet.cols
        {
            return Ok(Value::Error(ErrorKind::Ref));
        }
        let new_row = new_row_i as usize;
        let new_col = new_col_i as usize;
        // Also bounds-check the height/width corner.
        if new_row + height as usize > self.spreadsheet.rows
            || new_col + width as usize > self.spreadsheet.cols
        {
            return Ok(Value::Error(ErrorKind::Ref));
        }
        if height == 1 && width == 1 {
            let cell = self.spreadsheet.get_cell(new_row, new_col);
            if let Ok(n) = cell.value.parse::<f64>() {
                return Ok(Value::Number(n));
            }
            return Ok(Value::String(cell.value));
        }
        // Build a List of the resulting block.
        let mut list = Vec::with_capacity((height * width) as usize);
        for r in new_row..new_row + height as usize {
            for c in new_col..new_col + width as usize {
                let cell = self.spreadsheet.get_cell(r, c);
                if let Ok(n) = cell.value.parse::<f64>() {
                    list.push(Value::Number(n));
                } else {
                    list.push(Value::String(cell.value));
                }
            }
        }
        Ok(Value::List(list))
    }
    
    /// Evaluates function arguments. Ranges produce a single `Value::List`
    /// so range-aware functions (SUMIF, VLOOKUP, ...) can preserve structure.
    /// Scalar aggregate functions (SUM, AVG, ...) call `flatten_args`.
    fn evaluate_function_args(&self, args: &[Expr]) -> Result<Vec<Value>, String> {
        let mut values = Vec::new();
        for arg in args {
            // Resolution priority for bare names:
            // 1. Local scope (LET / LAMBDA params) — use the bound value as-is.
            // 2. Named range — substitute its parsed Expr (lets ranges expand).
            // 3. Fall through to scalar eval (will error if truly unknown).
            let effective = if let Expr::NamedRef(name) = arg {
                if self.lookup_local(name).is_some() {
                    None // handled by evaluate(target)
                } else {
                    Some(self.resolve_name(name)?)
                }
            } else {
                None
            };
            let target = effective.as_ref().unwrap_or(arg);
            match target {
                Expr::Range(start_cell, end_cell) => {
                    // 3-D range (Sheet1:Sheet3!A1) — both endpoints carry the
                    // same `<S1>..<S3>!<cell>` marker. Expand across sheets.
                    if let Some((s1, s2, cell)) =
                        crate::domain::models::Spreadsheet::parse_three_d_marker(start_cell)
                    {
                        let wb = self.workbook.ok_or_else(|| {
                            "3-D range needs workbook context".to_string()
                        })?;
                        let i1 = wb
                            .sheet_names
                            .iter()
                            .position(|n| n.eq_ignore_ascii_case(&s1))
                            .ok_or_else(|| format!("Unknown sheet: {}", s1))?;
                        let i2 = wb
                            .sheet_names
                            .iter()
                            .position(|n| n.eq_ignore_ascii_case(&s2))
                            .ok_or_else(|| format!("Unknown sheet: {}", s2))?;
                        let (lo, hi) = if i1 <= i2 { (i1, i2) } else { (i2, i1) };
                        let (row, col) =
                            crate::domain::models::Spreadsheet::parse_cell_reference(&cell)
                                .ok_or_else(|| format!("Invalid cell: {}", cell))?;
                        let mut list = Vec::new();
                        for i in lo..=hi {
                            let c = wb.sheets[i].get_cell(row, col);
                            if let Ok(n) = c.value.parse::<f64>() {
                                list.push(Value::Number(n));
                            } else {
                                list.push(Value::String(c.value));
                            }
                        }
                        values.push(Value::List(list));
                        continue;
                    }
                    // Regular range. Endpoints may be sheet-qualified;
                    // both must address the same sheet.
                    let sp = crate::domain::models::Spreadsheet::parse_qualified_reference(start_cell)
                        .ok_or_else(|| format!("Invalid cell reference: {}", start_cell))?;
                    let ep = crate::domain::models::Spreadsheet::parse_qualified_reference(end_cell)
                        .ok_or_else(|| format!("Invalid cell reference: {}", end_cell))?;
                    let sheet = if let Some(name) = &sp.0 {
                        let wb = self.workbook.ok_or_else(|| {
                            "Cross-sheet range needs workbook context".to_string()
                        })?;
                        let idx = wb
                            .sheet_names
                            .iter()
                            .position(|n| n.eq_ignore_ascii_case(name))
                            .ok_or_else(|| format!("Unknown sheet: {}", name))?;
                        &wb.sheets[idx]
                    } else {
                        self.spreadsheet
                    };
                    let start = (sp.1, sp.2);
                    let end = (ep.1, ep.2);
                    let r0 = start.0.min(end.0);
                    let r1 = start.0.max(end.0);
                    let c0 = start.1.min(end.1);
                    let c1 = start.1.max(end.1);
                    let rows = r1 - r0 + 1;
                    let cols = c1 - c0 + 1;
                    let mut data = Vec::with_capacity(rows * cols);
                    for row in r0..=r1 {
                        for col in c0..=c1 {
                            let cell = sheet.get_cell(row, col);
                            if let Ok(num) = cell.value.parse::<f64>() {
                                data.push(Value::Number(num));
                            } else {
                                data.push(Value::String(cell.value));
                            }
                        }
                    }
                    values.push(Value::Array { rows, cols, data });
                }
                _ => {
                    values.push(self.evaluate(target)?);
                }
            }
        }
        Ok(values)
    }
}
