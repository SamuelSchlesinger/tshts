//! Submodule of `parser` — see parser/mod.rs.

#![allow(unused_imports)]
use super::*;

/// Walk an AST and return the maximal [`FunctionPurity`] of any function
/// it transitively calls. A formula is parallel-safe iff its purity is
/// `Pure`, `VolatileClock`, or `VolatileRandom` — the executor uses
/// this to partition a level into parallel-dispatchable cells and
/// serial-only cells.
///
/// Lookup priority for a function call:
/// 1. **Inline structural functions** (INDIRECT, OFFSET, CHOOSE) —
///    hard-coded as `VolatileStructural` because they're handled inline
///    in `ExpressionEvaluator::evaluate` rather than via the registry.
/// 2. **Registry purity** for everything else.
///
/// The walker is conservative: when in doubt about a function's purity,
/// returning the higher-impurity value is always safe (only constrains
/// parallel dispatch more). Returning a lower-impurity value than the
/// truth would be a correctness bug.
pub fn formula_purity(expr: &Expr, registry: &FunctionRegistry) -> FunctionPurity {
    let mut acc = FunctionPurity::Pure;
    walk_purity(expr, registry, &mut acc);
    acc
}

fn walk_purity(expr: &Expr, registry: &FunctionRegistry, acc: &mut FunctionPurity) {
    match expr {
        Expr::Number(_)
        | Expr::String(_)
        | Expr::CellRef(_)
        | Expr::NamedRef(_)
        | Expr::ErrorLit(_) => {}
        Expr::Range(_, _) => {}
        Expr::Binary { left, right, .. } => {
            walk_purity(left, registry, acc);
            walk_purity(right, registry, acc);
        }
        Expr::Unary { operand, .. } => {
            walk_purity(operand, registry, acc);
        }
        Expr::FunctionCall { name, args } => {
            let upper = name.to_uppercase();
            // INDIRECT and OFFSET have inline handling in
            // `ExpressionEvaluator::evaluate` rather than in the
            // registry. The other classical Excel volatile-structural
            // functions (CHOOSE, ROW, COLUMN, FORMULATEXT, CELL, INFO,
            // HYPERLINK) aren't yet implemented in tshts — they'd
            // error before reaching here. Keep the list narrow to
            // reflect actual coverage; widen as those functions land.
            let inline_volatile_structural = matches!(
                upper.as_str(),
                "INDIRECT" | "OFFSET"
            );
            if inline_volatile_structural {
                *acc = acc.join(FunctionPurity::VolatileStructural);
            } else {
                *acc = acc.join(registry.purity(&upper));
            }
            for a in args {
                walk_purity(a, registry, acc);
            }
        }
        Expr::Lambda { body, .. } => walk_purity(body, registry, acc),
        Expr::Let { bindings, body } => {
            for (_, v) in bindings {
                walk_purity(v, registry, acc);
            }
            walk_purity(body, registry, acc);
        }
        Expr::ArrayLiteral { rows } => {
            for row in rows {
                for cell in row {
                    walk_purity(cell, registry, acc);
                }
            }
        }
    }
}

/// Compare two values for ordering. Numbers compare numerically (with the
/// same float-tolerance as `=`); strings compare lexicographically and
/// case-insensitively (Excel convention). Mixed numeric/text uses Excel's
/// ranking: numbers < strings < booleans. Tolerance-based equality means
/// nearly-equal numbers compare `Equal` (so `<` is strict at the same
/// precision used by `=`).
/// Is `name` the case-insensitive identifier `TRUE` or `FALSE`? These
/// are boolean literals when they appear bare (no parens), matching
/// Excel's reserved-word semantics. The dedicated `TRUE()` and
/// `FALSE()` zero-arg function forms still work via the FunctionCall
/// path; this shortcut only kicks in for the bare-identifier form.
fn bare_bool_literal(name: &str) -> Option<bool> {
    match name.to_uppercase().as_str() {
        "TRUE" => Some(true),
        "FALSE" => Some(false),
        _ => None,
    }
}

fn cmp_ord(left: &Value, right: &Value) -> std::cmp::Ordering {
    use std::cmp::Ordering;
    let rank = |v: &Value| -> u8 {
        match v {
            Value::Number(_) => 0,
            Value::String(_) => 1,
            Value::Bool(_) => 2,
            // Lists/arrays implicit-intersect to first element; errors
            // shouldn't reach here (caller propagates first_error).
            _ => 1,
        }
    };
    match (left, right) {
        (Value::Number(a), Value::Number(b)) => {
            if super::numbers_equal(*a, *b) {
                Ordering::Equal
            } else {
                a.partial_cmp(b).unwrap_or(Ordering::Equal)
            }
        }
        (Value::String(a), Value::String(b)) => {
            a.to_lowercase().cmp(&b.to_lowercase())
        }
        (Value::Bool(a), Value::Bool(b)) => a.cmp(b),
        _ => rank(left).cmp(&rank(right)),
    }
}

/// Cap on evaluator recursion. Mirrors `MAX_PARSE_DEPTH` but applies to the
/// evaluator — protects against runaway LAMBDA recursion (`=FIB(50)` where
/// FIB is recursive) and pathological INDIRECT chains.
pub(crate) const MAX_EVAL_DEPTH: u32 = 256;

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
    /// Per-call evaluator-recursion counter. RefCell so `evaluate(&self)` can
    /// bump it.
    depth: std::cell::Cell<u32>,
}

impl<'a> ExpressionEvaluator<'a> {
    /// Single constructor — pass `None` for `names` and/or `workbook` when
    /// the caller doesn't have them. Previously this was three near-identical
    /// constructors that diverged only in which fields they set to `None`.
    pub fn new(
        spreadsheet: &'a Spreadsheet,
        function_registry: &'a FunctionRegistry,
        names: Option<&'a HashMap<String, String>>,
        workbook: Option<&'a crate::domain::models::Workbook>,
    ) -> Self {
        Self {
            spreadsheet,
            function_registry,
            named_ranges: names,
            workbook,
            local_scope: std::cell::RefCell::new(Vec::new()),
            depth: std::cell::Cell::new(0),
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

    /// Look up a name and parse its value into an `Expr`. The value may be
    /// a single cell ref or a range like `A1:B10`. Bare TRUE/FALSE are
    /// handled via `bare_bool_literal` BEFORE callers reach this site —
    /// they need to return a `Value::Bool`, not an `Expr`.
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
        let d = self.depth.get();
        if d >= MAX_EVAL_DEPTH {
            return Err("Formula recursion too deep".to_string());
        }
        self.depth.set(d + 1);
        let result = self.evaluate_inner(expr);
        self.depth.set(d);
        result
    }

    fn evaluate_inner(&self, expr: &Expr) -> Result<Value, String> {
        match expr {
            Expr::Number(value) => Ok(Value::Number(*value)),

            Expr::String(text) => Ok(Value::String(text.clone())),

            // Source-level error literal: surface as Value::Error so the
            // existing first-error propagation in Binary/Unary/FunctionCall
            // cascades it through containing expressions exactly like a
            // runtime-produced error would.
            Expr::ErrorLit(kind) => Ok(Value::Error(*kind)),

            Expr::NamedRef(name) => {
                // Local scope (LET / LAMBDA params) takes priority over
                // workbook-level named ranges.
                if let Some(v) = self.lookup_local(name) {
                    return Ok(v);
                }
                if let Some(b) = bare_bool_literal(name) {
                    return Ok(Value::Bool(b));
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

            // A bare LAMBDA outside an invocation produces a placeholder
            // display value. The underlying formula source is preserved in
            // `cell.formula`, so the lambda survives save/load and can be
            // invoked by name (after being assigned to a named range or
            // used inside a LET). Invoking a lambda stored in another cell
            // via `=A1(5)` is NOT supported — that would require a dedicated
            // `Value::Lambda` variant carrying the AST, which is a larger
            // change. The Excel-equivalent error here would be `#CALC!`;
            // we use a non-error string so the cell display is friendly.
            Expr::Lambda { params, .. } => {
                let summary = if params.is_empty() {
                    "[LAMBDA]".to_string()
                } else {
                    format!("[LAMBDA({})]", params.join(","))
                };
                Ok(Value::String(summary))
            }

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
                let (sheet_opt, row, col, _, _) =
                    crate::domain::models::Spreadsheet::parse_qualified_reference(cell_ref)
                        .ok_or_else(|| format!("Invalid cell reference: {}", cell_ref))?;
                let sheet = self.resolve_sheet(sheet_opt.as_deref())?;
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
                    BinaryOp::Less => Ok(Value::Number(if cmp_ord(&left_val, &right_val) == std::cmp::Ordering::Less { 1.0 } else { 0.0 })),
                    BinaryOp::LessEqual => Ok(Value::Number(if cmp_ord(&left_val, &right_val) != std::cmp::Ordering::Greater { 1.0 } else { 0.0 })),
                    BinaryOp::Greater => Ok(Value::Number(if cmp_ord(&left_val, &right_val) == std::cmp::Ordering::Greater { 1.0 } else { 0.0 })),
                    BinaryOp::GreaterEqual => Ok(Value::Number(if cmp_ord(&left_val, &right_val) != std::cmp::Ordering::Less { 1.0 } else { 0.0 })),
                    BinaryOp::Equal => {
                        // Excel rule: values of different types are never
                        // equal, even when one's string form matches the
                        // other's (e.g. =1="1" is FALSE). Numbers are
                        // compared with a small epsilon; strings ignore
                        // ASCII case to match other string ops.
                        let result = match (&left_val, &right_val) {
                            (Value::Number(l), Value::Number(r)) => numbers_equal(*l, *r),
                            (Value::String(l), Value::String(r)) => l.eq_ignore_ascii_case(r),
                            (Value::Bool(l), Value::Bool(r)) => l == r,
                            // Empty cells (Value::String("")) compare equal
                            // to numeric 0 the way Excel does.
                            (Value::Number(n), Value::String(s))
                            | (Value::String(s), Value::Number(n))
                                if s.is_empty() => *n == 0.0,
                            _ => false,
                        };
                        Ok(Value::Number(if result { 1.0 } else { 0.0 }))
                    }
                    BinaryOp::NotEqual => {
                        let result = match (&left_val, &right_val) {
                            (Value::Number(l), Value::Number(r)) => !numbers_equal(*l, *r),
                            (Value::String(l), Value::String(r)) => !l.eq_ignore_ascii_case(r),
                            (Value::Bool(l), Value::Bool(r)) => l != r,
                            (Value::Number(n), Value::String(s))
                            | (Value::String(s), Value::Number(n))
                                if s.is_empty() => *n != 0.0,
                            _ => true,
                        };
                        Ok(Value::Number(if result { 1.0 } else { 0.0 }))
                    }
                }
            }
            
            Expr::Unary { operator, operand } => {
                let operand_val = self.evaluate(operand)?;

                // Error propagation: an error-carrying operand returns the
                // error directly, same as Binary does. Without this,
                // `=-#REF!` would silently become `0` because to_number()
                // converts Value::Error to 0.0.
                if let Some(e) = operand_val.first_error() {
                    return Ok(Value::Error(e));
                }
                match operator {
                    UnaryOp::Plus => Ok(Value::Number(operand_val.to_number())),
                    UnaryOp::Minus => Ok(Value::Number(-operand_val.to_number())),
                }
            }
            
            Expr::FunctionCall { name, args } => {
                let upper = name.to_uppercase();
                // Short-circuit control-flow functions: evaluate args lazily
                // so e.g. `=IFERROR(1/0, "n/a")` doesn't surface #DIV/0!,
                // and `=IF(A1=0, 0, 1/A1)` doesn't divide when the guard
                // is true. Must run BEFORE evaluate_function_args.
                if let Some(v) = self.try_short_circuit(&upper, args)? {
                    return Ok(v);
                }
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
                if let Some(names) = self.named_ranges
                    && let Some(value) = names
                        .get(&upper)
                        .or_else(|| names.get(name))
                        && let Ok(mut p) = Parser::new(value)
                            && let Ok(Expr::Lambda { params, body }) = p.parse() {
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
                let func = self.function_registry.get_function(name)
                    .ok_or_else(|| format!("Unknown function: {}", name))?;
                let arg_values = self.evaluate_function_args(args)?;
                func(&arg_values)
            }
        }
    }

    /// Lazy dispatch for control-flow functions that must NOT eagerly
    /// evaluate every argument. Returns `Ok(Some(value))` when the call was
    /// dispatched (and the value is the result), `Ok(None)` when this
    /// function name isn't one of the short-circuit cases (so the caller
    /// should fall through to the normal eager dispatch), or `Err` if
    /// dispatch was attempted but failed.
    fn try_short_circuit(&self, upper: &str, args: &[Expr]) -> Result<Option<Value>, String> {
        match upper {
            "IF" => {
                if args.len() != 3 && args.len() != 2 {
                    return Err("IF requires 2 or 3 arguments".to_string());
                }
                let cond = self.evaluate(&args[0])?;
                if let Some(e) = cond.first_error() {
                    return Ok(Some(Value::Error(e)));
                }
                if cond.is_truthy() {
                    Ok(Some(self.evaluate(&args[1])?))
                } else if let Some(else_branch) = args.get(2) {
                    Ok(Some(self.evaluate(else_branch)?))
                } else {
                    Ok(Some(Value::Bool(false)))
                }
            }
            "IFERROR" => {
                if args.len() != 2 {
                    return Err("IFERROR requires 2 arguments".to_string());
                }
                match self.evaluate(&args[0]) {
                    Ok(v) if v.is_error() => Ok(Some(self.evaluate(&args[1])?)),
                    Ok(v) => Ok(Some(v)),
                    Err(_) => Ok(Some(self.evaluate(&args[1])?)),
                }
            }
            "IFNA" => {
                if args.len() != 2 {
                    return Err("IFNA requires 2 arguments".to_string());
                }
                match self.evaluate(&args[0]) {
                    Ok(v) if matches!(v.first_error(), Some(ErrorKind::NA)) => {
                        Ok(Some(self.evaluate(&args[1])?))
                    }
                    Ok(v) => Ok(Some(v)),
                    Err(msg) => {
                        // String-classified #N/A: fall through to fallback.
                        if msg.to_lowercase().contains("not found") {
                            Ok(Some(self.evaluate(&args[1])?))
                        } else {
                            Err(msg)
                        }
                    }
                }
            }
            "IFS" => {
                if args.len() < 2 || args.len() % 2 != 0 {
                    return Err("IFS requires pairs (cond, value), at least one pair".to_string());
                }
                let mut i = 0;
                while i < args.len() {
                    let cond = self.evaluate(&args[i])?;
                    if let Some(e) = cond.first_error() {
                        return Ok(Some(Value::Error(e)));
                    }
                    if cond.is_truthy() {
                        return Ok(Some(self.evaluate(&args[i + 1])?));
                    }
                    i += 2;
                }
                Ok(Some(Value::Error(ErrorKind::NA)))
            }
            "SWITCH" => {
                if args.len() < 3 {
                    return Err("SWITCH requires expr + at least one match pair".to_string());
                }
                let expr_v = self.evaluate(&args[0])?;
                let expr_s = expr_v.to_string();
                let mut i = 1;
                while i + 1 < args.len() {
                    let case = self.evaluate(&args[i])?;
                    if case.to_string() == expr_s {
                        return Ok(Some(self.evaluate(&args[i + 1])?));
                    }
                    i += 2;
                }
                if i < args.len() {
                    Ok(Some(self.evaluate(&args[i])?))
                } else {
                    Ok(Some(Value::Error(ErrorKind::NA)))
                }
            }
            // AND/OR can also short-circuit, though Excel evaluates eagerly.
            // We keep them eager for now (existing registry impl handles them).
            _ => Ok(None),
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
                let rows_f = self.evaluate(&args[0])?.to_number();
                let cols_f = self.evaluate(&args[1])?.to_number();
                if rows_f < 1.0 || cols_f < 1.0 || !rows_f.is_finite() || !cols_f.is_finite() {
                    return Ok(Value::Error(ErrorKind::Num));
                }
                let rows = rows_f as usize;
                let cols = cols_f as usize;
                let total = rows.checked_mul(cols).unwrap_or(usize::MAX);
                if total > super::registry_fns::dynamic_array::MAX_DYNAMIC_ARRAY_CELLS {
                    return Ok(Value::Error(ErrorKind::Num));
                }
                let mut out = Vec::with_capacity(total);
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
    /// Errors if the string can't be parsed as a cell reference. Supports
    /// sheet-qualified addresses (`"Sheet2!A1"`) when a workbook is in scope.
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
        // Qualified `Sheet!A1` resolves through the workbook; unqualified
        // refs read the current sheet. Match Expr::CellRef semantics.
        // Bad reference text is returned as #REF! (Excel-equivalent) so
        // callers can trap it with IFERROR.
        let Some((sheet_opt, row, col, _, _)) =
            Spreadsheet::parse_qualified_reference(&ref_text)
        else {
            return Ok(Value::Error(ErrorKind::Ref));
        };
        let sheet = match self.resolve_sheet(sheet_opt.as_deref()) {
            Ok(s) => s,
            Err(_) => return Ok(Value::Error(ErrorKind::Ref)),
        };
        // Capture the resolved target so the executor can record this
        // as a dynamic dep — without it, a change to the target cell
        // wouldn't trigger this INDIRECT cell to re-evaluate. The
        // workbook auto-seeds VolatileStructural cells whose captured
        // targets overlap the dirty closure.
        crate::domain::parser::push_dynamic_target(sheet_opt.as_deref(), row, col);
        let cell = sheet.get_cell(row, col);
        if let Ok(n) = cell.value.parse::<f64>() {
            Ok(Value::Number(n))
        } else {
            Ok(Value::String(cell.value))
        }
    }

    /// Map an optional sheet name (case-insensitive) to a `&Spreadsheet`.
    /// Returns the current sheet for `None`. Returns `Err` with a typed
    /// message when the name is unknown (or no workbook is in scope).
    /// The single point where the dual "workbook-present / workbook-absent"
    /// modes are handled — every cross-sheet ref path funnels through here
    /// so the conditional doesn't have to be repeated at every call site.
    fn resolve_sheet(&self, sheet_name: Option<&str>) -> Result<&Spreadsheet, String> {
        let Some(name) = sheet_name else { return Ok(self.spreadsheet); };
        let wb = self
            .workbook
            .ok_or_else(|| format!("Cross-sheet ref to '{}' requires workbook context", name))?;
        let idx = wb
            .sheet_names
            .iter()
            .position(|n| n.eq_ignore_ascii_case(name))
            .ok_or_else(|| format!("Unknown sheet: {}", name))?;
        Ok(&wb.sheets[idx])
    }

    /// Walk a Sheet1:Sheet3 range and produce every sheet in `[lo..=hi]`.
    /// Errors if either endpoint is unknown or no workbook is in scope.
    fn resolve_3d_range(&self, s1: &str, s2: &str) -> Result<Vec<&Spreadsheet>, String> {
        let wb = self
            .workbook
            .ok_or_else(|| format!("3-D range '{}:{}' requires workbook context", s1, s2))?;
        let i1 = wb
            .sheet_names
            .iter()
            .position(|n| n.eq_ignore_ascii_case(s1))
            .ok_or_else(|| format!("Unknown sheet: {}", s1))?;
        let i2 = wb
            .sheet_names
            .iter()
            .position(|n| n.eq_ignore_ascii_case(s2))
            .ok_or_else(|| format!("Unknown sheet: {}", s2))?;
        let (lo, hi) = if i1 <= i2 { (i1, i2) } else { (i2, i1) };
        Ok(wb.sheets[lo..=hi].iter().collect())
    }

    /// OFFSET(base, rows, cols, [height], [width]) → value or range starting
    /// at (base.row + rows, base.col + cols) of the given size. Height and
    /// width default to 1. Returns implicit-intersection value when size is 1×1.
    fn eval_offset(&self, args: &[Expr]) -> Result<Value, String> {
        if args.len() < 3 || args.len() > 5 {
            return Err("OFFSET requires 3-5 arguments".to_string());
        }
        // Base must be a CellRef (extract sheet/row/col without evaluating).
        // `parse_qualified_reference` accepts both `A1` and `Sheet2!A1`.
        let parse_base = |r: &str| -> Result<(Option<String>, usize, usize), String> {
            Spreadsheet::parse_qualified_reference(r)
                .map(|(s, row, col, _, _)| (s, row, col))
                .ok_or_else(|| format!("OFFSET: invalid base: {}", r))
        };
        let (base_sheet, base_row, base_col) = match &args[0] {
            Expr::CellRef(r) => parse_base(r)?,
            Expr::NamedRef(name) => {
                let resolved = self.resolve_name(name)?;
                match resolved {
                    Expr::CellRef(r) => parse_base(&r)?,
                    Expr::Range(start, _) => parse_base(&start)?,
                    _ => return Err("OFFSET: base must be a cell ref".to_string()),
                }
            }
            _ => return Err("OFFSET: base must be a cell ref".to_string()),
        };
        let sheet = self.resolve_sheet(base_sheet.as_deref())?;
        // Validate every numeric arg is finite before the lossy `as i64`
        // cast. NaN/Inf would saturate to 0 / huge silently.
        let to_int = |v: Value| -> Result<i64, String> {
            let n = v.to_number();
            if !n.is_finite() {
                return Err("OFFSET: arg must be finite".to_string());
            }
            Ok(n as i64)
        };
        let row_off_raw = self.evaluate(&args[1])?;
        let col_off_raw = self.evaluate(&args[2])?;
        if let Some(e) = row_off_raw.first_error().or_else(|| col_off_raw.first_error()) {
            return Ok(Value::Error(e));
        }
        let row_off = match to_int(row_off_raw) {
            Ok(v) => v,
            Err(_) => return Ok(Value::Error(ErrorKind::Value)),
        };
        let col_off = match to_int(col_off_raw) {
            Ok(v) => v,
            Err(_) => return Ok(Value::Error(ErrorKind::Value)),
        };
        let height = args
            .get(3)
            .map(|e| self.evaluate(e).and_then(|v| to_int(v)))
            .transpose()
            .map_err(|_| "OFFSET: height must be finite".to_string())?
            .unwrap_or(1)
            .max(1);
        let width = args
            .get(4)
            .map(|e| self.evaluate(e).and_then(|v| to_int(v)))
            .transpose()
            .map_err(|_| "OFFSET: width must be finite".to_string())?
            .unwrap_or(1)
            .max(1);
        let new_row_i = base_row as i64 + row_off;
        let new_col_i = base_col as i64 + col_off;
        // Negative offsets and rows/cols beyond the sheet are `#REF!`.
        if new_row_i < 0
            || new_col_i < 0
            || new_row_i as usize >= sheet.rows
            || new_col_i as usize >= sheet.cols
        {
            return Ok(Value::Error(ErrorKind::Ref));
        }
        let new_row = new_row_i as usize;
        let new_col = new_col_i as usize;
        // Also bounds-check the height/width corner.
        if new_row + height as usize > sheet.rows
            || new_col + width as usize > sheet.cols
        {
            return Ok(Value::Error(ErrorKind::Ref));
        }
        // Capture every cell in the offset region as a dynamic dep
        // target. The executor reads these post-eval and stores them
        // on the workbook so the auto-seed can skip this OFFSET cell
        // when none of its targets are in the next recalc's dirty set.
        for r in new_row..new_row + height as usize {
            for c in new_col..new_col + width as usize {
                crate::domain::parser::push_dynamic_target(
                    base_sheet.as_deref(),
                    r,
                    c,
                );
            }
        }
        if height == 1 && width == 1 {
            let cell = sheet.get_cell(new_row, new_col);
            if let Ok(n) = cell.value.parse::<f64>() {
                return Ok(Value::Number(n));
            }
            return Ok(Value::String(cell.value));
        }
        // Build a 2-D Array result so callers that index by (row, col)
        // (e.g. INDEX) see the original shape. A previous version
        // returned a flat List, which made `INDEX(OFFSET(...),r,c)`
        // mis-resolve any non-square block.
        let h = height as usize;
        let w = width as usize;
        let mut data = Vec::with_capacity(h * w);
        for r in new_row..new_row + h {
            for c in new_col..new_col + w {
                let cell = sheet.get_cell(r, c);
                if let Ok(n) = cell.value.parse::<f64>() {
                    data.push(Value::Number(n));
                } else {
                    data.push(Value::String(cell.value));
                }
            }
        }
        Ok(Value::Array { rows: h, cols: w, data })
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
                } else if bare_bool_literal(name).is_some() {
                    // TRUE/FALSE as a function arg — let evaluate() handle
                    // the shortcut via the NamedRef branch. Returning
                    // None falls through to `evaluate(target)` below.
                    None
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
                        let sheets = self.resolve_3d_range(&s1, &s2)?;
                        let (row, col) =
                            crate::domain::models::Spreadsheet::parse_cell_reference(&cell)
                                .ok_or_else(|| format!("Invalid cell: {}", cell))?;
                        let mut list = Vec::with_capacity(sheets.len());
                        for sheet in sheets {
                            let c = sheet.get_cell(row, col);
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
                    let sheet = self.resolve_sheet(sp.0.as_deref())?;
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

#[cfg(test)]
mod tests {
    use super::*;
    use crate::domain::{Spreadsheet, CellData, FormulaEvaluator};
    use crate::domain::parser::*;

    fn create_test_spreadsheet() -> Spreadsheet {
        let mut sheet = Spreadsheet::default();
        sheet.set_cell(0, 0, CellData { value: "10".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        sheet.set_cell(0, 1, CellData { value: "20".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        sheet.set_cell(0, 2, CellData { value: "30".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        sheet.set_cell(1, 0, CellData { value: "5".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        sheet.set_cell(1, 1, CellData { value: "15".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        sheet.set_cell(1, 2, CellData { value: "25".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        sheet
    }


    /// Bare TRUE / FALSE inside a function argument list — the
    /// scenarios were hitting #NAME? on `=VLOOKUP(E2,$A$2:$B$6,2,FALSE)`.
    /// This test pins the function-arg case specifically.
    #[test]
    fn test_bare_true_false_inside_function_args() {
        use crate::domain::services::FormulaEvaluator;
        let mut sheet = crate::domain::Spreadsheet::default();
        sheet.cells.insert((0, 0), CellData {
            value: "apple".to_string(), formula: None, format: None,
            comment: None, spill_anchor: None,
        });
        sheet.cells.insert((0, 1), CellData {
            value: "5".to_string(), formula: None, format: None,
            comment: None, spill_anchor: None,
        });
        sheet.cells.insert((1, 0), CellData {
            value: "banana".to_string(), formula: None, format: None,
            comment: None, spill_anchor: None,
        });
        sheet.cells.insert((1, 1), CellData {
            value: "10".to_string(), formula: None, format: None,
            comment: None, spill_anchor: None,
        });
        let ev = FormulaEvaluator::new(&sheet);
        // VLOOKUP with bare FALSE as the exact-match flag.
        assert_eq!(ev.evaluate_formula(
            "=VLOOKUP(\"apple\",A1:B2,2,FALSE)"), "5",
            "VLOOKUP should accept bare FALSE as the exact-match arg");
        assert_eq!(ev.evaluate_formula(
            "=VLOOKUP(\"banana\",A1:B2,2,FALSE)"), "10");
    }

    /// Bare TRUE / FALSE (no parens) are boolean literals — most users
    /// write `=IF(A1>0, FALSE, TRUE)` without `()` after the bools.
    /// Pre-fix this returned `#NAME?` because TRUE/FALSE registered as
    /// zero-arg functions and bare identifiers became NamedRef which
    /// failed to resolve.
    #[test]
    fn test_bare_true_false_evaluate_as_booleans() {
        use crate::domain::services::FormulaEvaluator;
        let sheet = crate::domain::Spreadsheet::default();
        let ev = FormulaEvaluator::new(&sheet);
        // Direct.
        assert_eq!(ev.evaluate_formula("=TRUE"), "TRUE");
        assert_eq!(ev.evaluate_formula("=FALSE"), "FALSE");
        // Case-insensitive (Excel convention).
        assert_eq!(ev.evaluate_formula("=true"), "TRUE");
        assert_eq!(ev.evaluate_formula("=False"), "FALSE");
        // Inside IF as branch values.
        assert_eq!(ev.evaluate_formula("=IF(1=1, TRUE, FALSE)"), "TRUE");
        assert_eq!(ev.evaluate_formula("=IF(1=2, TRUE, FALSE)"), "FALSE");
        // Boolean arithmetic still works.
        assert_eq!(ev.evaluate_formula("=TRUE+TRUE"), "2");
        assert_eq!(ev.evaluate_formula("=FALSE*5"), "0");
        // Zero-arg function form ALSO still works.
        assert_eq!(ev.evaluate_formula("=TRUE()"), "TRUE");
        assert_eq!(ev.evaluate_formula("=FALSE()"), "FALSE");
    }

    #[test]
    fn test_expression_evaluator_numbers() {
        let sheet = create_test_spreadsheet();
        let registry = FunctionRegistry::new();
        let evaluator = ExpressionEvaluator::new(&sheet, &registry, None, None);
        
        let expr = Expr::Number(42.5);
        match evaluator.evaluate(&expr).unwrap() {
            Value::Number(n) => assert_eq!(n, 42.5),
            _ => panic!("Expected number"),
        }
        
        let expr = Expr::String("Hello".to_string());
        match evaluator.evaluate(&expr).unwrap() {
            Value::String(s) => assert_eq!(s, "Hello"),
            _ => panic!("Expected string"),
        }
    }

    #[test]
    fn test_expression_evaluator_cell_refs() {
        let sheet = create_test_spreadsheet();
        let registry = FunctionRegistry::new();
        let evaluator = ExpressionEvaluator::new(&sheet, &registry, None, None);
        
        let expr = Expr::CellRef("A1".to_string());
        match evaluator.evaluate(&expr).unwrap() {
            Value::Number(n) => assert_eq!(n, 10.0),
            _ => panic!("Expected number"),
        }
        
        let expr = Expr::CellRef("B1".to_string());
        match evaluator.evaluate(&expr).unwrap() {
            Value::Number(n) => assert_eq!(n, 20.0),
            _ => panic!("Expected number"),
        }
    }

    #[test]
    fn test_expression_evaluator_string_cells() {
        let mut sheet = Spreadsheet::default();
        sheet.set_cell(0, 0, CellData { value: "Hello".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        sheet.set_cell(0, 1, CellData { value: "World".to_string(), formula: None, format: None, comment: None, spill_anchor: None });
        sheet.set_cell(0, 2, CellData { value: "123".to_string(), formula: None, format: None, comment: None, spill_anchor: None }); // Number as string
        
        let registry = FunctionRegistry::new();
        let evaluator = ExpressionEvaluator::new(&sheet, &registry, None, None);
        
        // String cell reference
        let expr = Expr::CellRef("A1".to_string());
        match evaluator.evaluate(&expr).unwrap() {
            Value::String(s) => assert_eq!(s, "Hello"),
            _ => panic!("Expected string"),
        }
        
        // Numeric string cell reference  
        let expr = Expr::CellRef("C1".to_string());
        match evaluator.evaluate(&expr).unwrap() {
            Value::Number(n) => assert_eq!(n, 123.0),
            _ => panic!("Expected number"),
        }
    }

    #[test]
    fn test_expression_evaluator_binary_ops() {
        let sheet = create_test_spreadsheet();
        let registry = FunctionRegistry::new();
        let evaluator = ExpressionEvaluator::new(&sheet, &registry, None, None);
        
        let expr = Expr::Binary {
            left: Box::new(Expr::Number(10.0)),
            operator: BinaryOp::Add,
            right: Box::new(Expr::Number(5.0)),
        };
        match evaluator.evaluate(&expr).unwrap() {
            Value::Number(n) => assert_eq!(n, 15.0),
            _ => panic!("Expected number"),
        }
        
        let expr = Expr::Binary {
            left: Box::new(Expr::CellRef("A1".to_string())),
            operator: BinaryOp::Multiply,
            right: Box::new(Expr::CellRef("B1".to_string())),
        };
        match evaluator.evaluate(&expr).unwrap() {
            Value::Number(n) => assert_eq!(n, 200.0), // 10 * 20
            _ => panic!("Expected number"),
        }
    }

    #[test]
    fn test_expression_evaluator_unary_ops() {
        let sheet = create_test_spreadsheet();
        let registry = FunctionRegistry::new();
        let evaluator = ExpressionEvaluator::new(&sheet, &registry, None, None);
        
        let expr = Expr::Unary {
            operator: UnaryOp::Minus,
            operand: Box::new(Expr::Number(5.0)),
        };
        match evaluator.evaluate(&expr).unwrap() {
            Value::Number(n) => assert_eq!(n, -5.0),
            _ => panic!("Expected number"),
        }
        
        // NOT is now a function, not a unary operator
        let expr = Expr::FunctionCall {
            name: "NOT".to_string(),
            args: vec![Expr::Number(0.0)],
        };
        match evaluator.evaluate(&expr).unwrap() {
            Value::Number(n) => assert_eq!(n, 1.0),
            _ => panic!("Expected number"),
        }
    }

    #[test]
    fn test_expression_evaluator_functions() {
        let sheet = create_test_spreadsheet();
        let registry = FunctionRegistry::new();
        let evaluator = ExpressionEvaluator::new(&sheet, &registry, None, None);
        
        let expr = Expr::FunctionCall {
            name: "SUM".to_string(),
            args: vec![
                Expr::CellRef("A1".to_string()),
                Expr::CellRef("B1".to_string()),
            ],
        };
        match evaluator.evaluate(&expr).unwrap() {
            Value::Number(n) => assert_eq!(n, 30.0),
            _ => panic!("Expected number"),
        }
        
        let expr = Expr::FunctionCall {
            name: "IF".to_string(),
            args: vec![
                Expr::Number(1.0),
                Expr::Number(100.0),
                Expr::Number(200.0),
            ],
        };
        match evaluator.evaluate(&expr).unwrap() {
            Value::Number(n) => assert_eq!(n, 100.0),
            _ => panic!("Expected number"),
        }
    }

    #[test]
    fn test_expression_evaluator_string_functions() {
        let sheet = create_test_spreadsheet();
        let registry = FunctionRegistry::new();
        let evaluator = ExpressionEvaluator::new(&sheet, &registry, None, None);
        
        // Test CONCAT function
        let expr = Expr::FunctionCall {
            name: "CONCAT".to_string(),
            args: vec![
                Expr::String("Hello".to_string()),
                Expr::String(" ".to_string()),
                Expr::String("World".to_string()),
            ],
        };
        match evaluator.evaluate(&expr).unwrap() {
            Value::String(s) => assert_eq!(s, "Hello World"),
            _ => panic!("Expected string"),
        }
        
        // Test LEN function
        let expr = Expr::FunctionCall {
            name: "LEN".to_string(),
            args: vec![Expr::String("Hello".to_string())],
        };
        match evaluator.evaluate(&expr).unwrap() {
            Value::Number(n) => assert_eq!(n, 5.0),
            _ => panic!("Expected number"),
        }
        
        // Test UPPER function
        let expr = Expr::FunctionCall {
            name: "UPPER".to_string(),
            args: vec![Expr::String("hello".to_string())],
        };
        match evaluator.evaluate(&expr).unwrap() {
            Value::String(s) => assert_eq!(s, "HELLO"),
            _ => panic!("Expected string"),
        }
        
        // Test LEFT function
        let expr = Expr::FunctionCall {
            name: "LEFT".to_string(),
            args: vec![
                Expr::String("Hello World".to_string()),
                Expr::Number(5.0),
            ],
        };
        match evaluator.evaluate(&expr).unwrap() {
            Value::String(s) => assert_eq!(s, "Hello"),
            _ => panic!("Expected string"),
        }
        
        // Test FIND function
        let expr = Expr::FunctionCall {
            name: "FIND".to_string(),
            args: vec![
                Expr::String("lo".to_string()),
                Expr::String("Hello".to_string()),
            ],
        };
        match evaluator.evaluate(&expr).unwrap() {
            Value::Number(n) => assert_eq!(n, 4.0), // 1-based - "lo" starts at position 4
            _ => panic!("Expected number"),
        }
    }

    #[test]
    fn test_expression_evaluator_string_concatenation() {
        let sheet = create_test_spreadsheet();
        let registry = FunctionRegistry::new();
        let evaluator = ExpressionEvaluator::new(&sheet, &registry, None, None);
        
        let expr = Expr::Binary {
            left: Box::new(Expr::String("Hello".to_string())),
            operator: BinaryOp::Concatenate,
            right: Box::new(Expr::String(" World".to_string())),
        };
        match evaluator.evaluate(&expr).unwrap() {
            Value::String(s) => assert_eq!(s, "Hello World"),
            _ => panic!("Expected string"),
        }
        
        // Test mixed concatenation
        let expr = Expr::Binary {
            left: Box::new(Expr::String("Number: ".to_string())),
            operator: BinaryOp::Concatenate,
            right: Box::new(Expr::Number(42.0)),
        };
        match evaluator.evaluate(&expr).unwrap() {
            Value::String(s) => assert_eq!(s, "Number: 42"),
            _ => panic!("Expected string"),
        }
    }

    #[test]
    fn test_expression_evaluator_string_equality() {
        let sheet = create_test_spreadsheet();
        let registry = FunctionRegistry::new();
        let evaluator = ExpressionEvaluator::new(&sheet, &registry, None, None);
        
        // String equality
        let expr = Expr::Binary {
            left: Box::new(Expr::String("Hello".to_string())),
            operator: BinaryOp::Equal,
            right: Box::new(Expr::String("Hello".to_string())),
        };
        match evaluator.evaluate(&expr).unwrap() {
            Value::Number(n) => assert_eq!(n, 1.0),
            _ => panic!("Expected number"),
        }
        
        // String inequality
        let expr = Expr::Binary {
            left: Box::new(Expr::String("Hello".to_string())),
            operator: BinaryOp::NotEqual,
            right: Box::new(Expr::String("World".to_string())),
        };
        match evaluator.evaluate(&expr).unwrap() {
            Value::Number(n) => assert_eq!(n, 1.0),
            _ => panic!("Expected number"),
        }
    }

    #[test]
    fn test_expression_evaluator_ranges() {
        let sheet = create_test_spreadsheet();
        let registry = FunctionRegistry::new();
        let evaluator = ExpressionEvaluator::new(&sheet, &registry, None, None);
        
        let expr = Expr::FunctionCall {
            name: "SUM".to_string(),
            args: vec![Expr::Range("A1".to_string(), "B1".to_string())],
        };
        match evaluator.evaluate(&expr).unwrap() {
            Value::Number(n) => assert_eq!(n, 30.0),
            _ => panic!("Expected number"),
        }
        
        let expr = Expr::FunctionCall {
            name: "AVERAGE".to_string(),
            args: vec![Expr::Range("A1".to_string(), "C1".to_string())],
        };
        match evaluator.evaluate(&expr).unwrap() {
            Value::Number(n) => assert_eq!(n, 20.0),
            _ => panic!("Expected number"),
        }
    }

    #[test]
    fn test_function_registry() {
        let mut registry = FunctionRegistry::new();
        
        // Test that built-in functions are registered
        assert!(registry.get_function("SUM").is_some());
        assert!(registry.get_function("AVERAGE").is_some());
        assert!(registry.get_function("MIN").is_some());
        assert!(registry.get_function("MAX").is_some());
        assert!(registry.get_function("IF").is_some());
        
        // Test case insensitivity
        assert!(registry.get_function("sum").is_some());
        assert!(registry.get_function("Sum").is_some());
        
        // Test unknown function
        assert!(registry.get_function("UNKNOWN").is_none());
        
        // Test registering custom function
        registry.register_function("DOUBLE", |args| {
            if args.len() == 1 {
                Ok(Value::Number(args[0].to_number() * 2.0))
            } else {
                Err("DOUBLE requires exactly 1 argument".to_string())
            }
        });
        
        assert!(registry.get_function("DOUBLE").is_some());
        let double_func = registry.get_function("DOUBLE").unwrap();
        match double_func(&[Value::Number(5.0)]).unwrap() {
            Value::Number(n) => assert_eq!(n, 10.0),
            _ => panic!("Expected number"),
        }
    }

    #[test]
    fn test_complex_expression_parsing_and_evaluation() {
        let sheet = create_test_spreadsheet();
        let registry = FunctionRegistry::new();
        let evaluator = ExpressionEvaluator::new(&sheet, &registry, None, None);
        
        // Test complex expression: IF(SUM(A1:B1) > 25, MAX(A1:C1), MIN(A1:C1))
        let mut parser = Parser::new("IF(SUM(A1:B1) > 25, MAX(A1:C1), MIN(A1:C1))").unwrap();
        let ast = parser.parse().unwrap();
        let result = evaluator.evaluate(&ast).unwrap();
        
        // SUM(A1:B1) = 10 + 20 = 30, which is > 25, so we take MAX(A1:C1) = 30
        match result {
            Value::Number(n) => assert_eq!(n, 30.0),
            _ => panic!("Expected number"),
        }
        
        // Test arithmetic with functions: SUM(A1:B1) + 5
        let mut parser = Parser::new("SUM(A1:B1) + 5").unwrap();
        let ast = parser.parse().unwrap();
        let result = evaluator.evaluate(&ast).unwrap();
        match result {
            Value::Number(n) => assert_eq!(n, 35.0),
            _ => panic!("Expected number"),
        }
        
        // Test power operations: 2 ** 3 + 1
        let mut parser = Parser::new("2 ** 3 + 1").unwrap();
        let ast = parser.parse().unwrap();
        let result = evaluator.evaluate(&ast).unwrap();
        match result {
            Value::Number(n) => assert_eq!(n, 9.0),
            _ => panic!("Expected number"),
        }
    }

    #[test]
    fn test_error_handling() {
        let sheet = create_test_spreadsheet();
        let registry = FunctionRegistry::new();
        let evaluator = ExpressionEvaluator::new(&sheet, &registry, None, None);
        
        // Division by zero now yields a typed Value::Error(Div0).
        let expr = Expr::Binary {
            left: Box::new(Expr::Number(10.0)),
            operator: BinaryOp::Divide,
            right: Box::new(Expr::Number(0.0)),
        };
        let v = evaluator.evaluate(&expr).unwrap();
        assert_eq!(v, Value::Error(ErrorKind::Div0));
        
        // Test unknown function
        let expr = Expr::FunctionCall {
            name: "UNKNOWN".to_string(),
            args: vec![Expr::Number(5.0)],
        };
        assert!(evaluator.evaluate(&expr).is_err());
        
        // Test invalid cell reference
        let expr = Expr::CellRef("INVALID".to_string());
        assert!(evaluator.evaluate(&expr).is_err());
    }

}
