use super::models::Spreadsheet;
use std::collections::HashSet;

pub struct FormulaEvaluator<'a> {
    spreadsheet: &'a Spreadsheet,
}

impl<'a> FormulaEvaluator<'a> {
    pub fn new(spreadsheet: &'a Spreadsheet) -> Self {
        Self { spreadsheet }
    }

    pub fn evaluate_formula(&self, formula: &str) -> String {
        if formula.starts_with('=') {
            let expr = &formula[1..];
            
            match self.evaluate_expression(expr) {
                Ok(result) => result.to_string(),
                Err(_) => "#ERROR".to_string(),
            }
        } else {
            formula.to_string()
        }
    }

    pub fn would_create_circular_reference(&self, formula: &str, current_cell: (usize, usize)) -> bool {
        self.check_circular_reference_recursive(formula, current_cell, &mut HashSet::new())
    }

    fn check_circular_reference_recursive(&self, formula: &str, target_cell: (usize, usize), visited: &mut HashSet<(usize, usize)>) -> bool {
        if !formula.starts_with('=') {
            return false;
        }

        let expr = &formula[1..];
        let referenced_cells = self.extract_cell_references(expr);
        
        for (row, col) in referenced_cells {
            if (row, col) == target_cell {
                return true;
            }
            
            if visited.contains(&(row, col)) {
                continue;
            }
            
            visited.insert((row, col));
            
            let cell = self.spreadsheet.get_cell(row, col);
            if let Some(ref cell_formula) = cell.formula {
                if self.check_circular_reference_recursive(cell_formula, target_cell, visited) {
                    return true;
                }
            }
            
            visited.remove(&(row, col));
        }
        
        false
    }

    fn extract_cell_references(&self, expr: &str) -> Vec<(usize, usize)> {
        let mut references = Vec::new();
        let mut current_token = String::new();
        
        for ch in expr.chars() {
            if ch.is_alphanumeric() {
                current_token.push(ch);
            } else {
                if !current_token.is_empty() {
                    if let Some((row, col)) = Spreadsheet::parse_cell_reference(&current_token) {
                        references.push((row, col));
                    }
                    current_token.clear();
                }
            }
        }
        
        if !current_token.is_empty() {
            if let Some((row, col)) = Spreadsheet::parse_cell_reference(&current_token) {
                references.push((row, col));
            }
        }
        
        references
    }

    fn evaluate_expression(&self, expr: &str) -> Result<f64, String> {
        let expr = expr.trim();
        
        if let Ok(result) = expr.parse::<f64>() {
            return Ok(result);
        }
        
        if let Some((row, col)) = Spreadsheet::parse_cell_reference(expr) {
            return Ok(self.spreadsheet.get_cell_value_for_formula(row, col));
        }
        
        if expr.to_uppercase().starts_with("SUM(") || expr.to_uppercase().starts_with("AVERAGE(") || 
           expr.to_uppercase().starts_with("MIN(") || expr.to_uppercase().starts_with("MAX(") ||
           expr.to_uppercase().starts_with("IF(") || expr.to_uppercase().starts_with("AND(") ||
           expr.to_uppercase().starts_with("OR(") || expr.to_uppercase().starts_with("NOT(") {
            return self.evaluate_function(expr);
        }
        
        if expr.contains("**") {
            let parts: Vec<&str> = expr.split("**").collect();
            if parts.len() == 2 {
                let a = self.evaluate_operand(parts[0].trim())?;
                let b = self.evaluate_operand(parts[1].trim())?;
                return Ok(a.powf(b));
            }
        }
        
        if expr.contains("^") {
            let parts: Vec<&str> = expr.split("^").collect();
            if parts.len() == 2 {
                let a = self.evaluate_operand(parts[0].trim())?;
                let b = self.evaluate_operand(parts[1].trim())?;
                return Ok(a.powf(b));
            }
        }
        
        if expr.contains("%") {
            let parts: Vec<&str> = expr.split("%").collect();
            if parts.len() == 2 {
                let a = self.evaluate_operand(parts[0].trim())?;
                let b = self.evaluate_operand(parts[1].trim())?;
                if b == 0.0 {
                    return Err("Modulo by zero".to_string());
                }
                return Ok(a % b);
            }
        }
        
        if expr.contains("<=") {
            let parts: Vec<&str> = expr.split("<=").collect();
            if parts.len() == 2 {
                let a = self.evaluate_operand(parts[0].trim())?;
                let b = self.evaluate_operand(parts[1].trim())?;
                return Ok(if a <= b { 1.0 } else { 0.0 });
            }
        }
        
        if expr.contains(">=") {
            let parts: Vec<&str> = expr.split(">=").collect();
            if parts.len() == 2 {
                let a = self.evaluate_operand(parts[0].trim())?;
                let b = self.evaluate_operand(parts[1].trim())?;
                return Ok(if a >= b { 1.0 } else { 0.0 });
            }
        }
        
        if expr.contains("<>") {
            let parts: Vec<&str> = expr.split("<>").collect();
            if parts.len() == 2 {
                let a = self.evaluate_operand(parts[0].trim())?;
                let b = self.evaluate_operand(parts[1].trim())?;
                return Ok(if (a - b).abs() > f64::EPSILON { 1.0 } else { 0.0 });
            }
        }
        
        if expr.contains("<") {
            let parts: Vec<&str> = expr.split("<").collect();
            if parts.len() == 2 {
                let a = self.evaluate_operand(parts[0].trim())?;
                let b = self.evaluate_operand(parts[1].trim())?;
                return Ok(if a < b { 1.0 } else { 0.0 });
            }
        }
        
        if expr.contains(">") {
            let parts: Vec<&str> = expr.split(">").collect();
            if parts.len() == 2 {
                let a = self.evaluate_operand(parts[0].trim())?;
                let b = self.evaluate_operand(parts[1].trim())?;
                return Ok(if a > b { 1.0 } else { 0.0 });
            }
        }
        
        if expr.contains('+') {
            let parts: Vec<&str> = expr.split('+').collect();
            if parts.len() == 2 {
                let a = self.evaluate_operand(parts[0].trim())?;
                let b = self.evaluate_operand(parts[1].trim())?;
                return Ok(a + b);
            }
        }
        
        if expr.contains('-') && !expr.starts_with('-') {
            let parts: Vec<&str> = expr.split('-').collect();
            if parts.len() == 2 {
                let a = self.evaluate_operand(parts[0].trim())?;
                let b = self.evaluate_operand(parts[1].trim())?;
                return Ok(a - b);
            }
        }
        
        if expr.contains('*') && !expr.contains("**") {
            let parts: Vec<&str> = expr.split('*').collect();
            if parts.len() == 2 {
                let a = self.evaluate_operand(parts[0].trim())?;
                let b = self.evaluate_operand(parts[1].trim())?;
                return Ok(a * b);
            }
        }
        
        if expr.contains('/') {
            let parts: Vec<&str> = expr.split('/').collect();
            if parts.len() == 2 {
                let a = self.evaluate_operand(parts[0].trim())?;
                let b = self.evaluate_operand(parts[1].trim())?;
                if b == 0.0 {
                    return Err("Division by zero".to_string());
                }
                return Ok(a / b);
            }
        }
        
        Err("Invalid expression".to_string())
    }

    fn evaluate_function(&self, expr: &str) -> Result<f64, String> {
        let expr = expr.trim().to_uppercase();
        
        if expr.starts_with("SUM(") && expr.ends_with(')') {
            let args = &expr[4..expr.len()-1];
            let values = self.parse_function_args(args)?;
            return Ok(values.iter().sum());
        }
        
        if expr.starts_with("AVERAGE(") && expr.ends_with(')') {
            let args = &expr[8..expr.len()-1];
            let values = self.parse_function_args(args)?;
            if values.is_empty() {
                return Err("No values for average".to_string());
            }
            return Ok(values.iter().sum::<f64>() / values.len() as f64);
        }
        
        if expr.starts_with("MIN(") && expr.ends_with(')') {
            let args = &expr[4..expr.len()-1];
            let values = self.parse_function_args(args)?;
            if values.is_empty() {
                return Err("No values for min".to_string());
            }
            return Ok(values.iter().fold(f64::INFINITY, |a, &b| a.min(b)));
        }
        
        if expr.starts_with("MAX(") && expr.ends_with(')') {
            let args = &expr[4..expr.len()-1];
            let values = self.parse_function_args(args)?;
            if values.is_empty() {
                return Err("No values for max".to_string());
            }
            return Ok(values.iter().fold(f64::NEG_INFINITY, |a, &b| a.max(b)));
        }
        
        if expr.starts_with("IF(") && expr.ends_with(')') {
            let args = &expr[3..expr.len()-1];
            let parts: Vec<&str> = args.split(',').collect();
            if parts.len() == 3 {
                let condition = self.evaluate_operand(parts[0].trim())?;
                let true_val = self.evaluate_operand(parts[1].trim())?;
                let false_val = self.evaluate_operand(parts[2].trim())?;
                return Ok(if condition != 0.0 { true_val } else { false_val });
            }
        }
        
        if expr.starts_with("AND(") && expr.ends_with(')') {
            let args = &expr[4..expr.len()-1];
            let values = self.parse_function_args(args)?;
            return Ok(if values.iter().all(|&x| x != 0.0) { 1.0 } else { 0.0 });
        }
        
        if expr.starts_with("OR(") && expr.ends_with(')') {
            let args = &expr[3..expr.len()-1];
            let values = self.parse_function_args(args)?;
            return Ok(if values.iter().any(|&x| x != 0.0) { 1.0 } else { 0.0 });
        }
        
        if expr.starts_with("NOT(") && expr.ends_with(')') {
            let args = &expr[4..expr.len()-1];
            let value = self.evaluate_operand(args.trim())?;
            return Ok(if value == 0.0 { 1.0 } else { 0.0 });
        }
        
        Err("Unknown function".to_string())
    }

    fn parse_function_args(&self, args: &str) -> Result<Vec<f64>, String> {
        let mut values = Vec::new();
        for arg in args.split(',') {
            let arg = arg.trim();
            if arg.contains(':') {
                let range_values = self.parse_range(arg)?;
                values.extend(range_values);
            } else {
                values.push(self.evaluate_operand(arg)?);
            }
        }
        Ok(values)
    }

    fn parse_range(&self, range: &str) -> Result<Vec<f64>, String> {
        let parts: Vec<&str> = range.split(':').collect();
        if parts.len() != 2 {
            return Err("Invalid range".to_string());
        }
        
        let start = Spreadsheet::parse_cell_reference(parts[0].trim())
            .ok_or("Invalid start cell")?;
        let end = Spreadsheet::parse_cell_reference(parts[1].trim())
            .ok_or("Invalid end cell")?;
        
        let mut values = Vec::new();
        for row in start.0..=end.0 {
            for col in start.1..=end.1 {
                values.push(self.spreadsheet.get_cell_value_for_formula(row, col));
            }
        }
        
        Ok(values)
    }

    fn evaluate_operand(&self, operand: &str) -> Result<f64, String> {
        let operand = operand.trim();
        
        if let Ok(num) = operand.parse::<f64>() {
            return Ok(num);
        }
        
        if let Some((row, col)) = Spreadsheet::parse_cell_reference(operand) {
            return Ok(self.spreadsheet.get_cell_value_for_formula(row, col));
        }
        
        Err("Invalid operand".to_string())
    }
}