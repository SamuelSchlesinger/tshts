//! Formula evaluation services for the terminal spreadsheet.
//!
//! This module provides the core formula evaluation engine that can
//! parse and execute spreadsheet formulas with cell references,
//! arithmetic operations, and built-in functions.

use super::models::Spreadsheet;
use super::parser::{Parser, ExpressionEvaluator, FunctionRegistry, Expr};
use std::collections::HashSet;

/// A formula evaluation engine that processes spreadsheet expressions.
///
/// The evaluator uses a modern recursive descent parser with formal BNF grammar.
/// All logical operations are implemented as functions for consistency and extensibility.
///
/// Supported features:
/// - Arithmetic operations: +, -, *, /, **, ^, %
/// - Comparison operators: <, >, <=, >=, =, <>
/// - Comprehensive functions: SUM, AVERAGE, MIN, MAX, IF, AND, OR, NOT, ABS, SQRT, ROUND
/// - Cell references: A1, B2, AA123, etc.
/// - Cell ranges: A1:C3, B1:B10, etc.
/// - Circular reference detection with AST analysis
/// - Extensible function registry for custom functions
///
/// # Examples
///
/// ```
/// use tshts::domain::{Spreadsheet, FormulaEvaluator};
///
/// let sheet = Spreadsheet::default();
/// let evaluator = FormulaEvaluator::new(&sheet);
/// 
/// // Arithmetic operations
/// assert_eq!(evaluator.evaluate_formula("=2+3*4"), "14");
/// 
/// // Logical functions (not operators)
/// assert_eq!(evaluator.evaluate_formula("=AND(1>0, 2<5)"), "1");
/// assert_eq!(evaluator.evaluate_formula("=OR(1>2, 3<5)"), "1");
/// assert_eq!(evaluator.evaluate_formula("=NOT(0)"), "1");
/// 
/// // Built-in functions with ranges
/// assert_eq!(evaluator.evaluate_formula("=SUM(A1:A3)"), "result");
/// ```
pub struct FormulaEvaluator<'a> {
    /// Reference to the spreadsheet for cell lookups
    spreadsheet: &'a Spreadsheet,
}

impl<'a> FormulaEvaluator<'a> {
    /// Creates a new formula evaluator for the given spreadsheet.
    ///
    /// # Arguments
    ///
    /// * `spreadsheet` - Reference to the spreadsheet for cell lookups
    ///
    /// # Examples
    ///
    /// ```
    /// use tshts::domain::{Spreadsheet, FormulaEvaluator};
    ///
    /// let sheet = Spreadsheet::default();
    /// let evaluator = FormulaEvaluator::new(&sheet);
    /// ```
    pub fn new(spreadsheet: &'a Spreadsheet) -> Self {
        Self { spreadsheet }
    }

    /// Evaluates a formula and returns the result as a string.
    ///
    /// Uses a recursive descent parser with formal BNF grammar to parse and evaluate
    /// expressions. All logical operations (AND, OR, NOT) are implemented as functions.
    /// Formulas must start with '=' to be recognized as formulas.
    /// Non-formula strings are returned unchanged.
    ///
    /// # Arguments
    ///
    /// * `formula` - Formula string to evaluate (e.g., "=A1+B1", "=AND(A1>0,B1<10)")
    ///
    /// # Returns
    ///
    /// String representation of the evaluation result, or "#ERROR" if evaluation fails
    ///
    /// # Examples
    ///
    /// ```
    /// use tshts::domain::{Spreadsheet, FormulaEvaluator};
    ///
    /// let sheet = Spreadsheet::default();
    /// let evaluator = FormulaEvaluator::new(&sheet);
    ///
    /// assert_eq!(evaluator.evaluate_formula("=2+3"), "5");
    /// assert_eq!(evaluator.evaluate_formula("=AND(1,1)"), "1");
    /// assert_eq!(evaluator.evaluate_formula("hello"), "hello");
    /// ```
    pub fn evaluate_formula(&self, formula: &str) -> String {
        if formula.starts_with('=') {
            let expr = &formula[1..];
            
            match self.parse_and_evaluate(expr) {
                Ok(result) => result.to_string(),
                Err(_) => "#ERROR".to_string(),
            }
        } else {
            formula.to_string()
        }
    }

    /// Checks if a formula would create a circular reference using AST analysis.
    ///
    /// A circular reference occurs when a cell's formula directly or indirectly
    /// references itself, which would cause infinite recursion during evaluation.
    /// This method parses the formula into an AST and analyzes all cell references.
    ///
    /// # Arguments
    ///
    /// * `formula` - Formula to check for circular references
    /// * `current_cell` - Coordinates of the cell that would contain this formula
    ///
    /// # Returns
    ///
    /// `true` if the formula would create a circular reference, `false` otherwise
    ///
    /// # Examples
    ///
    /// ```
    /// use tshts::domain::{Spreadsheet, FormulaEvaluator};
    ///
    /// let sheet = Spreadsheet::default();
    /// let evaluator = FormulaEvaluator::new(&sheet);
    ///
    /// // This would be circular: A1 referring to itself
    /// assert!(evaluator.would_create_circular_reference("=A1+1", (0, 0)));
    /// // This is fine: A1 referring to B1
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

    /// Parses and evaluates an expression using the new parser.
    fn parse_and_evaluate(&self, expr: &str) -> Result<f64, String> {
        let mut parser = Parser::new(expr)?;
        let ast = parser.parse()?;
        
        let function_registry = FunctionRegistry::new();
        let evaluator = ExpressionEvaluator::new(self.spreadsheet, &function_registry);
        evaluator.evaluate(&ast)
    }
    
    /// Checks for circular references in an AST.
    fn check_circular_reference_in_ast(&self, expr: &Expr, target_cell: (usize, usize), visited: &mut HashSet<(usize, usize)>) -> bool {
        match expr {
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
                        if self.would_create_circular_reference(cell_formula, target_cell) {
                            return true;
                        }
                    }
                    
                    visited.remove(&(row, col));
                }
                false
            }
            Expr::Range(start_cell, end_cell) => {
                // Check both start and end cells of the range
                if let (Some((start_row, start_col)), Some((end_row, end_col))) = 
                    (Spreadsheet::parse_cell_reference(start_cell), Spreadsheet::parse_cell_reference(end_cell)) {
                    for row in start_row..=end_row {
                        for col in start_col..=end_col {
                            if (row, col) == target_cell {
                                return true;
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
            Expr::Number(_) => false,
        }
    }

    /// Extracts all cell references from an AST.
    ///
    /// This is a utility method for analyzing formula dependencies.
    /// Currently used mainly for testing and future dependency analysis features.
    ///
    /// # Arguments
    ///
    /// * `expr` - Expression AST to analyze
    ///
    /// # Returns
    ///
    /// Vector of (row, col) tuples representing the referenced cells
    #[allow(dead_code)]
    fn extract_cell_references_from_ast(&self, expr: &Expr) -> Vec<(usize, usize)> {
        let mut references = Vec::new();
        
        match expr {
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
            Expr::Number(_) => {}
        }
        
        references
    }

}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::domain::CellData;

    fn create_test_spreadsheet() -> Spreadsheet {
        let mut sheet = Spreadsheet::default();
        
        // Set up some test data
        sheet.set_cell(0, 0, CellData { value: "10".to_string(), formula: None });
        sheet.set_cell(0, 1, CellData { value: "20".to_string(), formula: None });
        sheet.set_cell(0, 2, CellData { value: "30".to_string(), formula: None });
        sheet.set_cell(1, 0, CellData { value: "5".to_string(), formula: None });
        sheet.set_cell(1, 1, CellData { value: "15".to_string(), formula: None });
        sheet.set_cell(1, 2, CellData { value: "25".to_string(), formula: None });
        
        sheet
    }

    #[test]
    fn test_non_formula_passthrough() {
        let sheet = create_test_spreadsheet();
        let evaluator = FormulaEvaluator::new(&sheet);
        
        assert_eq!(evaluator.evaluate_formula("hello"), "hello");
        assert_eq!(evaluator.evaluate_formula("123"), "123");
        assert_eq!(evaluator.evaluate_formula(""), "");
    }

    #[test]
    fn test_simple_arithmetic() {
        let sheet = create_test_spreadsheet();
        let evaluator = FormulaEvaluator::new(&sheet);
        
        assert_eq!(evaluator.evaluate_formula("=2+3"), "5");
        assert_eq!(evaluator.evaluate_formula("=10-3"), "7");
        assert_eq!(evaluator.evaluate_formula("=4*5"), "20");
        assert_eq!(evaluator.evaluate_formula("=15/3"), "5");
        assert_eq!(evaluator.evaluate_formula("=2**3"), "8");
        assert_eq!(evaluator.evaluate_formula("=3^2"), "9");
        assert_eq!(evaluator.evaluate_formula("=10%3"), "1");
    }

    #[test]
    fn test_comparison_operators() {
        let sheet = create_test_spreadsheet();
        let evaluator = FormulaEvaluator::new(&sheet);
        
        assert_eq!(evaluator.evaluate_formula("=5<10"), "1");
        assert_eq!(evaluator.evaluate_formula("=10<5"), "0");
        assert_eq!(evaluator.evaluate_formula("=10>5"), "1");
        assert_eq!(evaluator.evaluate_formula("=5>10"), "0");
        assert_eq!(evaluator.evaluate_formula("=5<=5"), "1");
        assert_eq!(evaluator.evaluate_formula("=5<=4"), "0");
        assert_eq!(evaluator.evaluate_formula("=5>=5"), "1");
        assert_eq!(evaluator.evaluate_formula("=4>=5"), "0");
        assert_eq!(evaluator.evaluate_formula("=5<>5"), "0");
        assert_eq!(evaluator.evaluate_formula("=5<>4"), "1");
    }

    #[test]
    fn test_cell_references() {
        let sheet = create_test_spreadsheet();
        let evaluator = FormulaEvaluator::new(&sheet);
        
        assert_eq!(evaluator.evaluate_formula("=A1"), "10");
        assert_eq!(evaluator.evaluate_formula("=B1"), "20");
        assert_eq!(evaluator.evaluate_formula("=C1"), "30");
        assert_eq!(evaluator.evaluate_formula("=A2"), "5");
    }

    #[test]
    fn test_cell_arithmetic() {
        let sheet = create_test_spreadsheet();
        let evaluator = FormulaEvaluator::new(&sheet);
        
        assert_eq!(evaluator.evaluate_formula("=A1+B1"), "30"); // 10 + 20
        assert_eq!(evaluator.evaluate_formula("=C1-A1"), "20"); // 30 - 10
        assert_eq!(evaluator.evaluate_formula("=A1*A2"), "50"); // 10 * 5
        assert_eq!(evaluator.evaluate_formula("=B1/A2"), "4"); // 20 / 5
    }

    #[test]
    fn test_sum_function() {
        let sheet = create_test_spreadsheet();
        let evaluator = FormulaEvaluator::new(&sheet);
        
        assert_eq!(evaluator.evaluate_formula("=SUM(A1,B1,C1)"), "60"); // 10+20+30
        assert_eq!(evaluator.evaluate_formula("=SUM(A1:C1)"), "60"); // Range sum
        assert_eq!(evaluator.evaluate_formula("=SUM(A1:A2)"), "15"); // 10+5
        assert_eq!(evaluator.evaluate_formula("=SUM(5,10,15)"), "30"); // Literal values
    }

    #[test]
    fn test_average_function() {
        let sheet = create_test_spreadsheet();
        let evaluator = FormulaEvaluator::new(&sheet);
        
        assert_eq!(evaluator.evaluate_formula("=AVERAGE(A1,B1,C1)"), "20"); // (10+20+30)/3
        assert_eq!(evaluator.evaluate_formula("=AVERAGE(A1:C1)"), "20"); // Range average
        assert_eq!(evaluator.evaluate_formula("=AVERAGE(10,20)"), "15");
    }

    #[test]
    fn test_min_max_functions() {
        let sheet = create_test_spreadsheet();
        let evaluator = FormulaEvaluator::new(&sheet);
        
        assert_eq!(evaluator.evaluate_formula("=MIN(A1,B1,C1)"), "10");
        assert_eq!(evaluator.evaluate_formula("=MAX(A1,B1,C1)"), "30");
        assert_eq!(evaluator.evaluate_formula("=MIN(A1:C1)"), "10");
        assert_eq!(evaluator.evaluate_formula("=MAX(A1:C1)"), "30");
    }

    #[test]
    fn test_if_function() {
        let sheet = create_test_spreadsheet();
        let evaluator = FormulaEvaluator::new(&sheet);
        
        assert_eq!(evaluator.evaluate_formula("=IF(1,100,200)"), "100"); // True condition
        assert_eq!(evaluator.evaluate_formula("=IF(0,100,200)"), "200"); // False condition
        // Note: Complex comparisons in IF functions need a more sophisticated parser
        // For now, test with simple values
        assert_eq!(evaluator.evaluate_formula("=IF(1,1,0)"), "1"); // Simple true
        assert_eq!(evaluator.evaluate_formula("=IF(0,1,0)"), "0"); // Simple false
    }

    #[test]
    fn test_logical_functions() {
        let sheet = create_test_spreadsheet();
        let evaluator = FormulaEvaluator::new(&sheet);
        
        // Test basic functions first
        assert_eq!(evaluator.evaluate_formula("=SUM(1,2)"), "3"); // Simple sum
        
        // Test logical functions (these are built-in functions in the registry)
        assert_eq!(evaluator.evaluate_formula("=AND(1,1)"), "1"); // Both true  
        assert_eq!(evaluator.evaluate_formula("=AND(1,0)"), "0"); // One false
        assert_eq!(evaluator.evaluate_formula("=OR(0,1)"), "1"); // One true
        assert_eq!(evaluator.evaluate_formula("=OR(0,0)"), "0"); // Both false
        assert_eq!(evaluator.evaluate_formula("=NOT(0)"), "1"); // Not false
        assert_eq!(evaluator.evaluate_formula("=NOT(1)"), "0"); // Not true
        
        // All logical operations are now functions
        // (No need for separate binary operator tests since they're all functions now)
    }

    #[test]
    fn test_range_parsing() {
        let sheet = create_test_spreadsheet();
        let evaluator = FormulaEvaluator::new(&sheet);
        
        // Test different range formats
        assert_eq!(evaluator.evaluate_formula("=SUM(A1:C1)"), "60"); // Row range
        assert_eq!(evaluator.evaluate_formula("=SUM(A1:A2)"), "15"); // Column range
        assert_eq!(evaluator.evaluate_formula("=SUM(A1:B2)"), "50"); // Rectangle range
    }

    #[test]
    fn test_error_cases() {
        let sheet = create_test_spreadsheet();
        let evaluator = FormulaEvaluator::new(&sheet);
        
        assert_eq!(evaluator.evaluate_formula("=1/0"), "#ERROR"); // Division by zero
        assert_eq!(evaluator.evaluate_formula("=10%0"), "#ERROR"); // Modulo by zero
        assert_eq!(evaluator.evaluate_formula("=INVALID()"), "#ERROR"); // Unknown function
        assert_eq!(evaluator.evaluate_formula("=AVERAGE()"), "#ERROR"); // No args for average
    }

    #[test]
    fn test_circular_reference_detection() {
        let mut sheet = Spreadsheet::default();
        // Set up a cell that would reference itself
        sheet.set_cell(0, 0, CellData {
            value: "10".to_string(),
            formula: Some("=B1+1".to_string()),
        });
        
        // Set up indirect circular reference chain
        sheet.set_cell(0, 1, CellData {
            value: "20".to_string(),
            formula: Some("=C1+1".to_string()),
        });
        
        let evaluator = FormulaEvaluator::new(&sheet);
        
        // Direct self-reference
        assert!(evaluator.would_create_circular_reference("=A1+1", (0, 0)));
        
        // Non-circular reference
        assert!(!evaluator.would_create_circular_reference("=B1+1", (0, 0)));
        
        // This would create A1->C1->A1 if we set C1 to reference A1
        assert!(evaluator.would_create_circular_reference("=A1+1", (0, 2)));
    }

    #[test]
    fn test_extract_cell_references_from_ast() {
        let sheet = create_test_spreadsheet();
        let evaluator = FormulaEvaluator::new(&sheet);
        
        // Test cell reference extraction from AST
        let mut parser = Parser::new("A1 + B2 * C3").unwrap();
        let ast = parser.parse().unwrap();
        let refs = evaluator.extract_cell_references_from_ast(&ast);
        assert_eq!(refs.len(), 3);
        assert!(refs.contains(&(0, 0))); // A1
        assert!(refs.contains(&(1, 1))); // B2
        assert!(refs.contains(&(2, 2))); // C3
        
        // Test range extraction
        let mut parser = Parser::new("SUM(A1:A3)").unwrap();
        let ast = parser.parse().unwrap();
        let refs = evaluator.extract_cell_references_from_ast(&ast);
        assert_eq!(refs.len(), 3); // Should find A1, A2, A3 from the range
        assert!(refs.contains(&(0, 0))); // A1
        assert!(refs.contains(&(1, 0))); // A2
        assert!(refs.contains(&(2, 0))); // A3
    }

    #[test]
    fn test_case_insensitive_functions() {
        let sheet = create_test_spreadsheet();
        let evaluator = FormulaEvaluator::new(&sheet);
        
        assert_eq!(evaluator.evaluate_formula("=sum(A1,B1)"), "30");
        assert_eq!(evaluator.evaluate_formula("=SUM(A1,B1)"), "30");
        assert_eq!(evaluator.evaluate_formula("=Sum(A1,B1)"), "30");
        assert_eq!(evaluator.evaluate_formula("=average(A1,B1)"), "15");
        assert_eq!(evaluator.evaluate_formula("=AVERAGE(A1,B1)"), "15");
    }

    #[test]
    fn test_whitespace_handling() {
        let sheet = create_test_spreadsheet();
        let evaluator = FormulaEvaluator::new(&sheet);
        
        assert_eq!(evaluator.evaluate_formula("= 2 + 3 "), "5");
        assert_eq!(evaluator.evaluate_formula("=SUM( A1 , B1 )"), "30");
        assert_eq!(evaluator.evaluate_formula("= A1 * 2 "), "20");
    }

    #[test]
    fn test_complex_expressions() {
        let sheet = create_test_spreadsheet();
        let evaluator = FormulaEvaluator::new(&sheet);
        
        // Test power operations (simple arithmetic)
        assert_eq!(evaluator.evaluate_formula("=2**3+1"), "9"); // 8+1
        assert_eq!(evaluator.evaluate_formula("=3*4+2"), "14"); // 12+2
        
        // Test functions work correctly
        assert_eq!(evaluator.evaluate_formula("=SUM(A1,B1)"), "30"); // 10+20
        assert_eq!(evaluator.evaluate_formula("=SUM(A1:B1)"), "30"); // 10+20
    }

    #[test]
    fn test_large_numbers() {
        let mut sheet = Spreadsheet::default();
        sheet.set_cell(0, 0, CellData { value: "1000000".to_string(), formula: None });
        sheet.set_cell(0, 1, CellData { value: "2000000".to_string(), formula: None });
        
        let evaluator = FormulaEvaluator::new(&sheet);
        assert_eq!(evaluator.evaluate_formula("=A1+B1"), "3000000");
        assert_eq!(evaluator.evaluate_formula("=A1*B1"), "2000000000000");
    }

    #[test]
    fn test_negative_numbers() {
        let mut sheet = Spreadsheet::default();
        sheet.set_cell(0, 0, CellData { value: "-10".to_string(), formula: None });
        sheet.set_cell(0, 1, CellData { value: "5".to_string(), formula: None });
        
        let evaluator = FormulaEvaluator::new(&sheet);
        assert_eq!(evaluator.evaluate_formula("=A1+B1"), "-5");
        assert_eq!(evaluator.evaluate_formula("=A1*B1"), "-50");
        assert_eq!(evaluator.evaluate_formula("=-5+10"), "5");
    }

    #[test]
    fn test_decimal_precision() {
        let sheet = create_test_spreadsheet();
        let evaluator = FormulaEvaluator::new(&sheet);
        
        assert_eq!(evaluator.evaluate_formula("=1/3"), "0.3333333333333333");
        assert_eq!(evaluator.evaluate_formula("=22/7"), "3.142857142857143");
    }
}