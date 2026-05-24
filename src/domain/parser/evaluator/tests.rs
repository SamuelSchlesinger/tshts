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

