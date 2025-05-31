#[derive(Debug, Clone, PartialEq)]
pub enum DomainError {
    InvalidCellReference(String),
    CircularReference,
    FormulaEvaluationError(String),
    InvalidFormula(String),
}

impl std::fmt::Display for DomainError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            DomainError::InvalidCellReference(ref_str) => {
                write!(f, "Invalid cell reference: {}", ref_str)
            }
            DomainError::CircularReference => {
                write!(f, "Circular reference detected")
            }
            DomainError::FormulaEvaluationError(msg) => {
                write!(f, "Formula evaluation error: {}", msg)
            }
            DomainError::InvalidFormula(formula) => {
                write!(f, "Invalid formula: {}", formula)
            }
        }
    }
}

impl std::error::Error for DomainError {}

pub type DomainResult<T> = Result<T, DomainError>;