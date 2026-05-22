//! Submodule of `models` — see models/mod.rs.

#![allow(unused_imports)]
use serde::{Deserialize, Serialize};
use std::collections::{HashMap, HashSet, VecDeque};
use super::*;

/// Represents the data contained within a single spreadsheet cell.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[derive(Default)]
pub struct CellData {
    /// The display value of the cell (either user input or formula result)
    pub value: String,
    /// Optional formula that generates the value (starts with '=')
    pub formula: Option<String>,
    /// Optional cell format
    pub format: Option<CellFormat>,
    /// Optional cell comment
    pub comment: Option<String>,
    /// When set, this cell is a SPILL ghost — its value derives from the
    /// anchor cell's array result. Edits land at the anchor; clearing the
    /// anchor sweeps these up. `None` for normal cells and the anchor itself.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub spill_anchor: Option<(usize, usize)>,
}


impl CellData {
    /// True if this cell is a spill ghost (derived from another cell's
    /// array formula). Ghosts are read-only and cleared when the anchor
    /// changes.
    #[allow(dead_code)]
    pub fn is_spill_ghost(&self) -> bool {
        self.spill_anchor.is_some()
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::domain::{CellData};
    #[test]
    fn test_cell_data_default() {
        let cell = CellData::default();
        assert!(cell.value.is_empty());
        assert!(cell.formula.is_none());
    }

    #[test]
    fn test_cell_data_creation() {
        let cell = CellData {
            value: "42".to_string(),
            formula: Some("=6*7".to_string()),
            format: None,
            comment: None,
        spill_anchor: None,
        };
        assert_eq!(cell.value, "42");
        assert_eq!(cell.formula.unwrap(), "=6*7");
    }

}
