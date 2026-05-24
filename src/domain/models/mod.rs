//! Domain models for the terminal spreadsheet application.
//!
//! Split into focused submodules:
//! - style — NumberFormat, TerminalColor, CellStyle, CellFormat, format_cell_value
//! - refs — sheet-name rewriting helpers
//! - cell — CellData
//! - spreadsheet — Spreadsheet struct (the workhorse) plus Table and ConditionalFormat (tightly coupled types)
//! - workbook — Workbook (multi-sheet container) + cross-sheet dependency graph

mod style;
mod refs;
mod cell;
mod spreadsheet;
mod workbook;
mod dep_graph;

pub use style::{NumberFormat, TerminalColor, CellStyle, CellFormat, format_cell_value};
pub use refs::{
    replace_sheet_refs_with_ref_error,
    rewrite_sheet_refs,
    rewrite_sheet_refs_for_name_value,
};
pub use cell::CellData;
pub use spreadsheet::{Spreadsheet, Table, ConditionalFormat, SheetViewState};
pub use workbook::{Workbook, WORKBOOK_SCHEMA_VERSION};
#[allow(unused_imports)]
pub use workbook::CrossSheetKey;
pub use dep_graph::{NodeKey, SheetId, WorkbookGraph};
pub use workbook::with_recalc_context;

#[cfg(test)]
mod tests {
}
