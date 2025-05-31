//! Presentation layer handling terminal UI and user input.
//!
//! This module manages the terminal user interface using ratatui,
//! handles keyboard input, and renders the spreadsheet display.

pub mod ui;
pub mod input;

pub use ui::*;
pub use input::*;