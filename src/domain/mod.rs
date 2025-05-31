//! Domain layer containing core business logic and data structures.
//!
//! This module contains the essential spreadsheet functionality including:
//! - Data models for cells and spreadsheets
//! - Formula evaluation services with comprehensive function support
//! - Expression parser with formal BNF grammar and extensible function registry
//! - All logical operations implemented as functions (AND, OR, NOT)

pub mod models;
pub mod services;
pub mod parser;

pub use models::*;
pub use services::*;