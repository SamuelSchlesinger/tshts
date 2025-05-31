//! Infrastructure layer providing external service integrations.
//!
//! This module contains implementations for external concerns like
//! file I/O, persistence, and other system-level operations.

pub mod persistence;

pub use persistence::*;