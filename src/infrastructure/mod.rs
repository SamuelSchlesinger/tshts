//! Infrastructure layer providing external service integrations.
//!
//! This module contains implementations for external concerns like
//! file I/O, persistence, and other system-level operations.

pub mod atomic;
pub mod persistence;
pub mod fetcher;
pub mod recent;
pub mod autosave;
pub mod sidecar;
pub mod xlsx;

pub use persistence::*;
