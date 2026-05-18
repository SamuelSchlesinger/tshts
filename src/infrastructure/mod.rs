//! Infrastructure layer providing external service integrations.

pub mod persistence;
pub mod fetcher;
pub mod recent;
pub mod autosave;
pub mod sidecar;
pub mod xlsx;

pub use persistence::*;
