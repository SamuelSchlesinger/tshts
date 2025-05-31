//! Application layer managing state and business workflows.
//!
//! This module coordinates between the domain layer and presentation layer,
//! managing application state, user interactions, and business workflows.

pub mod state;

pub use state::*;