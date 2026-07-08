//! Unified error types for Litchi library.
//!
//! This module provides a unified error type that encompasses errors from both
//! OLE2 and OOXML parsing, presenting a consistent API to users.

// Submodule declarations
pub mod conversions;
pub mod types;

// Re-exports
pub use types::{Error, Result};
