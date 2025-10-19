//! Unified error types for Litchi library.
//!
//! This module provides a unified error type that encompasses errors from both
//! OLE2 and OOXML parsing, presenting a consistent API to users.

// Submodule declarations
pub mod types;
pub mod conversions;

// Re-exports
pub use types::{Error, Result};
