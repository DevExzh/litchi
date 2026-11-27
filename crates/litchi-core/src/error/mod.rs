//! Format-neutral error vocabulary shared by Litchi crates.
//!
//! Concrete crates map parser and container failures at their own ownership
//! boundary instead of adding dependency-specific conversions to core.

pub mod types;

// Re-exports
pub use types::{Error, Result};
