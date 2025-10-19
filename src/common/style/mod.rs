//! Common style and formatting types.
//!
//! This module provides unified style types used across different Office formats.

// Submodule declarations
pub mod color;
pub mod len;
pub mod text;

// Re-exports
pub use color::RGBColor;
pub use len::Length;
pub use text::VerticalPosition;