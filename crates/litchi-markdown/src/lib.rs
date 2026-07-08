//! Format-agnostic Markdown emission helpers for the Litchi office-formats library.
//!
//! This crate provides the building blocks used by Litchi's higher-level format
//! crates (and the `litchi` umbrella crate) to render Office documents and
//! presentations as Markdown:
//!
//! - The [`ToMarkdown`] trait for converting types to Markdown.
//! - [`MarkdownOptions`] and related enums for configuring the output.
//! - Unicode helpers for rendering super- and subscript characters.
//!
//! The crate intentionally has no knowledge of any concrete document format;
//! per-format `impl ToMarkdown for ...` blocks live alongside their respective
//! format crates.
//!
//! # Re-exports
//!
//! The most commonly used items are re-exported at the crate root for
//! convenience.
pub mod config;
pub mod traits;
pub mod unicode;

pub use config::{FormulaStyle, MarkdownOptions, ScriptStyle, StrikethroughStyle, TableStyle};
pub use traits::ToMarkdown;
