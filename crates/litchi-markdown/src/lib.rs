//! Format-agnostic Markdown emission helpers for the Litchi office-formats library.
//!
//! This crate provides the building blocks used by Litchi's higher-level format
//! crates (and the `litchi` umbrella crate) to render Office documents and
//! presentations as Markdown:
//!
//! - The [`ToMarkdown`] trait for converting types to Markdown.
//! - [`MarkdownOptions`] and related enums for configuring the output.
//! - [`escape`] helpers for embedding literal text safely.
//! - Unicode helpers for rendering super- and subscript characters.
//!
//! The crate intentionally has no knowledge of any concrete document format;
//! concrete format adapters currently live in the `litchi` umbrella crate's
//! focused `markdown` module.
//!
//! # Re-exports
//!
//! The most commonly used items are re-exported at the crate root for
//! convenience.
pub mod config;
pub mod escape;
pub mod traits;
pub mod unicode;

pub use config::{FormulaStyle, MarkdownOptions, ScriptStyle, StrikethroughStyle, TableStyle};
pub use traits::ToMarkdown;
