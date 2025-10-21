//! ODF XML element classes.
//!
//! This module provides a comprehensive set of classes for parsing and manipulating
//! OpenDocument XML elements, inspired by odfdo/odfpy libraries.

/// Drawing elements (shapes, frames, images)
pub mod draw;
/// Core element functionality
pub mod element;
/// Metadata elements
pub mod meta;
/// Namespace handling utilities
pub mod namespace;
/// Office document elements
pub mod office;
/// Style elements
pub mod style;
/// Table-related elements (tables, rows, cells)
pub mod table;
/// Text-related elements (paragraphs, spans, headings)
pub mod text;
