//! ODF XML element classes.
//!
//! This module provides a comprehensive set of classes for parsing and manipulating
//! OpenDocument XML elements, inspired by odfdo/odfpy libraries.

/// Core element functionality
pub mod element;
/// Namespace handling utilities
pub mod namespace;
/// Text-related elements (paragraphs, spans, headings)
pub mod text;
/// Table-related elements (tables, rows, cells)
pub mod table;
/// Drawing elements (shapes, frames, images)
pub mod draw;
/// Style elements
pub mod style;
/// Metadata elements
pub mod meta;
/// Office document elements
pub mod office;



