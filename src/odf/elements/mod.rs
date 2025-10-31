//! ODF XML element classes.
//!
//! This module provides a comprehensive set of classes for parsing and manipulating
//! OpenDocument XML elements, inspired by odfdo/odfpy libraries.

/// Bookmark elements for marking locations in documents
pub mod bookmark;
/// Drawing elements (shapes, frames, images)
pub mod draw;
/// Core element functionality
pub mod element;
/// Field elements for dynamic content
pub mod field;
/// Metadata elements
pub mod meta;
/// Namespace handling utilities
pub mod namespace;
/// Office document elements
pub mod office;
/// Generic ODF document parser (shared across ODT/ODS/ODP)
pub mod parser;
/// Style elements
pub mod style;
/// Table-related elements (tables, rows, cells)
pub mod table;
/// Table expansion utilities for repeated cells/rows
pub mod table_expansion;
/// Text-related elements (paragraphs, spans, headings)
pub mod text;
