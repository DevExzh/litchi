//! OpenDocument Text (.odt) implementation.
//!
//! This module provides comprehensive support for parsing, creating, and manipulating
//! OpenDocument Text documents (.odt files), which are the open standard
//! equivalent of Microsoft Word documents.
//!
//! # Implementation Progress
//!
//! ## ✅ Reading (`document.rs`, `parser.rs`) - COMPLETE
//! - ✅ `Document::open()` - Load from file path
//! - ✅ `Document::from_bytes()` - Load from memory
//! - ✅ `text()` - Extract all text content
//! - ✅ `paragraphs()` - Parse all paragraphs with formatting
//! - ✅ `tables()` - Parse tables with nested structure support
//! - ✅ `lists()` - Parse ordered and unordered lists
//! - ✅ `headings()` - Extract heading hierarchy
//! - ✅ `metadata()` - Extract document metadata
//! - ✅ `hyperlinks()` - Extract all hyperlinks
//! - ✅ `bookmarks()` - Parse bookmarks and references
//! - ✅ `footnotes()` / `endnotes()` - Parse notes
//! - ✅ `comments()` - Parse document comments
//! - ✅ `track_changes()` - Parse tracked changes
//! - ✅ `sections()` - Parse document sections
//! - ✅ Section protection keys, visibility conditions, linked sources, and inert DDE
//! - ✅ `text_indexes()` - Parse all generated index sources and cached bodies
//! - ✅ `text_index_marks()` - Parse TOC, user, alphabetical, and bibliography marks
//! - ✅ `reference_marks()` - Parse point/range cross-reference targets
//! - ✅ `rubies()` - Parse ruby base/pronunciation pairs and styles
//! - ✅ Style parsing and resolution with registry
//!
//! ## ✅ Writing (`builder.rs`, `mutable.rs`) - COMPLETE
//! - ✅ `DocumentBuilder::new()` - Create new documents
//! - ✅ `add_paragraph()` - Add paragraphs with text
//! - ✅ `add_table()` - Add tables with rows/cells
//! - ✅ `add_list()` - Add lists
//! - ✅ `add_heading()` - Add headings with levels
//! - ✅ `set_title()` / `set_author()` - Set metadata
//! - ✅ `save()` / `to_bytes()` - Write to file or bytes
//! - ✅ `MutableDocument` - Modify existing documents
//!
//! ## 🚧 TODO - Advanced Features
//! - ⚠️ Table of contents generation
//! - ⚠️ Index creation (alphabetical, figure, table indexes)
//! - ⚠️ Mail merge and field replacement
//! - ⚠️ Advanced drawing objects (complex shapes, connectors)
//! - ⚠️ Form controls (text fields, checkboxes, dropdowns)
//! - ⚠️ Master page editing
//! - ⚠️ Header and footer manipulation
//! - ⚠️ Page numbering and sections
//! - ⚠️ Document protection
//!
//! # References
//! - ODF Specification: §4-5 (Text Content)
//! - odfpy: `odf/text.py`, `odf/table.py`
//! - ODF Toolkit: Simple API - Document class

mod builder;
mod document;
mod header_footer;
mod header_footer_content;
mod index;
mod index_mark;
mod mutable;
mod note;
pub(crate) mod page_layout;
mod parser;
mod reference_mark;
mod ruby;
mod tracked_changes;

pub use builder::DocumentBuilder;
pub use document::Document;
pub use header_footer::{HeaderFooter, HeaderFooterKind, MasterPage};
#[allow(unused_imports)] // Library public API
pub use header_footer_content::{
    HeaderFooterBlock, HeaderFooterField, HeaderFooterFieldKind, HeaderFooterInline,
};
pub use index::{TextIndex, TextIndexAttribute, TextIndexContent, TextIndexElement, TextIndexKind};
pub use index_mark::{TextIndexMark, TextIndexMarkKind};
pub use mutable::MutableDocument;
pub use note::{Note, NoteClass};
pub use page_layout::{PageLayout, PageLayoutAttribute, PageLayoutProperties, PageUsage};
pub use reference_mark::ReferenceMark;
pub use ruby::Ruby;

// Re-export ODT-specific types for external use
#[allow(unused_imports)] // Library public API
pub use parser::{
    ChangeType, Comment, Section, SectionDdeSource, SectionDisplay, SectionSource, TrackChange,
    TrackedChanges,
};
