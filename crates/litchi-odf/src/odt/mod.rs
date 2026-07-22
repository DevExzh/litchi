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
//! - ✅ `ruby_annotations()` / `rubies()` - Parse structure-preserving and simplified ruby pairs
//! - ✅ Style parsing and resolution with registry
//!
//! ## ✅ Writing (`builder.rs`, `mutable.rs`) - COMPLETE
//! - ✅ `DocumentBuilder::new()` - Create new documents
//! - ✅ `add_paragraph()` - Add paragraphs with text
//! - ✅ `add_hyperlink()` / `add_hyperlink_element()` - Add inert simple hyperlinks
//! - ✅ `add_note()` - Author plain-text footnotes and endnotes
//! - ✅ `add_ruby_annotation()` / `add_ruby_style()` - Author typed ruby annotations and styles
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
mod dynamic_text;
pub(crate) mod header_footer;
mod header_footer_content;
mod index;
pub(crate) use index::expanded_attributes;
mod index_mark;
mod mutable;
mod note;
pub(crate) use note::parse_notes;
pub(crate) mod page_layout;
mod parser;
mod reference_mark;
mod ruby;
mod section;
mod tracked_changes;

pub use builder::DocumentBuilder;
pub use document::Document;
#[allow(unused_imports)]
pub use dynamic_text::{
    insert_database_field_xml, insert_dynamic_text_field_xml, remove_database_field_xml,
    remove_dynamic_text_field_xml, replace_database_field_xml, replace_dynamic_text_field_xml,
};
pub use header_footer::{
    HeaderFooter, HeaderFooterKind, MasterPage, MasterPageChild, MasterPageChildKind,
};
#[allow(unused_imports)] // Library public API
pub use header_footer_content::{
    HeaderFooterBlock, HeaderFooterField, HeaderFooterFieldKind, HeaderFooterInline,
};
pub use index::{
    AlphabeticalIndexSource, BibliographyIndexSource, IllustrationIndexSource, ObjectIndexSource,
    TableIndexSource, TableOfContentsSource, TextAlphabeticalIndexEntryTemplate,
    TextAlphabeticalIndexLevel, TextBibliographyEntryTemplate, TextBibliographyEntryToken,
    TextBibliographyType, TextIndex, TextIndexAttribute, TextIndexBody, TextIndexBodyParagraph,
    TextIndexBodyTitle, TextIndexCaptionSequenceFormat, TextIndexChapterDisplay, TextIndexContent,
    TextIndexElement, TextIndexEntryTemplate, TextIndexEntryToken, TextIndexKind, TextIndexScope,
    TextIndexSimpleEntryTemplate, TextIndexSourceStyles, TextIndexTabStop, TextIndexTitleTemplate,
    UserIndexSource, insert_text_index_xml, remove_text_index_xml, replace_text_index_xml,
};
pub use index_mark::{
    TextAlphabeticalMarkMetadata, TextIndexMark, TextIndexMarkFragments, TextIndexMarkKind,
    insert_text_index_mark_xml, remove_text_index_mark_xml, replace_text_index_mark_xml,
};
pub use mutable::MutableDocument;
pub use note::{Note, NoteClass, insert_note_xml, remove_note_xml, replace_note_xml};
pub use page_layout::{PageLayout, PageLayoutAttribute, PageLayoutProperties, PageUsage};
pub use reference_mark::{
    ReferenceMark, ReferenceMarkFragments, insert_reference_mark_xml, remove_reference_mark_xml,
    replace_reference_mark_xml,
};
pub use ruby::Ruby;
pub use section::{
    OdtSectionBlock, add_section_xml, clear_sections_xml, remove_section_xml,
    unwrap_section_xml, update_section_xml, wrap_section_xml,
};
pub use tracked_changes::{
    OdtTrackedPosition, OdtTrackedStory, mark_tracked_change_range_xml,
    mark_tracked_deletion_xml, set_tracked_changes_xml, unmark_tracked_change_xml,
};

// Re-export ODT-specific types for external use
#[allow(unused_imports)] // Library public API
pub use parser::{
    ChangeType, Comment, Section, SectionDdeSource, SectionDisplay, SectionSource, TrackChange,
    TrackedChanges,
};
