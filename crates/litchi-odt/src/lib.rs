//! `OpenDocument` Text (.odt) implementation.
//!
//! This module provides comprehensive support for parsing, creating, and manipulating
//! `OpenDocument` Text documents (.odt files), which are the open standard
//! equivalent of Microsoft Word documents.
//!
//! # Implementation Progress
//!
//! ## ✅ Reading (`document.rs`, `parser.rs`) - COMPLETE
//! - ✅ `Package::open().and_then(|package| package.document())` - Load from file path
//! - ✅ `Package::from_bytes(&).and_then(|package| package.document())` - Load from memory
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
//! - ✅ `forms()` - Inert classic form controls (`office:forms`): text, textarea,
//!   checkbox, button, combobox/listbox with options, radio, fixed-text, hidden,
//!   and number/date/time controls with ids, labels, current state, and inert
//!   event-listener metadata
//! - ✅ `page_sequence()` - Parse explicit `text:page-sequence` master-page assignments
//! - ✅ Style parsing and resolution with registry
//!
//! ## ✅ Writing (`builder.rs`, `mutable.rs`) - COMPLETE
//! - ✅ `Builder::new()` - Create new documents
//! - ✅ `add_paragraph()` - Add paragraphs with text
//! - ✅ `add_hyperlink()` / `add_hyperlink_element()` - Add inert simple hyperlinks
//! - ✅ `add_note()` - Author plain-text or validated structured footnotes and endnotes
//! - ✅ `add_ruby_annotation()` / `add_ruby_style()` - Author typed ruby annotations and styles
//! - ✅ `add_table()` - Add tables with rows/cells
//! - ✅ `add_list()` - Add lists
//! - ✅ `add_heading()` - Add headings with levels
//! - ✅ `add_control_form()` and family APIs - Author typed forms and controls
//! - ✅ `set_title()` / `set_author()` - Set metadata
//! - ✅ `save()` / `to_bytes()` - Write to file or bytes
//! - ✅ `MutableDocument` - Modify existing documents
//! - ✅ Insert, replace, and remove form controls and form properties
//! - ✅ `set_page_sequence()` - Author, replace, and remove explicit page sequences
//!
//! ## 🚧 TODO - Advanced Features
//! - ⚠️ Table of contents generation
//! - ⚠️ Index creation (alphabetical, figure, table indexes)
//! - ⚠️ Mail merge and field replacement
//! - ⚠️ Advanced drawing objects (complex shapes, connectors)
//! - ⚠️ Master page editing
//! - ⚠️ Header and footer manipulation
//! - ✅ Document protection metadata (`Document::protection()` and
//!   `Document::set_protection()`); policy is typed and inert, not enforced
//!
//! # References
//! - ODF Specification: §4-5 (Text Content)
//! - odfpy: `odf/text.py`, `odf/table.py`
//! - ODF Toolkit: Simple API - Document class

#![forbid(unsafe_code)]
#![allow(
    clippy::allow_attributes_without_reason,
    clippy::arbitrary_source_item_ordering,
    clippy::case_sensitive_file_extension_comparisons,
    clippy::cast_possible_truncation,
    clippy::collapsible_if,
    clippy::default_trait_access,
    clippy::doc_markdown,
    clippy::format_push_string,
    clippy::inline_always,
    clippy::items_after_statements,
    clippy::many_single_char_names,
    clippy::match_same_arms,
    clippy::match_wildcard_for_single_variants,
    clippy::missing_errors_doc,
    clippy::missing_panics_doc,
    clippy::module_inception,
    clippy::module_name_repetitions,
    clippy::must_use_candidate,
    clippy::needless_borrow,
    clippy::needless_pass_by_value,
    clippy::ref_option,
    clippy::redundant_closure_for_method_calls,
    clippy::return_self_not_must_use,
    clippy::shadow_reuse,
    clippy::shadow_same,
    clippy::shadow_unrelated,
    clippy::similar_names,
    clippy::struct_field_names,
    clippy::trivially_copy_pass_by_ref,
    clippy::unnecessary_wraps,
    clippy::unreadable_literal,
    clippy::useless_asref,
    clippy::wildcard_enum_match_arm,
    clippy::wildcard_imports,
    reason = "legacy ODT API shape and naming are compatibility constraints; broad style lints and legacy XML-position truncation remain quarantined"
)]
pub use litchi_odf_common::core;
pub use litchi_odf_common::{constants, coordinates, datatype, namespace};

pub use litchi_odf_common::detect;
pub mod odc;
pub use litchi_odf_common::rdf;

/// Typed values used by ODT table cells.
#[derive(Debug, Clone, PartialEq)]
pub enum CellValue {
    Empty,
    Text(String),
    Number(f64),
    Boolean(bool),
    Date(String),
    Currency(f64, String),
    Percentage(f64),
    Time(String),
}

pub mod auto_mark_file;
pub mod bibliography_configuration;
pub mod body;
pub mod chart_properties;
pub mod content_metadata;
pub mod content_validation;
pub mod dde_connection;
pub mod document_scripts;
pub mod drawing;
pub mod drawing_page_properties;
pub mod embedded;
pub mod flat;
pub mod font_face;
pub mod footnote_separator;
pub mod form;
pub mod generic;
pub mod graphic_properties;
mod graphic_property_specs;
pub mod image_map;
pub mod line_numbering;
pub mod list_label_alignment;
pub mod list_style;
pub mod master_page;
pub mod media;
pub mod metadata;
pub mod notes_configuration;
pub mod outline_style;
pub mod package;
pub mod protection;
pub mod ruby_family;
pub mod section_properties;
pub mod settings;
pub mod streaming;
pub mod style;
pub mod transaction;
pub mod variable_declaration;
pub mod xforms;

pub(crate) use auto_mark_file::AlphabeticalIndexAutoMarkFile;
pub(crate) use core::Metadata;
pub(crate) use document_scripts::Scripts;
pub(crate) use elements::bookmark::{
    BookmarkTarget, insert_bookmark_xml, parse_bookmark_targets, remove_bookmark_xml,
    replace_bookmark_xml,
};
pub(crate) use embedded::Object;
pub(crate) use generic::{Family, FlatDocument, Package as GenericPackage};
pub(crate) use litchi_odf_common::drawing::Part;
pub(crate) use media::Image;
#[cfg(test)]
pub(crate) use odc::{ChartClass, Definition, StyleElement};
pub(crate) use package::annotation::{
    Annotation, AnnotationAnchor, AnnotationInfo, AnnotationUpdate,
};
pub use package::scripts::{ScriptResource, ScriptResourceKind, ScriptResourceSpec};
pub(crate) use ruby_family::{
    insert_ruby_annotation_xml, parse_ruby_annotations, parse_ruby_styles,
    remove_ruby_annotation_xml, remove_ruby_style_xml, replace_ruby_annotation_xml,
    set_ruby_style_xml, wrap_ruby_annotation_xml,
};
pub(crate) use section_properties::SectionStyleProperties;
pub(crate) use section_properties::{BackgroundRepeat, SectionBackgroundImage};
pub(crate) use settings::Settings;
pub(crate) use style::text::TextProperty;
pub mod elements;
#[cfg(test)]
pub(crate) use elements::field::{
    MetaFieldAttribute, MetaFieldElement, MetaFieldNode, NoteBodyContent,
};
pub(crate) use elements::text::Hyperlink;
pub use elements::text::{NumberLabel, SoftPageBreak};
#[cfg(test)]
pub(crate) use elements::text::{TextHyperlinkActuate, TextHyperlinkShow};

mod builder;
mod document;
mod dynamic_text;
pub mod frame;
pub mod header_footer;
pub mod index;
#[cfg(test)]
mod internal_tests;
pub(crate) use index::expanded_attributes;
pub mod index_mark;
pub(crate) mod mutable;
pub mod note;
pub(crate) use note::parse_notes;
pub mod page_layout;
pub mod page_sequence;
pub mod parser;
pub mod reference_mark;
pub mod ruby;
pub mod section;
pub mod tracked_changes;

pub use builder::Builder;
pub use content_metadata::{
    ContentMetadata, MetadataAttribute, MetadataPart, RdfaAttributes, RdfaHost, RdfaOccurrence,
    TextMeta,
};
pub use document::{
    Document, SourceBackedDocument, SourceBackedDocumentCatalog, TextBlockCatalogEntry,
};
#[allow(unused_imports)]
pub use dynamic_text::{
    insert_database_field_xml, insert_dynamic_text_field_xml, remove_database_field_xml,
    remove_dynamic_text_field_xml, replace_database_field_xml, replace_dynamic_text_field_xml,
};
pub(crate) use index::{
    TextBibliographyType, TextIndex, insert_text_index_xml, remove_text_index_xml,
    replace_text_index_xml,
};
pub(crate) use index_mark::{
    TextIndexMark, insert_text_index_mark_xml, remove_text_index_mark_xml,
    replace_text_index_mark_xml,
};
pub use litchi_odf_common::{
    Error, LimitKind, OpenOptions, SemanticLimitKind, SemanticLimits, SemanticLimitsBuilder,
    SourceLimitKind, SourceLimits, SourceLimitsBuilder,
};
pub(crate) use note::{Note, NoteClass, insert_note_xml, remove_note_xml, replace_note_xml};
pub(crate) use page_layout::PageLayout;
pub(crate) use page_sequence::Sequence;
pub use protection::{Key, Policy};
pub(crate) use reference_mark::{
    ReferenceMark, insert_reference_mark_xml, remove_reference_mark_xml, replace_reference_mark_xml,
};
pub(crate) use section::{
    Block, add_section_xml, clear_sections_xml, remove_section_xml, unwrap_section_xml,
    update_section_xml, wrap_section_xml,
};
pub(crate) use tracked_changes::{
    Position, mark_tracked_change_range_xml, mark_tracked_deletion_xml, set_tracked_changes_xml,
    unmark_tracked_change_xml,
};
pub use transaction::Package;
pub use xforms::{
    Attribute as XFormsAttribute, Bind, Extension as XFormsExtension, Instance,
    Model as XFormsModel, ModelChild as XFormsModelChild,
    NamespaceBinding as XFormsNamespaceBinding, Submission, SubmissionMethod, SubmissionReplace,
};

// Re-export ODT-specific types for external use
pub(crate) use parser::{Section, TrackChange, TrackedChanges};
