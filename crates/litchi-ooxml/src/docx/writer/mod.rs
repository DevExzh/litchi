//! Mutable document writer components for DOCX.
//!
//! This module provides the writer API for creating and modifying Word documents.

pub mod bookmark;
pub mod comment;
pub mod content_control;
pub mod doc;
pub mod field;
pub mod hyperlink;
pub mod image;
pub mod note;
pub mod paragraph;
pub mod revision;
pub(crate) mod relmap;
pub mod run;
pub mod section;
pub mod smart_tag;
pub mod style;
pub mod table;
pub mod theme;
pub mod toc;
pub mod watermark;

// Re-export main document type
pub use doc::{DocumentProtection, MutableDocument};

// Re-export note types
pub use note::Note;

// Re-export section types
pub use section::{
    DocumentGridType, NoteNumberRestart, PageNumberFormat, PageOrientation, SectionColumn,
    SectionColumns, SectionDocumentGrid, SectionHeaderFooterPart, SectionHeaderFooterReference,
    SectionNoteProperties, SectionPageNumbering, SectionProperties, SectionTextDirection,
};

// Re-export smart-tag types
pub use smart_tag::{MutableSmartTag, MutableSmartTagAttribute};

// Re-export hyperlink types
pub use hyperlink::MutableHyperlink;

// Re-export image types
pub use image::{ImageFormat, MutableInlineImage};

// Re-export paragraph types
pub use paragraph::{ListType, MutableParagraph};

// Re-export run types
pub use run::{MutableRun, RunContent};
pub use revision::{
    CellRevisionKind, MutableRevision, RevisionContentControl, RevisionKind, RevisionMetadata,
    RowRevisionKind, TableCellMergeRevisionState, TableRevisionKind,
};

// Re-export table types
pub use table::{CellProperties, MutableCell, MutableRow, MutableTable, TableBorder, TableBorders};

// Re-export comment types
pub use comment::MutableComment;

// Re-export bookmark types
pub use bookmark::MutableBookmark;

// Re-export field types
pub use field::{CitationFieldSpec, CitationSource, MutableField};

// Re-export style types
pub use style::{MutableStyle, generate_styles_xml};

// Re-export content control types
pub use content_control::{ContentControlType, MutableContentControl};

// Re-export theme types
pub use theme::{ColorScheme, MutableTheme};

// Re-export watermark types
pub use watermark::Watermark;

// Re-export TOC types
pub use toc::TableOfContents;
