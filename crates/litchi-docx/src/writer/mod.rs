//! Mutable document writer components for DOCX.
//!
//! This module provides the writer API for creating and modifying Word documents.

pub(crate) mod bookmark;
pub(crate) mod comment;
pub(crate) mod content_control;
pub(crate) mod doc;
pub(crate) mod field;
pub(crate) mod hyperlink;
pub(crate) mod image;
pub(crate) mod note;
pub(crate) mod ole_object;
pub(crate) mod paragraph;
pub(crate) mod relmap;
pub(crate) mod revision;
pub(crate) mod run;
pub(crate) mod section;
pub(crate) mod smart_tag;
pub(crate) mod smartart;
pub(crate) mod style;
pub(crate) mod table;
pub(crate) mod textbox;
pub(crate) mod theme;
pub(crate) mod toc;
pub(crate) mod vml_shape;
pub(crate) mod watermark;

// Re-export main document type
pub use doc::Protection;

pub use note::Note;

pub use section::borders::{Art, Border, Borders, Color, Display, OffsetFrom, Style, ZOrder};
pub use section::{
    ChapterSep, EndnotePos, Endnotes, FootnotePos, Footnotes, GridType, LineNumberRestart,
    NoteNumberRestart, PageNumberFormat, PageOrientation, SectionColumn, SectionColumns,
    SectionDocumentGrid, SectionHeaderFooterPart, SectionHeaderFooterReference,
    SectionLineNumbering, SectionPageNumbering, SectionPaperSource, SectionProperties,
    SectionTextDirection, SectionVerticalAlignment,
};

pub use image::ImageFormat;

pub use paragraph::ListType;

pub use revision::{
    CellRevisionKind, ConflictKind, RevisionContentControl, RevisionKind, RevisionMetadata,
    RowRevisionKind, TableCellMergeRevisionState,
};
pub use run::RunContent;

pub use table::{CellProperties, TableBorder, TableBorders};

pub use ole_object::MAX_OLE_PAYLOAD_BYTES;

pub use smartart::MAX_SMART_ARTS;

pub use vml_shape::{VmlShapeKind, VmlShapePosition};

pub use comment::{Date, DateError, MAX_DATE_BYTES};

pub use field::{BibliographyFieldSpec, BibliographyFilter, CitationFieldSpec, CitationSource};

pub use style::generate_styles_xml;

pub use content_control::{Checksum, ContentControlType, DataBinding, FormattingAllowed};

pub use theme::ColorScheme;

pub use watermark::{
    ImageWatermark, ImageWatermarkAnchor, MAX_WATERMARK_IMAGE_BYTES, Watermark, WatermarkLayout,
};

pub use toc::TableOfContents;
