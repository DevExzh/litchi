//! Mutable document writer components for DOCX.
//!
//! This module provides the writer API for creating and modifying Word documents.

pub mod bookmark;
pub mod comment;
pub mod doc;
pub mod field;
pub mod hyperlink;
pub mod image;
pub mod note;
pub mod paragraph;
pub(crate) mod relmap;
pub mod run;
pub mod section;
pub mod table;

// Re-export main document type
pub use doc::MutableDocument;

// Re-export note types
pub use note::Note;

// Re-export section types
pub use section::{PageNumberFormat, PageOrientation, SectionProperties};

// Re-export hyperlink types
pub use hyperlink::MutableHyperlink;

// Re-export image types
pub use image::{ImageFormat, MutableInlineImage};

// Re-export paragraph types
pub use paragraph::{ListType, MutableParagraph};

// Re-export run types
pub use run::{MutableRun, RunContent};

// Re-export table types
pub use table::{CellProperties, MutableCell, MutableRow, MutableTable, TableBorder, TableBorders};

// Re-export comment types
pub use comment::MutableComment;

// Re-export bookmark types
pub use bookmark::MutableBookmark;

// Re-export field types
pub use field::MutableField;
