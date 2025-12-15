//! Word (.docx) document support.
//!
//! This module provides parsing and manipulation of Microsoft Word documents
//! in the Office Open XML (OOXML) format (.docx files).
//!
//! # Architecture
//!
//! The module is organized around these key types:
//! - `Package`: The overall .docx file package
//! - `Document`: The main document content and API
//! - `Paragraph`: A paragraph with runs
//! - `Run`: A text run with formatting
//! - `Table`: A table with rows and cells
//! - `Section`: A document section with page properties
//! - `Styles`: Collection of document styles
//! - `DocumentPart`: The core document.xml part
//!
//! # Example
//!
//! ```rust,no_run
//! use litchi::ooxml::docx::Package;
//!
//! // Open a document
//! let package = Package::open("document.docx")?;
//! let doc = package.document()?;
//!
//! // Access paragraphs and runs
//! for para in doc.paragraphs()? {
//!     println!("Paragraph: {}", para.text()?);
//!     for run in para.runs()? {
//!         println!("  Run: {} (bold: {:?})", run.text()?, run.bold()?);
//!     }
//! }
//!
//! // Access tables
//! for table in doc.tables()? {
//!     for row in table.rows()? {
//!         for cell in row.cells()? {
//!             println!("Cell: {}", cell.text()?);
//!         }
//!     }
//! }
//!
//! // Access sections
//! let mut sections = doc.sections()?;
//! for section in sections.iter_mut() {
//!     println!("Page width: {:?}", section.page_width());
//! }
//! # Ok::<(), Box<dyn std::error::Error>>(())
//! ```

pub mod bookmark;
pub mod comment;
pub mod content_control;
pub mod custom_xml;
pub mod document;
pub mod drawing;
pub mod enums;
pub mod field;
pub mod footnote;
pub mod format;
pub mod header_footer;
pub mod hyperlink;
pub mod image;
pub mod numbering;
pub mod package;
pub mod paragraph;
pub mod parts;
pub mod revision;
pub mod section;
pub mod settings;
pub mod statistics;
pub mod styles;
pub mod table;
pub mod template;
pub mod theme;
pub mod variables;
pub mod writer;

pub use bookmark::Bookmark;
pub use comment::Comment;
pub use content_control::ContentControl;
pub use custom_xml::CustomXmlPart;
pub use document::Document;
pub use drawing::{DrawingObject, ShapeType};
pub use enums::{WdHeaderFooter, WdOrientation, WdSectionStart, WdStyleType};
pub use field::Field;
pub use footnote::{Note, NoteType};
pub use header_footer::HeaderFooter;
pub use hyperlink::Hyperlink;
pub use image::InlineImage;
pub use numbering::{AbstractNum, Num, Numbering};
pub use package::Package;
pub use paragraph::{Paragraph, Run, RunProperties};
pub use revision::{Revision, RevisionType};
pub use section::{Emu, Margins, PageSize, Section, Sections};
pub use settings::{DocumentSettings, ProtectionType};
pub use statistics::DocumentStatistics;
pub use styles::{Style, Styles};
pub use table::{Cell, Row, Table, VMergeState};
pub use theme::Theme;
pub use variables::DocumentVariables;
// Re-export shared formatting types
pub use format::{ImageFormat, LineSpacing, ParagraphAlignment, TableBorderStyle, UnderlineStyle};
// Re-export writer types
pub use writer::{
    CellProperties, ColorScheme, ContentControlType, DocumentProtection, ListType, MutableBookmark,
    MutableComment, MutableContentControl, MutableDocument, MutableField, MutableHyperlink,
    MutableInlineImage, MutableParagraph, MutableRun, MutableStyle, MutableTable, MutableTheme,
    PageNumberFormat, PageOrientation, RunContent, SectionProperties, TableBorder, TableBorders,
    TableOfContents, Watermark, generate_styles_xml,
};
// Note: writer::Note is not re-exported to avoid naming conflict with footnote::Note
// Use writer::Note explicitly if needed
