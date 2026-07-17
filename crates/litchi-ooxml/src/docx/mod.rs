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
//! use litchi_ooxml::docx::Package;
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
pub mod alt_chunk;
pub mod comment;
pub mod content_control;
pub mod custom_xml;
pub mod document;
pub mod drawing;
pub mod enums;
pub mod field;
pub mod font_table;
pub mod glossary;
pub mod footnote;
pub mod format;
pub mod header_footer;
pub mod hyperlink;
pub mod image;
mod namespace;
pub mod numbering;
pub mod list;
pub mod mail_merge;
pub mod modern_comments;
pub mod package;
pub mod paragraph;
pub mod parts;
pub mod revision;
pub mod section;
pub mod settings;
pub mod smart_tag;
pub mod statistics;
pub mod styles;
pub mod table;
pub mod template;
pub mod theme;
pub mod variables;
pub mod web_settings;
pub mod writer;

pub use bookmark::Bookmark;
pub use alt_chunk::{AltChunk, AlternativeFormatKind, AlternativeFormatPart};
pub use comment::Comment;
pub use content_control::ContentControl;
pub use custom_xml::CustomXmlPart;
pub use document::Document;
pub use drawing::{DrawingObject, ShapeType};
pub use enums::{WdHeaderFooter, WdOrientation, WdSectionStart, WdStyleType};
pub use field::Field;
pub use font_table::{
    EmbeddedFont, EmbeddedFontResource, EmbeddedFontStyle, Font, FontCharacterSet, FontFamily,
    FontPitch, FontSignature, FontTable, FontTableConformance, FontTableExtensionAttribute,
    parse_font_table, write_font_table,
};
pub use footnote::{Note, NoteType};
pub use header_footer::HeaderFooter;
pub use hyperlink::Hyperlink;
pub use mail_merge::{
    MailMergeConformance, MailMergeDataSourceObject, MailMergeDataType, MailMergeDestination,
    MailMergeFieldMap, MailMergeFieldMappingType, MailMergeMainDocumentType, MailMergeRecipient,
    MailMergeRecipients, MailMergeSettings,
};
pub use modern_comments::{
    CommentExtension, CommentIdMapping, CommentReaction, CommentReactionInfo,
    CommentReactionUser, ExtensibleComment, ModernCommentConformance, ModernCommentMetadata,
    ModernCommentRelationshipIds, Person, PresenceInfo, load_modern_comment_metadata,
    parse_comments_extended, parse_comments_extensible, parse_comments_ids, parse_people,
    store_modern_comment_metadata, write_comments_extended, write_comments_extensible,
    write_comments_ids, write_people,
};
pub use image::InlineImage;
pub use numbering::{AbstractNum, Num, Numbering};
pub use package::Package;
pub use paragraph::{
    Paragraph, Run, RunBreak, RunBreakClear, RunBreakType, RunProperties, RunUnderline,
    RunUnderlineColor,
};
pub use revision::{Revision, RevisionType};
pub use section::{Emu, Margins, PageSize, Section, Sections};
pub use settings::{DocumentSettings, ProtectionType, SmartTagType};
pub use smart_tag::{SmartTag, SmartTagAttribute};
pub use statistics::DocumentStatistics;
pub use styles::{Style, Styles};
pub use table::{Cell, Row, Table, VMergeState};
pub use theme::Theme;
pub use variables::DocumentVariables;
pub use web_settings::{
    Frame, FrameLayout, FrameScrollbarVisibility, Frameset, FramesetChild, FramesetColor,
    FramesetSplitBar, HtmlDiv, HtmlDivBorder, HtmlDivBorders, TargetScreenSize, ThemeColor,
    WebSettings, WebSettingsConformance,
};
// Re-export shared formatting types
pub use format::{ImageFormat, LineSpacing, ParagraphAlignment, TableBorderStyle, UnderlineStyle};
// Re-export writer types
pub use writer::{
    CellProperties, ColorScheme, ContentControlType, DocumentProtection, ListType, MutableBookmark,
    MutableComment, MutableContentControl, MutableDocument, MutableField, MutableHyperlink,
    MutableInlineImage, MutableParagraph, MutableRun, MutableSmartTag, MutableSmartTagAttribute,
    MutableStyle, MutableTable, MutableTheme, PageNumberFormat, PageOrientation, RunContent,
    SectionProperties, TableBorder, TableBorders, TableOfContents, Watermark, generate_styles_xml,
};
// Note: writer::Note is not re-exported to avoid naming conflict with footnote::Note
// Use writer::Note explicitly if needed

/// Crate-native ordered document element returned by [`Document::elements`].
///
/// The umbrella `litchi` crate maps this into its public `DocumentElement`
/// variants. Keeping it crate-local avoids a reverse dependency from
/// `litchi-ooxml` back to the umbrella's `document` types.
#[derive(Debug, Clone)]
pub enum DocxElement {
    /// A paragraph element.
    Paragraph(Box<Paragraph>),
    /// A table element.
    Table(Box<Table>),
}

/// Ordered main-document block, including opaque alternative-format anchors.
///
/// This is separate from [`DocxElement`] so existing format-agnostic callers
/// that exhaustively match paragraphs and tables remain source-compatible.
#[derive(Debug, Clone)]
pub enum DocumentBlock {
    Paragraph(Box<Paragraph>),
    Table(Box<Table>),
    AltChunk(Box<AltChunk>),
}
