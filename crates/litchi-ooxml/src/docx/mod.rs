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

pub mod bibliography;
pub mod bibliography_writer;
pub mod bookmark;
pub mod chart;
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
pub mod list;
pub mod mail_merge;
pub mod math;
mod namespace;
mod numbering;
pub mod package;
pub mod paragraph;
pub mod parts;
pub mod revision;
pub mod section;
pub mod settings;
pub mod smart_tag;
pub mod smartart;
pub mod statistics;
pub mod styles;
pub mod table;
pub mod template;
pub mod textbox;
pub mod theme;
pub mod vba_project;
pub mod writer;

pub use bibliography::{
    BibliographySource, BibliographySourceStore, BibliographySourceValue,
    LEGACY_WORD_BIBLIOGRAPHY_NAMESPACE, OOXML_BIBLIOGRAPHY_NAMESPACE,
    STRICT_OOXML_BIBLIOGRAPHY_NAMESPACE,
};
pub use bibliography_writer::{
    BibliographyPerson, BibliographySourceBuilder, BibliographySourceKind,
};
pub use bookmark::Bookmark;
pub use chart::{
    DocxChartCompanionResource, DocxChartConformance, DocxChartEmbeddedWorkbookContentType,
    DocxChartEmbeddedWorkbookResource, DocxChartGraph, DocxChartResource, load_chart_graph,
    store_chart_graph,
};
pub use comment::Comment;
pub use content_control::{ContentControl, Kind as ContentControlKind};
pub use custom_xml::{Binding, NewStore, Part};
pub use document::{Document, ImageWatermarkPart};
pub use drawing::DrawingObject;
pub use enums::{WdHeaderFooter, WdOrientation, WdSectionStart, WdStyleType};
pub use field::Compare;
pub use field::{
    ActiveContent, ActiveContentKind, Advance, AdvanceAdjustment, AdvanceOperation, AutoNumber,
    AutoNumberKind, AutoText, AutoTextKind, AutoTextList, AutoTextListOption, Barcode,
    Bibliography, BidiOutline, Citation, Context, ContextKind, CountryInclusion, Database, Dde,
    DdeFormat, DdeKind, Embed, Equation, Field, Formula, GoToButton, If, Include, IncludeKind,
    IncludeOption, Index, IndexEntry, IndexOrder, Info, Information, InformationKind, LegacyForm,
    LegacyFormKind, Link, LinkFormat, LinkResult, ListNumber, MacroButton, Merge, MergeControl,
    MergeControlKind, MergeCounter, MergeCounterKind, MergeData, MergeNext, Print, Private, Prompt,
    PromptKind, Property, Quote, RecipientKind, Reference, ReferenceKind, ReferenceOption,
    Sequence, Set, Shape, StyleOption, StyleReference, SubDocument, Switch, Symbol, Toa, ToaEntry,
    Toc, TocEntry, TocLevelRange, UserIdentity, UserIdentityFormat, UserIdentityKind, Variable,
};
pub use footnote::{Note, NoteType};
pub use header_footer::HeaderFooter;
pub use hyperlink::Hyperlink;
pub use image::InlineImage;
pub use litchi_docx::{color, glossary, web};
pub use litchi_opc::FontEmbedding;
pub use mail_merge::{
    Conformance, DataSourceObject, DataType, Destination, FieldMap, FieldMappingType,
    MainDocumentType, RECIPIENT_CONTENT_TYPE, Recipient, Recipients, Settings, Source, Target,
};
pub use math::{OfficeMath, OfficeMathParagraph};
pub use package::Package;
pub use paragraph::{
    Paragraph, Run, RunBreak, RunBreakClear, RunBreakType, RunProperties, RunUnderline,
    RunUnderlineColor,
};
pub use revision::{Revision, RevisionType};
pub use section::{Emu, Margins, PageSize, Section, Sections};
pub use settings::{
    AttachedTemplate, ColorSchemeIndex, ColorSchemeMapping, ColorSchemeSlot, CompatFlag,
    CompatibilityOption, CompatibilitySetting, DocumentSettings, MAX_LANGUAGE_TAG_LENGTH,
    NoteNumberingProperties, NoteNumberingRestart, NotePosition, ParseCompatFlagError,
    ParseNotePositionError, ProofState, ProofingState, ProtectionType, SmartTagType,
    ThemeFontLanguages, View,
};
pub use smart_tag::{SmartTag, SmartTagAttribute};
pub use smartart::{DocxDiagramConformance, DocxSmartArt, load_smart_arts};
// Re-export the shared semantic SmartArt model for authoring.
use litchi_docx::alt::Chunk;
pub use litchi_drawingml::diagram::{DiagramNode, DiagramType, SmartArt, SmartArtBuilder};
pub use litchi_drawingml::geom::{Preset, TextPreset};
pub use statistics::Statistics;
pub use styles::{Outline, Style, Styles};
pub use table::{Cell, Row, Table, VMergeState};
pub use textbox::{
    Columns, Coordinate32, DocxTextBox, TextBoxAnchor, TextBoxAutofit, TextBoxBodyProperties,
    TextBoxInsets, TextBoxParagraph, TextBoxRun, TextDirection, TextUnderline, TextVerticalAnchor,
    TextWrap, WordArt, load_text_boxes,
};
pub use theme::Theme;
pub use vba_project::{VbaDocumentEvent, VbaMacroDescriptor, VbaProject, VbaSupplementalData};
// Re-export shared formatting types
pub use format::{ImageFormat, LineSpacing, ParagraphAlignment, TableBorderStyle, UnderlineStyle};
// Re-export writer types
pub use writer::{
    BibliographyFieldSpec, BibliographyFilter, BorderColor, CellProperties, CellRevisionKind,
    ChapterSep, CitationFieldSpec, CitationSource, ColorScheme, ContentControlType,
    DocumentGridType, DocumentProtection, EndnotePos, Endnotes, FootnotePos, Footnotes,
    ImageWatermark, ImageWatermarkAnchor, LineNumberRestart, ListType, MAX_OLE_PAYLOAD_BYTES,
    MAX_SMART_ARTS, MAX_WATERMARK_IMAGE_BYTES, MutableBookmark, MutableComment,
    MutableContentControl, MutableDocument, MutableField, MutableHyperlink, MutableInlineImage,
    MutableOleObject, MutableParagraph, MutableRevision, MutableRun, MutableSmartArt,
    MutableSmartTag, MutableSmartTagAttribute, MutableStyle, MutableTable, MutableTextBox,
    MutableTheme, MutableVmlShape, NoteNumberRestart, PageBorderArt, PageBorderDisplay,
    PageBorderOffsetFrom, PageBorderStyle, PageBorderZOrder, PageNumberFormat, PageOrientation,
    RevisionContentControl, RevisionKind, RevisionMetadata, RowRevisionKind, RunContent,
    SectionColumn, SectionColumns, SectionDocumentGrid, SectionHeaderFooterPart,
    SectionHeaderFooterReference, SectionLineNumbering, SectionPageBorder, SectionPageBorders,
    SectionPageNumbering, SectionPaperSource, SectionProperties, SectionTextDirection,
    SectionVerticalAlignment, TableBorder, TableBorders, TableCellMergeRevisionState,
    TableOfContents, TableRevisionKind, VmlShapeKind, VmlShapePosition, Watermark, WatermarkLayout,
    generate_styles_xml,
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
    Alt(Box<Chunk>),
}
