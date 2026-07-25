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

pub mod alt_chunk;
pub mod bibliography;
pub mod bookmark;
pub mod chart;
pub mod comment;
pub mod content_control;
pub mod custom_xml;
pub mod document;
pub mod drawing;
pub mod enums;
pub mod field;
pub mod font_table;
pub mod footnote;
pub mod format;
pub mod glossary;
pub mod header_footer;
pub mod hyperlink;
pub mod image;
pub mod list;
pub mod mail_merge;
pub mod math;
pub mod modern_comments;
mod namespace;
pub mod numbering;
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
pub mod variables;
pub mod vba_project;
pub mod web_settings;
pub mod writer;

pub use alt_chunk::{
    AltChunk, AltChunkNamespace, AlternativeFormatData, AlternativeFormatImport,
    AlternativeFormatKind, AlternativeFormatPart, AlternativeFormatTarget,
};
pub use bibliography::{
    BibliographySource, BibliographySourceStore, BibliographySourceValue,
    LEGACY_WORD_BIBLIOGRAPHY_NAMESPACE, OOXML_BIBLIOGRAPHY_NAMESPACE,
    STRICT_OOXML_BIBLIOGRAPHY_NAMESPACE,
};
pub use bookmark::Bookmark;
pub use chart::{
    DocxChartCompanionResource, DocxChartConformance, DocxChartEmbeddedWorkbookContentType,
    DocxChartEmbeddedWorkbookResource, DocxChartGraph, DocxChartResource, load_chart_graph,
    store_chart_graph,
};
pub use comment::Comment;
pub use content_control::ContentControl;
pub use custom_xml::{CustomXmlBinding, CustomXmlPart, NewCustomXmlDataStore};
pub use document::{Document, ImageWatermarkPart};
pub use drawing::{DrawingObject, ShapeType};
pub use enums::{WdHeaderFooter, WdOrientation, WdSectionStart, WdStyleType};
pub use field::CompareField;
pub use field::{
    ActiveContentField, ActiveContentFieldKind, AddressBlockCountryInclusion, AdvanceField,
    AdvanceFieldAdjustment, AdvanceFieldOperation, AutoNumberField, AutoNumberFieldKind,
    AutoTextField, AutoTextFieldKind,
    AutoTextListField, AutoTextListOption, BarcodeField, BibliographyField, BidiOutlineField,
    CitationField, DatabaseField, DdeField, DdeFieldKind, DdeRepresentation,
    DocumentContextField, DocumentContextFieldKind,
    DocumentInformationField, DocumentInformationFieldKind, DocumentPropertyField,
    DocumentVariableField, ExternalIncludeField, ExternalIncludeOption, Field, FieldSwitch,
    EmbedField, EquationField, FormulaField, GoToButtonField, HyperlinkField, IfField,
    IncludeFieldKind, IndexEntryField,
    IndexField, IndexSortOrder, InfoField, LegacyFormField, LegacyFormFieldKind, LinkField,
    LinkFormatting, LinkResultOption, ListNumberField, MacroButtonField,
    MailMergeConditionalControlField, MailMergeConditionalControlKind, MailMergeCounterField,
    MailMergeCounterKind, MailMergeDataField, MailMergeNextField, MailMergeRecipientField,
    MailMergeRecipientFieldKind, MergeField, PrintField, PrivateField, PromptField,
    PromptFieldKind,
    QuoteField, ReferencedDocumentField, ReferenceField, ReferenceFieldKind, ReferenceFieldOption,
    SequenceField,
    SetField, ShapeField, StyleReferenceField, StyleReferenceFieldOption,
    SymbolField,
    TableOfAuthoritiesEntryField, TableOfAuthoritiesField, TableOfContentsEntryField,
    TableOfContentsField, TableOfContentsLevelRange,
    TableOfContentsSwitch, UserIdentityField, UserIdentityFieldKind, UserIdentityFormatting,
};
pub use font_table::{
    EmbeddedFont, EmbeddedFontLicensing, EmbeddedFontResource, EmbeddedFontStyle, Font,
    FontCharacterSet, FontFamily, FontPitch, FontSignature, FontTable, FontTableConformance,
    FontTableExtensionAttribute, add_font, deobfuscate_embedded_font_data, find_font,
    load_font_table, obfuscate_embedded_font_data, parse_font_table, remove_font, reorder_fonts,
    replace_font, store_font_table, update_font, validate_embedded_font_usage, write_font_table,
};
pub use footnote::{Note, NoteType};
pub use glossary::{
    DocPartCategory, DocPartGallery, DocPartName, DocPartProperties, DocPartType,
    GlossaryAuxiliaryPart, GlossaryDocument, GlossaryEntry, GlossaryPackage,
    GlossaryRelationship, InsertionBehavior,
};
pub use header_footer::HeaderFooter;
pub use hyperlink::Hyperlink;
pub use image::InlineImage;
pub use mail_merge::{
    MailMergeConformance, MailMergeDataSourceObject, MailMergeDataType, MailMergeDestination,
    MailMergeFieldMap, MailMergeFieldMappingType, MailMergeMainDocumentType, MailMergeRecipient,
    MailMergeRecipients, MailMergeSettings, MailMergeSource, MailMergeTarget,
};
pub use math::{OfficeMath, OfficeMathParagraph};
pub use modern_comments::{
    CommentExtension, CommentIdMapping, CommentReaction, CommentReactionInfo, CommentReactionUser,
    ExtensibleComment, ModernCommentConformance, ModernCommentMetadata,
    ModernCommentRelationshipIds, Person, PresenceInfo, load_modern_comment_metadata,
    parse_comments_extended, parse_comments_extensible, parse_comments_ids, parse_people,
    store_modern_comment_metadata, write_comments_extended, write_comments_extensible,
    write_comments_ids, write_people,
};
pub use numbering::{AbstractNum, Num, Numbering};
pub use package::Package;
pub use paragraph::{
    Paragraph, Run, RunBreak, RunBreakClear, RunBreakType, RunProperties, RunUnderline,
    RunUnderlineColor,
};
pub use revision::{Revision, RevisionType};
pub use section::{Emu, Margins, PageSize, Section, Sections};
pub use settings::{AttachedTemplate, DocumentSettings, ProtectionType, SmartTagType};
pub use smart_tag::{SmartTag, SmartTagAttribute};
pub use smartart::{DocxDiagramConformance, DocxSmartArt, load_smart_arts};
// Re-export the shared semantic SmartArt model for authoring.
pub use crate::diagrams::{DiagramNode, DiagramType, SmartArt, SmartArtBuilder};
pub use statistics::DocumentStatistics;
pub use styles::{Style, Styles};
pub use table::{Cell, Row, Table, VMergeState};
pub use textbox::{
    DocxTextBox, TextBoxAnchor, TextBoxAutofit, TextBoxBodyProperties, TextBoxInsets,
    TextBoxParagraph, TextBoxRun, TextDirection, TextVerticalAnchor, TextWarpPreset, TextWrap,
    WordArt, load_text_boxes,
};
pub use theme::Theme;
pub use variables::DocumentVariables;
pub use vba_project::VbaProject;
pub use web_settings::{
    Frame, FrameLayout, FrameScrollbarVisibility, Frameset, FramesetChild, FramesetColor,
    FramesetSplitBar, HtmlDiv, HtmlDivBorder, HtmlDivBorders, TargetScreenSize, ThemeColor,
    WebSettings, WebSettingsConformance,
};
// Re-export shared formatting types
pub use format::{ImageFormat, LineSpacing, ParagraphAlignment, TableBorderStyle, UnderlineStyle};
// Re-export writer types
pub use writer::{
    BibliographyFieldSpec, BibliographyFilter, CellProperties, CitationFieldSpec, CitationSource,
    ColorScheme, ContentControlType, DocumentProtection, ListType, MutableBookmark, MutableComment,
    MutableContentControl, MutableDocument, MutableField, MutableHyperlink,
    MutableInlineImage, MutableParagraph, MutableRevision, MutableRun, MutableSmartTag,
    MutableSmartTagAttribute, MutableStyle, MutableTable, MutableTextBox, MutableTheme,
    MutableOleObject, MAX_OLE_PAYLOAD_BYTES, MutableSmartArt, MAX_SMART_ARTS, PageNumberFormat,
    MutableVmlShape, VmlShapeKind, VmlShapePosition,
    PageOrientation, RevisionContentControl, RevisionKind, RevisionMetadata, RowRevisionKind,
    CellRevisionKind, TableCellMergeRevisionState, TableRevisionKind, RunContent,
    DocumentGridType, NoteNumberRestart, SectionColumn, SectionColumns, SectionDocumentGrid,
    SectionHeaderFooterPart, SectionHeaderFooterReference, SectionNoteProperties,
    SectionPageNumbering, SectionProperties, SectionTextDirection, TableBorder, TableBorders,
    TableOfContents, Watermark, WatermarkLayout, ImageWatermark, ImageWatermarkAnchor,
    MAX_WATERMARK_IMAGE_BYTES,
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
    AltChunk(Box<AltChunk>),
}
