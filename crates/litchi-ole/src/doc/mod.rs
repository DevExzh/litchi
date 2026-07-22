/// Word (.doc) document support.
///
/// This module provides parsing of Microsoft Word documents in the legacy
/// binary format (.doc files), which uses OLE2 structured storage.
///
/// # Architecture
///
/// The module is organized around these key types:
/// - `Package`: The overall .doc file package (OLE container)
/// - `Document`: The main document content and API
/// - `Paragraph`: A paragraph with runs (formatted text)
/// - `Run`: A text run with formatting
/// - `Table`: A table with rows and cells
/// - `Comment`: A comment with author metadata and body paragraphs
///
/// # DOC File Structure
///
/// A .doc file is an OLE2 structured storage containing several streams:
/// - **WordDocument**: Main document stream containing the FIB and text
/// - **1Table** or **0Table**: Contains formatting and structure information
/// - **Data**: Contains embedded objects and images
/// - **\x05SummaryInformation**: Document metadata
///
/// # Example
///
/// ```rust,no_run
/// use litchi_ole::doc::Package;
///
/// // Open a document
/// let mut package = Package::open("document.doc")?;
/// let doc = package.document()?;
///
/// // Extract all text
/// let text = doc.text()?;
/// println!("Document text: {}", text);
///
/// // Access paragraphs
/// for para in doc.paragraphs()? {
///     println!("Paragraph: {}", para.text()?);
/// }
///
/// // Access tables
/// for table in doc.tables()? {
///     for row in table.rows()? {
///         for cell in row.cells()? {
///             println!("Cell: {}", cell.text()?);
///         }
///     }
/// }
/// # Ok::<(), Box<dyn std::error::Error>>(())
/// ```
pub mod bookmark;
pub mod comment;
pub mod document;
mod encryption;
pub mod footnote;
pub mod header_footer;
pub mod hyperlink;
pub mod image;
pub mod package;
pub mod paragraph;
pub mod parts;
pub mod revision;
pub mod section;
pub mod shapes;
pub mod table;
pub mod tracked_revision;
pub mod vba;

/// DOC file writing
pub mod writer;
pub mod embedded_object;

pub use bookmark::Bookmark;
pub use comment::{Comment, CommentDateTime, CommentExtendedMetadata};
pub use document::Document;
pub use encryption::DocEncryptionProfile;
pub use footnote::{Endnote, Footnote};
pub use header_footer::HeaderFooter;
pub use hyperlink::Hyperlink;
pub use image::{Image, ImageError};
pub use package::{DocEncryptionKind, DocError, DocOpenOptions, Package};
pub use paragraph::{Paragraph, Run};
pub use parts::associated_strings::{AssociatedStringSlot, DocumentAssociatedStrings};
pub use parts::chp::CharacterConditionalFormatting;
pub use parts::document_properties::{
    CompatibilityOptions60, DocumentProperties, DocumentPropertiesBase, DocumentPropertyVersion,
    DocumentStatistics, DocumentTimestamp, EndnotePlacement, FootnotePlacement,
    NoteNumberingRestart, ProtectionSettings, SavedView, SavedViewKind, SavedZoomKind,
};
pub use parts::document_properties_97::{
    AutoSummaryState, AutoSummaryView, CompatibilityOptions80, CustomKinsokuLanguage,
    DocumentClassification, DocumentEventFlags, DocumentTypography, Dop95, Dop97,
    DopExtensionError, DrawingGrid, KinsokuLevel, MacroSecurityMetadata, OutlineDisplayLevel,
    TypographyJustification,
};
pub use parts::document_properties_2000::{
    CompatibilityOptions, Dop2000, LegacyFeatureSet, WebExportOptions, WebScreenSize,
};
pub use parts::document_properties_2002::{
    DocumentFeatureSet, Dop2002, RevisionBoundaries, StoryCharacterCounts, StylePaneFormatFilter,
    TextCodePage, TextLineEnding,
};
pub use parts::document_properties_2003::{
    DocumentProtectionMode, DocumentStateToolbars, Dop2003, ReadingModePageLock,
};
pub use parts::fields::{
    Field, FieldBoundary, FieldDescriptor, FieldEndFlags, FieldMarker, FieldMarkerValue,
    FieldStory, FieldStoryTable, FieldText, FieldType, FieldsTable, IfField, MacroButtonField,
    MailMergeConditionalControlField, MailMergeConditionalControlKind, MailMergeCounterField,
    MailMergeCounterKind, MailMergeNextField, MergeField, MergeFieldSwitch,
    PromptField, PromptFieldKind,
};
pub use parts::glossary::{
    GlossaryItem, GlossaryItemKind, GlossaryMetadata, GlossaryStyle, GlossaryTables,
};
pub use parts::list_names::ListNamesTable;
pub use parts::list_templates::{
    BuiltInListTemplate, ListTemplateCode, ListTemplateLanguageId, ListTemplateTable,
};
pub use parts::numbering::{ListLevel, ListTables, NumberFormat, ParagraphListBinding};
pub use parts::pap::ParagraphConditionalFormatting;
pub use parts::proofing::{
    ProofingEntry, ProofingFeature, ProofingRange, ProofingState, ProofingStateTable,
    ProofingStatus, ProofingTables,
};
pub use parts::saved_by::{SavedByEntry, SavedByTable};
pub use parts::styles::{
    StyleDefinition, StyleFlags, StyleKind, StylePost2000, StyleRevisionMark, StyleSheet,
    StyleSheetHeader,
};
pub use parts::tap::TableStyleCondition;
pub use revision::{
    DisplayFieldRevisionMark, NumberingRevisionMark, RevisionKind, RevisionMark, RevisionReason,
    SectionRevisionMark,
};
pub use section::{
    ChapterNumberSeparator, DocSection, LineNumberRestart, NoteNumberRestart, PageOrientation,
    SectionBehavior, SectionBreakKind, SectionColumn, SectionColumnError, SectionColumnLayout,
    SectionFootnotePosition,
    SectionLineNumbering, SectionMargins, SectionNoteSettings, SectionPageBorder,
    SectionPageBorderError,
    SectionPageBorderApplyTo, SectionPageBorderArt, SectionPageBorderColor, SectionPageBorderDepth,
    SectionPageBorderOffsetFrom, SectionPageBorderStyle, SectionPageBorders, SectionPageGrid,
    SectionPageGridMode, SectionPageLayout, SectionPageNumbering, SectionPaperSettings,
    SectionProtection, SectionTextFlow, SectionVerticalJustification, VerticalMargin,
};
pub use shapes::DocShape;
pub use table::{Cell, Row, Table};
pub use tracked_revision::{
    DocTrackedRevision, DocTrackedRevisionEditor, DocTrackedRevisionKind,
    DocTrackedRevisionMetadata,
};
pub use vba::VbaProjectStorage;
pub use writer::{
    AutoNumberAlignment, BookmarkEntry, CharacterFormatting, CommentEntry, DisplayFieldRevision,
    DocStyleDefinition, DocStyleRevision, DocWriteError, DocWriter, DropCap, DropCapType,
    FontAlignment, FormattingRevision, FrameAnchor, FrameHeight, FrameHorizontalAnchor,
    FrameHorizontalPosition, FrameTextFlow, FrameTextWrap, FrameVerticalAnchor,
    FrameVerticalPosition, HeaderFooterParagraph, LegacyAutoNumbering, LegacyBorderPosition,
    LegacyBorderStyle, LineSpacing, NumberingRevision, ParagraphBorder, ParagraphBorderStyle,
    ParagraphBorders, ParagraphFormatting, ParagraphShading, PhysicalJustification,
    StyleWriteError, TabAlignment, TabLeader, TabStop, TextBoxTightWrap, TextRevision,
};
pub use embedded_object::{
    DocEmbeddedObjectEditor, DocEmbeddedObjectReference, DocEmbeddedObjectWriteOptions,
};

/// Crate-native ordered document element returned by [`Document::elements`].
///
/// The umbrella `litchi` crate maps this into its public `DocumentElement`
/// variants. Keeping it crate-local avoids a reverse dependency from
/// `litchi-ole` back to the umbrella's `document` types.
#[derive(Debug, Clone)]
pub enum DocElement {
    /// A paragraph element.
    Paragraph(Box<Paragraph>),
    /// A table element.
    Table(Box<Table>),
}
