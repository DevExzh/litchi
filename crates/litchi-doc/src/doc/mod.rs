pub mod bookmark;
pub mod comment;
pub mod document;
mod encryption;
pub mod footnote;
pub mod header_footer;
pub mod hyperlink;
pub mod image;
pub mod leniency;
pub mod package;
pub mod paragraph;
pub mod parts;
mod plcf;
pub mod revision;
pub mod section;
pub mod shapes;
pub mod table;
pub mod tracked_revision;
pub mod vba;

pub mod embedded_object;
pub mod equation;
/// DOC file writing
pub mod writer;

pub use bookmark::Bookmark;
pub use comment::{Comment, CommentDateTime, CommentExtendedMetadata};
pub use document::Document;
pub use embedded_object::{
    DocEmbeddedObjectEditor, DocEmbeddedObjectReference, DocEmbeddedObjectWriteOptions,
};
pub use encryption::DocEncryptionProfile;
pub use equation::{DocMtefEquationWriteOptions, EQUATION_3_CLSID, MtefEquation};
pub use footnote::{Endnote, Footnote};
pub use header_footer::HeaderFooter;
pub use hyperlink::Hyperlink;
pub use image::{Image, ImageError};
pub use leniency::{DocLeniency, DocStylesheetDefect, DocToleranceReport, DocToleratedDefect};
pub use package::{DocEncryptionKind, DocError, DocOpenOptions, Package, Result};
pub use paragraph::{Paragraph, Run};
pub use parts::associated_strings::{AssociatedStringSlot, DocumentAssociatedStrings};
pub use parts::auto_summary::{AutoSummaryRange, DocumentAutoSummary};
pub use parts::captions::{
    AutoCaptionEntry, AutoCaptionTable, CaptionDefinition, CaptionInfo, CaptionLabelTable,
    CaptionLocation, CaptionTables, ChapterHeading, ChapterNumbering, ChapterSeparator,
};
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
pub use parts::embedded_fonts::{DocumentEmbeddedFonts, EmbeddedFont};
pub use parts::fields::BarcodeField;
pub use parts::fields::BidiOutlineField;
pub use parts::fields::CompareField;
pub use parts::fields::DocumentPropertyField;
pub use parts::fields::DocumentVariableField;
pub use parts::fields::EmbedField;
pub use parts::fields::EquationField;
pub use parts::fields::FormulaField;
pub use parts::fields::HyperlinkField;
pub use parts::fields::InfoField;
pub use parts::fields::ListNumberField;
pub use parts::fields::PrintField;
pub use parts::fields::PrivateField;
pub use parts::fields::QuoteField;
pub use parts::fields::SequenceField;
pub use parts::fields::SetField;
pub use parts::fields::ShapeField;
pub use parts::fields::SymbolField;
pub use parts::fields::{ActiveContentField, ActiveContentFieldKind};
pub use parts::fields::{
    AddressBlockCountryInclusion, Field, FieldBoundary, FieldDescriptor, FieldEndFlags,
    FieldMarker, FieldMarkerValue, FieldStory, FieldStoryTable, FieldText, FieldType, FieldsTable,
    GoToButtonField, IfField, MacroButtonField, MailMergeConditionalControlField,
    MailMergeConditionalControlKind, MailMergeCounterField, MailMergeCounterKind,
    MailMergeDataField, MailMergeNextField, MailMergeRecipientField, MailMergeRecipientFieldKind,
    MergeField, MergeFieldSwitch, NonPlcfFields, PromptField, PromptFieldKind,
    TableOfAuthoritiesEntryField, TableOfAuthoritiesEntryOption, TableOfAuthoritiesField,
    TableOfAuthoritiesOption, TableOfContentsEntryField, TableOfContentsEntryOption,
    TableOfContentsField, TableOfContentsOption, UserIdentityField, UserIdentityFieldKind,
    UserIdentityFormatting,
};
pub use parts::fields::{AdvanceField, AdvanceFieldAdjustment, AdvanceFieldOperation};
pub use parts::fields::{AutoNumberField, AutoNumberFieldKind};
pub use parts::fields::{AutoTextField, AutoTextFieldKind, AutoTextListField, AutoTextListOption};
pub use parts::fields::{DdeField, DdeFieldKind, DdeRepresentation};
pub use parts::fields::{
    DocumentContextField, DocumentContextFieldKind, DocumentInformationField,
    DocumentInformationFieldKind,
};
pub use parts::fields::{ExternalIncludeField, ExternalIncludeOption, IncludeFieldKind};
pub use parts::fields::{IndexEntryField, IndexEntryOption, IndexField, IndexOption};
pub use parts::fields::{LegacyFormField, LegacyFormFieldKind};
pub use parts::fields::{LinkField, LinkFormatting, LinkResultOption};
pub use parts::fields::{
    ReferenceField, ReferenceFieldKind, ReferenceFieldOption, ReferencedDocumentField,
};
pub use parts::fields::{StyleReferenceField, StyleReferenceFieldOption};
pub use parts::form_fields::{
    CheckBoxState, FormFieldData, FormFieldDataKind, FormFieldTextKind, NilPicfAndBinData,
};
pub use parts::format_consistency::{
    DocumentFormatConsistencyMarks, FormatConsistencyInfo, FormatConsistencyKind,
    FormatConsistencyMark, FormatConsistencyProperties,
};
pub use parts::glossary::{
    AttachedGlossary, GlossaryItem, GlossaryItemKind, GlossaryMetadata, GlossaryStyle,
    GlossaryTables,
};
pub use parts::grammar_cookies::{
    CookieElement, CookieEntry, CookieErrorType, GrammarCookie, GrammarCookiePlc,
    GrammarCookieTable, GrammarCookieTables, LegacyGrammarCookie, LegacyGrammarCookieTable,
};
pub use parts::headers::HeaderFooterType;
pub use parts::list_names::ListNamesTable;
pub use parts::list_templates::{
    BuiltInListTemplate, ListTemplateCode, ListTemplateLanguageId, ListTemplateTable,
};
pub use parts::mail_merge::{
    DocumentMailMerge, FieldMapInfo, FieldMapping, FilterComparison, FilterCondition,
    FilterDataItem, Fnpi, MailMergeDestination, MailMergeDocumentType, MailMergeType,
    MergeDataSourceKind, MergeErrorCheck, MergeFileToken, OdsoProperty, Pmfs, Pms, RecipientEntry,
    RecipientInfo, Rfs, SortColumnAndDirection, SortDirection, SttbfRfs, Wpms,
};
pub use parts::numbering::{ListLevel, ListTables, NumberFormat, ParagraphListBinding};
pub use parts::ole_controls::{DocumentOleControls, OleControlInfo};
pub use parts::pap::ParagraphConditionalFormatting;
pub use parts::proofing::{
    ProofingEntry, ProofingFeature, ProofingRange, ProofingState, ProofingStateTable,
    ProofingStatus, ProofingTables,
};
pub use parts::protection::{
    DocumentProtectedRanges, ProtectedRange, ProtectionUser, ProtectionUserRole, UidSel,
};
pub use parts::repair_bookmarks::{DocumentRepairBookmarks, RepairBookmark};
pub use parts::rmd_threading::{DocumentRmdThreading, MessageDisplayProperties, ThreadingMessage};
pub use parts::rsids::DocumentRsids;
pub use parts::saved_by::{SavedByEntry, SavedByTable};
pub use parts::smart_tags::{
    DocumentSmartTag, DocumentSmartTags, SmartTagBookmarkInfo, SmartTagOrigin,
    SmartTagRecognizerRange, SmartTagRecognizerState,
};
pub use parts::spa::{
    ShapeAnchor, ShapeHorizontalOrigin, ShapeTextWrap, ShapeVerticalOrigin, ShapeWrapSide, Spa,
};
pub use parts::structured_tags::{
    DocumentStructuredTag, DocumentStructuredTags, StructuredTagAttribute, StructuredTagInfo,
    StructuredTagKind, StructuredTagName,
};
pub use parts::styles::{
    StyleDefinition, StyleFlags, StyleKind, StylePost2000, StyleRevisionMark, StyleSheet,
    StyleSheetHeader,
};
pub use parts::subdocuments::{
    DocumentSubdocuments, ReferencedFileKind, ReferencedFileName, Subdocument,
};
pub use parts::table_char_cache::{TableCharEntry, TableCharInfo, TableCharacterCache};
pub use parts::tap::TableStyleCondition;
pub use parts::text_services::{TextServicesTables, Uim, UimEntry, UimGuidTable, UimTable};
pub use parts::textbox::DocTextBox;
pub use parts::textbox_breaks::{
    TextBoxBreak, TextBoxBreakEntry, TextBoxBreakKind, TextBoxBreakTable, TextBoxBreakTables,
};
pub use revision::{
    DisplayFieldRevisionMark, NumberingRevisionMark, RevisionKind, RevisionMark, RevisionReason,
    SectionRevisionMark,
};
pub use section::borders::{ApplyTo, Art, Border, Borders, Color, Depth, Error, Offset, Style};
pub use section::{
    ChapterNumberSeparator, DocSection, LineNumberRestart, NoteNumberRestart, PageOrientation,
    SectionBehavior, SectionBreakKind, SectionFootnotePosition, SectionLineNumbering,
    SectionMargins, SectionNoteSettings, SectionPageGrid, SectionPageGridMode, SectionPageLayout,
    SectionPageNumbering, SectionPaperSettings, SectionProtection, SectionTextFlow,
    SectionVerticalJustification, VerticalMargin,
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
    DocDrawingShape, DocHeaderKind, DocPicture, DocShapeKind, DocSmartTagEntry, DocStyleDefinition,
    DocStyleRevision, DocWriteError, DocWriter, DropCap, DropCapType, FloatingPosition,
    FontAlignment, FormattingRevision, FrameAnchor, FrameHeight, FrameHorizontalAnchor,
    FrameHorizontalPosition, FrameTextFlow, FrameTextWrap, FrameVerticalAnchor,
    FrameVerticalPosition, HeaderFooterParagraph, LegacyAutoNumbering, LegacyBorderPosition,
    LegacyBorderStyle, LineSpacing, NumberingRevision, ParagraphBorder, ParagraphBorderStyle,
    ParagraphBorders, ParagraphFormatting, ParagraphShading, PhysicalJustification,
    StyleWriteError, TabAlignment, TabLeader, TabStop, TextBoxTightWrap, TextRevision,
};

/// Crate-native ordered document element returned by [`Document::elements`].
///
/// The umbrella `litchi` crate maps this into its public `DocumentElement`
/// variants. Keeping it crate-local avoids a reverse dependency from
/// `litchi-doc` back to the umbrella's `document` types.
#[derive(Debug, Clone)]
pub enum DocElement {
    /// A paragraph element.
    Paragraph(Box<Paragraph>),
    /// A table element.
    Table(Box<Table>),
}
