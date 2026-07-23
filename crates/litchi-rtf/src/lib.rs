#![allow(missing_docs)]

//! RTF (Rich Text Format) parser module.
//!
//! This module provides high-performance parsing of RTF documents with support
//! for the RTF 1.9.1 specification. It uses arena allocation (bumpalo) for efficient
//! memory management during parsing and zero-copy patterns where possible.
//!
//! # Architecture
//!
//! The parser is organized into several components:
//! - **Lexer**: Tokenizes RTF input into control words, symbols, and text
//! - **Parser**: Builds a structured document from tokens
//! - **Document**: High-level document representation with paragraphs, runs, and tables
//!
//! # Example
//!
//! ```rust,no_run
//! use litchi_rtf::RtfDocument;
//!
//! # fn main() -> Result<(), litchi_rtf::RtfError> {
//! let rtf_text = r#"{\rtf1\ansi{\fonttbl\f0\fswiss Helvetica;}\f0\pard Hello World!\par}"#;
//! let doc = RtfDocument::parse(rtf_text)?;
//! let text = doc.text();
//! # Ok::<(), litchi_rtf::RtfError>(())
//! # }
//! ```

mod annotation;
mod bookmark;
mod border;
mod character_positioning;
mod compressed;
mod data_store;
mod document;
mod document_asian_grid_compatibility;
mod document_booklet_printing;
mod document_compatibility_policy;
mod document_default_formatting;
mod document_drawing_grid;
mod document_east_asian_compatibility;
mod document_embedding_policies;
mod document_file_settings;
mod document_legacy_layout_compatibility;
mod document_line_spacing_compatibility;
mod document_origin;
mod document_output_settings;
mod document_print_layout_settings;
mod document_privacy_policies;
mod document_processing_settings;
mod document_rendering_settings;
mod document_revision_policies;
mod document_save_preferences;
mod document_style_policies;
mod document_style_restrictions;
mod document_table_layout_compatibility;
mod document_theme_languages;
mod document_variable;
mod document_view;
mod document_word_2003_compatibility;
mod document_xml_policies;
mod error;
mod external_reference;
mod field;
mod file_table;
mod form_field;
mod generated_list_marker;
mod generator;
mod hyphenation;
mod info;
mod language;
mod latent_style;
mod legacy_drawing;
mod legacy_numbering;
mod legacy_paragraph_numbering;
mod legacy_text_box;
mod lexer;
mod list;
mod mail_merge;
mod math_properties;
mod navigation_entry;
mod note_options;
mod note_separator;
mod object;
mod page_border;
mod paragraph_group;
mod parser;
mod picture;
mod picture_compatibility;
mod protection_user;
mod review_display;
mod revision_save;
mod section;
mod shape;
mod style_list_filter;
mod stylesheet;
mod table;
mod theme;
mod types;
mod user_property;
mod window_caption;
mod write_reservation;
mod writer;
mod xml_namespace;
mod xsl_transform;

// Re-exports
pub use annotation::{Annotation, AnnotationType, Revision, RevisionAuthor, RevisionType};
pub use bookmark::{Bookmark, BookmarkTable};
pub use border::{
    Border, BorderStyle, Borders, CharacterBorder, CharacterBorderStyle, CharacterShading,
    MAX_PARAGRAPH_TAB_STOPS, Shading, ShadingPattern, TabAlignment, TabLeader, TabStop, TabStops,
};
pub use character_positioning::{
    CharacterBaseline, CharacterExpansion, CharacterPositioning,
    MAX_CHARACTER_BASELINE_HALF_POINTS, MAX_CHARACTER_EXPANSION, MAX_CHARACTER_KERNING_HALF_POINTS,
    MAX_CHARACTER_SCALE_PERCENT,
};
pub use compressed::{compress, decompress, is_compressed_rtf};
pub use data_store::DocumentDataStore;
pub use document::RtfDocument;
pub use document_asian_grid_compatibility::DocumentAsianGridCompatibility;
pub use document_booklet_printing::DocumentBookletPrinting;
pub use document_compatibility_policy::{DocumentCompatibilityPolicy, DocumentFeatureThrottle};
pub use document_default_formatting::{
    DefaultCharacterProperties, DefaultFormattingDestination, DefaultParagraphProperties,
    DocumentDefaultFonts, DocumentDefaultFormatting,
};
pub use document_drawing_grid::{DocumentDrawingGrid, DrawingGridLineInterval, DrawingGridSpacing};
pub use document_east_asian_compatibility::DocumentEastAsianCompatibility;
pub use document_embedding_policies::DocumentEmbeddingPolicies;
pub use document_file_settings::DocumentFileSettings;
pub use document_legacy_layout_compatibility::DocumentLegacyLayoutCompatibility;
pub use document_line_spacing_compatibility::DocumentLineSpacingCompatibility;
pub use document_origin::{
    DocumentAutoFormatType, DocumentOrigin, DocumentOriginMetadata, HtmlEmailVersion,
};
pub use document_output_settings::DocumentOutputSettings;
pub use document_print_layout_settings::{
    DocumentPrintLayoutSettings, MAX_DOCUMENT_GUTTER_TWIPS,
};
pub use document_privacy_policies::DocumentPrivacyPolicies;
pub use document_processing_settings::{
    AbstractNumberingCleanupStatus, DocumentEventMask, DocumentProcessingSettings,
};
pub use document_rendering_settings::{
    DocumentJustificationMode, DocumentRenderingOrientation, DocumentRenderingSettings,
};
pub use document_revision_policies::DocumentRevisionPolicies;
pub use document_save_preferences::{
    DocumentReadOnlyRecommendation, DocumentSavePreferences, DocumentThumbnailPreference,
};
pub use document_style_policies::DocumentStylePolicies;
pub use document_style_restrictions::DocumentStyleRestrictions;
pub use document_table_layout_compatibility::DocumentTableLayoutCompatibility;
pub use document_theme_languages::DocumentThemeLanguages;
pub use document_variable::DocumentVariable;
pub use document_view::{
    DocumentView, DocumentViewKind, DocumentZoomKind, MAX_DOCUMENT_VIEW_SCALE_PERCENT,
};
pub use document_word_2003_compatibility::DocumentWord2003Compatibility;
pub use document_xml_policies::DocumentXmlPolicies;
pub use error::{RtfError, RtfResult};
pub use external_reference::{
    DocumentExternalReferences, MAX_DOCUMENT_EXTERNAL_REFERENCE_BYTES,
    MAX_DOCUMENT_EXTERNAL_REFERENCE_TOTAL_BYTES,
};
pub use field::CompareField;
pub use field::{
    ActiveContentField, ActiveContentFieldKind, AddressBlockCountryInclusion, AutoNumberField,
    AutoNumberFieldKind, AutoTextField, AutoTextFieldKind, AutoTextListField, AutoTextListOption,
    BibliographyField,
    BibliographyOption, BodyStoryEvent, CitationField, CitationOption, DdeField, DdeFieldKind,
    DdeRepresentation, DocumentContextField, DocumentContextFieldKind, DocumentInformationField,
    DocumentInformationFieldKind, DocumentPropertyField, DocumentVariableField, EquationField,
    ExternalIncludeField, ExternalIncludeOption, Field, FieldCodeError, FieldCodeToken,
    FieldOwner, FieldStatus, FieldSwitch, FieldType, FormulaField, GoToButtonField,
    HyperlinkCode, IfField, InfoField,
    IncludeFieldKind, IndexEntryField, IndexEntryOption, IndexField, IndexOption, LinkField,
    LinkFormatting, LinkResultOption, ListNumberField, MacroButtonField,
    MailMergeConditionalControlField,
    MailMergeConditionalControlKind, MailMergeCounterField, MailMergeCounterKind,
    MailMergeNextField, MailMergeRecipientField, MailMergeRecipientFieldKind, MergeField,
    PageBreak, ParsedFieldCode, PrintField, PromptField, PromptFieldKind, QuoteField,
    ReferenceCode,
    SectionBreak, SequenceField, SetField, StoryEvent, StoryField, StyleReferenceField,
    StyleReferenceFieldOption, SymbolField, TableOfAuthoritiesEntryField,
    TableOfAuthoritiesEntryOption,
    TableOfAuthoritiesField, TableOfAuthoritiesOption, TableOfContentsEntryField,
    TableOfContentsEntryOption, TableOfContentsField, TableOfContentsOption, UserIdentityField,
    UserIdentityFieldKind, UserIdentityFormatting, parse_field_code,
};
pub use field::{AdvanceField, AdvanceFieldAdjustment, AdvanceFieldOperation};
pub use file_table::{FileLocation, FileSystemValidity, FileTable, FileTableEntry};
pub use form_field::{FormField, FormFieldType, FormTextType};
pub use generated_list_marker::{GeneratedListMarker, GeneratedListMarkerKind};
pub use generator::DocumentGenerator;
pub use hyphenation::{
    DocumentHyphenation, MAX_HYPHENATION_CONSECUTIVE_LINES, MAX_HYPHENATION_HOT_ZONE_TWIPS,
};
pub use info::{DocumentInfo, DocumentProtection, ProtectionLevel, ProtectionType, RtfTimestamp};
pub use language::{DocumentLanguageDefaults, LanguageId};
pub use latent_style::{LatentStyleException, LatentStyles};
pub use legacy_drawing::{
    LegacyCallout, LegacyCalloutAttachment, LegacyCalloutType, LegacyDrawing, LegacyDrawingArrow,
    LegacyDrawingArrowFill, LegacyDrawingArrowSize, LegacyDrawingColor, LegacyDrawingFill,
    LegacyDrawingFillPattern, LegacyDrawingGeometry, LegacyDrawingLine, LegacyDrawingLineStyle,
    LegacyDrawingPoint, LegacyDrawingPrimitive, LegacyDrawingProperties, LegacyDrawingShadow,
    MAX_LEGACY_DRAWING_DEPTH, MAX_LEGACY_DRAWING_POINTS, MAX_LEGACY_DRAWING_PRIMITIVES,
    MAX_LEGACY_DRAWING_TOTAL_POINTS, MAX_LEGACY_DRAWINGS,
};
pub use legacy_numbering::{
    LegacyNumberingAlignment, LegacyNumberingFormat, LegacySectionNumbering,
    LegacySectionNumberingLevel,
};
pub use legacy_paragraph_numbering::{
    LegacyParagraphNumbering, LegacyParagraphNumberingAlignment, LegacyParagraphNumberingBidi,
    LegacyParagraphNumberingFormat, LegacyParagraphNumberingLevel,
    LegacyParagraphNumberingRevision, LegacyParagraphNumberingUnderline,
    MAX_LEGACY_PARAGRAPH_NUMBERING_RECORDS, MAX_LEGACY_PARAGRAPH_NUMBERING_TEXT_BYTES,
};
pub use legacy_text_box::{
    LegacyHorizontalAnchor, LegacyTextBox, LegacyTextDirection, LegacyVerticalAnchor,
};
pub use lexer::CharacterSet;
pub use list::{
    List, ListFollow, ListJustification, ListLevel, ListLevelType, ListOverride, ListOverrideLevel,
    ListOverrideTable, ListTable,
};
pub use mail_merge::{
    MAX_MAIL_MERGE_FIELD_MAPPINGS, MAX_MAIL_MERGE_NESTING_DEPTH, MAX_MAIL_MERGE_RECIPIENT_DATA,
    MAX_MAIL_MERGE_STRING_BYTES, MAX_MAIL_MERGE_TOTAL_BYTES, MailMerge, MailMergeColumnIndex,
    MailMergeDataSourceObject, MailMergeDataSourceType, MailMergeFieldMapping,
};
pub use math_properties::{
    DocumentMathProperties, MathBinaryOperatorBreak, MathBinarySubtractionBreak, MathFlag,
    MathJustification, MathLimitPlacement,
};
pub use navigation_entry::{IndexEntry, IndexPageReference, NavigationEntry, TableOfContentsEntry};
pub use note_options::{
    EndnoteRestart, FootnoteRestart, NoteNumberingStyle, NoteOptions, NotePlacement,
    PresentNoteKinds,
};
pub use note_separator::{
    NoteSeparator, NoteSeparatorElement, NoteSeparatorKind, NoteSeparatorTable,
};
pub use object::{
    EmbeddedObject, MAX_EMBEDDED_OBJECTS, MAX_OBJECT_DATA_BYTES, MAX_OBJECT_METADATA_BYTES,
    ObjectKind, ObjectResultKind, OleObjectHeader,
};
pub use page_border::{
    PageBorder, PageBorderAppliesTo, PageBorderDepth, PageBorderOffset, PageBorderSide,
    PageBorderStyle, PageBorders,
};
pub use paragraph_group::{ParagraphGroupProperty, ParagraphGroupPropertyTable};
pub use picture::{
    ImageType, MAX_PICTURE_SHAPE_PROPERTIES, MAX_PICTURE_SHAPE_PROPERTY_BYTES, Picture,
    PictureBitmapMetadata, PictureCrop, PictureIdentity, PictureShapeProperties, detect_image_type,
};
pub use picture_compatibility::{
    MAX_PICTURE_COMPATIBILITY_RECORDS, PictureCompatibilityKind, PictureCompatibilityRecord,
};
pub use protection_user::{
    MAX_PROTECTION_USER_BYTES, MAX_PROTECTION_USER_TOTAL_BYTES, MAX_PROTECTION_USERS,
    ProtectionUser, ProtectionUserTable,
};
pub use review_display::DocumentReviewDisplay;
pub use revision_save::RevisionSaveMetadata;
pub use section::{
    HeaderFooter, HeaderFooterParagraph, HeaderFooterType, MAX_SECTION_COLUMN_TWIPS,
    MAX_SECTION_COLUMNS, MAX_SECTION_LINE_DISTANCE, MAX_SECTION_LINE_INCREMENT,
    MAX_SECTION_LINE_START, Note, PageNumberFormat, PageOrientation, Section, SectionBreakType,
    SectionColumn, SectionColumns, SectionLineNumberRestart, SectionLineNumbering,
    SectionProperties, VerticalAlignment,
};
pub use section::{SectionFootnotePlacement, SectionNoteOptions};
pub use shape::{
    Fill, FillType, GradientDirection, MAX_SHAPE_PROPERTY_BINARY_BYTES, OfficeArtColor,
    OfficeArtOpacity, Shape, ShapeGeometry, ShapeGroup, ShapeGroupChild, ShapeGroupInfo,
    ShapeHorizontalAnchor, ShapeLine, ShapeProperty, ShapeResult, ShapeRotationDegrees,
    ShapeThemeColor, ShapeThemeValue, ShapeTwips, ShapeType, ShapeVerticalAnchor, ShapeWrapSide,
    ShapeWrapStyle, ShapeZOrder, StoryDrawing, WrapMode,
};
pub use style_list_filter::{DocumentStyleListFilter, DocumentStyleSortMethod};
pub use stylesheet::{Style, StyleSheet, StyleType};
pub use table::{
    Cell, CellNestedTable, CellStoryEvent, CellStoryReference, FloatingTablePosition,
    TableCellCoordinate, TableCellPath,
    MAX_FLOATING_TABLE_DISTANCE_TWIPS, MAX_TABLE_CELLS_PER_ROW, MAX_TABLE_DISTANCE_TWIPS,
    MAX_TABLE_GEOMETRY_TWIPS, MAX_TABLE_NESTING_DEPTH, MAX_TABLE_ROW_INDEX,
    MAX_TABLE_WIDTH_PERCENT, Row, Table, TableAutoformatFlag, TableAutoformatFlags,
    TableCellBorderSide, TableCellBorders, TableCellLayout, TableCellMergeAxis, TableCellMergeRole,
    TableCellMergeState, TableCellTextFlow, TableCellVerticalAlignment, TableDistanceKind,
    TableDistanceScope, TableDistanceTarget, TableDistanceUnit, TableEdge, TableEdgeDistances,
    TableHorizontalPosition, TableHorizontalReference, TableIndent, TableIndentUnit,
    TablePreferredWidth, TablePreferredWidthUnit, TableRowAlignment, TableRowBandIndex,
    TableRowBanding, TableRowBorderSide, TableRowBorders, TableRowGeometry, TableRowHeight,
    TableRowLayout, TableShading, TableSideDistance, TableVerticalPosition, TableVerticalReference,
    TableWrapDistances,
};
pub use theme::DocumentTheme;
pub use types::{
    Alignment, AssociatedCharacterBaseline, AssociatedCharacterFormatting,
    AssociatedUnderlineStyle, CharacterGrid, CharacterType, Color, ColorRef, ColorTable,
    DocumentElement, EmbeddedFont, EmbeddedFontFormat, Font, FontFamily, FontPitch, FontRef,
    FontTable, Formatting, Indentation, MAX_PARAGRAPH_DROP_CAP_LINES, Paragraph,
    ParagraphContent, ParagraphDropCap, ParagraphDropCapKind, ParagraphFontAlignment,
    ParagraphLineBreaking, ParagraphLogicalIndentation, ParagraphSpacingPolicy,
    ParagraphWrapping, Run, Spacing, StyleBlock, TextDirection, UnderlineStyle,
};
pub use user_property::{UserProperty, UserPropertyDateTime, UserPropertyType, UserPropertyValue};
pub use window_caption::{DocumentWindowCaption, MAX_WINDOW_CAPTION_BYTES};
pub use write_reservation::{
    DocumentWriteReservations, LegacyWriteReservation, MAX_WRITE_RESERVATION_BYTES,
    WriteReservationHash,
};
pub use writer::{
    DEFAULT_TAB_WIDTH_TWIPS, DefaultTabWidthPolicy, MAX_DEFAULT_TAB_WIDTH_TWIPS, RtfWriter,
    WriterOptions,
};
pub use xml_namespace::XmlNamespace;
pub use xsl_transform::{
    DocumentXslTransform, DocumentXslTransformUsage, MAX_DOCUMENT_XSL_TRANSFORM_LOCATION_BYTES,
};
