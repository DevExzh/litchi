#![forbid(unsafe_code)]
#![cfg_attr(
    not(test),
    deny(
        clippy::expect_used,
        clippy::indexing_slicing,
        clippy::panic,
        clippy::panic_in_result_fn,
        clippy::string_slice,
        clippy::todo,
        clippy::unimplemented,
        clippy::unreachable,
        clippy::unwrap_used
    )
)]
#![allow(
    missing_docs,
    reason = "the retained model vocabulary mirrors self-explanatory RTF specification names"
)]
#![allow(
    non_ascii_idents,
    reason = "zerocopy's wire-layout derives emit compiler-generated identifiers"
)]

//! Typed Rich Text Format documents.
//!
//! [`Document`] is the ordinary immutable, lifetime-free snapshot. Parsing and
//! opening live in [`read`], streaming serialization lives in [`mod@write`], and
//! Outlook's compressed-RTF transport lives in [`transport`]. The retained
//! mutable parser model is an advanced interface under [`raw`].
//!
//! # Architecture
//!
//! The crate keeps format mechanics behind responsibility-focused boundaries:
//!
//! - `api` owns the immutable [`Document`] facade and borrowed [`text::Story`]
//!   views.
//! - `codec` owns bounded transport decoding, tokenization, parsing, and
//!   serialization.
//! - `model` owns the retained lossless RTF snapshot used by [`raw`].
//! - `resource` owns borrowed font and color facades.
//! - `text`, `content`, `drawing`, `review`, `metadata`, `numbering`, and
//!   `policy` group the corresponding native RTF vocabularies.
//!
//! These private ownership modules keep dependency direction explicit. Stable
//! public entry points use the concise facade modules below rather than
//! mirroring internal file paths.
//!
//! # Example
//!
//! ```rust,no_run
//! use litchi_rtf::{Document, Result};
//!
//! # fn main() -> Result<()> {
//! let rtf_text = r#"{\rtf1\ansi{\fonttbl{\f0\fswiss Helvetica;}}\f0\pard Hello World!\par}"#;
//! let doc = Document::parse(rtf_text)?;
//! assert_eq!(doc.text(), "Hello World!\n");
//! assert_eq!(doc.body().paragraphs().count(), 1);
//! # Ok(())
//! # }
//! ```

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "items stay grouped by RTF feature area rather than by item kind"
)]
mod api;
mod codec;
mod content;
mod drawing;
pub mod edit;
pub mod metadata;
mod model;
mod numbering;
mod policy;
mod resource;
pub mod review;
pub mod streaming;
pub mod tail_append;
pub mod text;
pub mod validation;

pub use content::{field, math, section, table};
pub use drawing::{picture, shape};
pub use numbering::list;

// Crate-private compatibility aliases keep the retained model isolated from
// the source layout. New code should depend on its responsibility module.
use api::story;
use codec::{compressed, error, lexer, limits, parser, writer};
use content::{equation, form_field, math_properties};
use drawing::{legacy_drawing, legacy_text_box, object, page_border, picture_compatibility};
use metadata::{
    custom_xml, data_store, document_origin, document_variable, external_reference, file_table,
    generator, info, mail_merge, theme, user_property, window_caption, write_reservation,
    xml_namespace, xsl_transform,
};
use model::{document, types};
use numbering::{
    generated_list_marker, legacy_numbering, legacy_paragraph_numbering, navigation_entry,
};
use policy::{
    document_asian_grid_compatibility, document_booklet_printing, document_compatibility_policy,
    document_drawing_grid, document_east_asian_compatibility, document_embedding_policies,
    document_file_settings, document_legacy_layout_compatibility,
    document_line_spacing_compatibility, document_output_settings, document_print_layout_settings,
    document_privacy_policies, document_processing_settings, document_rendering_settings,
    document_revision_policies, document_save_preferences, document_style_policies,
    document_style_restrictions, document_table_layout_compatibility, document_theme_languages,
    document_view, document_word_2003_compatibility, document_xml_policies,
};
use review::{
    annotation, bookmark, editable_region, note_options, note_separator, protection_range,
    protection_user, review_display, revision_save,
};
use text::{
    border, character_positioning, document_default_formatting, hyphenation, kinsoku, language,
    latent_style, paragraph_group, style_list_filter, stylesheet,
};

/// Concise document-opening facade.
pub mod read {
    pub use crate::api::Document;
    pub use crate::codec::limits::ParseLimits as Limits;
}

/// Concise streaming writer facade.
pub mod write {
    pub use crate::codec::writer::{
        Charset, DEFAULT_TAB_WIDTH_TWIPS, DefaultTabWidthPolicy as TabWidth,
        MAX_DEFAULT_TAB_WIDTH_TWIPS, RtfWriter as Writer, WriterOptions as Options,
    };
    pub use crate::model::types::Formatting as Format;
    pub use crate::streaming::{
        StreamingRtfError, StreamingRtfLimits, StreamingRtfOptions, StreamingRtfWriter,
    };
}

/// Bounded Outlook/MAPI compressed-RTF transport codec.
pub mod transport {
    pub use crate::codec::compressed::{
        DEFAULT_MAX_DECOMPRESSED_RTF_BYTES, DecompressionLimits as Limits, compress, decompress,
        decompress_with_limits, is_compressed_rtf,
    };
}

/// Advanced retained RTF model.
///
/// This interface exposes format-specific structure. Ordinary read-only code
/// should prefer [`crate::Document`].
pub mod raw {
    pub use crate::model::document::RtfDocument as Document;
    pub use crate::native::*;
}

/// Unsupported syntax retained as bounded inert data.
pub mod opaque {
    pub use crate::model::opaque::{Anchor, Context, Kind, Node};
}

/// Font resources and checked references.
pub mod font {
    pub use crate::resource::font::{Catalog, Embedded, Font, Iter, LookupError};
    pub use crate::types::{
        EmbeddedFontFormat as EmbeddedFormat, FontCharset as Charset, FontFamily as Family,
        FontPage as Page, FontPitch as Pitch, FontTheme as Theme,
    };
}

/// Color resources and checked references.
pub mod color {
    pub use crate::resource::color::{Color, Iter, Palette, Value};
}

/// Named document styles and stylesheet policy values.
pub mod style {
    pub use crate::latent_style::{LatentStyleException as LatentException, LatentStyles};
    pub use crate::style_list_filter::{
        DocumentStyleListFilter as ListFilter, DocumentStyleSortMethod as SortMethod,
    };
    pub use crate::stylesheet::{
        Style, StyleType as Kind, TableStyleConditionalFormatting as TableConditional,
    };
}

pub use api::Document;
pub use codec::error::{RtfError as Error, RtfResult as Result};
pub use tail_append::{
    DurableTailAppendPatch, PlainParagraph, PlainRun, TailAppendCommit, TailAppendDiagnostics,
    TailAppendEdit, TailAppendError, TailAppendLimits, TailAppendOutputProgress, TailAppendPatch,
    TailAppendPublicationError, TailAppendPublicationLimits, TailAppendPublicationPlan,
    TailAppendPublicationReport, TailSelector,
};
pub use validation::{
    ValidationCheck, ValidationCounts, ValidationDependency, ValidationLimits, ValidationReport,
    ValidationStatus,
};

// Canonical native RTF vocabulary used by the retained model and writer. The
// ordinary facade selects a smaller contextual subset from this module.
mod native {
    use super::{
        annotation, bookmark, border, character_positioning, compressed, custom_xml, data_store,
        document, document_asian_grid_compatibility, document_booklet_printing,
        document_compatibility_policy, document_default_formatting, document_drawing_grid,
        document_east_asian_compatibility, document_embedding_policies, document_file_settings,
        document_legacy_layout_compatibility, document_line_spacing_compatibility, document_origin,
        document_output_settings, document_print_layout_settings, document_privacy_policies,
        document_processing_settings, document_rendering_settings, document_revision_policies,
        document_save_preferences, document_style_policies, document_style_restrictions,
        document_table_layout_compatibility, document_theme_languages, document_variable,
        document_view, document_word_2003_compatibility, document_xml_policies, editable_region,
        equation, error, external_reference, field, file_table, form_field, generated_list_marker,
        generator, hyphenation, info, kinsoku, language, latent_style, legacy_drawing,
        legacy_numbering, legacy_paragraph_numbering, legacy_text_box, limits, list, mail_merge,
        math, math_properties, navigation_entry, note_options, note_separator, object, page_border,
        paragraph_group, picture, picture_compatibility, protection_range, protection_user,
        review_display, revision_save, section, shape, style_list_filter, stylesheet, table, theme,
        types, user_property, window_caption, write_reservation, writer, xml_namespace,
        xsl_transform,
    };

    pub use annotation::{Annotation, AnnotationType, Revision, RevisionAuthor, RevisionType};
    pub use bookmark::{Bookmark, BookmarkTable};
    pub use border::{
        Border, BorderStyle, Borders, CharacterBorder, CharacterBorderStyle, CharacterShading,
        MAX_PARAGRAPH_TAB_STOPS, Shading, ShadingPattern, TabAlignment, TabLeader, TabStop,
        TabStops,
    };
    pub use character_positioning::{
        CharacterBaseline, CharacterExpansion, CharacterPositioning,
        MAX_CHARACTER_BASELINE_HALF_POINTS, MAX_CHARACTER_EXPANSION,
        MAX_CHARACTER_KERNING_HALF_POINTS, MAX_CHARACTER_SCALE_PERCENT,
    };
    pub use compressed::{
        DEFAULT_MAX_DECOMPRESSED_RTF_BYTES, DecompressionLimits, compress, decompress,
        decompress_with_limits, is_compressed_rtf,
    };
    pub use custom_xml::{CustomXmlAttribute, CustomXmlTag};
    pub use data_store::DocumentDataStore;
    pub use document::RtfDocument;
    pub use document_asian_grid_compatibility::DocumentAsianGridCompatibility;
    pub use document_booklet_printing::DocumentBookletPrinting;
    pub use document_compatibility_policy::{DocumentCompatibilityPolicy, DocumentFeatureThrottle};
    pub use document_default_formatting::{
        DefaultCharacterProperties, DefaultFormattingDestination, DefaultParagraphProperties,
        DocumentDefaultFonts, DocumentDefaultFormatting,
    };
    pub use document_drawing_grid::{
        DocumentDrawingGrid, DrawingGridLineInterval, DrawingGridSpacing,
    };
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
    pub use editable_region::EditableRegion;
    pub use equation::{
        EquationAlignment, EquationArray, EquationBox, EquationBracket, EquationDisplace,
        EquationGroup, EquationIntegral, EquationIntegralSymbol, EquationModel, EquationOverstrike,
        EquationScript, EquationSegment, EquationSpacing, EquationSwitch,
    };
    pub use error::{RtfError, RtfResult};
    pub use external_reference::{
        DocumentExternalReferences, MAX_DOCUMENT_EXTERNAL_REFERENCE_BYTES,
        MAX_DOCUMENT_EXTERNAL_REFERENCE_TOTAL_BYTES,
    };
    pub use field::CompareField;
    pub use field::{
        ActiveContentField, ActiveContentFieldKind, AddressBlockCountryInclusion, AutoNumberField,
        AutoNumberFieldKind, AutoTextField, AutoTextFieldKind, AutoTextListField,
        AutoTextListOption, BarcodeDisplayField, BarcodeDisplayFieldKind, BarcodeField,
        BibliographyField, BibliographyOption, BidiOutlineField, BodyStoryEvent, CitationField,
        CitationOption, ColumnBreak, DatabaseField, DdeField, DdeFieldKind, DdeRepresentation,
        DocumentContextField, DocumentContextFieldKind, DocumentInformationField,
        DocumentInformationFieldKind, DocumentPropertyField, DocumentVariableField, EmbedField,
        EquationField, ExternalIncludeField, ExternalIncludeOption, Field, FieldCodeError,
        FieldCodeToken, FieldOwner, FieldStatus, FieldSwitch, FieldType, FormulaField,
        GoToButtonField, HyperlinkCode, HyperlinkField, IfField, IncludeFieldKind, IndexEntryField,
        IndexEntryOption, IndexField, IndexOption, InfoField, LegacyFormField, LegacyFormFieldKind,
        LinkField, LinkFormatting, LinkResultOption, ListNumberField, MacroButtonField,
        MailMergeConditionalControlField, MailMergeConditionalControlKind, MailMergeCounterField,
        MailMergeCounterKind, MailMergeDataField, MailMergeNextField, MailMergeRecipientField,
        MailMergeRecipientFieldKind, MergeField, PageBreak, ParsedFieldCode, PrintField,
        PrivateField, PromptField, PromptFieldKind, QuoteField, ReferenceCode, ReferenceField,
        ReferenceFieldKind, ReferencedDocumentField, SectionBreak, SequenceField, SetField,
        ShapeField, SoftBreak, SoftBreakKind, StoryEvent, StoryField, StyleReferenceField,
        StyleReferenceFieldOption, SymbolField, TableOfAuthoritiesEntryField,
        TableOfAuthoritiesEntryOption, TableOfAuthoritiesField, TableOfAuthoritiesOption,
        TableOfContentsEntryField, TableOfContentsEntryOption, TableOfContentsField,
        TableOfContentsOption, UserIdentityField, UserIdentityFieldKind, UserIdentityFormatting,
        parse_field_code,
    };
    pub use field::{AdvanceField, AdvanceFieldAdjustment, AdvanceFieldOperation};
    pub use file_table::{FileLocation, FileSystemValidity, FileTable, FileTableEntry};
    pub use form_field::{FormField, FormFieldType, FormTextType};
    pub use generated_list_marker::{GeneratedListMarker, GeneratedListMarkerKind};
    pub use generator::DocumentGenerator;
    pub use hyphenation::{
        DocumentHyphenation, MAX_HYPHENATION_CONSECUTIVE_LINES, MAX_HYPHENATION_HOT_ZONE_TWIPS,
    };
    pub use info::{
        DocumentInfo, DocumentProtection, ProtectionLevel, ProtectionType, RtfTimestamp,
    };
    pub use kinsoku::DocumentKinsoku;
    pub use language::{DocumentLanguageDefaults, LanguageId};
    pub use latent_style::{LatentStyleException, LatentStyles};
    pub use legacy_drawing::{
        LegacyCallout, LegacyCalloutAttachment, LegacyCalloutType, LegacyDrawing,
        LegacyDrawingArrow, LegacyDrawingArrowFill, LegacyDrawingArrowSize, LegacyDrawingColor,
        LegacyDrawingFill, LegacyDrawingFillPattern, LegacyDrawingGeometry, LegacyDrawingLine,
        LegacyDrawingLineStyle, LegacyDrawingPoint, LegacyDrawingPrimitive,
        LegacyDrawingProperties, LegacyDrawingShadow, MAX_LEGACY_DRAWING_DEPTH,
        MAX_LEGACY_DRAWING_POINTS, MAX_LEGACY_DRAWING_PRIMITIVES, MAX_LEGACY_DRAWING_TOTAL_POINTS,
        MAX_LEGACY_DRAWINGS,
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
    pub use limits::ParseLimits;
    pub use list::{
        List, ListFollow, ListJustification, ListLevel, ListLevelType, ListOverride,
        ListOverrideLevel, ListOverrideTable, ListTable,
    };
    pub use mail_merge::{
        MAX_MAIL_MERGE_FIELD_MAPPINGS, MAX_MAIL_MERGE_NESTING_DEPTH, MAX_MAIL_MERGE_RECIPIENT_DATA,
        MAX_MAIL_MERGE_STRING_BYTES, MAX_MAIL_MERGE_TOTAL_BYTES, MailMerge, MailMergeColumnIndex,
        MailMergeDataSourceObject, MailMergeDataSourceType, MailMergeFieldMapping,
    };
    pub use math::{
        MathElement, MathElementRole, MathMatrixColumn, MathMatrixRow, MathObject, MathProperties,
        MathPropertiesKind, MathProperty, MathPropertyName, MathRun, MathStructure,
        MathStructureChild, MathStructureKind, MathZone, MathZoneKind,
    };
    pub use math_properties::{
        DocumentMathProperties, MathBinaryOperatorBreak, MathBinarySubtractionBreak, MathFlag,
        MathJustification, MathLimitPlacement,
    };
    pub use navigation_entry::{
        IndexEntry, IndexPageReference, NavigationEntry, TableOfContentsEntry,
    };
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
        ImageType, MAX_PICTURE_SHAPE_PROPERTIES, MAX_PICTURE_SHAPE_PROPERTY_BYTES, MAX_PICTURES,
        Picture, PictureBitmapMetadata, PictureCrop, PictureIdentity, PictureShapeProperties,
        detect_image_type,
    };
    pub use picture_compatibility::{
        MAX_PICTURE_COMPATIBILITY_RECORDS, PictureCompatibilityKind, PictureCompatibilityRecord,
    };
    pub use protection_range::ProtectionRange;
    pub use protection_user::{
        MAX_PROTECTION_USER_BYTES, MAX_PROTECTION_USER_TOTAL_BYTES, MAX_PROTECTION_USERS,
        ProtectionUser, ProtectionUserTable,
    };
    pub use review_display::DocumentReviewDisplay;
    pub use revision_save::RevisionSaveMetadata;
    pub use section::{
        HeaderFooter, HeaderFooterParagraph, HeaderFooterType, MAX_PAGE_NUMBER_HEADING_LEVEL,
        MAX_SECTION_COLUMN_TWIPS, MAX_SECTION_COLUMNS, MAX_SECTION_LINE_DISTANCE,
        MAX_SECTION_LINE_GRID_TWIPS, MAX_SECTION_LINE_INCREMENT, MAX_SECTION_LINE_START, Note,
        PageNumberFormat, PageNumberHeadingSeparator, PageNumberRestart, PageOrientation, Section,
        SectionBreakType, SectionColumn, SectionColumns, SectionDocumentGrid,
        SectionDocumentGridType, SectionLineNumberRestart, SectionLineNumbering,
        SectionPageNumberHeading, SectionProperties, SectionRendering, VerticalAlignment,
    };
    pub use section::{SectionFootnotePlacement, SectionNoteOptions};
    pub use shape::{
        Fill, FillType, GradientDirection, MAX_SHAPE_PROPERTY_BINARY_BYTES, OfficeArtColor,
        OfficeArtOpacity, Shape, ShapeGeometry, ShapeGroup, ShapeGroupChild, ShapeGroupInfo,
        ShapeHorizontalAnchor, ShapeHyperlink, ShapeLine, ShapeProperty, ShapeResult,
        ShapeRotationDegrees, ShapeThemeColor, ShapeThemeValue, ShapeTwips, ShapeType,
        ShapeVerticalAnchor, ShapeWrapSide, ShapeWrapStyle, ShapeZOrder, StoryDrawing, WrapMode,
    };
    pub use style_list_filter::{DocumentStyleListFilter, DocumentStyleSortMethod};
    pub use stylesheet::{Style, StyleSheet, StyleType, TableStyleConditionalFormatting};
    pub use table::{
        Cell, CellNestedTable, CellRevision, CellRevisionKind, CellStoryEvent, CellStoryReference,
        FloatingTablePosition, MAX_FLOATING_TABLE_DISTANCE_TWIPS, MAX_TABLE_CELLS_PER_ROW,
        MAX_TABLE_DISTANCE_TWIPS, MAX_TABLE_GEOMETRY_TWIPS, MAX_TABLE_NESTING_DEPTH,
        MAX_TABLE_ROW_INDEX, MAX_TABLE_WIDTH_PERCENT, Row, Table, TableAutoformatFlag,
        TableAutoformatFlags, TableCellBorderSide, TableCellBorders, TableCellCoordinate,
        TableCellLayout, TableCellMergeAxis, TableCellMergeRole, TableCellMergeState,
        TableCellPath, TableCellTextFlow, TableCellVerticalAlignment, TableDistanceKind,
        TableDistanceScope, TableDistanceTarget, TableDistanceUnit, TableEdge, TableEdgeDistances,
        TableHorizontalPosition, TableHorizontalReference, TableIndent, TableIndentUnit,
        TablePreferredWidth, TablePreferredWidthUnit, TableRowAlignment, TableRowBandIndex,
        TableRowBanding, TableRowBorderSide, TableRowBorders, TableRowCellDefaults,
        TableRowGeometry, TableRowHeight, TableRowLayout, TableShading, TableSideDistance,
        TableStyleBorderSide, TableStyleDefaultBorders, TableVerticalPosition,
        TableVerticalReference, TableWrapDistances,
    };
    pub use theme::DocumentTheme;
    pub use types::{
        Alignment, AnimatedTextEffect, AssociatedCharacterBaseline, AssociatedCharacterFormatting,
        AssociatedUnderlineStyle, CharacterGrid, CharacterType, Color, ColorRef, ColorTable,
        DocumentElement, EmbeddedFont, EmbeddedFontFormat, EmphasisMark, FitText, Font,
        FontCharset, FontFamily, FontPage, FontPitch, FontRef, FontTable, FontTheme, Formatting,
        Indentation, MAX_PARAGRAPH_DROP_CAP_LINES, Paragraph, ParagraphContent, ParagraphDropCap,
        ParagraphDropCapKind, ParagraphFontAlignment, ParagraphLineBreaking,
        ParagraphLogicalIndentation, ParagraphSpacingPolicy, ParagraphWrapping, RevisionMetadata,
        Run, Spacing, StyleBlock, TextDirection, UnderlineStyle,
    };
    pub use user_property::{
        UserProperty, UserPropertyDateTime, UserPropertyType, UserPropertyValue,
    };
    pub use window_caption::{DocumentWindowCaption, MAX_WINDOW_CAPTION_BYTES};
    pub use write_reservation::{
        DocumentWriteReservations, LegacyWriteReservation, MAX_WRITE_RESERVATION_BYTES,
        WriteReservationHash,
    };
    pub use writer::{
        Charset, DEFAULT_TAB_WIDTH_TWIPS, DefaultTabWidthPolicy, MAX_DEFAULT_TAB_WIDTH_TWIPS,
        RtfWriter, WriterOptions,
    };
    pub use xml_namespace::XmlNamespace;
    pub use xsl_transform::{
        DocumentXslTransform, DocumentXslTransformUsage, MAX_DOCUMENT_XSL_TRANSFORM_LOCATION_BYTES,
    };
}

#[cfg(doc)]
pub(crate) use native::*;

#[cfg(not(doc))]
#[doc(hidden)]
pub use native::*;
