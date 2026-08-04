//! Canonical WordprocessingML (`.docx`) APIs.
//!
//! The concise modules own format semantics while [`litchi_opc`] remains the
//! explicit low-level package graph.

#![forbid(unsafe_code)]

mod error;

pub mod alt;
pub mod color;
pub mod enums;
pub mod field;
pub mod font;
pub mod format;
pub mod glossary;
pub mod hyperlink;
pub mod mail_merge;
pub mod modern_comments;
pub mod settings;
pub mod statistics;
pub mod web;

pub use enums::{WdHeaderFooter, WdOrientation, WdSectionStart, WdStyleType};
pub use error::{Error, Result};
pub use field::{
    ActiveContentField, ActiveContentFieldKind, AddressBlockCountryInclusion, AdvanceField,
    AdvanceFieldAdjustment, AdvanceFieldOperation, AutoNumberField, AutoNumberFieldKind,
    AutoTextField, AutoTextFieldKind, AutoTextListField, AutoTextListOption, BarcodeField,
    BibliographyField, BidiOutlineField, CitationField, CompareField, DatabaseField, DdeField,
    DdeFieldKind, DdeRepresentation, DocumentContextField, DocumentContextFieldKind,
    DocumentInformationField, DocumentInformationFieldKind, DocumentPropertyField,
    DocumentVariableField, EmbedField, EquationField, ExternalIncludeField, ExternalIncludeOption,
    Field, FieldSwitch, FormulaField, GoToButtonField, HyperlinkField, IfField, IncludeFieldKind,
    IndexEntryField, IndexField, IndexSortOrder, InfoField, LegacyFormField, LegacyFormFieldKind,
    LinkField, LinkFormatting, LinkResultOption, ListNumberField, MacroButtonField,
    MailMergeConditionalControlField, MailMergeConditionalControlKind, MailMergeCounterField,
    MailMergeCounterKind, MailMergeDataField, MailMergeNextField, MailMergeRecipientField,
    MailMergeRecipientFieldKind, MergeField, PrintField, PrivateField, PromptField,
    PromptFieldKind, QuoteField, ReferenceField, ReferenceFieldKind, ReferenceFieldOption,
    ReferencedDocumentField, SequenceField, SetField, ShapeField, StyleReferenceField,
    StyleReferenceFieldOption, SymbolField, TableOfAuthoritiesEntryField, TableOfAuthoritiesField,
    TableOfContentsEntryField, TableOfContentsField, TableOfContentsLevelRange,
    TableOfContentsSwitch, UserIdentityField, UserIdentityFieldKind, UserIdentityFormatting,
};
pub use format::{ImageFormat, LineSpacing, ParagraphAlignment, TableBorderStyle, UnderlineStyle};
pub use hyperlink::Hyperlink;
pub use mail_merge::{
    MailMergeConformance, MailMergeDataSourceObject, MailMergeDataType, MailMergeDestination,
    MailMergeFieldMap, MailMergeFieldMappingType, MailMergeMainDocumentType, MailMergeRecipient,
    MailMergeRecipients, MailMergeSettings, MailMergeSource, MailMergeTarget,
    RECIPIENT_CONTENT_TYPE, parse_settings_mail_merge,
};
pub use modern_comments::{
    CommentExtension, CommentIdMapping, CommentReaction, CommentReactionInfo, CommentReactionUser,
    ExtensibleComment, ModernCommentConformance, ModernCommentMetadata,
    ModernCommentRelationshipIds, Person, PresenceInfo, load_modern_comment_metadata,
    parse_comments_extended, parse_comments_extensible, parse_comments_ids, parse_people,
    store_modern_comment_metadata, write_comments_extended, write_comments_extensible,
    write_comments_ids, write_people,
};
pub use settings::{
    ColorSchemeIndex, ColorSchemeMapping, ColorSchemeSlot, CompatFlag, CompatibilityOption,
    CompatibilitySetting, DocumentView, MAX_LANGUAGE_TAG_LENGTH, MAX_SETTINGS_XML_BYTES,
    MAX_SETTINGS_XML_DEPTH, MAX_SETTINGS_XML_NODES, NoteNumberFormat, NoteNumberingProperties,
    NoteNumberingRestart, NotePosition, ParseCompatFlagError, ParseNoteNumberFormatError,
    ParseNotePositionError, ProofState, ProofingState, ProtectionType, Settings,
    ThemeFontLanguages,
};
pub use statistics::{
    DocumentStatistics, count_characters, count_characters_no_spaces, count_words,
    estimate_line_count, estimate_page_count,
};
