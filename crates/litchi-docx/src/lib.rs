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
pub use statistics::{
    DocumentStatistics, count_characters, count_characters_no_spaces, count_words,
    estimate_line_count, estimate_page_count,
};
