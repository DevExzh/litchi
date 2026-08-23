//! Shared implementation imports for the semantic subdomains.

pub(super) use super::super::state::Document;
pub(super) use crate::bookmark::Bookmark;
pub(super) use crate::comment::Comment;
pub(super) use crate::footnote::Footnote;
pub(super) use crate::header_footer::HeaderFooter;
pub(super) use crate::hyperlink::Hyperlink;
pub(super) use crate::package::{Error as PackageError, Result};
pub(super) use crate::paragraph::{Paragraph, Run};
pub(super) use crate::parts::associated_strings::DocumentAssociatedStrings;
pub(super) use crate::parts::auto_summary::DocumentAutoSummary;
pub(super) use crate::parts::captions::CaptionTables;
pub(super) use crate::parts::document_properties::DocumentProperties;
pub(super) use crate::parts::embedded_fonts::DocumentEmbeddedFonts;
pub(super) use crate::parts::fib::FileInformationBlock;
pub(super) use crate::parts::fields::{
    ActiveContentField, AdvanceField, AutoNumberField, AutoTextField, AutoTextListField,
    BarcodeField, BidiOutlineField, CompareField, DdeField, DocumentContextField,
    DocumentInformationField, DocumentPropertyField, DocumentVariableField, EmbedField,
    EquationField, ExternalIncludeField, Field, FieldStory, FieldText, FieldsTable, FormulaField,
    GoToButtonField, HyperlinkField, IfField, IndexEntryField, IndexField, InfoField,
    LegacyFormField, LinkField, ListNumberField, MacroButtonField,
    MailMergeConditionalControlField, MailMergeCounterField, MailMergeDataField,
    MailMergeNextField, MailMergeRecipientField, MergeField, PrintField, PrivateField, PromptField,
    QuoteField, ReferenceField, ReferencedDocumentField, SequenceField, SetField, ShapeField,
    StyleReferenceField, SymbolField, TableOfAuthoritiesEntryField, TableOfAuthoritiesField,
    TableOfContentsEntryField, TableOfContentsField, UserIdentityField, non_plcf_field_texts,
};
pub(super) use crate::parts::form_fields::FormFieldData;
pub(super) use crate::parts::format_consistency::DocumentFormatConsistencyMarks;
pub(super) use crate::parts::glossary::{AttachedGlossary, GlossaryMetadata};
pub(super) use crate::parts::grammar_cookies::GrammarCookieTables;
pub(super) use crate::parts::list_names::ListNamesTable;
pub(super) use crate::parts::list_templates::ListTemplateTable;
pub(super) use crate::parts::mail_merge::DocumentMailMerge;
pub(super) use crate::parts::numbering::{ListTables, ParagraphListBinding};
pub(super) use crate::parts::ole_controls::RgxOcxInfo;
pub(super) use crate::parts::paragraph_extractor::{ExtractedParagraph, ParagraphExtractor};
pub(super) use crate::parts::proofing::ProofingTables;
pub(super) use crate::parts::protection::Ranges;
pub(super) use crate::parts::repair_bookmarks::DocumentRepairBookmarks;
pub(super) use crate::parts::rmd_threading::DocumentRmdThreading;
pub(super) use crate::parts::rsids::DocumentRsids;
pub(super) use crate::parts::saved_by::SavedByTable;
pub(super) use crate::parts::smart_tags::DocumentSmartTags;
pub(super) use crate::parts::structured_tags::DocumentStructuredTags;
pub(super) use crate::parts::styles::StyleSheet;
pub(super) use crate::parts::subdocuments::Collection;
pub(super) use crate::parts::table_char_cache::TableCharacterCache;
pub(super) use crate::parts::text_services::TextServicesTables;
pub(super) use crate::parts::textbox_breaks::TextBoxBreakTables;
pub(super) use crate::table::Table;
pub(super) use litchi_core::Position;
pub(super) use std::sync::Arc;
