//! Canonical `WordprocessingML` (`.docx`) APIs.
//!
//! The concise modules own format semantics while [`litchi_opc`] remains the
//! explicit low-level package graph.

#![forbid(unsafe_code)]
// quick-xml's checked attribute iteration is quadratic on hostile tags; read
// attributes through `BytesStartExt` (record 0770, workspace `clippy.toml`).
#![cfg_attr(not(test), deny(clippy::disallowed_methods))]

mod error;

pub mod alt;
pub mod bibliography;
pub mod bookmark;
pub mod chart;
pub mod color;
pub mod comment;
pub mod content_control;
pub mod custom_xml;
pub mod document;
pub mod drawing;
pub mod field;
pub mod font;
pub mod footnote;
pub mod format;
pub mod glossary;
pub mod header_footer;
pub mod hyperlink;
pub mod image;
pub mod ink;
pub mod list;
pub mod mail_merge;
pub mod math;
pub mod modern_comments;
mod namespace;
pub mod numbering;
pub mod package;
pub mod paragraph;
pub mod parts;
pub mod redact;
pub mod revision;
pub mod run_effects;
pub mod run_symbols;
pub mod sanitize;
pub mod section;
pub mod settings;
pub mod smart_tag;
pub mod smartart;
pub mod source_backed;
pub mod statistics;
/// Source-backed package-wide inventory and forward-only redaction of
/// relationship-owned story hyperlinks.
pub mod story_hyperlinks;
/// Bounded, forward-only creation of plain DOCX paragraphs and runs.
pub mod streaming;
pub mod styles;
pub mod table;
pub mod template;
pub mod textbox;
pub mod theme;
/// Read-only, bounded DOCX package and main-document validation.
pub mod validation;
pub mod variables;
#[cfg(feature = "vba-inspection")]
pub mod vba_project;
pub mod web;
/// Inert Office Add-in and persisted task-pane models, limits and graph patches.
///
/// These shared types are used by [`Package::task_panes`] and
/// [`Package::plan_task_panes`]. DOCX web-output settings remain in [`web`].
///
/// ```
/// use litchi_core::patch::{BlobLimits, Patch as DurablePatch, PatchLimits, Reversible};
/// use litchi_docx::{Package, web_extensions as extensions};
///
/// # fn main() -> Result<(), Box<dyn std::error::Error>> {
/// let mut package = Package::new()?;
/// let reference = extensions::Reference::new("addin", "1.0", extensions::Store::Omex)?;
/// let mut add_in = extensions::AddIn::new("addin-instance", reference)?;
/// let mut functions = extensions::CustomFunctions::new();
/// functions.set_background_app_data(Some(extensions::BackgroundAppData::new(
///     3, "runtime-instance",
/// )?));
/// add_in.set_custom_functions(Some(functions))?;
/// let mut panes = extensions::Panes::new();
/// panes.push(extensions::Pane::new(add_in))?;
/// let patch = package.plan_task_panes(panes, extensions::Conformance::Transitional)?;
/// package.apply_task_panes_patch(&patch)?;
/// package.apply_task_panes_patch(&patch.inverse())?;
///
/// // Choose finite wire limits for the expected document workload.
/// let wire_limits = PatchLimits::new(
///     BlobLimits::new(2, 1_048_576, 2_097_152),
///     4_194_304, 1, 8, 1_024, 1_024,
/// );
/// let wire = patch.to_durable(wire_limits)?.to_deterministic_json()?;
/// let durable = DurablePatch::<Reversible>::from_deterministic_json(&wire, wire_limits)?;
/// package.apply_durable_task_panes_patch(&durable)?;
/// package.apply_durable_task_panes_patch(&durable.inverse())?;
/// # Ok(())
/// # }
/// ```
pub use litchi_ooxml_common::web as web_extensions;
pub mod writer;

#[cfg(feature = "encryption")]
pub use litchi_crypto::ooxml as encryption;

pub use bibliography::{
    BibliographySource, BibliographySourceStore, BibliographySourceValue,
    LEGACY_WORD_BIBLIOGRAPHY_NAMESPACE, MAX_BIBLIOGRAPHY_DEPTH, MAX_BIBLIOGRAPHY_SOURCES,
    MAX_BIBLIOGRAPHY_TEXT_BYTES, MAX_BIBLIOGRAPHY_VALUES, MAX_BIBLIOGRAPHY_XML_BYTES,
    OOXML_BIBLIOGRAPHY_NAMESPACE, STRICT_OOXML_BIBLIOGRAPHY_NAMESPACE, is_bibliography_namespace,
    is_bibliography_node, is_bibliography_root, parse_bibliography_source_store,
};
pub use error::{Error, Result};
pub use field::{
    ActiveContent, ActiveContentKind, Advance, AdvanceAdjustment, AdvanceOperation, AutoNumber,
    AutoNumberKind, AutoText, AutoTextKind, AutoTextList, AutoTextListOption, Barcode,
    Bibliography, BidiOutline, Citation, Compare, Context, ContextKind, CountryInclusion, Database,
    Dde, DdeFormat, DdeKind, Embed, Equation, Field, Formula, GoToButton, If, Include, IncludeKind,
    IncludeOption, Index, IndexEntry, IndexOrder, Info, Information, InformationKind, LegacyForm,
    LegacyFormKind, Link, LinkFormat, LinkResult, ListNumber, MacroButton, Merge, MergeControl,
    MergeControlKind, MergeCounter, MergeCounterKind, MergeData, MergeNext, Print, Private, Prompt,
    PromptKind, Property, Quote, RecipientKind, Reference, ReferenceKind, ReferenceOption,
    Sequence, Set, Shape, StyleOption, StyleReference, SubDocument, Switch, Symbol, Toa, ToaEntry,
    Toc, TocEntry, TocLevelRange, UserIdentity, UserIdentityFormat, UserIdentityKind, Variable,
};
pub use format::{ImageFormat, LineSpacing, ParagraphAlignment, TableBorderStyle, UnderlineStyle};
pub use hyperlink::Hyperlink;
/// Resource policy for package ingestion through [`Package`].
pub use litchi_opc::ReadLimits;
pub use mail_merge::{
    DataSourceObject, DataType, Destination, FieldMap, FieldMappingType, MainDocumentType,
    RECIPIENT_CONTENT_TYPE, Recipient, Recipients, RelationshipId, Source, Target,
    parse_settings_mail_merge,
};
pub use modern_comments::{
    Comment, Conformance, Extended, Extension, ExtensionList, IdMapping, Metadata, Person,
    Presence, Reaction, ReactionInfo, ReactionUser, RelationshipIds, load_modern_comment_metadata,
    parse_comments_extended, parse_comments_extensible, parse_comments_ids, parse_people,
    store_modern_comment_metadata, write_comments_extended, write_comments_extensible,
    write_comments_ids, write_people,
};
pub use numbering::{
    Collection, Definition, Format, Instance, Level, MultiLevel, Override, ParseFormatError,
    ParseMultiLevelError, PictureBullet, Restart, Suffix, parse_numbering,
};
pub use settings::{
    ColorSchemeIndex, ColorSchemeMapping, ColorSchemeSlot, CompatFlag, CompatibilityOption,
    CompatibilitySetting, MAX_LANGUAGE_TAG_LENGTH, MAX_SETTINGS_XML_BYTES, MAX_SETTINGS_XML_DEPTH,
    MAX_SETTINGS_XML_NODES, MAX_SMART_TAG_NAME_CHARS, MAX_SMART_TAG_NAMESPACE_URI_CHARS,
    MAX_SMART_TAG_URL_CHARS, NoteNumberFormat, NoteNumberingProperties, NoteNumberingRestart,
    NotePosition, ParseCompatFlagError, ParseNoteNumberFormatError, ParseNotePositionError,
    ProofState, ProofingState, ProtectionType, Settings, SmartTagType, ThemeFontLanguages, View,
    validate_smart_tag_type,
};
pub use statistics::{
    Statistics, count_characters, count_characters_no_spaces, count_words, estimate_line_count,
    estimate_page_count,
};
pub use validation::{
    DEFAULT_DOCX_VALIDATION_LIMITS, DocxValidationError, DocxValidationLimits, validate_read_at,
    validate_read_at_with_limits,
};
pub use variables::{Variables, parse_variables};

// Concrete document entry points are available through their contextual
// modules. These root exports keep the standalone facade concise without
// collapsing the owner modules back into host aliases.
pub use document::{Block, Document, Element, ImageWatermarkPart, OpaqueBlock};
pub use math::{OfficeMath, OfficeMathParagraph};
pub use package::Package;
pub use paragraph::{
    Collapsed, Inline, InlineHyperlink, OpaqueInline, OpaqueRunContent, Paragraph, Run, RunBreak,
    RunBreakClear, RunBreakType, RunContent, RunProperties, RunUnderline, RunUnderlineColor,
};
pub use run_effects::{Effect, Effects, OpaqueExtension};
pub use section::{Emu, Margins, PageSize, Section, Sections};
pub use streaming::{
    StreamingDocumentError, StreamingDocumentErrorSource, StreamingDocumentLimits,
    StreamingDocumentWriter,
};
pub use table::{Cell, Row, Table, VMergeState};
pub use writer::*;
