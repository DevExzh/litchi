//! Canonical WordprocessingML (`.docx`) APIs.
//!
//! The concise modules own format semantics while [`litchi_opc`] remains the
//! explicit low-level package graph.

#![forbid(unsafe_code)]

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
pub mod run_effects;
pub mod run_symbols;
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
#[cfg(feature = "vba-inspection")]
pub mod vba_project;
pub mod web;
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
    RECIPIENT_CONTENT_TYPE, Recipient, Recipients, Source, Target, parse_settings_mail_merge,
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
pub use variables::{Variables, parse_variables};

// Concrete document entry points are available through their contextual
// modules. These root exports keep the standalone facade concise without
// collapsing the owner modules back into host aliases.
pub use document::{Block, Document, Element, ImageWatermarkPart};
pub use math::{OfficeMath, OfficeMathParagraph};
pub use package::Package;
pub use paragraph::{
    Collapsed, Paragraph, Run, RunBreak, RunBreakClear, RunBreakType, RunProperties, RunUnderline,
    RunUnderlineColor,
};
pub use run_effects::{Effect, Effects, OpaqueExtension};
pub use section::{Emu, Margins, PageSize, Section, Sections};
pub use table::{Cell, Row, Table, VMergeState};
pub use writer::*;
