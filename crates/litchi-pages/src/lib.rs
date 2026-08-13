//! Pages semantic models and bounded native package parsing.
//!
//! [`Package`] owns the Pages-specific native archive adapter while
//! [`Document`] remains an immutable semantic snapshot. Callers can parse and
//! inspect Pages files without importing the umbrella iWork crate, raw
//! protobuf schemas, or package identifiers.
//!
//! Settings stored directly on a section use the complete, presence-preserving
//! [`section::Settings`] value. [`Package::section_settings`] and
//! [`Package::edit_section_settings`] accept a semantic [`SectionSelector`];
//! the transaction vocabulary lives in [`section::settings`] and exposes no
//! native object identifiers. Assigning `None` to one of the four Boolean
//! setters removes that field, while `Some(false)` retains an explicit native
//! false. A changed exact-source commit verifies the selected section's
//! template dependencies, rewrites only that Section component, preserves
//! ViewState, the derived layout cache, and root previews exactly, fully
//! reopens the candidate, and returns an exact-source inverse.
//! The focused section-name and pagination entry points remain convenient
//! projections when only one of those value families is being changed.
//!
//! Section body text is read and edited through [`SectionSelector`] and
//! checked [`TextSpan`] or [`TextPosition`] values. A text transaction stages
//! one UTF-16 splice, preserves unrelated native records, fully reopens its
//! candidate under the retained limits, and returns an exact-source reversible
//! patch. Private Buffa lazy views validate known body-graph fields while the
//! original raw records remain the preservation authority.
//!
//! ```no_run
//! use litchi_pages::{Package, SectionSelector, TextSpan};
//!
//! # fn edit(source: &[u8]) -> Result<(), Box<dyn std::error::Error>> {
//! let package = Package::from_bytes(source)?;
//! let mut edit = package.edit_section_text(SectionSelector::name("Introduction"))?;
//! edit.replace(TextSpan::from_utf16_indexes(0, 0)?, "Preface: ")?;
//! let commit = edit.commit()?;
//!
//! let inverse = commit.patch().inverse();
//! let restored = commit.package().apply_section_text(&inverse)?;
//! assert_eq!(restored.package().source_bytes(), package.source_bytes());
//! # Ok(())
//! # }
//! ```

#![forbid(unsafe_code)]

pub mod audio;
mod document;
pub mod document_options;
/// Combined document and footnote formatter settings.
pub mod document_settings;
pub mod footnote;
pub mod header_footer;
pub mod image;
pub mod movie;
mod package;
pub mod page_layout;
pub mod section;
pub mod selector;

pub use document::{
    Body, DEFAULT_MAX_TEXT_BYTES, Document, DocumentReadOptions, DocumentSourceLimitKind,
    DocumentSourceLimits, DocumentSourceLimitsError, Error, IoKind, MAX_BODY_STORAGES,
    MAX_SECTIONS, ReadError, ReadLimitKind, Result, Root, SemanticLimitKind, SemanticLimits,
    SemanticLimitsError,
};
pub use litchi_core::Position;
/// Maximum bytes retained for each canonical Pages metadata sidecar.
pub use litchi_iwa_detect::MAX_PROPERTIES_BYTES as MAX_DOCUMENT_PROPERTIES_BYTES;
pub use litchi_iwa_text::{TextPosition, TextSpan};
#[cfg(feature = "internal-iwork-source")]
#[doc(hidden)]
pub use package::__semantic_document_from_prepared_source;
pub use package::{Limits, MAX_OBJECTS, Package, PackageError, PackageResult, Stats};
pub use package::{
    PageLayoutCommit, PageLayoutDiagnostics, PageLayoutEdit, PageLayoutError, PageLayoutLimitKind,
    PageLayoutPatch,
};
pub use package::{
    SectionNameCommit, SectionNameDiagnostics, SectionNameEdit, SectionNameError,
    SectionNameLimitKind, SectionNamePatch,
};
pub use package::{
    SectionPaginationCommit, SectionPaginationDiagnostics, SectionPaginationEdit,
    SectionPaginationError, SectionPaginationLimitKind, SectionPaginationPatch,
};
pub use package::{
    SectionTextCommit, SectionTextDiagnostics, SectionTextEdit, SectionTextError,
    SectionTextLimitKind, SectionTextPatch,
};
pub use section::{Section, SectionType};
pub use selector::{SectionSelector, SelectorError, SelectorResult};
