//! Pages semantic models and bounded native package parsing.
//!
//! [`Package`] owns the Pages-specific native archive adapter while
//! [`Document`] remains an immutable semantic snapshot. Callers can parse and
//! inspect Pages files without importing the umbrella iWork crate, raw
//! protobuf schemas, or package identifiers.

#![forbid(unsafe_code)]

pub mod audio;
mod document;
pub mod document_options;
pub mod footnote;
pub mod header_footer;
pub mod image;
pub mod movie;
mod package;
pub mod page_layout;
pub mod section;
pub mod selector;

pub use document::{
    Body, DEFAULT_MAX_TEXT_BYTES, Document, Error, MAX_BODY_STORAGES, MAX_SECTIONS, Result, Root,
};
pub use litchi_core::Position;
#[cfg(feature = "internal-iwork-source")]
#[doc(hidden)]
pub use package::__semantic_document_from_prepared_source;
pub use package::{Limits, Package, PackageError, PackageResult, Stats};
pub use package::{
    SectionNameCommit, SectionNameDiagnostics, SectionNameEdit, SectionNameError,
    SectionNameLimitKind, SectionNamePatch,
};
pub use section::{Section, SectionType};
pub use selector::{SectionSelector, SelectorError, SelectorResult};
