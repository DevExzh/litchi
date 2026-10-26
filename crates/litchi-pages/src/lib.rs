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
//! let mut original_bytes = Vec::new();
//! package.write_to(&mut original_bytes)?;
//! let mut restored_bytes = Vec::new();
//! restored.package().write_to(&mut restored_bytes)?;
//! assert_eq!(restored_bytes, original_bytes);
//! # Ok(())
//! # }
//! ```

#![forbid(unsafe_code)]

pub mod audio;
pub mod builder;
pub mod chart;
mod document;
pub mod document_options;
/// Combined document and footnote formatter settings.
pub mod document_settings;
pub mod drawable;
mod drawable_order;
/// Field pass between the stored drawing values of
/// `litchi_iwa_protos::shape_style_codec::native` and
/// `litchi_iwa_common::shape::native` (record 0375; W1-E hunk).
#[cfg(not(feature = "internal-iwork-source"))]
#[allow(
    dead_code,
    unreachable_pub,
    reason = "consumed by the W2 drawing families"
)]
mod drawing_native;
/// Field pass between the stored drawing values of
/// `litchi_iwa_protos::shape_style_codec::native` and
/// `litchi_iwa_common::shape::native` (record 0375; W1-E hunk), visible to
/// the `litchi-iwa` exit parity harness only.
#[cfg(feature = "internal-iwork-source")]
#[doc(hidden)]
pub mod drawing_native;
pub mod footnote;
pub mod header_footer;
pub mod image;
/// Exact-source transactions for the document properties (title, author,
/// keywords, description).
pub mod metadata {
    pub use crate::package::metadata_edit::{
        Commit, Diagnostics, Edit, Error, LimitKind, Patch, Refusal,
    };
}
mod limits;
pub mod movie;
mod package;
pub mod page_layout;
pub mod section;
pub mod selector;
pub mod table;

pub use document::{
    Body, DEFAULT_MAX_TEXT_BYTES, Document, DocumentStats, Error, MAX_BODY_STORAGES, MAX_SECTIONS,
    Result, Root,
};
pub use drawable_order::{BodyDrawableHandle, BodyDrawableSelector, DrawableLayerMove};
pub use header_footer::{HeaderFooter, HeaderFooterSelector, Kind, Template};
pub use image::{ImageAdjustment, ImageAdjustments, ImageEnhancement};
pub use limits::{
    LimitKind, OpenOptions, SemanticLimitKind, SemanticLimits, SemanticLimitsBuilder,
    SourceLimitKind, SourceLimits, SourceLimitsBuilder,
};
pub use litchi_core::Position;
pub use litchi_iwa_common::chart::arrangement::ChartArrangement;
pub use litchi_iwa_common::chart::data::ChartData;
pub use litchi_iwa_common::chart::metadata::ChartMetadata;
/// Maximum bytes retained for each canonical Pages metadata sidecar.
pub use litchi_iwa_detect::MAX_PROPERTIES_BYTES as MAX_DOCUMENT_PROPERTIES_BYTES;
pub use litchi_iwa_text::{TextPosition, TextSpan};
pub use package::MergeReader;
#[cfg(feature = "internal-iwork-source")]
#[doc(hidden)]
pub use package::{
    __CatalogBodyTableAppearanceSource, __catalog_body_table_appearance,
    __catalog_body_table_style_edges,
};
#[cfg(feature = "internal-iwork-source")]
#[doc(hidden)]
pub use package::{__decode_image_adjustments_payload, ImageAdjustmentsError};
#[cfg(feature = "internal-iwork-source")]
#[doc(hidden)]
pub use package::{
    __decode_movie_playback_payload, __rewrite_movie_playback_payload, MoviePlaybackError,
};
#[cfg(feature = "internal-iwork-source")]
#[doc(hidden)]
pub use package::{
    __is_valid_pages_text_storage, __pages_body_storage_discovery,
    __semantic_document_from_prepared_source, BodyStorageDiscovery,
};
pub use package::{
    MAX_OBJECTS, Package, PackageError, PackageResult, SaveError, Stats, WriteError,
};
pub use section::{Section, SectionType};
pub use selector::{BodyChartSelector, BodyTableSelector, ImageSelector};
pub use selector::{SectionSelector, SelectorError, SelectorResult};
pub use table::merge::Region;
