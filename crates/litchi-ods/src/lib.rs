//! `OpenDocument` Spreadsheet (`.ods`) support.
//!
//! The crate is organized by responsibility: immutable spreadsheet vocabulary
//! in [`model`], XML codecs in [`codec`], package access in [`package`],
//! construction in [`authoring`], and the concise entry points in [`facade`].

#![forbid(unsafe_code)]
// ODF element models and the established spreadsheet facade deliberately use
// specification-shaped names, ownership, and enum layouts. Changing them for
// generic API heuristics would be a breaking migration rather than a codec fix.
#![allow(
    clippy::large_enum_variant,
    clippy::module_name_repetitions,
    clippy::needless_pass_by_value,
    clippy::ref_option,
    clippy::return_self_not_must_use,
    clippy::struct_field_names,
    clippy::unused_self,
    reason = "the public ODS API and ODF vocabulary retain specification-shaped names and ownership"
)]
// XML event parsers intentionally reuse record-local names as raw attributes
// become decoded values and active semantic nodes.
#![allow(
    clippy::shadow_reuse,
    clippy::shadow_same,
    clippy::shadow_unrelated,
    clippy::similar_names,
    reason = "short-lived streaming XML bindings mirror successive wire events and their decoded projections"
)]
// ODF codecs are ordered by document traversal and keep schema cases explicit.
// These style exceptions retain that audit order while all correctness lints
// continue to be warning-denied.
#![allow(
    clippy::allow_attributes_without_reason,
    clippy::arbitrary_source_item_ordering,
    clippy::assigning_clones,
    clippy::doc_markdown,
    clippy::float_cmp,
    clippy::items_after_statements,
    clippy::match_same_arms,
    clippy::missing_panics_doc,
    clippy::needless_for_each,
    clippy::redundant_closure_for_method_calls,
    clippy::single_match_else,
    clippy::unnecessary_sort_by,
    clippy::unnecessary_wraps,
    clippy::unnested_or_patterns,
    clippy::unreadable_literal,
    reason = "ODF codec declarations follow document order and retain explicit schema branches for reviewability"
)]

pub mod advanced;
pub mod annotations;
pub mod authoring;
pub mod charts;
pub mod codec;
pub mod content_validation;
pub mod data_pilot;
pub mod definitions;
pub mod document;
pub mod drawing;
pub mod embedded;
pub mod facade;
pub mod flat;
pub mod media;
pub mod metadata;
pub mod metadata_graphs;
pub mod model;
mod open_parse;
pub mod package;
pub mod protection;
pub mod settings;
pub mod source_features;
pub mod styles;
pub mod worksheet;

pub use charts::Chart;
pub use drawing::{Frame, Part};
pub use embedded::{Kind, Object, Parameter, Root};
pub use facade::{
    Builder, CellSelector, MAX_CELL_SELECTORS, MutableSpreadsheet, ReadLimits,
    SourceBackedSpreadsheet, SourceCellCommit, SourceCellEdit, SourceCellPatch,
    SourceCellPublicationReport, SourceCellSnapshot, Spreadsheet,
};
pub use flat::{
    FlatCommit, FlatEdit, FlatSpreadsheet, Limits as FlatLimits, Patch as FlatPatch,
    SheetSelector as FlatSheetSelector, Snapshot as FlatSnapshot,
};
pub use litchi_core::Metadata;
pub use litchi_odf_common::rdf;
pub use media::Image;
pub use model::names;
pub use model::tracked_changes;
pub use model::{dde, scenario};
pub use settings::{Iteration, IterationStatus, NullDate, Settings};
pub use source_features::{
    Drawing, DrawingKind, Hyperlink as SourceHyperlink, Limits as SourceFeatureLimits,
    Sheet as SourceSheet, Snapshot as SourceFeatures,
};
pub use worksheet::{Cell, CellValue, CellView, Merge, Row, Sheet};
