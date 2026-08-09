//! OpenDocument Spreadsheet (`.ods`) support.
//!
//! The crate is organized by responsibility: immutable spreadsheet vocabulary
//! in [`model`], XML codecs in [`codec`], package access in [`package`],
//! construction in [`authoring`], and the concise entry points in [`facade`].

pub mod annotations;
pub mod authoring;
pub mod charts;
pub mod codec;
pub mod data_pilot;
pub mod drawing;
pub mod embedded;
pub mod facade;
pub mod flat;
pub mod media;
pub mod metadata;
pub mod model;
pub mod package;
pub mod protection;
pub mod settings;
pub mod source_features;
pub mod styles;
pub mod worksheet;

pub use charts::Chart;
pub use drawing::{Frame, Part};
pub use embedded::{Kind, Object, Parameter, Root};
pub use facade::{Builder, MutableSpreadsheet, Spreadsheet};
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
    Drawing, DrawingKind, Hyperlink as SourceHyperlink, Sheet as SourceSheet,
    Snapshot as SourceFeatures,
};
pub use worksheet::{Cell, CellValue, CellView, Merge, Row, Sheet};
