//! Layered BIFF8 chart-sheet and embedded-chart owner.
//!
//! The public surface is contextual to an XLS chart: the model lives in the
//! model module, record parsing and replay in the codec module, and workbook
//! hosting in the package module. Chart formulas and cached values remain
//! inert. Unsupported records are retained rather than normalized or evaluated.

pub mod chart_area;
pub mod frt;

mod codec;
mod inventory;
pub(crate) mod model;
mod package;
mod wire;

#[cfg(test)]
mod tests;

pub use litchi_ograph::chart::{format, group};
pub use litchi_ograph::record::{chart3d, frame, line, marker, pie, series};

/// The chart family error: the workbook error of the source or republished
/// package.
pub use crate::error::Error;
pub use inventory::{ChartInventory, SemanticCompleteness};
pub use model::{
    AreaFormat, Axis, AxisKind, AxisLine, AxisLineKind, Cache, CacheKind, CellRef, DataKind,
    DataLink, Format, Formula, Group, GroupKind, Kind, Label, Legend, Limits, LineFormat, Location,
    Raw, Role, Scale, Selector, Series, Source, Tick, Value,
};
pub use package::build_workbook;
pub use package::{Commit, Diagnostics, Edit, EditOptions, Entry, LimitKind, Patch};
