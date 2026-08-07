//! Strict, bounded XLSB sparkline semantics.
//!
//! This owner implements `[MS-XLSB]` sections 2.4.228--2.4.230,
//! 2.4.581--2.4.583, 2.4.806, 2.5.59--2.5.61, and 2.5.66. Formula bytes are
//! retained and structurally validated but never compiled or evaluated.
//! Source-bound worksheet snapshots remain package-neutral; [`Workbook`]
//! APIs additionally prove name and `Xti` indexes against inert workbook
//! metadata without following an external target.
//!
//! [`Workbook`]: crate::Workbook

mod codec;
mod model;
pub(crate) mod workbook;
pub mod worksheet;

#[cfg(test)]
mod tests;
#[cfg(test)]
mod worksheet_tests;

pub use litchi_sheet::sparkline::{AxisType, EmptyCells, SparklineType};
pub use model::{
    Area, Axis, Color, ColorType, Colors, Error, Formula, FormulaKind, FrtState, Group, Groups,
    Limits, Location, Options, Reference, Result, Sparkline,
};

pub(crate) use codec::{encode_block, parse_block};
pub use worksheet::{Commit, Edit, Snapshot};
