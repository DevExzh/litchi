//! Typed `SpreadsheetML` conditional-formatting support.
//!
//! Conditional formatting is embedded in worksheet XML and refers to frozen
//! differential formats in the workbook styles part. Semantic declarations
//! live in [`model`], bounded SpreadsheetML conversion and differential-format
//! association in [`codec`], and guarded package/source publication in the
//! private package, snapshot, patch, and source layers. The public editor
//! replaces the complete ordered core collection for one existing worksheet
//! while binding its workbook, relationship, and styles dependency closure.

mod codec;
mod model;
mod package;
mod patch;
mod snapshot;
mod source;
#[cfg(test)]
mod tests;

pub use codec::{parse_conditional_formattings, parse_differential_formats};
pub use model::{
    Association, Axis, Color, ColorRole, ColorScale, Component, DataBar, Differential,
    DifferentialRef, Direction, Formatting, IconSet, IconSet14, Icons, Kind, NamedColor,
    NumberFormat, Operator, Payload, Period, Range, Rule, Source, TokenError, Value, ValueKind,
};
pub use package::replace_conditional_formattings;
pub use patch::{Commit, Diagnostics, Patch};
pub use snapshot::Snapshot;
pub use source::{SourceBackedEditor, SourceEdit};
