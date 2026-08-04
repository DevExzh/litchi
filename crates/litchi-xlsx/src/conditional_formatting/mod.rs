//! Typed SpreadsheetML conditional-formatting support.
//!
//! The owner is layered by responsibility: semantic declarations in
//! [`model`], bounded SpreadsheetML/MCE conversion and differential-format
//! association in [`codec`], and regression coverage in [`tests`]. There is
//! no package layer because conditional formatting is embedded in worksheet
//! and styles parts rather than owning an OPC part or relationship graph.

mod codec;
mod model;
#[cfg(test)]
mod tests;

pub use codec::{parse_conditional_formattings, parse_differential_formats};
pub use model::{
    Association, Axis, Color, ColorRole, ColorScale, Component, DataBar, Differential,
    DifferentialRef, Direction, Formatting, IconSet, IconSet14, Icons, Kind, NamedColor,
    NumberFormat, Operator, Payload, Period, Range, Rule, Source, TokenError, Value, ValueKind,
};

// Historical public names remain aliases. The canonical declarations above
// are contextual to this owner and avoid repeating the module or format name.
pub type CellRangeRef = Range;
pub type ConditionalFormatPayload = Payload;
pub type ConditionalFormatValue = Value;
pub type ConditionalFormatting = Formatting;
pub type ConditionalFormattingRule = Rule;
pub type ConditionalFormattingSource = Source;
pub type DifferentialFormat = Differential;
pub type DifferentialFormatComponent = Component;
pub type DifferentialFormatRef = DifferentialRef;
pub type DifferentialNumberFormat = NumberFormat;
pub type ExtensionAssociation = Association;
pub type SpreadsheetColor = Color;
