//! Layered XLSB conditional-formatting model and Brt* codec.
//!
//! The owner exposes the complete concise semantic vocabulary and bounded
//! record codec. Worksheet/package traversal remains in the host.

mod codec;
mod model;
#[cfg(test)]
mod tests;

pub use codec::{
    Error, Result, icon_count14, parse_classic_header, parse_rule_extension_guid,
    serialize_rule_extension_guid, validate_formula_count, validate_template,
    write_conditional_formattings,
};
pub use model::{
    AxisPosition14, Bar, Bar14, BarHeader14, Color, Direction14, Formatting, Icon, IconHeader14,
    IconSet, IconSet14, RecordKind, Rule, RuleMetadata, RuleType, Scale, Value,
};
