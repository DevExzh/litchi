//! Layered XLSB conditional-formatting model and Brt* codec.
//!
//! The owner exposes concise semantic names while retaining the historical
//! public names as aliases. Worksheet/package traversal remains in the host.

mod codec;
mod model;

pub use codec::{
    Error, Result, XlsbError, XlsbResult, icon_count14, parse_classic_header,
    parse_rule_extension_guid, serialize_rule_extension_guid, validate_formula_count,
    validate_template, write_conditional_formattings,
};
pub use model::{
    AxisPosition14, Bar, Bar14, BarHeader14, Color, Direction14, Formatting, Icon, IconHeader14,
    IconSet, IconSet14, RecordKind, Rule, RuleMetadata, RuleType, Scale, Value,
};

// Compatibility aliases for the historical host/API vocabulary.
pub use model::AxisPosition14 as DataBarAxisPosition14;
pub use model::Bar as DataBar;
pub use model::Bar14 as DataBar14;
pub use model::BarHeader14 as DataBar14Header;
pub use model::Color as ConditionalFormatColor;
pub use model::Direction14 as DataBarDirection14;
pub use model::Formatting as Collection;
pub use model::Formatting as ConditionalFormatting;
pub use model::Icon as ConditionalFormatIcon;
pub use model::IconHeader14 as IconSet14Header;
pub use model::RecordKind as ConditionalFormattingRecordKind;
pub use model::Rule as ConditionalFormattingRule;
pub use model::RuleMetadata as ConditionalFormattingRule14Metadata;
pub use model::RuleType as CfRuleType;
pub use model::Scale as ColorScale;
pub use model::Value as Cfvo;
