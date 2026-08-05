//! Layered BIFF8 conditional-format owner.
///
/// The semantic model, BIFF record codecs, worksheet collection grammar, and
/// regression fixtures live in dedicated layers while the crate-level facade
/// keeps its established public names.
mod codec;
mod model;

#[cfg(test)]
mod tests;

pub(crate) use codec::Collector as ConditionalFormatCollector;

pub use model::{
    Alignment as ConditionalAlignment, Border as ConditionalBorder,
    Comparison as ConditionalComparison, Extension as ConditionalExtension,
    Font as ConditionalFont, Formatting as ConditionalFormatting,
    Formatting12 as ConditionalFormatting12, NumberFormat as ConditionalNumberFormat,
    Pattern as ConditionalPattern, Protection as ConditionalProtection,
    Range as ConditionalFormatRange, Rule as ConditionalRule, Rule12 as ConditionalRule12,
    Rule12Kind as ConditionalRule12Kind, RuleKind as ConditionalRuleKind,
    Style as ConditionalStyle,
};
