//! Bounded XLSB Brt* conditional-formatting codec facade.
//!
//! Wire primitives, typed record codecs, and record-stream writing live in
//! nested owners so the public conditional-formatting facade remains small.

#![allow(clippy::too_many_arguments)]

mod records;
mod wire;
mod writer;

pub use records::parse_classic_header;
#[cfg(test)]
pub(super) use wire::{CfCursor, parse_formula_header};
pub use wire::{parse_rule_extension_guid, serialize_rule_extension_guid};
pub use writer::write_conditional_formattings;
#[cfg(test)]
pub(super) use writer::{serialize_cf_rule, serialize_cond_formatting_header};
