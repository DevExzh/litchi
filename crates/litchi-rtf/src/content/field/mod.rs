#![cfg_attr(not(test), deny(clippy::indexing_slicing))]

//! Safe, structured RTF field-code support.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "items stay grouped by RTF feature area rather than by item kind"
)]
const MAX_INSTRUCTION_LEN: usize = 65_536;
const MAX_TOKENS: usize = 256;
pub(crate) const MAX_GENERIC_FIELDS: usize = 65_536;

mod codec;
mod model;
mod validation;

#[cfg(test)]
mod tests;

pub use codec::parse_field_code;
pub(crate) use codec::quoted_field_operand;
pub use model::*;
pub(crate) use validation::{push_story_page_break, validate_story_events};
