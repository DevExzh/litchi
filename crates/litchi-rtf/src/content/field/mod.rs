#![cfg_attr(not(test), deny(clippy::indexing_slicing))]

//! Safe, structured RTF field-code support.

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
