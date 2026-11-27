//! Word 97+ stylesheet (STSH) parsing.
//!
//! The owner is intentionally layered by responsibility:
//!
//! - model contains typed stylesheet records and the facade container;
//! - codec parses the binary STSH/STD/UPX representation;
//! - semantic validates style invariants and resolves inherited properties;
//! - tests keeps focused byte-level and semantic coverage beside the owner.

mod codec;
mod model;
mod semantic;

#[cfg(test)]
mod tests;

pub use model::*;
pub(crate) use semantic::{
    validate_character_style_sprms, validate_numbering_style_sprms, validate_paragraph_style_sprms,
    validate_table_style_sprms,
};
