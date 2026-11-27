//! Semantic style resolution and validation layers.

mod resolve;
mod validation;

#[cfg(test)]
pub(super) use resolve::flatten_conditional_style_sprms;
pub(super) use validation::{strip_paragraph_style_index, validate_style_sprms, validate_styles};
pub(crate) use validation::{
    validate_character_style_sprms, validate_numbering_style_sprms, validate_paragraph_style_sprms,
    validate_table_style_sprms,
};
