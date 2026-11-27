//! Bounded validation for the family content part.

mod content;
mod styles;

pub(crate) use content::{validate, validate_tree};
pub(crate) use styles::validate as validate_styles;
