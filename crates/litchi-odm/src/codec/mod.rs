//! Bounded validation for the family content part.

mod content;
mod styles;

pub(crate) use content::{Semantics, parse, validate};
pub(crate) use styles::parse_catalog;
