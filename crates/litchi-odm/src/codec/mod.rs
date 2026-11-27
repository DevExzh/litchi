//! Bounded validation for the family content part.

mod compact;
mod content;
mod styles;

pub(crate) use compact::compact_source_xml;
pub(crate) use content::{Semantics, parse, validate};
pub(crate) use styles::parse_catalog;
