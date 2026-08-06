//! Typed, inert worksheet smart-tag metadata.
//!
//! This owner is split by responsibility: semantic values in [`model`],
//! bounded XML conversion in [`codec`], worksheet package ownership in
//! [`package`], validation in [`validation`], and clone-staged edits in
//! [`transaction`]. Smart-tag actions are never loaded or executed.

pub mod codec;
pub mod model;
pub mod package;
pub mod transaction;
pub mod validation;

#[cfg(test)]
mod tests;

pub use codec::{parse, replace_worksheet, write};
pub use model::{Cell, Collection, Conformance, Property, Tag};
pub use transaction::Transaction;

/// Validate a complete worksheet smart-tag collection without mutating it.
pub fn validate(value: &Collection) -> crate::Result<()> {
    validation::collection(value)
}

pub(crate) const TRANSITIONAL: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
pub(crate) const STRICT: &str = "http://purl.oclc.org/ooxml/spreadsheetml/main";
pub(crate) const MAX_XML_BYTES: usize = 32 * 1024 * 1024;
pub(crate) const MAX_CELLS: usize = 65_536;
pub(crate) const MAX_TAGS: usize = 65_536;
pub(crate) const MAX_PROPERTIES: usize = 262_144;
pub(crate) const MAX_TEXT_BYTES: usize = 32_767;
pub(crate) const MAX_DEPTH: usize = 256;
