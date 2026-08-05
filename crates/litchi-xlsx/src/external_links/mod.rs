//! Layered SpreadsheetML external-link ownership.
//!
//! Typed link models and bounded XML conversion stay independent from OPC
//! relationship orchestration. External targets are always inert metadata.

mod codec;
mod model;
mod package;
#[cfg(test)]
mod tests;

pub use codec::parse_external_link;
pub use model::{
    Cell, CellType, Conformance, Dde, DdeItem, DdeValue, DdeValueType, DdeValues, DefinedName,
    ItemSource, Link, Ole, OleItem, Row, SheetData, Target, Workbook,
};
pub use package::{
    Entry, build_external_link_part, build_external_link_part_with_conformance, load_external_link,
};

pub(super) fn invalid(message: impl Into<String>) -> crate::error::Error {
    crate::error::Error::Invalid(message.into())
}

pub(super) fn limit(name: &str) -> crate::error::Error {
    invalid(format!("external-link {name} limit exceeded"))
}
