//! Layered `SpreadsheetML` external-link ownership.
//!
//! Typed link models and bounded XML conversion stay independent from OPC
//! relationship orchestration. External targets are always inert metadata.

mod codec;
mod model;
mod package;
pub mod patch;
pub mod snapshot;
#[cfg(test)]
mod tests;
pub mod transaction;
pub mod validation;

pub use codec::parse_external_link;
pub use model::{
    Cell, CellType, Conformance, Dde, DdeItem, DdeValue, DdeValueType, DdeValues, DefinedName,
    ItemSource, Link, Ole, OleItem, Row, SheetData, Target, Workbook,
};
pub use package::{
    Entry, add_external_link, build_external_link_part, build_external_link_part_with_conformance,
    load_external_link, load_external_links, remove_external_link, replace_external_link,
    store_external_links, validate_graph,
};
pub use patch::{Commit, Patch};
pub use snapshot::Snapshot;
pub use transaction::Transaction;

pub(super) fn invalid(message: impl Into<String>) -> crate::error::Error {
    crate::error::Error::Invalid(message.into())
}

pub(super) fn limit(name: &str) -> crate::error::Error {
    invalid(format!("external-link {name} limit exceeded"))
}
