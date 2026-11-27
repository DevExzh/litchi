//! Layered `SpreadsheetML` external-link ownership.
//!
//! Typed link models and bounded XML conversion stay independent from OPC
//! relationship orchestration. External targets are always inert metadata.
//!
//! [`crate::Package::edit_external_links`] starts an [`Edit`] of the
//! workbook's external-link catalog; [`Edit::commit`] returns a [`Commit`]
//! holding the new package, the exact [`Patch`] and the [`Diagnostics`], and
//! [`crate::Package::apply_external_links`] replays a patch. The source
//! package never changes.

mod codec;
mod edit;
mod model;
mod package;
mod patch;
mod snapshot;
#[cfg(test)]
mod tests;
pub mod validation;
mod workbook;

pub use codec::parse_external_link;
pub use edit::Edit;
#[cfg(test)]
pub(crate) use edit::Transaction;
pub use model::{
    AlternateUrl, AlternateUrls, Cell, CellType, Conformance, Dde, DdeItem, DdeValue, DdeValueType,
    DdeValues, DefinedName, ItemSource, Link, Ole, OleItem, Row, SheetData, Target, Workbook,
};
pub use package::{
    Entry, build_external_link_part, build_external_link_part_with_conformance, load_external_link,
    load_external_links, validate_graph,
};
pub use patch::{Commit, Diagnostics, Error, LimitKind, Patch};
pub use snapshot::Snapshot;

pub(super) fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}

pub(super) fn limit(kind: LimitKind, observed: usize, maximum: usize) -> Error {
    Error::limit(kind, observed, maximum)
}
