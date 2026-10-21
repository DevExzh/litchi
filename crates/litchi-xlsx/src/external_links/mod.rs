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

pub(super) fn limit(name: &str) -> Error {
    use model::{
        MAX_ALTERNATE_URL_METADATA_BYTES, MAX_CACHE_TEXT_BYTES, MAX_CACHED_CELLS, MAX_CACHED_ROWS,
        MAX_CACHED_SHEETS, MAX_CELL_COLUMN, MAX_CELL_ROW, MAX_COLUMN_LETTERS, MAX_DEFINED_NAMES,
        MAX_EXTERNAL_TARGET_BYTES, MAX_LINK_ITEMS, MAX_RELATIONSHIP_ID_BYTES, MAX_SHEET_NAMES,
    };

    let (kind, maximum) = if name.contains("target URI") {
        (LimitKind::TargetUri, MAX_EXTERNAL_TARGET_BYTES)
    } else if name.contains("sheet names") {
        (LimitKind::SheetNames, MAX_SHEET_NAMES)
    } else if name.contains("defined names") {
        (LimitKind::DefinedNames, MAX_DEFINED_NAMES)
    } else if name.contains("cached sheets") {
        (LimitKind::CachedSheets, MAX_CACHED_SHEETS)
    } else if name.contains("cached rows") {
        (LimitKind::CachedRows, MAX_CACHED_ROWS)
    } else if name.contains("cached cells") {
        (LimitKind::CachedCells, MAX_CACHED_CELLS)
    } else if name.contains("matrix") || name.contains("values") {
        (LimitKind::MatrixCells, MAX_CACHED_CELLS)
    } else if name.contains("items") {
        (LimitKind::Items, MAX_LINK_ITEMS)
    } else if name.contains("cache text") {
        (LimitKind::CacheTextBytes, MAX_CACHE_TEXT_BYTES)
    } else if name.contains("relationship ID") {
        (LimitKind::RelationshipIdBytes, MAX_RELATIONSHIP_ID_BYTES)
    } else if name.contains("alternate URL") {
        (
            LimitKind::AlternateUrlMetadataBytes,
            MAX_ALTERNATE_URL_METADATA_BYTES,
        )
    } else if name.contains("column letters") {
        (LimitKind::ColumnLetters, MAX_COLUMN_LETTERS)
    } else if name.contains("column") {
        (LimitKind::CellColumn, MAX_CELL_COLUMN as usize)
    } else if name.contains("row") {
        (LimitKind::CellRow, MAX_CELL_ROW as usize)
    } else {
        (LimitKind::XmlBytes, MAX_CACHE_TEXT_BYTES)
    };
    Error::limit(kind, maximum.saturating_add(1), maximum)
}
