//! Unified writes over the leaf ODT body edits.

use litchi_odt::body::{Attribute, Kind as LeafKind, Node as LeafNode};

use super::Odt;
use crate::Result;

/// The columns a table's first row spans (`table:number-columns-repeated`
/// counted).
pub(crate) fn table_columns(table: LeafNode<'_>) -> usize {
    let Some(row) = table.children().find(|child| child.kind() == LeafKind::Row) else {
        return 0;
    };
    row.children()
        .map(|cell| {
            cell.attribute(Attribute::ColumnsRepeated)
                .and_then(|repeat| repeat.parse::<usize>().ok())
                .unwrap_or(1)
                .max(1)
        })
        .fold(0usize, usize::saturating_add)
}

/// A fresh empty text document.
#[allow(dead_code, reason = "TEMP-0276: used by the write adapter")]
pub(crate) fn create(options: &crate::CreateOptions) -> Result<Odt> {
    let _ = options;
    let package = litchi_odt::Package::new().map_err(super::odf_error)?;
    Odt::from_package(package, None)
}
