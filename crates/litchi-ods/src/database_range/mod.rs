//! Contextual ODS database-range ownership.
//!
//! [`crate::model::database_range`] owns the ODF vocabulary and standalone
//! XML grammar. This module binds that vocabulary to the direct
//! `office:spreadsheet` owner and supplies immutable catalogs plus
//! clone-staged, source-checked CRUD transactions. Database source metadata
//! is retained as inert data; no database, query, refresh, or external I/O is
//! performed.

mod codec;
mod content;
mod model;
mod package;
mod snapshot;
mod transaction;
mod validation;

#[cfg(test)]
mod tests;

pub use content::{ContentCommit, ContentEdit, ContentEditor};
pub use model::{Catalog, Selector};
pub use snapshot::{Commit, Edit, OwnedEditor, Patch, Snapshot};
pub use transaction::{Commit as CatalogCommit, Editor, Transaction};

pub use crate::model::database_range::{
    Condition, ConditionSource, DataType, EmbeddedNumberBehavior, Expression, Field, Filter, Key,
    Order, Orientation, Range, Rule, Rules, Sort, SortGroups, Source,
};

/// Parse the direct spreadsheet database-range owner from a content XML
/// source. The returned boolean distinguishes an absent owner from an
/// explicitly empty `table:database-ranges` element.
pub(crate) fn parse_content(xml: &str) -> litchi_core::Result<(Vec<Range>, bool)> {
    let location = codec::locate(xml)?;
    let ranges = if let Some(container) = &location.container {
        let fragment = codec::owner_fragment(xml, container)?;
        crate::model::database_range::parse_database_ranges(&fragment)?
    } else {
        Vec::new()
    };
    validation::validate_snapshot(xml, &location, &ranges)?;
    Ok((ranges, location.container.is_some()))
}

/// Replace the direct spreadsheet database-range owner in one content XML
/// source. This helper is used by the package builder, where no package
/// object exists yet; it keeps the same source-preservation and opaque-owner
/// checks as the package facade.
pub(crate) fn replace_content(
    xml: &str,
    candidate: Option<&[Range]>,
) -> litchi_core::Result<String> {
    let location = codec::locate(xml)?;
    let original = if let Some(container) = &location.container {
        let fragment = codec::owner_fragment(xml, container)?;
        Some(crate::model::database_range::parse_database_ranges(
            &fragment,
        )?)
    } else {
        None
    };
    validation::validate_snapshot(xml, &location, original.as_deref().unwrap_or(&[]))?;
    let candidate_owned = candidate.map(<[Range]>::to_vec);
    validation::validate_candidate(&location, &original, &candidate_owned)?;
    if original.as_deref() == candidate_owned.as_deref() {
        return Ok(xml.to_owned());
    }
    codec::replace(xml, &location, candidate)
}
