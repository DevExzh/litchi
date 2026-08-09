//! Embedded ODF chart ownership for presentations.
//!
//! ODP owns drawing-page occurrence discovery and package topology. The
//! shared ODF chart crate owns the typed chart grammar and authoring model;
//! this module binds those capabilities to selector-first, clone-staged
//! presentation edits.

mod codec;
mod model;
mod package;
mod snapshot;
mod transaction;

pub(crate) use codec::{locate_pages, page_index};

#[cfg(test)]
mod tests;

pub use model::{Chart, Limits, Page, Part, Selector, Storage};
pub use snapshot::{Commit, Diagnostics, Edit, Patch, Snapshot};
pub use transaction::{Commit as InventoryCommit, Editor, Inventory, Transaction};

pub use litchi_odf_common::chart::authoring::{CachedCell, Definition, SeriesSpec};
pub use litchi_odf_common::chart::{Axis, DataPoint, Element, Grid, Legend, PlotArea, Series};

pub(crate) fn inventory(
    package: &crate::core::family::Package,
    limits: Limits,
) -> litchi_core::Result<Inventory<'_>> {
    Inventory::load(package.package(), limits)
}
