//! Embedded ODF chart ownership for presentations.
//!
//! ODP owns drawing-page occurrence discovery and package topology. The
//! shared ODF chart crate owns the typed chart grammar and authoring model;
//! this module binds those capabilities to selector-first, clone-staged
//! presentation edits.

mod codec;
mod model;
mod package;
mod transaction;

#[cfg(test)]
mod tests;

pub use model::{Chart, Limits, Page, Part, Selector, Storage};
pub use transaction::{Commit, Editor, Inventory, Transaction};

pub use litchi_odf_common::chart::authoring::Definition;
pub use litchi_odf_common::chart::{Axis, DataPoint, Element, Grid, Legend, PlotArea, Series};

pub(crate) fn inventory(
    package: &crate::core::family::Package,
    limits: Limits,
) -> litchi_core::Result<Inventory<'_>> {
    Inventory::load(package.package(), limits)
}
