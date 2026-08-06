//! Embedded ODF chart ownership for spreadsheets.
//!
//! The ODS host owns occurrence selection and package topology. The common
//! ODF chart crate owns the retained chart tree and typed authoring grammar;
//! this module only binds those capabilities to `table:shapes` occurrences.

mod codec;
mod model;
mod package;
mod transaction;

#[cfg(test)]
mod tests;

pub use model::{Chart, Limits, Part, Selector, Storage};
pub use transaction::{Commit, Editor, Inventory, Transaction};

pub use litchi_odf_common::chart::authoring::Definition;
pub use litchi_odf_common::chart::{Axis, DataPoint, Element, Grid, Legend, PlotArea, Series};

pub(crate) fn inventory(
    package: &crate::package::Package,
    limits: Limits,
) -> litchi_core::Result<Inventory<'_>> {
    Inventory::load(package, limits)
}
