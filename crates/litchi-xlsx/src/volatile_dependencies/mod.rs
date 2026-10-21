//! Layered, inert `SpreadsheetML` volatile-dependency records.
//!
//! Typed semantic values live in `model`, bounded XML conversion in
//! `codec`, and workbook relationship traversal in `package`.
//! [`crate::Package::edit_volatile_dependencies`] starts the [`Edit`] family.

mod codec;
mod edit;
mod model;
mod package;

#[cfg(test)]
mod tests;

pub use edit::{Commit, Diagnostics, Edit, Error, LimitKind, Patch};
pub use model::*;
pub use package::{load_from_package, load_from_package_with_conformance};
#[cfg(test)]
pub(crate) use package::{remove_from_package, store_in_package};

pub(super) fn invalid(message: impl Into<String>) -> Error {
    crate::error::invalid(message)
}
