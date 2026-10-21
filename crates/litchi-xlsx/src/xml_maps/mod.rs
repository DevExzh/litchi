//! Layered `SpreadsheetML` Custom XML Maps ownership.
//!
//! Typed contextual values live in `model`, bounded XML conversion lives in
//! `codec`, workbook relationship traversal lives in `package`, and focused
//! regression coverage lives in `tests`. Schema and binding payloads remain
//! inert, bounded XML rather than being resolved or executed.
//!
//! [`crate::Package::edit_xml_maps`] starts an [`Edit`] of the workbook's
//! Custom XML Maps owner; [`Edit::commit`] returns a [`Commit`] holding the
//! new package, the exact [`Patch`] and the [`Diagnostics`], and
//! [`crate::Package::apply_xml_maps`] replays a patch. The source package
//! never changes.

mod codec;
mod edit;
mod model;
mod package;
mod patch;
mod snapshot;

#[cfg(test)]
mod tests;

pub use codec::{parse_xml_map_info, serialize_xml_map_info};
#[cfg(test)]
pub(crate) use edit::Transaction;
pub use edit::{Edit, EditOptions};
pub use model::{
    XmlMap, XmlMapConformance, XmlMapDataBinding, XmlMapInfo, XmlMapLimits, XmlMapSchema,
};
pub use package::{load_from_package, load_from_package_with_conformance};
#[cfg(test)]
pub(crate) use package::{remove_from_package, store_in_package};
pub use patch::{Commit, Diagnostics, Error, LimitKind, Patch};
pub use snapshot::Snapshot;

fn invalid(message: impl Into<String>) -> Error {
    crate::error::invalid(message)
}
