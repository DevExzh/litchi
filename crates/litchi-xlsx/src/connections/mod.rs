//! Layered `SpreadsheetML` external data connection owner.
//!
//! Typed connection models live in the model module, bounded XML conversion in the codec
//! module, and workbook relationship integration in the package module. External targets
//! remain inert data and are never dereferenced.
//!
//! [`crate::Package::edit_connections`] starts an [`Edit`] of the workbook's
//! connection catalog; [`Edit::commit`] returns a [`Commit`] holding the new
//! package, the exact [`Patch`] and the [`Diagnostics`], and
//! [`crate::Package::apply_connections`] replays a patch. The source package
//! never changes.

mod codec;
pub(crate) mod embedded_data;
mod model;
mod namespace;
mod package;

/// Typed and package-level validation for inert connection metadata.
pub mod validation {
    use litchi_opc::{OpcPackage, SourceLimits};

    use super::model::Connections;
    use super::package;
    use crate::error::Result;

    pub fn connections(value: &Connections, limits: &SourceLimits) -> Result<()> {
        value.to_xml(false, limits).map(|_| ())
    }

    pub fn graph(package: &OpcPackage, limits: &SourceLimits) -> Result<Option<Connections>> {
        package::validate_graph(package, limits)
    }
}

#[cfg(test)]
mod tests;

pub use model::*;
#[cfg(test)]
pub(crate) use package::Transaction;
pub use package::{
    Commit, Diagnostics, Edit, Error, LimitKind, Patch, Snapshot, load_from_package, validate_graph,
};

fn invalid(v: impl Into<String>) -> Error {
    crate::error::invalid(v.into())
}

pub(crate) fn limit(kind: LimitKind, observed: usize, maximum: usize) -> Error {
    Error::limit(kind, observed, maximum)
}
