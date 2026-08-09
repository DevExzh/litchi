//! Layered `SpreadsheetML` external data connection owner.
//!
//! Typed connection models live in the model module, bounded XML conversion in the codec
//! module, and workbook relationship integration in the package module. External targets
//! remain inert data and are never dereferenced.

mod codec;
mod model;
mod package;

/// Source snapshots for the contextual connections owner.
pub mod snapshot {
    pub use super::package::Snapshot;
}

/// Reversible source-checked connection patches.
pub mod patch {
    pub use super::package::{Commit, Patch};
}

/// Failure-atomic semantic transactions over the workbook connection owner.
pub mod transaction {
    pub use super::package::Transaction;
}

/// Typed and package-level validation for inert connection metadata.
pub mod validation {
    use litchi_opc::OpcPackage;

    use super::model::Connections;
    use super::package;
    use litchi_core::sheet::Result;

    pub fn connections(value: &Connections) -> Result<()> {
        value.to_xml(false).map(|_| ())
    }

    pub fn graph(package: &OpcPackage) -> Result<Option<Connections>> {
        package::validate_graph(package)
    }
}

#[cfg(test)]
mod tests;

pub use model::*;
pub use package::{
    Commit, Patch, Snapshot, Transaction, load_from_package, remove_from_package, store_in_package,
    store_in_package_with_query_table_validator, validate_graph,
};

fn invalid(v: impl Into<String>) -> Box<dyn std::error::Error + Send + Sync> {
    std::io::Error::new(std::io::ErrorKind::InvalidData, v.into()).into()
}
