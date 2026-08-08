//! Contextual OPC entry points for workbook calculation metadata.

use litchi_opc::OpcPackage;

use super::{Limits, Patch, Snapshot, Transaction};
use crate::error::Result;

/// Load a source-bound calculation-metadata snapshot.
pub fn load(package: &OpcPackage) -> Result<Snapshot> {
    Snapshot::load(package)
}

/// Load with a caller-supplied calculation-metadata resource policy.
pub fn load_with_limits(package: &OpcPackage, limits: &Limits) -> Result<Snapshot> {
    Snapshot::load_with_limits(package, limits)
}

/// Start a source-bound calculation-metadata transaction.
pub fn edit(package: &mut OpcPackage) -> Result<Transaction<'_>> {
    Transaction::new(package)
}

/// Start a transaction with a caller-supplied resource policy.
pub fn edit_with_limits<'a>(
    package: &'a mut OpcPackage,
    limits: &Limits,
) -> Result<Transaction<'a>> {
    Transaction::with_limits(package, limits)
}

/// Apply an exact source-bound calculation-metadata patch.
pub fn apply_patch(package: &mut OpcPackage, patch: &Patch) -> Result<()> {
    patch.apply(package)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn contextual_entry_points_share_the_source_bound_lifecycle() {
        let mut package = crate::package::build_minimal_package().unwrap();
        assert!(load(&package).unwrap().properties().is_none());
        let commit = edit(&mut package).unwrap().commit().unwrap();
        assert!(!commit.changed());
        apply_patch(&mut package, commit.patch()).unwrap();
    }
}
