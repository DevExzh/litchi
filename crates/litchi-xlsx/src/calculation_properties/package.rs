//! Contextual OPC read entry point for workbook calculation metadata.

use litchi_opc::OpcPackage;

use super::Snapshot;
use crate::error::Result;

/// Load a source-bound calculation-metadata snapshot.
pub fn load(package: &OpcPackage) -> Result<Snapshot> {
    Snapshot::load(package)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn the_package_entry_points_share_the_source_bound_lifecycle() {
        let package = crate::Package::new().unwrap();
        assert!(load(package.opc()).unwrap().properties().is_none());
        let commit = package
            .edit_calculation_metadata()
            .unwrap()
            .commit()
            .unwrap();
        assert!(!commit.diagnostics().changed());
        let replayed = package.apply_calculation_metadata(commit.patch()).unwrap();
        assert!(!replayed.diagnostics().changed());
    }
}
