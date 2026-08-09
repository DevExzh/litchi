//! Contextual OPC entry points for worksheet page breaks.

use litchi_opc::OpcPackage;

use super::{Patch, Snapshot, Transaction};
use crate::Selector;
use crate::error::Result;

/// Load one source-bound worksheet page-break snapshot.
///
/// # Errors
///
/// Returns an error when the package or selector is invalid, the selected
/// sheet is not a worksheet, or its page-break XML is invalid.
pub fn load<'a>(package: &OpcPackage, selector: impl Into<Selector<'a>>) -> Result<Snapshot> {
    Snapshot::load(package, selector)
}

/// Start one source-bound worksheet page-break transaction.
///
/// # Errors
///
/// Returns an error when loading the selected worksheet fails.
pub fn edit<'package, 'selector>(
    package: &'package mut OpcPackage,
    selector: impl Into<Selector<'selector>>,
) -> Result<Transaction<'package>> {
    Transaction::new(package, selector)
}

/// Apply an exact source-bound page-break patch.
///
/// # Errors
///
/// Returns an error when the patch source is stale, the package is signed, or
/// publication/readback fails.
pub fn apply_patch(package: &mut OpcPackage, patch: &Patch) -> Result<()> {
    patch.apply(package)
}
