//! Regression coverage for the raw package snapshot transaction boundary.
//!
//! The test intentionally uses only the legacy raw facade.  A no-op edit must
//! publish an empty, reversible patch while retaining the exact source bytes.

use std::error::Error;

#[allow(deprecated)]
use litchi_iwa::raw::package::IWorkPackage;

#[test]
#[allow(deprecated)]
fn noop_snapshot_commit_replays_and_reverts_exact_source_bytes() -> Result<(), Box<dyn Error>> {
    let mut package = IWorkPackage::new();
    package.insert_entry("Data/source", b"source bytes".to_vec())?;
    let source_bytes = package.to_bytes()?;

    let source = IWorkPackage::from_bytes(&source_bytes)?.snapshot();
    let commit = source.edit_with(|_| Ok(()))?;
    assert!(commit.patch().is_empty());

    let replayed = source.apply(commit.patch())?;
    assert_eq!(replayed.to_bytes()?, source_bytes);

    let inverse = commit.patch().inverse();
    assert!(inverse.is_empty());
    let reverted = replayed.apply(&inverse)?;
    assert_eq!(reverted.to_bytes()?, source_bytes);

    Ok(())
}
