#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "integration tests use panic-on-failure extraction"
)]

//! Merged ranges through the primary workbook transaction (record 0261).

use litchi_xlsb::Package;
use litchi_xlsb::builder::{self, Builder};
use litchi_xlsb::cell_values::TransferLimits;
use litchi_xlsb::cell_values::structure::Patch;
use litchi_xlsb::merged_cells::MergedCell;
use litchi_xlsb::package::PackageError;

fn package() -> Package {
    let mut first = builder::Worksheet::new("First");
    first.set_cell(0, 0, "title");
    first.set_cell(0, 1, 2.0_f64);
    let mut builder = Builder::new();
    builder.add_worksheet(first);
    builder.add_worksheet(builder::Worksheet::new("Second"));
    builder.build().unwrap()
}

fn merges(package: &Package, sheet: usize) -> Vec<MergedCell> {
    package
        .workbook()
        .unwrap()
        .merged_cell_ranges(sheet)
        .unwrap()
}

#[test]
fn merges_and_unmerges_publish_and_replay_exactly() {
    let source = package();
    let mut edit = source.edit().unwrap();
    assert_eq!(
        edit.merge_cells("First", MergedCell::new(0, 0, 0, 2))
            .unwrap(),
        Some(true)
    );
    assert_eq!(
        edit.merge_cells("First", MergedCell::new(2, 3, 0, 0))
            .unwrap(),
        Some(true)
    );
    assert_eq!(
        edit.merge_cells("Second", MergedCell::new(4, 5, 4, 5))
            .unwrap(),
        Some(true)
    );
    assert_eq!(
        edit.merge_cells("Missing", MergedCell::new(0, 1, 0, 1))
            .unwrap(),
        None
    );
    let commit = edit.commit().unwrap();
    assert!(commit.diagnostics().changed());
    let (merged, patch) = commit.into_parts();
    assert_eq!(
        merges(&merged, 0),
        vec![MergedCell::new(0, 0, 0, 2), MergedCell::new(2, 3, 0, 0)]
    );
    assert_eq!(merges(&merged, 1), vec![MergedCell::new(4, 5, 4, 5)]);

    // The durable patch round-trips and replays; its inverse restores bytes.
    let wire = patch.to_bytes(TransferLimits::DEFAULT).unwrap();
    let decoded = Patch::from_bytes(
        &wire,
        TransferLimits::DEFAULT,
        litchi_xlsb::raw::Limits::DEFAULT,
    )
    .unwrap();
    assert_eq!(
        source
            .apply(&decoded)
            .map(litchi_xlsb::cell_values::structure::Commit::into_package)
            .unwrap()
            .to_bytes()
            .unwrap(),
        merged.to_bytes().unwrap()
    );
    assert_eq!(
        merged
            .apply(&patch.inverse())
            .map(litchi_xlsb::cell_values::structure::Commit::into_package)
            .unwrap()
            .to_bytes()
            .unwrap(),
        source.to_bytes().unwrap()
    );

    let mut edit = merged.edit().unwrap();
    assert_eq!(
        edit.unmerge_cells(0, MergedCell::new(2, 3, 0, 0)).unwrap(),
        Some(true)
    );
    assert_eq!(
        edit.unmerge_cells(0, MergedCell::new(2, 3, 0, 0)).unwrap(),
        Some(false)
    );
    let unmerged = edit.commit().unwrap().into_package();
    assert_eq!(merges(&unmerged, 0), vec![MergedCell::new(0, 0, 0, 2)]);
}

#[test]
fn exact_no_ops_keep_the_package_bytes() {
    let source = package();
    let mut edit = source.edit().unwrap();
    assert_eq!(
        edit.unmerge_cells(0, MergedCell::new(5, 6, 5, 6)).unwrap(),
        Some(false)
    );
    let commit = edit.commit().unwrap();
    assert!(!commit.diagnostics().changed());
    assert_eq!(
        commit.into_package().to_bytes().unwrap(),
        source.to_bytes().unwrap()
    );

    let mut edit = source.edit().unwrap();
    edit.merge_cells(0, MergedCell::new(0, 1, 0, 1)).unwrap();
    edit.unmerge_cells(0, MergedCell::new(0, 1, 0, 1)).unwrap();
    let commit = edit.commit().unwrap();
    assert!(commit.patch().is_noop());
}

#[test]
fn overlapping_and_single_cell_ranges_are_refused_before_staging() {
    let source = package();
    let mut edit = source.edit().unwrap();
    edit.merge_cells(0, MergedCell::new(0, 1, 0, 1)).unwrap();
    assert!(matches!(
        edit.merge_cells(0, MergedCell::new(1, 2, 1, 2)),
        Err(PackageError::Refused(_))
    ));
    assert!(matches!(
        edit.unmerge_cells(0, MergedCell::new(1, 2, 1, 2)),
        Err(PackageError::Refused(_))
    ));
    assert!(matches!(
        edit.merge_cells(0, MergedCell::new(3, 3, 3, 3)),
        Err(PackageError::InvalidFormat(_))
    ));
    assert!(edit.merge_cells(0, MergedCell::new(2, 1, 0, 0)).is_err());
    let merged = edit.commit().unwrap().into_package();
    assert_eq!(merges(&merged, 0), vec![MergedCell::new(0, 1, 0, 1)]);
}
