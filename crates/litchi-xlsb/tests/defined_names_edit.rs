#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "integration tests use panic-on-failure extraction"
)]

//! Defined names through the primary workbook transaction (record 0261).

use std::path::PathBuf;

use litchi_core::sheet::WorkbookTrait as _;

use litchi_xlsb::Package;
use litchi_xlsb::cell_values::TransferLimits;
use litchi_xlsb::cell_values::structure::Patch;
use litchi_xlsb::formula::Range;
use litchi_xlsb::named_ranges::Entry;
use litchi_xlsb::package::PackageError;

fn fixture(name: &str) -> Package {
    let bytes = std::fs::read(
        PathBuf::from(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/ooxml/xlsb")
            .join(name),
    )
    .expect("XLSB fixture");
    Package::from_bytes(&bytes).unwrap()
}

fn entries(package: &Package) -> Vec<Entry> {
    package.workbook().unwrap().defined_name_entries().unwrap()
}

fn area(row_first: u32, row_last: u32, col_first: u32, col_last: u32) -> Range {
    Range::new(row_first, row_last, col_first, col_last).unwrap()
}

#[test]
fn a_name_is_added_to_a_workbook_without_an_externals_block() {
    let source = fixture("Simple.xlsb");
    assert!(entries(&source).is_empty());
    let sheet = source.workbook().unwrap().worksheet_names()[0].clone();
    let mut edit = source.edit().unwrap();
    edit.set_defined_name("Totals", None, 0, area(0, 4, 1, 1))
        .unwrap()
        .unwrap();
    let commit = edit.commit().unwrap();
    assert!(commit.diagnostics().changed());
    let (named, patch) = commit.into_parts();
    let names = entries(&named);
    assert_eq!(names.len(), 1);
    assert_eq!(names[0].name, "Totals");
    assert_eq!(names[0].scope, None);
    assert_eq!(
        names[0].formula.as_deref(),
        Some(format!("{sheet}!$B$1:$B$5").as_str())
    );

    // The durable patch replays and inverts exactly.
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
        named.to_bytes().unwrap()
    );
    let restored = named
        .apply(&patch.inverse())
        .map(litchi_xlsb::cell_values::structure::Commit::into_package)
        .unwrap();
    assert_eq!(restored.to_bytes().unwrap(), patch.before());
    assert_eq!(patch.before(), source.to_bytes().unwrap());
    assert!(entries(&restored).is_empty());
}

#[test]
fn existing_names_keep_their_records_and_new_names_append() {
    let source = fixture("bug66682.xlsb");
    let before = entries(&source);
    assert_eq!(before.len(), 2);
    let first = before[0].clone();
    let mut edit = source.edit().unwrap();
    edit.set_defined_name(
        &first.name,
        first.scope.map(|scope| scope.into()),
        0,
        area(9, 9, 2, 3),
    )
    .unwrap()
    .unwrap();
    edit.set_defined_name("Appended", None, 0, area(0, 0, 0, 0))
        .unwrap()
        .unwrap();
    let after = edit.commit().unwrap().into_package();
    let names = entries(&after);
    assert_eq!(names.len(), 3);
    assert_eq!(names[0].name, first.name);
    assert!(
        names[0]
            .formula
            .as_deref()
            .unwrap()
            .ends_with("!$C$10:$D$10")
    );
    assert_eq!(names[1], before[1], "an untouched name is unchanged");
    assert_eq!(names[2].name, "Appended");
}

#[test]
fn scoped_names_removal_and_no_ops() {
    let source = Package::new().unwrap();
    let mut edit = source.edit().unwrap();
    edit.set_defined_name("Local", Some(0.into()), 0, area(1, 2, 1, 2))
        .unwrap()
        .unwrap();
    edit.set_defined_name("Global", None, 0, area(0, 0, 0, 0))
        .unwrap()
        .unwrap();
    let named = edit.commit().unwrap().into_package();
    let names = entries(&named);
    assert_eq!(names[0].scope, Some(0));
    assert_eq!(names[1].scope, None);

    // Setting the same reference is an exact no-op.
    let mut edit = named.edit().unwrap();
    edit.set_defined_name("global", None, 0, area(0, 0, 0, 0))
        .unwrap()
        .unwrap();
    let commit = edit.commit().unwrap();
    assert!(!commit.diagnostics().changed());

    // Only the last name can be removed; a missing name is a no-op.
    let mut edit = named.edit().unwrap();
    assert!(matches!(
        edit.remove_defined_name("Local", Some(0.into())),
        Err(PackageError::Refused(_))
    ));
    edit.remove_defined_name("Missing", None).unwrap().unwrap();
    edit.remove_defined_name("Global", None).unwrap().unwrap();
    let removed = edit.commit().unwrap().into_package();
    let names = entries(&removed);
    assert_eq!(names.len(), 1);
    assert_eq!(names[0].name, "Local");
}

#[test]
fn invalid_names_and_targets_are_refused_before_staging() {
    let source = Package::new().unwrap();
    let mut edit = source.edit().unwrap();
    for name in ["", "A1", "TRUE", "has space"] {
        assert!(
            edit.set_defined_name(name, None, 0, area(0, 0, 0, 0))
                .is_err(),
            "{name:?}"
        );
    }
    assert_eq!(
        edit.set_defined_name("Valid", None, "Missing", area(0, 0, 0, 0))
            .unwrap(),
        None
    );
    assert!(edit.commit().unwrap().patch().is_noop());
}
