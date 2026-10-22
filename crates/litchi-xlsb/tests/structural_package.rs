#![allow(
    clippy::expect_used,
    reason = "integration assertions report fixture failures"
)]

use litchi_core::sheet::traits::WorkbookTrait as _;
use litchi_xlsb::Package;
use litchi_xlsb::cell_values::TransferLimits;
use litchi_xlsb::cell_values::structure::Patch;

#[test]
fn package_edit_commit_apply_and_inverse_are_immutable() {
    let source = Package::new().expect("source");
    let original = source.to_bytes().expect("original bytes");
    let mut edit = source.edit().expect("edit");
    edit.set_sheet_name(0, "Renamed".to_owned())
        .expect("rename")
        .expect("resolved worksheet selector");
    let commit = edit.commit().expect("commit");
    assert!(commit.diagnostics().changed());
    assert_eq!(source.to_bytes().expect("unchanged source"), original);
    assert_eq!(
        commit.package().to_bytes().expect("candidate bytes"),
        commit.patch().after()
    );
    assert_eq!(
        commit.package().workbook().expect("view").worksheet_names(),
        &["Renamed".to_owned()]
    );
    let bytes = commit
        .patch()
        .to_bytes(TransferLimits::DEFAULT)
        .expect("wire");
    let patch = Patch::from_bytes(
        &bytes,
        TransferLimits::DEFAULT,
        litchi_xlsb::raw::Limits::DEFAULT,
    )
    .expect("decode");
    let applied = source
        .apply(&patch)
        .map(litchi_xlsb::cell_values::structure::Commit::into_package)
        .expect("immutable apply");
    assert_eq!(
        applied.to_bytes().expect("applied bytes"),
        commit.patch().after()
    );
    assert!(
        applied
            .apply(&patch)
            .map(litchi_xlsb::cell_values::structure::Commit::into_package)
            .is_err()
    );
    let reverted = applied
        .apply(&patch.inverse())
        .map(litchi_xlsb::cell_values::structure::Commit::into_package)
        .expect("inverse");
    assert_eq!(reverted.to_bytes().expect("restored bytes"), original);
    let (package, patch) = commit.into_parts();
    assert_eq!(package.to_bytes().expect("consumed bytes"), patch.after());
}

#[test]
fn noop_and_failed_edits_keep_the_package_source() {
    let source = Package::new().expect("source");
    let original = source.to_bytes().expect("original bytes");
    let noop = source.edit().expect("edit").commit().expect("commit");
    assert!(!noop.diagnostics().changed());
    assert!(noop.patch().is_noop());
    assert_eq!(
        noop.into_package().to_bytes().expect("no-op bytes"),
        original
    );
    let mut edit = source.edit().expect("edit");
    assert!(edit.set_sheet_name(0, "invalid/name".to_owned()).is_err());
    let commit = edit.commit().expect("failed staging leaves edit unchanged");
    assert!(commit.patch().is_noop());
    assert_eq!(
        source
            .apply(commit.patch())
            .map(litchi_xlsb::cell_values::structure::Commit::into_package)
            .expect("apply no-op")
            .to_bytes()
            .expect("bytes"),
        original
    );
    assert_eq!(source.to_bytes().expect("source bytes"), original);
}

#[test]
fn retained_input_budget_is_checked_before_commit_and_allows_bounded_chaining() {
    let bytes = Package::new().expect("source").to_bytes().expect("bytes");
    let size = u64::try_from(bytes.len()).expect("size");
    let exact = litchi_opc::SourceLimits::builder()
        .max_input_bytes(size)
        .expect("input")
        .build()
        .expect("limits");
    let source = Package::from_bytes_with_options(
        &bytes,
        litchi_xlsb::OpenOptions::new().source_limits(exact),
    )
    .expect("bounded source");
    let original = source.to_bytes().expect("source bytes");
    let mut edit = source.edit().expect("edit");
    edit.set_sheet_name(0, "A significantly longer name".to_owned())
        .expect("stage rename")
        .expect("resolved worksheet selector");
    assert!(
        edit.commit().is_err(),
        "growth beyond retained input budget must fail before publication"
    );
    assert_eq!(source.to_bytes().expect("unchanged"), original);
    assert!(
        source
            .edit()
            .expect("no-op edit")
            .commit()
            .expect("no-op commit")
            .patch()
            .is_noop()
    );

    let roomy = litchi_opc::SourceLimits::builder()
        .max_input_bytes(size + 4096)
        .expect("input")
        .build()
        .expect("limits");
    let mut package = Package::from_bytes_with_options(
        &bytes,
        litchi_xlsb::OpenOptions::new().source_limits(roomy),
    )
    .expect("bounded source");
    for name in ["A significantly longer name", "Second edit"] {
        let mut edit = package.edit().expect("edit");
        edit.set_sheet_name(0, name.to_owned())
            .expect("stage rename")
            .expect("resolved worksheet selector");
        let commit = edit.commit().expect("bounded commit");
        let applied = package
            .apply(commit.patch())
            .map(litchi_xlsb::cell_values::structure::Commit::into_package)
            .expect("same-policy apply");
        assert_eq!(
            applied.to_bytes().expect("apply bytes"),
            commit.package().to_bytes().expect("commit bytes")
        );
        assert_eq!(applied.opc_package().read_limits(), roomy);
        package = commit.into_package();
        assert_eq!(package.opc_package().read_limits(), roomy);
    }
    assert!(
        package
            .edit()
            .expect("no-op edit")
            .commit()
            .expect("no-op commit")
            .patch()
            .is_noop()
    );
}
