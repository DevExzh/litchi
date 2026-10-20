//! The ADR 0034 owner contract: `Package` owns every state verb, `Workbook`
//! is the read-only projection, and commits carry the package, patch and
//! diagnostics triad.

use litchi_xlsx::workbook::edit::LimitKind;
use litchi_xlsx::{Error, OpenOptions, Package, SourceLimitKind, SourceLimits};

fn stored(package: &Package, address: &str) -> bool {
    package
        .workbook()
        .sheet(0_usize)
        .unwrap()
        .expect("first worksheet")
        .cell(address)
        .unwrap()
        .stored()
        .is_some()
}

#[test]
fn a_changed_commit_publishes_a_new_package_and_leaves_the_source_intact() {
    let source = Package::new().unwrap();
    let source_bytes = source.to_bytes().unwrap();
    let mut edit = source.edit().unwrap();
    edit.sheet(0_usize)
        .unwrap()
        .expect("first worksheet")
        .set("B2", 7)
        .unwrap();
    let commit = edit.commit().unwrap();

    assert!(commit.diagnostics().changed());
    assert!(!commit.patch().is_noop());
    assert!(commit.diagnostics().changed_parts() >= 1);
    assert_eq!(commit.patch().operation_count(), 1);
    assert!(stored(commit.package(), "B2"));

    // Snapshot immutability: the source and its projection are untouched.
    assert!(!stored(&source, "B2"));
    assert_eq!(source.to_bytes().unwrap(), source_bytes);

    let (published, patch) = commit.into_parts();
    let inverse = published.apply(&patch.inverse()).unwrap();
    assert!(inverse.diagnostics().changed());
    assert!(!stored(inverse.package(), "B2"));
}

#[test]
fn an_exact_no_op_commit_is_the_source_byte_for_byte() {
    let bytes = Package::new().unwrap().to_bytes().unwrap();
    let source = Package::from_bytes(&bytes).unwrap();
    let commit = source.edit().unwrap().commit().unwrap();

    assert!(!commit.diagnostics().changed());
    assert!(commit.patch().is_noop());
    assert_eq!(commit.diagnostics().changed_parts(), 0);
    assert!(commit.package().same_content(&source).unwrap());
    assert_eq!(commit.into_package().to_bytes().unwrap(), bytes);
}

#[test]
fn configured_ingress_refuses_oversized_input_before_copying_it() {
    let bytes = Package::new().unwrap().to_bytes().unwrap();
    let limit = u64::try_from(bytes.len()).unwrap() - 1;
    let limits = SourceLimits::builder()
        .max_input_bytes(limit)
        .unwrap()
        .build()
        .unwrap();
    let error = Package::from_bytes_with_options(&bytes, OpenOptions::new().source_limits(limits))
        .expect_err("input over the limit");
    assert!(matches!(
        error,
        Error::Package(litchi_opc::OpcError::Limit {
            kind: SourceLimitKind::InputBytes,
            maximum,
            ..
        }) if maximum == limit
    ));
}

#[test]
fn transaction_limits_report_their_kind() {
    let source = Package::new().unwrap();
    let mut edit = source.edit().unwrap();
    let error = edit
        .copy_cells(0_usize, "A1:Z65536", 0_usize, "AA1")
        .expect_err("transfer above the cell bound");
    assert!(matches!(
        error,
        Error::Limit {
            kind: litchi_xlsx::LimitKind::Workbook(LimitKind::TransferCells),
            observed: 1_703_936,
            maximum: 65_536,
        }
    ));
}

#[test]
fn scoped_package_edits_publish_a_validated_snapshot() {
    let source = Package::new().unwrap();
    let mut transaction = source.edit_task_panes().unwrap();
    let reference = litchi_ooxml_common::web::Reference::new(
        "wa1",
        "1.0.0.0",
        litchi_ooxml_common::web::Store::Omex,
    )
    .unwrap();
    let add_in = litchi_ooxml_common::web::AddIn::new("add-in-1", reference).unwrap();
    let mut panes = litchi_ooxml_common::web::Panes::new();
    panes
        .push(litchi_ooxml_common::web::Pane::new(add_in))
        .unwrap();
    transaction.replace(panes).unwrap();
    let package = transaction.commit().unwrap().into_package();

    // The published package and its projection observe the new graph; the
    // source package does not.
    assert!(package.task_panes().unwrap().is_some());
    assert!(package.workbook().task_panes().unwrap().is_some());
    assert!(source.task_panes().unwrap().is_none());
}
