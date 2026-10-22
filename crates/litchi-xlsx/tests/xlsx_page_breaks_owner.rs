#![allow(
    clippy::unwrap_used,
    reason = "focused integration tests use panic-on-failure assertions"
)]

use litchi_xlsx::page_breaks::{Break, Collection, PageBreaks};
use litchi_xlsx::{Error, Package};

const ROW_XML: &[u8] = b"<rowBreaks count=\"1\" manualBreakCount=\"1\"><brk id=\"12\" max=\"16383\" man=\"1\"/></rowBreaks>";

#[test]
fn package_transaction_round_trips_and_inverts_exactly() {
    let source_package = Package::new().unwrap();
    let source = source_package
        .page_breaks("Sheet1")
        .unwrap()
        .source_arc()
        .unwrap();

    let mut edit = source_package.edit_page_breaks("Sheet1").unwrap();
    assert!(
        edit.set_horizontal(
            Collection::horizontal([Break::new(12, 0, 16_383).unwrap().with_manual(true)]).unwrap()
        )
        .unwrap()
    );
    assert!(
        edit.set_vertical(Collection::vertical([Break::new(4, 2, 90).unwrap()]).unwrap())
            .unwrap()
    );
    let commit = edit.commit().unwrap();
    assert!(commit.diagnostics().changed());
    let patch = commit.patch().clone();
    let xml = patch.after().source_xml();
    assert!(xml.windows(ROW_XML.len()).any(|window| window == ROW_XML));
    // The source package is unchanged.
    assert_eq!(
        source_package.page_breaks("Sheet1").unwrap().page_breaks(),
        &PageBreaks::new()
    );

    let package = commit.into_package();
    let bytes = package.to_plain_bytes().unwrap();
    let reopened = Package::from_bytes(&bytes).unwrap();
    let value = reopened.page_breaks(0usize).unwrap();
    assert_eq!(value.page_breaks(), patch.after().page_breaks());

    let package = package
        .apply_page_breaks(&patch.inverse())
        .unwrap()
        .into_package();
    let restored = package.page_breaks("Sheet1").unwrap();
    assert_eq!(restored.page_breaks(), &PageBreaks::new());
    assert_eq!(restored.source_xml(), source.as_slice());
}

#[test]
fn no_op_shares_source_and_stale_patch_is_rejected() {
    let package = Package::new().unwrap();
    let source = package.page_breaks(0usize).unwrap().source_arc().unwrap();
    let commit = package.edit_page_breaks(0usize).unwrap().commit().unwrap();
    assert!(!commit.diagnostics().changed());
    assert!(commit.patch().is_noop());
    assert!(std::sync::Arc::ptr_eq(
        &source,
        &commit
            .package()
            .page_breaks(0usize)
            .unwrap()
            .source_arc()
            .unwrap()
    ));

    let mut edit = package.edit_page_breaks(0usize).unwrap();
    edit.set_horizontal(Collection::horizontal([Break::new(1, 0, 1).unwrap()]).unwrap())
        .unwrap();
    let patch = edit.commit().unwrap().patch().clone();

    let mut divergent = package.edit_page_breaks(0usize).unwrap();
    divergent
        .set_vertical(Collection::vertical([Break::new(1, 0, 1).unwrap()]).unwrap())
        .unwrap();
    let divergent = divergent.commit().unwrap().into_package();
    assert!(matches!(
        divergent.apply_page_breaks(&patch),
        Err(Error::PatchConflict { .. })
    ));
    // The patch still replays onto its own source.
    assert!(
        package
            .apply_page_breaks(&patch)
            .unwrap()
            .diagnostics()
            .changed()
    );
}

#[test]
fn worksheet_facade_uses_the_canonical_owner() {
    let package = Package::new().unwrap();
    let workbook = package.workbook();
    let sheet = workbook.sheet("Sheet1").unwrap().unwrap();
    assert_eq!(sheet.page_breaks().unwrap(), PageBreaks::new());
}

#[test]
fn reads_real_poi_manual_breaks() {
    let package = Package::open(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../test-data/poi/test-data/spreadsheet/ConditionalFormattingSamples.xlsx"
    ))
    .unwrap();
    let workbook = package.workbook();
    let values = workbook
        .sheets()
        .map(|sheet| {
            sheet
                .page_breaks()
                .unwrap_or_else(|error| panic!("{}: {error}", sheet.name()))
        })
        .filter(|value| value.horizontal().is_some() || value.vertical().is_some())
        .collect::<Vec<_>>();
    assert_eq!(values.len(), 1);
    assert_eq!(values[0].horizontal().unwrap().breaks()[0].id(), 21);
    assert_eq!(values[0].vertical().unwrap().breaks()[0].id(), 20);
    assert!(values[0].horizontal().unwrap().breaks()[0].is_manual());
    assert!(values[0].vertical().unwrap().breaks()[0].is_manual());
}
