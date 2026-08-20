//! Transaction joins, conflicts, and cross-facet edit tests.

use super::super::*;
use super::support::part_text;

use crate::cell::{Number, Value};
use crate::{History, HistoryLimits, Margins, PageMargin};

fn strict_unknown_worksheet() -> Workbook {
    let source = Workbook::new().expect("source workbook");
    let mut package = source.inner.package.clone();
    package
        .get_part_mut(&source.inner.sheets[0].part_uri)
        .expect("worksheet part")
        .set_blob(
            br#"<x:worksheet xmlns:x="http://purl.oclc.org/ooxml/spreadsheetml/main" xmlns:u="urn:litchi:future"><x:sheetData><x:row r="1"><x:c r="A1"><x:v>7</x:v></x:c></x:row></x:sheetData><u:future marker="keep"/></x:worksheet>"#.to_vec(),
        );
    Workbook::from_package(package).expect("strict unknown worksheet")
}

fn test_margins() -> Margins {
    Margins::new(
        PageMargin::from_inches(0.7).expect("left margin"),
        PageMargin::from_inches(0.8).expect("right margin"),
        PageMargin::from_inches(1.0).expect("top margin"),
        PageMargin::from_inches(1.1).expect("bottom margin"),
        PageMargin::from_inches(0.3).expect("header margin"),
        PageMargin::from_inches(0.4).expect("footer margin"),
    )
}

#[test]
fn metadata_only_edit_preserves_strict_namespace_and_unknown_xml() {
    let source = strict_unknown_worksheet();
    let margins = test_margins();
    let mut edit = source.edit().expect("metadata edit");
    edit.put_page_margins("Sheet1", margins)
        .expect("put margins")
        .expect("worksheet selector");
    let committed = edit.commit().expect("metadata commit");
    let xml = part_text(committed.workbook(), "/xl/worksheets/sheet1.xml");

    assert!(xml.contains("<x:worksheet"));
    assert!(xml.contains("<u:future marker=\"keep\"/>"));
    assert!(xml.contains("<x:pageMargins"));
    assert_eq!(
        committed
            .workbook()
            .sheet("Sheet1")
            .expect("sheet lookup")
            .expect("worksheet")
            .page_margins()
            .expect("page margins read"),
        Some(margins)
    );
}

#[test]
fn metadata_only_edit_rejects_malformed_worksheet_like_the_source_parser() {
    let source = Workbook::new().expect("source workbook");
    let mut package = source.inner.package.clone();
    let uri = source.inner.sheets[0].part_uri.clone();
    let malformed = br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData></worksheet>"#;
    package
        .get_part_mut(&uri)
        .expect("worksheet part")
        .set_blob(malformed.to_vec());
    let malformed = Workbook::from_package(package).expect("catalog validation");
    let mut edit = malformed.edit().expect("metadata edit");
    edit.put_page_margins("Sheet1", test_margins())
        .expect("stage margins")
        .expect("worksheet selector");
    assert!(edit.commit().is_err());
}

#[test]
fn metadata_only_edit_with_the_existing_value_is_a_byte_no_op() {
    let source = Workbook::new().expect("source workbook");
    let margins = test_margins();
    let mut install = source.edit().expect("install edit");
    install
        .put_page_margins("Sheet1", margins)
        .expect("put margins")
        .expect("worksheet selector");
    let installed = install.commit().expect("install commit");
    let before = installed.workbook().to_plain_bytes().expect("before bytes");

    let mut no_op = installed.workbook().edit().expect("no-op edit");
    no_op
        .put_page_margins("Sheet1", margins)
        .expect("put existing margins")
        .expect("worksheet selector");
    let no_op = no_op.commit().expect("no-op commit");
    assert!(no_op.patch().is_empty());
    assert_eq!(
        no_op.workbook().to_plain_bytes().expect("after bytes"),
        before
    );
}

#[test]
fn page_margins_merge_replay_inverse_and_record_history() {
    let source = Workbook::new().expect("source workbook");
    let source_bytes = source.to_plain_bytes().expect("source bytes");
    let margins = Margins::new(
        PageMargin::from_inches(0.7).unwrap(),
        PageMargin::from_inches(0.8).unwrap(),
        PageMargin::from_inches(1.0).unwrap(),
        PageMargin::from_inches(1.1).unwrap(),
        PageMargin::from_inches(0.3).unwrap(),
        PageMargin::from_inches(0.4).unwrap(),
    );

    let mut margin_edit = source.edit().expect("margin edit");
    margin_edit
        .put_page_margins("Sheet1", margins)
        .expect("put margins")
        .expect("sheet selector");
    let mut cell_edit = source.edit().expect("cell edit");
    cell_edit
        .sheet("Sheet1")
        .expect("lookup")
        .expect("sheet")
        .set("H8", "disjoint")
        .expect("cell edit");
    let merged = margin_edit
        .plan_three_way(cell_edit, MergeLimits::new(2, 64, 128, 64))
        .expect("plan merge");
    assert!(merged.conflicts().is_empty());
    let commit = merged
        .finish()
        .expect("finish merge")
        .commit()
        .expect("commit margins");
    assert!(commit.patch().changes().iter().any(|change| {
        matches!(change.page_margins(), Some((None, Some(after))) if *after == margins)
    }));
    assert_eq!(
        commit
            .workbook()
            .sheet("Sheet1")
            .expect("lookup")
            .expect("sheet")
            .page_margins()
            .expect("read margins"),
        Some(margins)
    );

    let durable = commit.patch().durable().expect("durable margins patch");
    let json = durable
        .to_deterministic_json()
        .expect("deterministic margins patch");
    let durable = DurablePatch::from_deterministic_json(&json).expect("parse margins patch");
    let replayed = durable.apply(&source).expect("replay margins patch");
    assert_eq!(
        replayed.to_plain_bytes().expect("replayed bytes"),
        commit.workbook().to_plain_bytes().expect("committed bytes")
    );
    assert_eq!(
        durable
            .inverse()
            .apply(&replayed)
            .expect("inverse margins patch")
            .to_plain_bytes()
            .expect("restored bytes"),
        source_bytes
    );

    let mut stale_edit = source.edit().expect("stale edit");
    stale_edit
        .sheet("Sheet1")
        .expect("lookup")
        .expect("sheet")
        .set("A1", "stale")
        .expect("stale cell");
    let stale = stale_edit.commit().expect("stale commit").into_workbook();
    assert!(matches!(
        durable.apply(&stale),
        Err(Error::PatchConflict { .. })
    ));

    let mut history = History::new(source, HistoryLimits::new(2, 256 * 1024 * 1024));
    assert!(
        commit
            .record(&mut history)
            .expect("record margins")
            .is_empty()
    );
    assert!(history.undo());
    assert!(history.redo());
    assert_eq!(
        history
            .current()
            .sheet("Sheet1")
            .expect("lookup")
            .expect("sheet")
            .page_margins()
            .expect("history margins"),
        Some(margins)
    );
}

#[test]
fn independently_prepared_disjoint_edits_join_after_threaded_work() {
    fn assert_send<T: Send>() {}
    assert_send::<Edit>();

    let source = Workbook::new().expect("source workbook");
    let (mut left, right) = std::thread::scope(|scope| {
        let left_source = source.clone();
        let right_source = source.clone();
        let left = scope.spawn(move || {
            let mut edit = left_source.edit().expect("left edit");
            edit.sheet("Sheet1")
                .expect("lookup")
                .expect("sheet")
                .set("A1", "left")
                .expect("left cell");
            edit
        });
        let right = scope.spawn(move || {
            let mut edit = right_source.edit().expect("right edit");
            edit.sheet(0usize)
                .expect("lookup")
                .expect("sheet")
                .set("C3", 42_i32)
                .expect("right cell");
            edit
        });
        (
            left.join().expect("left worker"),
            right.join().expect("right worker"),
        )
    });

    left.join(right).expect("disjoint join");
    assert_eq!(left.len(), 2);
    let committed = left.commit().expect("joined commit");
    let sheet = committed
        .workbook()
        .sheet("Sheet1")
        .expect("lookup")
        .expect("sheet");
    assert!(matches!(
        sheet.cell("A1").expect("A1").stored(),
        Some(Cell::Value(Value::Text(text))) if text.as_str() == "left"
    ));
    assert!(matches!(
        sheet.cell("C3").expect("C3").stored(),
        Some(Cell::Value(Value::Number(number))) if number.as_str() == "42"
    ));

    let mut empty = source.edit().expect("empty edit");
    let mut incoming = source.edit().expect("incoming edit");
    incoming
        .sheet("Sheet1")
        .expect("lookup")
        .expect("sheet")
        .set("D4", true)
        .expect("incoming cell");
    empty.join(incoming).expect("adopt incoming sheet map");
    assert_eq!(empty.len(), 1);
}
#[test]
fn join_conflicts_are_structured_and_return_the_rejected_edit() {
    let source = Workbook::new().expect("source workbook");
    let mut left = source.edit().expect("left edit");
    let mut left_sheet = left.sheet(0usize).expect("lookup").expect("sheet");
    left_sheet
        .set("C3", "left tail")
        .expect("left tail cell")
        .set("A1", "left")
        .expect("left first cell");
    let mut right = source.edit().expect("right edit");
    let mut right_sheet = right.sheet("Sheet1").expect("lookup").expect("sheet");
    right_sheet
        .set("A1", "right")
        .expect("right first cell")
        .set("C3", "right tail")
        .expect("right tail cell");

    let error = match left.join(right) {
        Ok(_) => panic!("overlapping edits must not join"),
        Err(error) => error,
    };
    assert_eq!(left.len(), 2);
    let conflicts = error.conflicts().expect("overlap details");
    assert_eq!(conflicts.len(), 2);
    assert_eq!(conflicts.conflicts().len(), 1);
    assert_eq!(conflicts.conflicts()[0].sheet(), "Sheet1");
    assert_eq!(conflicts.conflicts()[0].position(), 0);
    assert_eq!(
        conflicts.conflicts()[0].cells().expect("cell conflicts"),
        &[
            Address::from_a1("A1").expect("first address"),
            Address::from_a1("C3").expect("tail address"),
        ]
    );
    let rejected = error.into_rejected();
    assert_eq!(rejected.len(), 2);

    let other_source = Workbook::new().expect("other source");
    let other = other_source.edit().expect("other edit");
    let error = match left.join(other) {
        Ok(_) => panic!("different snapshots must not join"),
        Err(error) => error,
    };
    assert!(matches!(error.failure(), JoinFailure::DifferentSnapshot));
    assert!(error.conflicts().is_none());
    assert_eq!(left.len(), 2);
}

#[test]
fn row_visibility_joins_with_cells_and_conflicts_by_row() {
    let source = Workbook::new().expect("source workbook");
    let mut cell = source.edit().expect("cell edit");
    cell.sheet(0usize)
        .expect("lookup")
        .expect("sheet")
        .set("A2", "same row")
        .expect("cell");
    let mut row = source.edit().expect("row edit");
    row.sheet("Sheet1")
        .expect("lookup")
        .expect("sheet")
        .row(1)
        .expect("row 2")
        .hide();
    cell.join(row).expect("orthogonal row and cell effects");
    let committed = cell.commit().expect("joined commit");
    let sheet = committed
        .workbook()
        .sheet(0usize)
        .expect("lookup")
        .expect("sheet");
    assert!(sheet.row(1).expect("row 2").hidden());
    assert!(matches!(
        sheet.cell("A2").expect("A2").stored(),
        Some(Cell::Value(Value::Text(text))) if text.as_str() == "same row"
    ));

    let mut left = source.edit().expect("left");
    left.sheet(0usize)
        .expect("lookup")
        .expect("sheet")
        .row(4)
        .expect("row 5")
        .hide();
    let mut right = source.edit().expect("right");
    right
        .sheet(0usize)
        .expect("lookup")
        .expect("sheet")
        .row(4)
        .expect("row 5")
        .show();
    let error = left.join(right).expect_err("same row must conflict");
    let conflicts = error.conflicts().expect("row conflicts");
    assert_eq!(conflicts.len(), 1);
    assert_eq!(conflicts.conflicts().len(), 1);
    assert_eq!(
        conflicts.conflicts()[0].rows().expect("row conflict"),
        &[RowIndex::new(4).expect("row 5")]
    );
    assert!(conflicts.conflicts()[0].cells().is_none());
}
