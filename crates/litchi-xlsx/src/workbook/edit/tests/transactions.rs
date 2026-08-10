//! Transaction joins, conflicts, and cross-facet edit tests.

use super::super::*;

use crate::cell::{Number, Value};
use crate::{History, HistoryLimits, Margins, PageMargin};

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
