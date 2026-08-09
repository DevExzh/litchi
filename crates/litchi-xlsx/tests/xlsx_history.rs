#![allow(
    clippy::unwrap_used,
    reason = "focused integration tests use panic-on-failure assertions"
)]

use litchi_xlsx::cell::{Cell, Value};
use litchi_xlsx::{History, HistoryLimits, Workbook};

#[test]
fn budgeted_history_tracks_immutable_workbook_snapshots() {
    let original = Workbook::new().unwrap();
    let mut edit = original.edit().unwrap();
    edit.sheet("Sheet1")
        .unwrap()
        .unwrap()
        .set("A1", "changed")
        .unwrap();
    let commit = edit.commit().unwrap();

    let mut history = History::new(original, HistoryLimits::new(4, 1024));
    assert!(
        history
            .record(commit.workbook().clone(), 1)
            .unwrap()
            .is_empty()
    );
    let current = history.current().sheet("Sheet1").unwrap().unwrap();
    assert!(matches!(
        current.cell("A1").unwrap().stored(),
        Some(Cell::Value(Value::Text(text))) if text.as_str() == "changed"
    ));
    assert!(history.undo());
    assert!(
        history
            .current()
            .sheet("Sheet1")
            .unwrap()
            .unwrap()
            .cell("A1")
            .unwrap()
            .stored()
            .is_none()
    );
    assert!(history.redo());
}
