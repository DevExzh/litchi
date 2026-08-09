#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "integration tests use panic-on-failure assertions"
)]

use litchi_core::sheet::traits::WorkbookTrait;
use litchi_xlsb::Workbook;
use litchi_xlsb::cell_values::{Reference, TransferLimits, Value, WorkbookHistory, WorkbookPatch};
use litchi_xlsb::writer::{MutableWorksheet, WorkbookWriter};
use std::fs::File;
use std::io::Cursor;
use std::path::PathBuf;

fn fixture(relative: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../..")
        .join(relative)
}

fn producer_workbook(name: &str, initial_value: Option<&str>) -> Workbook {
    let mut sheet = MutableWorksheet::new(name);
    if let Some(value) = initial_value {
        sheet.set_cell(0, 0, value);
    }
    let mut producer = WorkbookWriter::new();
    producer.add_worksheet(sheet);
    let mut bytes = Cursor::new(Vec::new());
    producer.save(&mut bytes).expect("producer save");
    Workbook::new(Cursor::new(bytes.into_inner())).expect("producer reopen")
}

#[test]
fn producer_reopen_transfers_sst_and_renames_on_a_durable_root() {
    let source = producer_workbook("Source", Some("shared transfer value"));
    let mut target = producer_workbook("Target", None);
    let destination = Reference::new(4, 3).expect("destination");

    let mut edit = target
        .edit_workbook_structure()
        .expect("detached workbook edit");
    edit.rename_sheet(0, "Imported".to_string())
        .expect("rename");
    edit.transfer_cell(
        &source,
        0,
        Reference::new(0, 0).expect("source reference"),
        0,
        destination,
    )
    .expect("dependency-managed transfer");
    let commit = edit.commit().expect("workbook commit");
    let encoded = commit
        .patch()
        .to_bytes(TransferLimits::DEFAULT)
        .expect("durable patch");
    let decoded = WorkbookPatch::from_bytes(&encoded, TransferLimits::DEFAULT)
        .expect("validated durable patch");
    decoded.apply(&mut target).expect("atomic publication");

    assert_eq!(target.worksheet_names(), &["Imported".to_string()]);
    let snapshot = target.cell_values(0).expect("target cells");
    let cell = snapshot
        .cell(destination)
        .expect("unique destination")
        .expect("transferred destination");
    let Value::SharedStringIndex(index) = cell.value() else {
        panic!("writer-produced string should remain SST-backed")
    };
    assert_eq!(
        target.shared_strings()[usize::try_from(*index).expect("SST index")].text,
        "shared transfer value"
    );

    let mut bytes = Cursor::new(Vec::new());
    target.save(&mut bytes).expect("save transferred workbook");
    let reopened = Workbook::new(Cursor::new(bytes.into_inner())).expect("consumer reopen");
    assert_eq!(reopened.worksheet_names(), &["Imported".to_string()]);
    assert!(
        reopened
            .cell_values(0)
            .expect("reopened cells")
            .cell(destination)
            .expect("unique destination")
            .is_some()
    );
}

#[test]
fn third_party_fixture_cell_transfers_with_its_style_dependency_closure() {
    let fixtures = [
        "test-data/ooxml/xlsb/sample.xlsb",
        "test-data/ooxml/xlsb/Simple.xlsb",
        "test-data/ooxml/xlsb/date.xlsb",
        "test-data/ooxml/xlsb/51519.xlsb",
    ];
    let (source, sheet, source_reference) = fixtures
        .iter()
        .find_map(|relative| {
            let workbook = Workbook::new(File::open(fixture(relative)).ok()?).ok()?;
            let selected = (0..workbook.worksheet_count()).find_map(|sheet| {
                workbook.cell_values(sheet).ok().and_then(|snapshot| {
                    snapshot
                        .cells()
                        .find(|cell| cell.formula().is_none())
                        .map(|cell| (sheet, cell.reference()))
                })
            });
            selected.map(|(sheet, reference)| (workbook, sheet, reference))
        })
        .expect("a third-party fixture contains a direct stored cell");
    let destination = Reference::new(10, 7).expect("destination");
    let mut target = producer_workbook("Target", None);
    let mut edit = target.edit_workbook_structure().expect("root edit");
    edit.transfer_cell(&source, sheet, source_reference, 0, destination)
        .expect("third-party transfer plan");
    let commit = edit.commit().expect("third-party transfer commit");
    target
        .apply_workbook_structure(&commit)
        .expect("publish third-party transfer");

    let transferred_snapshot = target.cell_values(0).expect("transferred snapshot");
    let transferred = transferred_snapshot
        .cell(destination)
        .expect("unique destination")
        .expect("transferred cell");
    assert!(
        target
            .styles()
            .get_cell_format(usize::try_from(transferred.style().get()).expect("style index"))
            .is_some()
    );
    let mut bytes = Cursor::new(Vec::new());
    target.save(&mut bytes).expect("save transferred fixture");
    let reopened = Workbook::new(Cursor::new(bytes.into_inner())).expect("reopen transfer");
    assert!(
        reopened
            .cell_values(0)
            .expect("reopened snapshot")
            .cell(destination)
            .expect("unique destination")
            .is_some()
    );
}

#[test]
fn three_way_history_and_adversarial_envelope_remain_atomic() {
    let source = producer_workbook("Source", Some("merge value"));
    let mut target = producer_workbook("Target", None);
    let destination = Reference::new(8, 2).expect("destination");

    let mut left_edit = target.edit_workbook_structure().expect("left plan");
    left_edit
        .rename_sheet(0, "Merged".to_string())
        .expect("rename");
    let left_commit = left_edit.commit().expect("left commit");

    let mut right_edit = target.edit_workbook_structure().expect("right plan");
    right_edit
        .transfer_cell(
            &source,
            0,
            Reference::new(0, 0).expect("source reference"),
            0,
            destination,
        )
        .expect("transfer");
    let right_commit = right_edit.commit().expect("right commit");
    let outcome = left_commit
        .patch()
        .merge_three_way(right_commit.patch())
        .expect("three-way merge");
    assert!(outcome.conflicts().is_empty());
    let merged = outcome.patch().expect("merged patch").clone();

    let mut corrupt = merged
        .to_bytes(TransferLimits::DEFAULT)
        .expect("encoded merge");
    let last = corrupt.len().checked_sub(1).expect("nonempty patch");
    corrupt[last] ^= 0x5a;
    assert!(WorkbookPatch::from_bytes(&corrupt, TransferLimits::DEFAULT).is_err());
    assert_eq!(target.worksheet_names(), &["Target".to_string()]);

    merged.apply(&mut target).expect("apply merge");
    assert_eq!(target.worksheet_names(), &["Merged".to_string()]);
    assert!(
        merged.apply(&mut target).is_err(),
        "stale reapply must fail"
    );

    let mut history = WorkbookHistory::new(TransferLimits::DEFAULT).expect("history");
    history.push(merged).expect("retain patch");
    history.undo(&mut target).expect("undo");
    assert_eq!(target.worksheet_names(), &["Target".to_string()]);
    assert!(
        target
            .cell_values(0)
            .expect("undone cells")
            .cell(destination)
            .expect("unique destination")
            .is_none()
    );
    history.redo(&mut target).expect("redo");
    assert_eq!(target.worksheet_names(), &["Merged".to_string()]);
}
