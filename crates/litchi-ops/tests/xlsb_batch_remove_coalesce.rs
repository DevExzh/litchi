//! Atomic XLSB remove batches retain one native cell-values edit for a large
//! sequence of distinct direct cell paths.
#![cfg(feature = "xlsb")]

use std::io::Cursor;

use litchi_ops::{BatchItem, BatchMode, File, ItemResult};

fn fixture() -> Vec<u8> {
    use litchi_xlsb::builder::{self, Builder};

    let mut writer = Builder::new();
    let mut sheet = builder::Worksheet::new("Data");
    for column in 0..1_000_u32 {
        sheet.set_cell(0, column, (column + 1) as i32);
    }
    writer.add_worksheet(sheet);

    let mut output = Cursor::new(Vec::new());
    writer
        .build()
        .expect("build XLSB fixture")
        .write_to(&mut output)
        .expect("serialize XLSB fixture");
    output.into_inner()
}

fn column_name(mut column: usize) -> String {
    let mut name = String::new();
    loop {
        let remainder = (column % 26) as u8;
        name.push((b'A' + remainder) as char);
        column /= 26;
        if column == 0 {
            break;
        }
        column -= 1;
    }
    name.chars().rev().collect()
}

fn remove(path: String) -> BatchItem<'static> {
    let mut item = BatchItem::new("remove");
    item.path = Some(path.into());
    item
}

fn open(bytes: Vec<u8>) -> File {
    litchi_ops::open_bytes(bytes).expect("open XLSB fixture")
}

#[test]
fn atomic_batch_coalesces_thousand_direct_removals_and_reports_virtual_noop() {
    let file = open(fixture());
    let items = (0..1_000)
        .map(|column| remove(format!("/sheet[1]/{}1", column_name(column))))
        .collect::<Vec<_>>();
    let commit = file
        .batch(items, BatchMode::Atomic)
        .expect("remove direct cells in one atomic batch");

    assert_eq!(commit.diagnostics().operations(), 1_000);
    assert_eq!(commit.diagnostics().commits(), 1);
    assert_eq!(commit.value().len(), 1_000);
    assert!(
        commit
            .value()
            .iter()
            .all(|result| matches!(result, ItemResult::Applied(applied) if applied.changed))
    );

    let missing = open(fixture())
        .batch([remove("/sheet[1]/Z999".to_owned())], BatchMode::Atomic)
        .expect("removing a virtual cell is an exact no-op");
    assert!(!missing.diagnostics().changed());
    assert!(matches!(
        &missing.value()[0],
        ItemResult::Applied(applied) if !applied.changed
    ));
}
