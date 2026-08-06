use super::phonetic::{Alignment, Info, Type};
use super::worksheet::read;
use super::{Watch, Watches};
use crate::raw::{Kind, Records, Writer, kind};

fn stream(records: &[(Kind, &[u8])]) -> Vec<u8> {
    let mut output = Vec::new();
    let mut writer = Writer::new(&mut output);
    for (record_kind, payload) in records {
        writer.write_record(*record_kind, payload).unwrap();
    }
    output
}

fn watch_payload(row: u32, column: u32) -> [u8; 8] {
    let mut payload = [0; 8];
    payload[..4].copy_from_slice(&row.to_le_bytes());
    payload[4..].copy_from_slice(&column.to_le_bytes());
    payload
}

fn phonetic_payload(font: u16, phonetic_type: u32, alignment: u32) -> [u8; 10] {
    let mut payload = [0; 10];
    payload[..2].copy_from_slice(&font.to_le_bytes());
    payload[2..6].copy_from_slice(&phonetic_type.to_le_bytes());
    payload[6..].copy_from_slice(&alignment.to_le_bytes());
    payload
}

fn worksheet_with_owner() -> Vec<u8> {
    let first = watch_payload(4, 7);
    let second = watch_payload(12, 2);
    let phonetic = phonetic_payload(9, 2, 3);
    let unknown_kind = Kind::new(0x1234).unwrap();
    stream(&[
        (kind::BEGIN_SHEET, &[]),
        (kind::BEGIN_SHEET_DATA, &[]),
        (kind::END_SHEET_DATA, &[]),
        (kind::BEGIN_CELL_WATCHES, &[]),
        (kind::CELL_WATCH, &first),
        (unknown_kind, &[0x80, 0x01, 0xFE]),
        (kind::CELL_WATCH, &second),
        (kind::END_CELL_WATCHES, &[]),
        (kind::PHONETIC_INFO, &phonetic),
        (kind::END_SHEET, &[]),
    ])
}

fn worksheet_with_empty_collection() -> Vec<u8> {
    stream(&[
        (kind::BEGIN_SHEET, &[]),
        (kind::BEGIN_SHEET_DATA, &[]),
        (kind::END_SHEET_DATA, &[]),
        (kind::BEGIN_CELL_WATCHES, &[]),
        (kind::END_CELL_WATCHES, &[]),
        (kind::END_SHEET, &[]),
    ])
}

#[test]
fn parses_typed_watches_and_phonetic_defaults() {
    let data = worksheet_with_owner();
    let snapshot = read(&data).unwrap();
    assert_eq!(snapshot.watches().len(), 2);
    assert_eq!(snapshot.watches()[0].row(), 4);
    assert_eq!(snapshot.watches()[1].column(), 2);
    assert_eq!(
        snapshot.phonetic(),
        Some(Info::new(9, Type::Hiragana, Alignment::Distribute))
    );
    assert!(snapshot.has_collection());
    assert!(snapshot.has_opaque_records());
    let opaque = snapshot.opaque_records().unwrap();
    assert_eq!(opaque.len(), 1);
    assert_eq!(opaque[0].kind(), 0x1234);
    assert_eq!(opaque[0].payload(), &[0x80, 0x01, 0xFE]);
}

#[test]
fn edit_preserves_opaque_order_and_commits_atomically() {
    let data = worksheet_with_owner();
    let snapshot = read(&data).unwrap();
    let mut edit = snapshot.edit();
    assert!(edit.remove(crate::cell_watches::Reference::new(4, 7).unwrap()));
    edit.add(Watch::new(99, 4).unwrap()).unwrap();
    edit.set_phonetic(Info::new(2, Type::AsEntered, Alignment::Center));
    let commit = edit.commit().unwrap();
    assert!(!commit.patch().is_empty());

    let output = commit.patch().apply(&data).unwrap();
    let updated = read(&output).unwrap();
    assert_eq!(
        updated.watches(),
        &[Watch::new(12, 2).unwrap(), Watch::new(99, 4).unwrap()]
    );
    assert_eq!(
        updated.phonetic(),
        Some(Info::new(2, Type::AsEntered, Alignment::Center))
    );
    assert_eq!(
        updated.opaque_records().unwrap()[0].payload(),
        &[0x80, 0x01, 0xFE]
    );

    let kinds: Vec<_> = Records::new(&output)
        .map(|record| record.unwrap().kind().get())
        .collect();
    let unknown = 0x1234;
    let first_watch = kind::CELL_WATCH.get();
    let last_watch = kinds
        .iter()
        .rposition(|value| *value == first_watch)
        .unwrap();
    let unknown_position = kinds.iter().position(|value| *value == unknown).unwrap();
    assert!(
        unknown_position
            < kinds
                .iter()
                .position(|value| *value == first_watch)
                .unwrap()
    );
    assert!(unknown_position < last_watch);
}

#[test]
fn clear_keeps_opaque_collection_and_inverse_restores_bytes() {
    let data = worksheet_with_owner();
    let snapshot = read(&data).unwrap();
    let mut edit = snapshot.edit();
    assert!(edit.clear());
    let commit = edit.commit().unwrap();
    let output = commit.patch().after().to_vec();
    let updated = read(&output).unwrap();
    assert!(updated.watches().is_empty());
    assert!(updated.has_collection());
    assert!(updated.has_opaque_records());
    assert_eq!(commit.patch().inverse().apply(&output).unwrap(), data);
}

#[test]
fn empty_collection_survives_noop_and_unrelated_phonetic_edits() {
    let data = worksheet_with_empty_collection();
    let snapshot = read(&data).unwrap();
    assert!(snapshot.has_collection());
    assert_eq!(snapshot.edit().commit().unwrap().patch().after(), data);

    let mut edit = snapshot.edit();
    edit.set_phonetic(Info::new(3, Type::AsEntered, Alignment::AllTextLeft));
    let output = edit.commit().unwrap().patch().after().to_vec();
    let updated = read(&output).unwrap();
    assert!(updated.has_collection());
    assert_eq!(
        updated.phonetic(),
        Some(Info::new(3, Type::AsEntered, Alignment::AllTextLeft))
    );
}

#[test]
fn rejects_invalid_coordinates_domains_and_duplicate_watches() {
    assert!(Watch::new(1_048_576, 0).is_err());
    assert!(Watch::new(0, 16_384).is_err());
    assert!(Watches::new(vec![Watch::new(1, 1).unwrap(), Watch::new(1, 1).unwrap(),]).is_err());

    let invalid_type = phonetic_payload(0, 4, 0);
    let data = stream(&[
        (kind::BEGIN_SHEET, &[]),
        (kind::BEGIN_SHEET_DATA, &[]),
        (kind::END_SHEET_DATA, &[]),
        (kind::PHONETIC_INFO, &invalid_type),
        (kind::END_SHEET, &[]),
    ]);
    assert!(read(&data).is_err());
}

#[test]
fn workbook_facade_publishes_only_a_valid_commit() {
    let package = crate::Package::create().unwrap();
    let mut workbook = package.into_workbook().unwrap();
    let mut edit = workbook.edit_cell_watches(0).unwrap();
    edit.add(Watch::new(3, 5).unwrap()).unwrap();
    let commit = edit.commit().unwrap();
    let snapshot = workbook.apply_cell_watches(0, &commit).unwrap();
    assert_eq!(snapshot.watches(), &[Watch::new(3, 5).unwrap()]);
    assert_eq!(
        workbook.cell_watches(0).unwrap().watches(),
        snapshot.watches()
    );

    let mut stale = workbook.edit_cell_watches(0).unwrap();
    stale.add(Watch::new(4, 4).unwrap()).unwrap();
    let stale_commit = stale.commit().unwrap();
    let mut other = workbook.edit_cell_watches(0).unwrap();
    other.add(Watch::new(5, 5).unwrap()).unwrap();
    let other_commit = other.commit().unwrap();
    workbook.apply_cell_watches(0, &other_commit).unwrap();
    assert!(workbook.apply_cell_watches(0, &stale_commit).is_err());
    assert_eq!(
        workbook.cell_watches(0).unwrap().watches(),
        &[Watch::new(3, 5).unwrap(), Watch::new(5, 5).unwrap()]
    );
}
