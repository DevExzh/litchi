//! Narrow cross-workbook transfer coverage for native BIFF8 scalar cells.

#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "fixture construction and assertions intentionally fail fast"
)]

use litchi_cfb::{OleFile, OleWriter};
use litchi_xls::Writer;
use litchi_xls::cell_values::{
    CellRange, Reference, Selector, Snapshot, Storage, StyleIndex, Value,
};
use std::io::Cursor;

fn workbook_bytes(numbers: &[(u32, u16, f64)], blanks: &[(u32, u16)]) -> Vec<u8> {
    let mut snapshot =
        Snapshot::from_bytes(neutralize_refresh_all(raw_workbook_bytes(numbers, blanks))).unwrap();
    if !blanks.is_empty() {
        let style = snapshot
            .worksheets()
            .flat_map(|sheet| sheet.cells())
            .next()
            .map(|cell| cell.style())
            .unwrap_or_else(|| StyleIndex::new(&snapshot, 0).unwrap());
        let mut edit = snapshot.edit();
        for &(row, column) in blanks {
            edit.insert_cell_with_style(
                Selector::Position(0),
                Reference::new(row, u32::from(column)).unwrap(),
                Value::Blank,
                style,
            )
            .unwrap();
        }
        snapshot = edit.commit().unwrap().snapshot().clone();
    }
    snapshot.bytes().to_vec()
}

fn raw_workbook_bytes(numbers: &[(u32, u16, f64)], _blanks: &[(u32, u16)]) -> Vec<u8> {
    let mut writer = Writer::new();
    let sheet = writer.add_worksheet("Sheet1").unwrap();
    for &(row, column, value) in numbers {
        writer.write_number(sheet, row, column, value).unwrap();
    }
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    output.into_inner()
}

fn rewrite_workbook_stream(bytes: Vec<u8>, edit: impl FnOnce(&mut Vec<u8>)) -> Vec<u8> {
    let workbook_path = vec!["Workbook".to_string()];
    let mut editor = litchi_ole_common::object::Editor::open(
        bytes,
        litchi_ole_common::object::Targets::default(),
        litchi_ole_common::object::Limits::default(),
    )
    .unwrap();
    let mut workbook = editor.stream(&workbook_path).unwrap().to_vec();
    edit(&mut workbook);
    editor.put_stream(&workbook_path, workbook).unwrap();
    editor.finish().unwrap()
}

fn neutralize_refresh_all(bytes: Vec<u8>) -> Vec<u8> {
    rewrite_workbook_stream(bytes, |workbook| {
        remove_global_record(workbook, 0x01B7);
    })
}

fn remove_global_record(workbook: &mut Vec<u8>, record_kind: u16) {
    let mut offset = 0_usize;
    let mut removed_len = None;
    while let Some((kind, end)) = record_bounds(workbook, offset) {
        if kind == record_kind {
            let length = end - offset;
            workbook.drain(offset..end);
            removed_len = Some(length);
            break;
        }
        offset = end;
        if kind == 0x000A {
            break;
        }
    }
    let Some(removed_len) = removed_len else {
        return;
    };
    let mut offset = 0_usize;
    while let Some((kind, end)) = record_bounds(workbook, offset) {
        if kind == 0x0085 && end - offset >= 8 {
            let payload = offset + 4;
            let old_position = u32::from_le_bytes([
                workbook[payload],
                workbook[payload + 1],
                workbook[payload + 2],
                workbook[payload + 3],
            ]);
            let new_position = old_position
                .checked_sub(u32::try_from(removed_len).unwrap())
                .unwrap();
            workbook[payload..payload + 4].copy_from_slice(&new_position.to_le_bytes());
        }
        offset = end;
        if kind == 0x000A {
            break;
        }
    }
}

fn replace_first_record_kind(workbook: &mut [u8], from: u16, to: u16) {
    let mut offset = 0_usize;
    while let Some((kind, end)) = record_bounds(workbook, offset) {
        if kind == from {
            workbook[offset..offset + 2].copy_from_slice(&to.to_le_bytes());
            return;
        }
        offset = end;
    }
}

fn replace_first_record(workbook: &mut [u8], from: u16, to: u16, payload: &[u8]) {
    let mut offset = 0_usize;
    while let Some((kind, end)) = record_bounds(workbook, offset) {
        if kind == from {
            assert_eq!(end - offset - 4, payload.len());
            workbook[offset..offset + 2].copy_from_slice(&to.to_le_bytes());
            workbook[offset + 4..end].copy_from_slice(payload);
            return;
        }
        offset = end;
    }
}

fn duplicate_first_number_owner(workbook: &mut [u8]) {
    let mut first_coordinates: Option<(u16, u16)> = None;
    let mut offset = 0_usize;
    while let Some((kind, end)) = record_bounds(workbook, offset) {
        if kind == 0x0203 && end - offset >= 8 {
            let payload = offset + 4;
            if let Some((row, column)) = first_coordinates {
                workbook[payload..payload + 2].copy_from_slice(&row.to_le_bytes());
                workbook[payload + 2..payload + 4].copy_from_slice(&column.to_le_bytes());
                return;
            }
            first_coordinates = Some((
                u16::from_le_bytes([workbook[payload], workbook[payload + 1]]),
                u16::from_le_bytes([workbook[payload + 2], workbook[payload + 3]]),
            ));
        }
        offset = end;
    }
}

fn record_bounds(bytes: &[u8], offset: usize) -> Option<(u16, usize)> {
    let header = bytes.get(offset..offset.checked_add(4)?)?;
    let kind = u16::from_le_bytes([header[0], header[1]]);
    let length = usize::from(u16::from_le_bytes([header[2], header[3]]));
    let end = offset.checked_add(4)?.checked_add(length)?;
    bytes.get(offset..end)?;
    Some((kind, end))
}

fn append_worksheet_records(workbook: &mut Vec<u8>, records: &[(u16, &[u8])]) {
    insert_worksheet_records_before(workbook, 0x000A, records);
}

fn insert_worksheet_records_before(
    workbook: &mut Vec<u8>,
    before_kind: u16,
    records: &[(u16, &[u8])],
) {
    let mut worksheet_start = None;
    let mut offset = 0_usize;
    while let Some((kind, end)) = record_bounds(workbook, offset) {
        if kind == 0x0085 && end - offset >= 8 {
            let payload = offset + 4;
            worksheet_start = Some(u32::from_le_bytes([
                workbook[payload],
                workbook[payload + 1],
                workbook[payload + 2],
                workbook[payload + 3],
            ]) as usize);
            break;
        }
        offset = end;
        if kind == 0x000A {
            break;
        }
    }
    let mut offset = worksheet_start.unwrap();
    while let Some((kind, end)) = record_bounds(workbook, offset) {
        if kind == before_kind {
            let mut appended = Vec::new();
            for &(record_kind, payload) in records {
                appended.extend_from_slice(&record_kind.to_le_bytes());
                appended.extend_from_slice(&(u16::try_from(payload.len()).unwrap()).to_le_bytes());
                appended.extend_from_slice(payload);
            }
            let tail = workbook.split_off(offset);
            workbook.extend_from_slice(&appended);
            workbook.extend_from_slice(&tail);
            return;
        }
        offset = end;
    }
}

fn add_opaque_stream(bytes: Vec<u8>) -> Vec<u8> {
    let mut editor = litchi_ole_common::object::Editor::open(
        bytes,
        litchi_ole_common::object::Targets::default(),
        litchi_ole_common::object::Limits::default(),
    )
    .unwrap();
    editor
        .add_stream(vec!["Opaque".to_string()], b"do-not-touch".to_vec())
        .unwrap();
    editor.finish().unwrap()
}

fn scalar_source() -> Snapshot {
    Snapshot::from_bytes(add_opaque_stream(workbook_bytes(
        &[(0, 0, 1.25), (2, 2, 9.5)],
        &[(1, 1)],
    )))
    .unwrap()
}

#[test]
fn copies_numbers_and_blanks_in_source_order_and_preserves_foreign_streams() {
    let donor = scalar_source();
    let donor_before = donor.bytes().to_vec();
    let target =
        Snapshot::from_bytes(add_opaque_stream(workbook_bytes(&[(10, 10, 42.0)], &[]))).unwrap();
    let target_opaque = OleFile::open(Cursor::new(target.bytes()))
        .unwrap()
        .open_stream(&["Opaque"])
        .unwrap()
        .to_vec();
    let range =
        CellRange::new(Reference::new(0, 0).unwrap(), Reference::new(2, 2).unwrap()).unwrap();

    let mut edit = target.edit();
    edit.copy_scalar_cells_from(
        &donor,
        Selector::Name("Sheet1"),
        range,
        Selector::Position(0),
        Reference::new(5, 4).unwrap(),
    )
    .unwrap();
    let commit = edit.commit().unwrap();
    let copied = commit.snapshot();
    let sheet = copied.worksheet(Selector::Position(0)).unwrap().unwrap();
    assert_eq!(
        sheet
            .cell(Reference::new(5, 4).unwrap())
            .unwrap()
            .unwrap()
            .value(),
        &Value::Number(1.25)
    );
    assert_eq!(
        sheet
            .cell(Reference::new(6, 5).unwrap())
            .unwrap()
            .unwrap()
            .storage(),
        Storage::Blank
    );
    assert_eq!(
        sheet
            .cell(Reference::new(7, 6).unwrap())
            .unwrap()
            .unwrap()
            .value(),
        &Value::Number(9.5)
    );
    assert_eq!(
        OleFile::open(Cursor::new(copied.bytes()))
            .unwrap()
            .open_stream(&["Opaque"])
            .unwrap(),
        target_opaque.as_slice()
    );
    assert_eq!(donor.bytes(), donor_before.as_slice());
    assert_eq!(
        commit.patch().inverse().apply(copied).unwrap().bytes(),
        target.bytes()
    );
}

#[test]
fn same_width_existing_targets_are_supported_but_occupied_mismatches_are_atomic() {
    let donor = scalar_source();
    let target = Snapshot::from_bytes(workbook_bytes(&[(5, 4, 77.0)], &[])).unwrap();
    let mut edit = target.edit();
    edit.copy_scalar_cell_from(
        &donor,
        Selector::Position(0),
        Reference::new(0, 0).unwrap(),
        Selector::Position(0),
        Reference::new(5, 4).unwrap(),
    )
    .unwrap();
    let commit = edit.commit().unwrap();
    assert_eq!(
        commit
            .snapshot()
            .worksheet(Selector::Position(0))
            .unwrap()
            .unwrap()
            .number(Reference::new(5, 4).unwrap())
            .unwrap()
            .unwrap()
            .value(),
        1.25
    );

    let occupied = Snapshot::from_bytes(workbook_bytes(&[(5, 4, 77.0)], &[])).unwrap();
    let mut refused = occupied.edit();
    let error = refused
        .copy_scalar_cell_from(
            &donor,
            Selector::Position(0),
            Reference::new(1, 1).unwrap(),
            Selector::Position(0),
            Reference::new(5, 4).unwrap(),
        )
        .unwrap_err();
    assert!(error.to_string().contains("record width or family"));
    assert_eq!(
        refused.commit().unwrap().snapshot().bytes(),
        occupied.bytes()
    );
}

#[test]
fn strings_formulas_and_unrelated_owners_are_refused_before_staging() {
    let mut string_writer = Writer::new();
    let string_sheet = string_writer.add_worksheet("Sheet1").unwrap();
    string_writer.write_number(string_sheet, 0, 0, 1.0).unwrap();
    string_writer
        .write_string(string_sheet, 0, 1, "not scalar")
        .unwrap();
    let mut string_output = Cursor::new(Vec::new());
    string_writer.write_to(&mut string_output).unwrap();
    let string_donor = Snapshot::from_bytes(string_output.into_inner()).unwrap();
    let target = Snapshot::from_bytes(workbook_bytes(&[], &[])).unwrap();
    let range =
        CellRange::new(Reference::new(0, 0).unwrap(), Reference::new(0, 1).unwrap()).unwrap();

    let mut edit = target.edit();
    assert!(
        edit.copy_scalar_cells_from(
            &string_donor,
            Selector::Position(0),
            range,
            Selector::Position(0),
            Reference::new(2, 2).unwrap(),
        )
        .is_err()
    );
    assert_eq!(edit.commit().unwrap().snapshot().bytes(), target.bytes());

    let mut formula_writer = Writer::new();
    let formula_sheet = formula_writer.add_worksheet("Sheet1").unwrap();
    formula_writer
        .write_formula(formula_sheet, 0, 0, "1+1")
        .unwrap();
    let mut formula_output = Cursor::new(Vec::new());
    formula_writer.write_to(&mut formula_output).unwrap();
    let formula_donor = Snapshot::from_bytes(formula_output.into_inner()).unwrap();
    let mut formula_edit = target.edit();
    assert!(
        formula_edit
            .copy_scalar_cell_from(
                &formula_donor,
                Selector::Position(0),
                Reference::new(0, 0).unwrap(),
                Selector::Position(0),
                Reference::new(2, 2).unwrap(),
            )
            .is_err()
    );
}

fn assert_scalar_transfer_refused(target: &Snapshot, donor: &Snapshot) {
    let before = target.bytes().to_vec();
    let mut edit = target.edit();
    assert!(
        edit.copy_scalar_cell_from(
            donor,
            Selector::Position(0),
            Reference::new(0, 0).unwrap(),
            Selector::Position(0),
            Reference::new(4, 4).unwrap(),
        )
        .is_err()
    );
    assert_eq!(edit.commit().unwrap().snapshot().bytes(), before.as_slice());
}

#[test]
fn hostile_duplicate_unknown_and_dependency_records_are_refused_atomically() {
    let target = Snapshot::from_bytes(workbook_bytes(&[(100, 100, 42.0)], &[])).unwrap();

    let duplicate = rewrite_workbook_stream(
        neutralize_refresh_all(raw_workbook_bytes(&[(0, 0, 1.25), (2, 2, 9.5)], &[])),
        |workbook| duplicate_first_number_owner(workbook),
    );
    let duplicate = Snapshot::from_bytes(duplicate).unwrap();
    assert_scalar_transfer_refused(&target, &duplicate);

    let unknown_worksheet = rewrite_workbook_stream(
        neutralize_refresh_all(raw_workbook_bytes(&[(0, 0, 1.25)], &[])),
        |workbook| replace_first_record_kind(workbook, 0x0055, 0x7FFE),
    );
    let unknown_worksheet = Snapshot::from_bytes(unknown_worksheet).unwrap();
    assert_scalar_transfer_refused(&target, &unknown_worksheet);

    let unknown_global = rewrite_workbook_stream(
        neutralize_refresh_all(raw_workbook_bytes(&[(0, 0, 1.25)], &[])),
        |workbook| replace_first_record_kind(workbook, 0x008C, 0x7FFD),
    );
    let unknown_global = Snapshot::from_bytes(unknown_global).unwrap();
    assert_scalar_transfer_refused(&target, &unknown_global);

    let supbook = rewrite_workbook_stream(
        neutralize_refresh_all(raw_workbook_bytes(&[(0, 0, 1.25)], &[])),
        |workbook| replace_first_record(workbook, 0x008C, 0x01AE, &[1, 0, 1, 4]),
    );
    let supbook = Snapshot::from_bytes(supbook).unwrap();
    assert_scalar_transfer_refused(&target, &supbook);

    let autofilter = rewrite_workbook_stream(
        neutralize_refresh_all(raw_workbook_bytes(&[(0, 0, 1.25)], &[])),
        |workbook| {
            let info = [1_u8, 0];
            let filter = [0_u8; 24];
            append_worksheet_records(workbook, &[(0x009D, &info), (0x009E, &filter)]);
        },
    );
    let autofilter = Snapshot::from_bytes(autofilter).unwrap();
    assert_scalar_transfer_refused(&target, &autofilter);

    let dval = rewrite_workbook_stream(
        neutralize_refresh_all(raw_workbook_bytes(&[(0, 0, 1.25)], &[])),
        |workbook| {
            let mut payload = [0_u8; 18];
            payload[10..14].copy_from_slice(&u32::MAX.to_le_bytes());
            append_worksheet_records(workbook, &[(0x01B2, &payload)]);
        },
    );
    let dval = Snapshot::from_bytes(dval).unwrap();
    assert_scalar_transfer_refused(&target, &dval);

    let custom_view = rewrite_workbook_stream(
        neutralize_refresh_all(raw_workbook_bytes(&[(0, 0, 1.25)], &[])),
        |workbook| {
            let mut begin = [0_u8; 64];
            begin[20..24].copy_from_slice(&100_u32.to_le_bytes());
            let end = [1_u8, 0];
            append_worksheet_records(workbook, &[(0x01AA, &begin), (0x01AB, &end)]);
        },
    );
    let custom_view = Snapshot::from_bytes(custom_view).unwrap();
    assert_scalar_transfer_refused(&target, &custom_view);

    let scenman = rewrite_workbook_stream(
        neutralize_refresh_all(raw_workbook_bytes(&[(0, 0, 1.25)], &[])),
        |workbook| {
            let payload = [0_u8, 0, 0xFF, 0xFF, 0xFF, 0xFF, 0, 0];
            insert_worksheet_records_before(workbook, 0x0200, &[(0x00AE, &payload)]);
        },
    );
    let scenman = Snapshot::from_bytes(scenman).unwrap();
    assert_scalar_transfer_refused(&target, &scenman);

    let duplicate_target = rewrite_workbook_stream(
        neutralize_refresh_all(raw_workbook_bytes(&[(4, 4, 77.0), (5, 5, 88.0)], &[])),
        |workbook| duplicate_first_number_owner(workbook),
    );
    let duplicate_target = Snapshot::from_bytes(duplicate_target).unwrap();
    let duplicate_target_before = duplicate_target.bytes().to_vec();
    let mut duplicate_target_edit = duplicate_target.edit();
    let range =
        CellRange::new(Reference::new(0, 0).unwrap(), Reference::new(2, 2).unwrap()).unwrap();
    assert!(
        duplicate_target_edit
            .copy_scalar_cells_from(
                &scalar_source(),
                Selector::Position(0),
                range,
                Selector::Position(0),
                Reference::new(4, 4).unwrap(),
            )
            .is_err()
    );
    assert_eq!(
        duplicate_target_edit.commit().unwrap().snapshot().bytes(),
        duplicate_target_before.as_slice()
    );

    let refresh_all = Snapshot::from_bytes(raw_workbook_bytes(&[(0, 0, 1.25)], &[])).unwrap();
    assert_scalar_transfer_refused(&target, &refresh_all);
}

#[test]
fn bounds_noop_and_stale_or_foreign_artifacts_are_checked() {
    let donor = scalar_source();
    let target = Snapshot::from_bytes(workbook_bytes(&[(100, 100, 42.0)], &[])).unwrap();
    let too_large = CellRange::new(
        Reference::new(0, 0).unwrap(),
        Reference::new(63, 64).unwrap(),
    )
    .unwrap();
    let mut over = target.edit();
    assert!(
        over.copy_scalar_cells_from(
            &donor,
            Selector::Position(0),
            too_large,
            Selector::Position(0),
            Reference::new(0, 0).unwrap(),
        )
        .is_err()
    );
    assert_eq!(over.commit().unwrap().snapshot().bytes(), target.bytes());

    let empty_range = CellRange::new(
        Reference::new(20, 20).unwrap(),
        Reference::new(20, 20).unwrap(),
    )
    .unwrap();
    let mut noop = target.edit();
    noop.copy_scalar_cells_from(
        &donor,
        Selector::Position(0),
        empty_range,
        Selector::Position(0),
        Reference::new(20, 20).unwrap(),
    )
    .unwrap();
    let noop_commit = noop.commit().unwrap();
    assert!(noop_commit.patch().is_empty());

    let mut copy = target.edit();
    copy.copy_scalar_cell_from(
        &donor,
        Selector::Position(0),
        Reference::new(0, 0).unwrap(),
        Selector::Position(0),
        Reference::new(4, 4).unwrap(),
    )
    .unwrap();
    let commit = copy.commit().unwrap();
    assert!(commit.patch().apply(&donor).is_err());
    assert!(commit.patch().apply(commit.snapshot()).is_err());
    assert_eq!(
        commit.patch().apply(&target).unwrap().bytes(),
        commit.snapshot().bytes()
    );
}

#[test]
fn signed_cfb_is_refused_before_cross_workbook_transfer() {
    let donor = scalar_source();
    let mut ole = OleFile::open(Cursor::new(donor.bytes())).unwrap();
    let workbook = ole.open_stream(&["Workbook"]).unwrap();
    let mut writer = OleWriter::new();
    writer.create_stream(&["Workbook"], &workbook).unwrap();
    writer
        .create_stream(&["DigitalSignature"], b"signature")
        .unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    assert!(Snapshot::from_bytes(output.into_inner()).is_err());
}
