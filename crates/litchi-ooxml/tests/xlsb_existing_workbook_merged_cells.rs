use litchi_ooxml::xlsb::merged_cells::MergedCell;
use litchi_ooxml::xlsb::writer::{MutableXlsbWorksheet, RecordWriter, XlsbWorkbookWriter};
use litchi_ooxml::xlsb::{XlsbRecord, XlsbWorkbook};
use litchi_opc::{OpcPackage, PackURI};
use std::io::Cursor;

const BEGIN_MERGE_CELLS: u16 = 0x00B1;
const MERGE_CELL: u16 = 0x00B0;
const END_MERGE_CELLS: u16 = 0x00B2;
const END_SHEET: u16 = 0x0082;

fn workbook_bytes(sheets: &[(&str, &[MergedCell])]) -> Vec<u8> {
    let mut workbook = XlsbWorkbookWriter::new();
    for (name, ranges) in sheets {
        let mut sheet = MutableXlsbWorksheet::new(*name);
        sheet.set_cell(0, 0, format!("{name} preserved"));
        for range in *ranges {
            sheet.add_merged_cell(range.clone());
        }
        workbook.add_worksheet(sheet);
    }
    let mut output = Cursor::new(Vec::new());
    workbook.save(&mut output).unwrap();
    output.into_inner()
}

fn save(workbook: &XlsbWorkbook) -> Vec<u8> {
    let mut output = Cursor::new(Vec::new());
    workbook.save(&mut output).unwrap();
    output.into_inner()
}

fn part_blob(package_bytes: &[u8], path: &str) -> Vec<u8> {
    let package = OpcPackage::from_reader(Cursor::new(package_bytes)).unwrap();
    package
        .get_part(&PackURI::new(path).unwrap())
        .unwrap()
        .blob()
        .to_vec()
}

fn records(data: &[u8]) -> Vec<(u16, usize, usize, Vec<u8>)> {
    let mut cursor = Cursor::new(data);
    let mut result = Vec::new();
    while cursor.position() < data.len() as u64 {
        let start = cursor.position() as usize;
        let record = XlsbRecord::read(&mut cursor).unwrap();
        result.push((
            record.header.record_type,
            start,
            cursor.position() as usize,
            record.data.to_vec(),
        ));
    }
    result
}

fn insert_before_end_sheet(data: &[u8], inserted: &[u8]) -> Vec<u8> {
    let offset = records(data)
        .into_iter()
        .find_map(|(record_type, start, _, _)| (record_type == END_SHEET).then_some(start))
        .unwrap();
    let mut output = Vec::with_capacity(data.len() + inserted.len());
    output.extend_from_slice(&data[..offset]);
    output.extend_from_slice(inserted);
    output.extend_from_slice(&data[offset..]);
    output
}

fn rewrite_first_sheet(package_bytes: &[u8], transform: impl FnOnce(&[u8]) -> Vec<u8>) -> Vec<u8> {
    let mut package = OpcPackage::from_reader(Cursor::new(package_bytes)).unwrap();
    let uri = PackURI::new("/xl/worksheets/sheet1.bin").unwrap();
    let changed = transform(package.get_part(&uri).unwrap().blob());
    package.get_part_mut(&uri).unwrap().set_blob(changed);
    let mut output = Cursor::new(Vec::new());
    package.to_stream(&mut output).unwrap();
    output.into_inner()
}

fn merge_block(declared_count: u32, ranges: &[MergedCell]) -> Vec<u8> {
    let mut output = Vec::new();
    let mut writer = RecordWriter::new(&mut output);
    writer
        .write_record(BEGIN_MERGE_CELLS, &declared_count.to_le_bytes())
        .unwrap();
    for range in ranges {
        writer
            .write_record(MERGE_CELL, &range.serialize())
            .unwrap();
    }
    writer.write_record(END_MERGE_CELLS, &[]).unwrap();
    output
}

fn strip_merge_block(data: &[u8]) -> Vec<u8> {
    let spans = records(data);
    let begin = spans
        .iter()
        .find_map(|(kind, start, _, _)| (*kind == BEGIN_MERGE_CELLS).then_some(*start));
    let end = spans
        .iter()
        .find_map(|(kind, _, end, _)| (*kind == END_MERGE_CELLS).then_some(*end));
    match (begin, end) {
        (Some(begin), Some(end)) => [&data[..begin], &data[end..]].concat(),
        (None, None) => data.to_vec(),
        _ => panic!("incomplete merge block"),
    }
}

#[test]
fn inserts_absent_block_preserving_unknown_records_parts_and_package_metadata() {
    let source = workbook_bytes(&[("First", &[]), ("Second", &[])]);
    let source = rewrite_first_sheet(&source, |sheet| {
        let mut unknown = Vec::new();
        RecordWriter::new(&mut unknown)
            .write_record(0x0FFE, b"unknown-record-payload")
            .unwrap();
        insert_before_end_sheet(sheet, &unknown)
    });
    let original_sheet = part_blob(&source, "/xl/worksheets/sheet1.bin");
    let original_other_sheet = part_blob(&source, "/xl/worksheets/sheet2.bin");
    let original_workbook = part_blob(&source, "/xl/workbook.bin");

    let mut workbook = XlsbWorkbook::new(Cursor::new(&source)).unwrap();
    workbook
        .set_merged_cell_ranges_by_name(
            "First",
            &[MergedCell::new(5, 6, 2, 3), MergedCell::new(0, 1, 0, 1)],
        )
        .unwrap();
    assert_eq!(
        workbook.merged_cell_ranges(0).unwrap(),
        [MergedCell::new(0, 1, 0, 1), MergedCell::new(5, 6, 2, 3)]
    );
    let output = save(&workbook);
    let changed_sheet = part_blob(&output, "/xl/worksheets/sheet1.bin");
    assert_eq!(strip_merge_block(&changed_sheet), original_sheet);
    assert_eq!(
        part_blob(&output, "/xl/worksheets/sheet2.bin"),
        original_other_sheet
    );
    assert_eq!(part_blob(&output, "/xl/workbook.bin"), original_workbook);

    let before_package = OpcPackage::from_reader(Cursor::new(&source)).unwrap();
    let after_package = OpcPackage::from_reader(Cursor::new(&output)).unwrap();
    for path in [
        "/xl/workbook.bin",
        "/xl/worksheets/sheet1.bin",
        "/xl/worksheets/sheet2.bin",
    ] {
        let uri = PackURI::new(path).unwrap();
        let before = before_package.get_part(&uri).unwrap();
        let after = after_package.get_part(&uri).unwrap();
        assert_eq!(after.content_type(), before.content_type());
        assert_eq!(after.rels().iter().count(), before.rels().iter().count());
    }
    let reparsed = XlsbWorkbook::new(Cursor::new(output)).unwrap();
    assert_eq!(
        reparsed.merged_cell_ranges_by_name("First").unwrap(),
        [MergedCell::new(0, 1, 0, 1), MergedCell::new(5, 6, 2, 3)]
    );
    assert!(reparsed.merged_cell_ranges_by_name("Second").unwrap().is_empty());
}

#[test]
fn replaces_adds_removes_and_clears_present_block_by_index_and_name() {
    let initial = [MergedCell::new(0, 1, 0, 1)];
    let source = workbook_bytes(&[("Data", &initial), ("Other", &[])]);
    let mut workbook = XlsbWorkbook::new(Cursor::new(source)).unwrap();
    workbook
        .add_merged_cell_range_by_name("Data", MergedCell::new(3, 4, 3, 4))
        .unwrap();
    assert_eq!(workbook.merged_cell_ranges(0).unwrap().len(), 2);
    assert!(workbook
        .remove_merged_cell_range(0, &MergedCell::new(0, 1, 0, 1))
        .unwrap());
    assert!(!workbook
        .remove_merged_cell_range_by_name("Data", &MergedCell::new(9, 10, 9, 10))
        .unwrap());
    assert_eq!(
        workbook.merged_cell_ranges_by_name("Data").unwrap(),
        [MergedCell::new(3, 4, 3, 4)]
    );
    workbook.clear_merged_cell_ranges_by_name("Data").unwrap();
    assert!(workbook.merged_cell_ranges(0).unwrap().is_empty());
    assert!(records(&part_blob(&save(&workbook), "/xl/worksheets/sheet1.bin"))
        .iter()
        .all(|(kind, _, _, _)| !matches!(*kind, BEGIN_MERGE_CELLS | MERGE_CELL | END_MERGE_CELLS)));
}

#[test]
fn rejects_bounds_duplicates_and_overlaps_atomically() {
    let source = workbook_bytes(&[("Data", &[])]);
    let mut workbook = XlsbWorkbook::new(Cursor::new(source)).unwrap();
    let invalid_sets = [
        vec![MergedCell::new(2, 1, 0, 1)],
        vec![MergedCell::new(0, 1_048_576, 0, 1)],
        vec![MergedCell::new(0, 1, 0, 16_384)],
        vec![MergedCell::new(0, 1, 0, 1), MergedCell::new(0, 1, 0, 1)],
        vec![MergedCell::new(0, 4, 0, 4), MergedCell::new(3, 6, 4, 8)],
    ];
    for ranges in invalid_sets {
        assert!(workbook.set_merged_cell_ranges(0, &ranges).is_err());
        assert!(workbook.merged_cell_ranges(0).unwrap().is_empty());
    }
    assert!(workbook.merged_cell_ranges(1).is_err());
    assert!(workbook.merged_cell_ranges_by_name("Missing").is_err());
}

#[test]
fn malformed_count_size_duplicate_and_out_of_order_blocks_roll_back() {
    let base = workbook_bytes(&[("Data", &[])]);
    let malformed = [
        rewrite_first_sheet(&base, |sheet| {
            insert_before_end_sheet(sheet, &merge_block(2, &[MergedCell::new(0, 1, 0, 1)]))
        }),
        rewrite_first_sheet(&base, |sheet| {
            let mut block = Vec::new();
            let mut writer = RecordWriter::new(&mut block);
            writer
                .write_record(BEGIN_MERGE_CELLS, &1u32.to_le_bytes())
                .unwrap();
            writer.write_record(MERGE_CELL, &[0; 15]).unwrap();
            writer.write_record(END_MERGE_CELLS, &[]).unwrap();
            insert_before_end_sheet(sheet, &block)
        }),
        rewrite_first_sheet(&base, |sheet| {
            let block = [merge_block(1, &[MergedCell::new(0, 1, 0, 1)]), merge_block(1, &[MergedCell::new(3, 4, 3, 4)])].concat();
            insert_before_end_sheet(sheet, &block)
        }),
        rewrite_first_sheet(&base, |sheet| {
            insert_before_end_sheet(
                sheet,
                &merge_block(
                    2,
                    &[MergedCell::new(5, 6, 5, 6), MergedCell::new(0, 1, 0, 1)],
                ),
            )
        }),
    ];
    for source in malformed {
        let original = part_blob(&source, "/xl/worksheets/sheet1.bin");
        let mut workbook = XlsbWorkbook::new(Cursor::new(source)).unwrap();
        assert!(workbook.merged_cell_ranges(0).is_err());
        assert!(workbook
            .set_merged_cell_ranges(0, &[MergedCell::new(8, 9, 8, 9)])
            .is_err());
        assert_eq!(
            part_blob(&save(&workbook), "/xl/worksheets/sheet1.bin"),
            original
        );
    }
}

#[test]
fn count_payload_and_record_order_are_strict_after_roundtrip() {
    let source = workbook_bytes(&[("Data", &[])]);
    let mut workbook = XlsbWorkbook::new(Cursor::new(source)).unwrap();
    workbook
        .set_merged_cell_ranges(
            0,
            &[MergedCell::new(8, 9, 2, 3), MergedCell::new(1, 2, 4, 5)],
        )
        .unwrap();
    let output = save(&workbook);
    let sheet = part_blob(&output, "/xl/worksheets/sheet1.bin");
    let records = records(&sheet);
    let begin = records
        .iter()
        .position(|(kind, _, _, _)| *kind == BEGIN_MERGE_CELLS)
        .unwrap();
    assert_eq!(records[begin].3, 2u32.to_le_bytes());
    assert_eq!(records[begin + 1].0, MERGE_CELL);
    assert_eq!(records[begin + 1].3.len(), 16);
    assert_eq!(records[begin + 2].0, MERGE_CELL);
    assert_eq!(records[begin + 2].3.len(), 16);
    assert_eq!(records[begin + 3].0, END_MERGE_CELLS);
    assert!(records[begin + 3].3.is_empty());

    let reparsed = XlsbWorkbook::new(Cursor::new(output)).unwrap();
    assert_eq!(
        reparsed.merged_cell_ranges(0).unwrap(),
        [MergedCell::new(1, 2, 4, 5), MergedCell::new(8, 9, 2, 3)]
    );
}
