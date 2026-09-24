#![allow(
    clippy::expect_used,
    clippy::pedantic,
    clippy::shadow_reuse,
    clippy::unwrap_used,
    reason = "binary-index integration tests use small, checked fixtures and panic-on-failure extraction"
)]

//! Integration coverage for worksheet binary-index lookup and maintenance.
//!
//! The index is an internal wire codec, so these tests exercise it through
//! the public source-backed lookup and workbook publication boundaries.  Raw
//! record inspection is limited to checking the offsets that the public API
//! promises to validate and maintain.

use std::io::{self, Cursor};
use std::path::Path;
use std::sync::Arc;
use std::sync::atomic::{AtomicU64, AtomicUsize, Ordering};

use litchi_core::sheet::CellValue;
use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionLimits, Limits as BudgetLimits,
    OwnedSource, ReadAt, SourceVersion,
};
use litchi_opc::{OpcError, OpcPackage, PackURI};
use litchi_xlsb::binary_index::Limits as BinaryIndexLimits;
use litchi_xlsb::cell_values::{Reference, StyleIndex, Value};
use litchi_xlsb::package::PackageError;
use litchi_xlsb::raw::{Header, Kind, Limits as RawLimits, Records, Writer, kind};
use litchi_xlsb::writer::{MutableWorksheet, WorkbookWriter};
use litchi_xlsb::{ReadLimits, SourceBackedWorkbook, Workbook};

const SHEET1: &str = "/xl/worksheets/sheet1.bin";
const INDEX1: &str = "/xl/worksheets/binaryIndex1.bin";
const OPAQUE_KIND: u16 = 0x1234;
const MAX_ROW: u32 = 1_048_575;
const MAX_COLUMN: u32 = 16_383;

fn fixture(relative: &str) -> Vec<u8> {
    std::fs::read(
        Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../..")
            .join(relative),
    )
    .expect("fixture")
}

fn source_workbook(relative: &str) -> SourceBackedWorkbook {
    SourceBackedWorkbook::from_reader(Cursor::new(fixture(relative))).expect("source workbook")
}

fn package_part(package_bytes: &[u8], path: &str) -> Vec<u8> {
    let package = OpcPackage::from_reader(Cursor::new(package_bytes)).expect("OPC package");
    package
        .get_part(&PackURI::new(path).expect("part URI"))
        .expect("part")
        .blob()
        .to_vec()
}

fn save_workbook(workbook: &Workbook) -> Vec<u8> {
    let mut output = Cursor::new(Vec::new());
    workbook.save(&mut output).expect("save workbook");
    output.into_inner()
}

#[derive(Debug, Clone)]
struct RawRecord {
    kind: Kind,
    start: usize,
    payload: Vec<u8>,
}

fn records(data: &[u8]) -> Vec<RawRecord> {
    let mut output = Vec::new();
    for record in Records::new(data) {
        let record = record.expect("record framing");
        output.push(RawRecord {
            kind: record.kind(),
            start: record.offset(),
            payload: record.payload().to_vec(),
        });
    }
    output
}

fn rewrite_index(package_bytes: &[u8], transform: impl FnOnce(&[u8]) -> Vec<u8>) -> Vec<u8> {
    let mut package = OpcPackage::from_reader(Cursor::new(package_bytes)).expect("OPC package");
    let uri = PackURI::new(INDEX1).expect("index URI");
    let changed = transform(package.get_part(&uri).expect("index part").blob());
    package
        .get_part_mut(&uri)
        .expect("index part")
        .set_blob(changed);
    let mut output = Vec::new();
    package.to_stream(&mut output).expect("rewrite package");
    output
}

fn remove_index_relationship(package_bytes: &[u8]) -> Vec<u8> {
    let mut package = OpcPackage::from_reader(Cursor::new(package_bytes)).expect("OPC package");
    let worksheet = PackURI::new(SHEET1).expect("worksheet URI");
    let relationship_id = package
        .get_part(&worksheet)
        .expect("worksheet part")
        .rels()
        .iter()
        .find(|relationship| {
            relationship.reltype()
                == "http://schemas.microsoft.com/office/2006/relationships/xlBinaryIndex"
        })
        .expect("binary-index relationship")
        .r_id()
        .to_owned();
    package
        .get_part_mut(&worksheet)
        .expect("worksheet part")
        .rels_mut()
        .remove(&relationship_id)
        .expect("remove binary-index relationship");
    let mut output = Vec::new();
    package.to_stream(&mut output).expect("rewrite package");
    output
}

fn insert_opaque_index_record(index: &[u8]) -> Vec<u8> {
    let insertion = {
        let mut bytes = Vec::new();
        Writer::new(&mut bytes)
            .write_record(
                Kind::new(OPAQUE_KIND).expect("opaque kind"),
                b"extension bytes",
            )
            .expect("opaque record");
        bytes
    };
    let offset = records(index)
        .into_iter()
        .find(|record| record.kind == kind::INDEX_BLOCK)
        .expect("index block")
        .start;
    let mut output = Vec::with_capacity(index.len() + insertion.len());
    output.extend_from_slice(&index[..offset]);
    output.extend_from_slice(&insertion);
    output.extend_from_slice(&index[offset..]);
    output
}

fn truncate_first_row_block(index: &[u8]) -> Vec<u8> {
    let first_row_block = records(index)
        .into_iter()
        .find(|record| record.kind == kind::INDEX_ROW_BLOCK)
        .expect("row block");
    let (_, header_len) = Header::parse(&index[first_row_block.start..], RawLimits::DEFAULT)
        .expect("row block header");
    let original_payload_end = first_row_block.start + header_len + first_row_block.payload.len();
    let new_payload_len = first_row_block.payload.len() - 4;
    let mut rebuilt = Vec::with_capacity(index.len() - 4);
    rebuilt.extend_from_slice(&index[..first_row_block.start]);
    let mut writer = Writer::new(&mut rebuilt);
    writer
        .write_record(
            kind::INDEX_ROW_BLOCK,
            &first_row_block.payload[..new_payload_len],
        )
        .expect("truncated row block");
    rebuilt.extend_from_slice(&index[original_payload_end..]);
    rebuilt
}

fn append_index_part_end(index: &[u8]) -> Vec<u8> {
    let mut output = index.to_vec();
    Writer::new(&mut output)
        .write_record(kind::INDEX_PART_END, &[])
        .expect("second end marker");
    output
}

fn append_empty_index_block(index: &[u8]) -> Vec<u8> {
    let end = records(index)
        .into_iter()
        .find(|record| record.kind == kind::INDEX_PART_END)
        .expect("part end marker")
        .start;
    let mut block = Vec::new();
    block.extend_from_slice(&32_u32.to_le_bytes());
    block.extend_from_slice(&64_u32.to_le_bytes());
    block.extend_from_slice(&[0; 8]);
    block.extend_from_slice(&[0; 8]);
    let mut encoded = Vec::new();
    Writer::new(&mut encoded)
        .write_record(kind::INDEX_BLOCK, &block)
        .expect("empty index block");
    let mut output = Vec::with_capacity(index.len() + encoded.len());
    output.extend_from_slice(&index[..end]);
    output.extend_from_slice(&encoded);
    output.extend_from_slice(&index[end..]);
    output
}

fn write_empty_index_block(writer: &mut Writer<&mut Vec<u8>>, row_start: u32) {
    let mut block = Vec::new();
    block.extend_from_slice(&row_start.to_le_bytes());
    block.extend_from_slice(&(row_start + 32).to_le_bytes());
    block.extend_from_slice(&[0; 8]);
    block.extend_from_slice(&[0; 8]);
    writer
        .write_record(kind::INDEX_BLOCK, &block)
        .expect("empty index block");
}

fn two_empty_index_blocks(_index: &[u8]) -> Vec<u8> {
    let mut output = Vec::new();
    let mut writer = Writer::new(&mut output);
    write_empty_index_block(&mut writer, 0);
    write_empty_index_block(&mut writer, 32);
    writer
        .write_record(kind::INDEX_PART_END, &[])
        .expect("empty index end marker");
    output
}

fn empty_index_with_end(_index: &[u8]) -> Vec<u8> {
    let mut output = Vec::new();
    Writer::new(&mut output)
        .write_record(kind::INDEX_PART_END, &[])
        .expect("empty index end marker");
    output
}

fn append_two_index_part_ends(index: &[u8]) -> Vec<u8> {
    append_index_part_end(&append_index_part_end(index))
}

fn overlap_second_index_block(index: &[u8]) -> Vec<u8> {
    let second = records(index)
        .into_iter()
        .filter(|record| record.kind == kind::INDEX_BLOCK)
        .nth(1)
        .expect("second index block");
    let (_, header_len) =
        Header::parse(&index[second.start..], RawLimits::DEFAULT).expect("index block header");
    let mut output = index.to_vec();
    output[second.start + header_len..second.start + header_len + 4]
        .copy_from_slice(&16_u32.to_le_bytes());
    output[second.start + header_len + 4..second.start + header_len + 8]
        .copy_from_slice(&48_u32.to_le_bytes());
    output
}

fn is_cell_kind(record_kind: Kind) -> bool {
    (kind::CELL_BLANK.get()..=kind::FMLA_ERROR.get()).contains(&record_kind.get())
        || record_kind == kind::CELL_R_STRING
}

fn worksheet_cell_offset(worksheet: &[u8], wanted_row: u32, wanted_column: u32) -> usize {
    let mut current_row = None;
    records(worksheet)
        .into_iter()
        .find_map(|record| {
            if record.kind == kind::ROW_HDR {
                current_row = Some(u32::from_le_bytes(
                    record.payload[..4].try_into().expect("row header payload"),
                ));
                return None;
            }
            if !is_cell_kind(record.kind) || current_row != Some(wanted_row) {
                return None;
            }
            let column = u32::from_le_bytes(record.payload[..4].try_into().expect("cell payload"));
            (column == wanted_column).then_some(record.start)
        })
        .expect("worksheet cell")
}

fn index_anchor_offset(index: &[u8], wanted_row: u32, wanted_column_block: u32) -> u64 {
    let mut block_start = None;
    for record in records(index) {
        match record.kind {
            kind::INDEX_BLOCK => {
                block_start = Some(u32::from_le_bytes(
                    record.payload[..4]
                        .try_into()
                        .expect("index block row start"),
                ));
            },
            kind::INDEX_ROW_BLOCK if block_start.is_some() => {
                let start = block_start.expect("block start");
                let row_mask =
                    u32::from_le_bytes(record.payload[..4].try_into().expect("row mask"));
                let base = u64::from_le_bytes(record.payload[4..12].try_into().expect("row base"));
                let row_count = usize::try_from(row_mask.count_ones()).expect("row count");
                let mut mask_offset = 12;
                let mut offset_cursor = 12 + row_count * 2;
                for row_offset in 0..32_u32 {
                    if row_mask & (1_u32 << row_offset) == 0 {
                        continue;
                    }
                    let row = start + row_offset;
                    let mask = u16::from_le_bytes(
                        record.payload[mask_offset..mask_offset + 2]
                            .try_into()
                            .expect("column mask"),
                    );
                    mask_offset += 2;
                    for column_block in 0..16_u32 {
                        if mask & (1_u16 << column_block) == 0 {
                            continue;
                        }
                        let relative = u32::from_le_bytes(
                            record.payload[offset_cursor..offset_cursor + 4]
                                .try_into()
                                .expect("column offset"),
                        );
                        offset_cursor += 4;
                        if row == wanted_row && column_block == wanted_column_block {
                            return base + u64::from(relative);
                        }
                    }
                }
                block_start = None;
            },
            _ => {},
        }
    }
    panic!("missing index anchor row={wanted_row} block={wanted_column_block}");
}

fn index_row_masks(index: &[u8]) -> Vec<(u32, u16)> {
    let mut block_start = None;
    let mut rows = Vec::new();
    for record in records(index) {
        match record.kind {
            kind::INDEX_BLOCK => {
                block_start = Some(u32::from_le_bytes(
                    record.payload[..4]
                        .try_into()
                        .expect("index block row start"),
                ));
            },
            kind::INDEX_ROW_BLOCK => {
                let start = block_start.take().expect("row block after index block");
                let row_mask =
                    u32::from_le_bytes(record.payload[..4].try_into().expect("index row mask"));
                let mut cursor = 12;
                for row_offset in 0..32_u32 {
                    if row_mask & (1_u32 << row_offset) == 0 {
                        continue;
                    }
                    let mask = u16::from_le_bytes(
                        record.payload[cursor..cursor + 2]
                            .try_into()
                            .expect("index column mask"),
                    );
                    cursor += 2;
                    rows.push((start + row_offset, mask));
                }
            },
            _ => {},
        }
    }
    rows
}

fn forge_first_index_offset(index: &[u8]) -> Vec<u8> {
    let row_block = records(index)
        .into_iter()
        .find(|record| record.kind == kind::INDEX_ROW_BLOCK)
        .expect("row block");
    let (_, header_len) =
        Header::parse(&index[row_block.start..], RawLimits::DEFAULT).expect("row block header");
    let row_mask = u32::from_le_bytes(row_block.payload[..4].try_into().expect("row mask"));
    let offsets_offset = header_len + 12 + usize::try_from(row_mask.count_ones()).unwrap() * 2;
    let absolute = row_block.start + offsets_offset;
    let mut output = index.to_vec();
    output[absolute..absolute + 4].copy_from_slice(&u32::MAX.to_le_bytes());
    output
}

#[test]
fn indexed_values_lookup_present_absent_and_formula_cache() {
    let workbook = source_workbook("test-data/ooxml/xlsb/sample.xlsb");
    let sheet = workbook
        .worksheet_by_index(0)
        .expect("worksheet selector")
        .expect("first worksheet");
    let indexed = sheet
        .indexed_values()
        .expect("indexed values")
        .expect("binary index relationship");

    assert_eq!(
        indexed.cached_value(0, 1).expect("numeric lookup"),
        Some(CellValue::Float(111.0))
    );
    assert!(
        indexed
            .cached_value(9, 1)
            .expect("formula cache lookup")
            .is_some_and(|value| matches!(value, CellValue::Float(_)))
    );
    assert_eq!(indexed.cached_value(0, 2).expect("absent lookup"), None);
    assert!(indexed.cached_value(1_048_576, 0).is_err());
    assert!(indexed.cached_value(0, 16_384).is_err());
}

#[test]
fn selected_shared_string_obeys_index_raw_string_limit() {
    let workbook = source_workbook("test-data/ooxml/xlsb/sample.xlsb");
    let sheet = workbook
        .worksheet_by_index(0)
        .expect("worksheet selector")
        .expect("first worksheet");
    let high_limits = BinaryIndexLimits::new(
        RawLimits::new(RawLimits::DEFAULT.payload(), 5),
        1_000_000,
        32_768,
        1_000_000,
    );
    let high = sheet
        .indexed_values_with_limits(high_limits)
        .expect("high-limit indexed values")
        .expect("binary index relationship");
    assert_eq!(
        high.cached_value(0, 0).expect("shared string lookup"),
        Some(CellValue::String("Lorem".to_string()))
    );

    let low_limits = BinaryIndexLimits::new(
        RawLimits::new(RawLimits::DEFAULT.payload(), 1),
        1_000_000,
        32_768,
        1_000_000,
    );
    let low = sheet
        .indexed_values_with_limits(low_limits)
        .expect("low-limit index binding")
        .expect("binary index relationship");
    assert!(low.cached_value(0, 0).is_err());
    assert_eq!(
        low.cached_value(0, 1).expect("numeric lookup"),
        Some(CellValue::Float(111.0))
    );
}

#[test]
fn indexed_values_accepts_real_zero_column_masks() {
    let workbook = source_workbook("test-data/poi/test-data/spreadsheet/testVarious.xlsb");
    let sheet = workbook
        .worksheet_by_index(0)
        .expect("worksheet selector")
        .expect("first worksheet");
    let indexed = sheet
        .indexed_values()
        .expect("indexed values")
        .expect("binary index relationship");

    assert!(
        indexed
            .cached_value(0, 0)
            .expect("first indexed cell")
            .is_some()
    );
    assert_eq!(indexed.cached_value(64, 0).expect("zero-column row"), None);
    assert_eq!(
        indexed.cached_value(65, 16_383).expect("zero-column row"),
        None
    );
}

#[test]
fn topology_change_preserves_formatting_only_index_rows() {
    let mut workbook = Workbook::new(Cursor::new(fixture(
        "test-data/poi/test-data/spreadsheet/testVarious.xlsb",
    )))
    .expect("workbook");
    let before_package = save_workbook(&workbook);
    let before_index = package_part(&before_package, INDEX1);
    let before_rows = index_row_masks(&before_index);
    for row in [22, 23, 27, 64, 65] {
        assert!(
            before_rows.contains(&(row, 0)),
            "fixture row {row} should be represented with a zero column mask"
        );
    }

    let reference = Reference::new(22, 0).expect("reference");
    let snapshot = workbook.cell_values(0).expect("cell snapshot");
    assert!(snapshot.cell(reference).expect("lookup").is_none());
    let mut edit = snapshot.edit();
    edit.insert(
        reference,
        StyleIndex::new(0).expect("style"),
        Value::Number(9.25),
    )
    .expect("insert cell into formatting-only row");
    let commit = edit.commit().expect("commit");
    let published = workbook
        .apply_cell_values(0, &commit)
        .expect("publish topology-changing edit");
    assert_eq!(
        published
            .cell(reference)
            .expect("published lookup")
            .expect("inserted cell")
            .value(),
        &Value::Number(9.25)
    );

    let after_package = save_workbook(&workbook);
    let after_index = package_part(&after_package, INDEX1);
    let after_rows = index_row_masks(&after_index);
    assert!(after_rows.contains(&(22, 1)), "new cell mask for row 22");
    for row in [23, 27, 64, 65] {
        assert!(
            after_rows.contains(&(row, 0)),
            "topology regeneration dropped formatting-only row {row}"
        );
    }

    let published_source = SourceBackedWorkbook::from_reader(Cursor::new(after_package))
        .expect("reopen published package");
    let indexed = published_source
        .worksheet_by_index(0)
        .expect("worksheet selector")
        .expect("worksheet")
        .indexed_values()
        .expect("indexed values")
        .expect("index");
    assert_eq!(
        indexed.cached_value(22, 0).expect("indexed inserted cell"),
        Some(CellValue::Float(9.25))
    );
    assert_eq!(indexed.cached_value(64, 0).expect("row-only lookup"), None);
}

#[test]
fn missing_index_relationship_returns_no_index_handle() {
    let source = remove_index_relationship(&fixture("test-data/ooxml/xlsb/sample.xlsb"));
    let workbook = SourceBackedWorkbook::from_reader(Cursor::new(source)).expect("catalog");
    let sheet = workbook
        .worksheet_by_index(0)
        .expect("worksheet selector")
        .expect("worksheet");
    assert!(
        sheet
            .indexed_values()
            .expect("missing index is not an error")
            .is_none()
    );
}

#[test]
fn writer_emits_lookupable_offsets_for_sparse_cells() {
    let mut sheet = MutableWorksheet::new("Indexed");
    sheet.set_cell(0, 0, 1.25_f64);
    sheet.set_cell(0, 1_024, 2.5_f64);
    sheet.set_cell(33, 2, 3.75_f64);
    sheet.set_cell(MAX_ROW, MAX_COLUMN, 4.0_f64);
    let mut writer = WorkbookWriter::new();
    writer.add_worksheet(sheet);
    let mut bytes = Cursor::new(Vec::new());
    writer.save(&mut bytes).expect("writer save");
    let package_bytes = bytes.into_inner();

    let index = package_part(&package_bytes, INDEX1);
    assert!(
        records(&index)
            .iter()
            .any(|record| record.kind == kind::INDEX_BLOCK)
    );
    let worksheet = package_part(&package_bytes, SHEET1);
    assert_eq!(
        index_anchor_offset(&index, 0, 0),
        u64::try_from(worksheet_cell_offset(&worksheet, 0, 0)).expect("offset")
    );
    assert_eq!(
        index_anchor_offset(&index, 0, 1),
        u64::try_from(worksheet_cell_offset(&worksheet, 0, 1_024)).expect("offset")
    );
    assert_eq!(
        index_anchor_offset(&index, 33, 0),
        u64::try_from(worksheet_cell_offset(&worksheet, 33, 2)).expect("offset")
    );

    let workbook = SourceBackedWorkbook::from_reader(Cursor::new(package_bytes)).expect("open");
    let indexed = workbook
        .worksheet_by_index(0)
        .expect("worksheet selector")
        .expect("worksheet")
        .indexed_values()
        .expect("indexed values")
        .expect("index");
    assert_eq!(
        indexed.cached_value(0, 1_024).expect("far-column lookup"),
        Some(CellValue::Float(2.5))
    );
    assert_eq!(
        indexed.cached_value(33, 2).expect("second block lookup"),
        Some(CellValue::Float(3.75))
    );
    assert_eq!(
        indexed
            .cached_value(MAX_ROW, MAX_COLUMN)
            .expect("maximum coordinate lookup"),
        Some(CellValue::Float(4.0))
    );
}

#[test]
fn writer_emits_an_empty_index_block_for_an_empty_sheet() {
    let mut writer = WorkbookWriter::new();
    writer.add_worksheet(MutableWorksheet::new("Empty"));
    let mut bytes = Cursor::new(Vec::new());
    writer.save(&mut bytes).expect("writer save");
    let package_bytes = bytes.into_inner();
    let index = package_part(&package_bytes, INDEX1);
    let index_records = records(&index);
    assert_eq!(
        index_records
            .iter()
            .filter(|record| record.kind == kind::INDEX_BLOCK)
            .count(),
        1
    );
    assert_eq!(
        index_records
            .iter()
            .filter(|record| record.kind == kind::INDEX_ROW_BLOCK)
            .count(),
        0
    );
    assert_eq!(
        index_records
            .iter()
            .filter(|record| record.kind == kind::INDEX_PART_END)
            .count(),
        1
    );
    let workbook = SourceBackedWorkbook::from_reader(Cursor::new(package_bytes)).expect("open");
    let indexed = workbook
        .worksheet_by_index(0)
        .expect("worksheet selector")
        .expect("worksheet")
        .indexed_values()
        .expect("indexed values")
        .expect("empty index");
    assert_eq!(indexed.cached_value(0, 0).expect("empty lookup"), None);
}

#[test]
fn index_accepts_optional_empty_blocks_and_two_end_markers() {
    let source = rewrite_index(&fixture("test-data/ooxml/xlsb/sample.xlsb"), |index| {
        append_index_part_end(&append_empty_index_block(index))
    });
    let workbook = SourceBackedWorkbook::from_reader(Cursor::new(source)).expect("catalog");
    let indexed = workbook
        .worksheet_by_index(0)
        .expect("worksheet selector")
        .expect("worksheet")
        .indexed_values()
        .expect("optional empty block and second end")
        .expect("index");
    assert_eq!(
        indexed.cached_value(0, 1).expect("indexed lookup"),
        Some(CellValue::Float(111.0))
    );
}

#[test]
fn index_accepts_an_empty_block_before_a_later_empty_block() {
    let mut writer = WorkbookWriter::new();
    writer.add_worksheet(MutableWorksheet::new("Empty"));
    let mut bytes = Cursor::new(Vec::new());
    writer.save(&mut bytes).expect("writer save");
    let source = rewrite_index(&bytes.into_inner(), two_empty_index_blocks);
    let workbook = SourceBackedWorkbook::from_reader(Cursor::new(source)).expect("catalog");
    let indexed = workbook
        .worksheet_by_index(0)
        .expect("worksheet selector")
        .expect("worksheet")
        .indexed_values()
        .expect("two empty blocks")
        .expect("index");
    assert_eq!(indexed.cached_value(32, 0).expect("empty lookup"), None);
}

#[test]
fn removing_an_earlier_cell_updates_later_anchor_offsets() {
    let source = fixture("test-data/ooxml/xlsb/sample.xlsb");
    let before_index = package_part(&source, INDEX1);
    let mut workbook = Workbook::new(Cursor::new(source)).expect("workbook");
    let snapshot = workbook.cell_values(0).expect("cell snapshot");
    let mut edit = snapshot.edit();
    edit.remove(Reference::new(0, 1).expect("reference"))
        .expect("remove second cell");
    let commit = edit.commit().expect("commit");
    workbook
        .apply_cell_values(0, &commit)
        .expect("publish edit");
    let saved = save_workbook(&workbook);
    let after_index = package_part(&saved, INDEX1);
    let after_sheet = package_part(&saved, SHEET1);

    assert_ne!(before_index, after_index);
    assert_eq!(
        index_anchor_offset(&after_index, 1, 0),
        u64::try_from(worksheet_cell_offset(&after_sheet, 1, 0)).expect("later cell offset")
    );
    let indexed = SourceBackedWorkbook::from_reader(Cursor::new(saved))
        .expect("reopen source workbook")
        .worksheet_by_index(0)
        .expect("worksheet selector")
        .expect("worksheet")
        .indexed_values()
        .expect("indexed values")
        .expect("index");
    assert_eq!(
        indexed.cached_value(1, 0).expect("later cell lookup"),
        Some(CellValue::String("ipsum".to_string()))
    );
}

#[test]
fn no_op_save_keeps_the_source_index_byte_exact() {
    let source = rewrite_index(
        &fixture("test-data/ooxml/xlsb/sample.xlsb"),
        insert_opaque_index_record,
    );
    let before_index = package_part(&source, INDEX1);
    let mut workbook = Workbook::new(Cursor::new(source)).expect("workbook");
    let snapshot = workbook.cell_values(0).expect("snapshot");
    let commit = snapshot.edit().commit().expect("empty commit");
    workbook
        .apply_cell_values(0, &commit)
        .expect("publish no-op");
    let after_index = package_part(&save_workbook(&workbook), INDEX1);
    assert_eq!(before_index, after_index);
}

#[test]
fn opaque_index_record_refuses_offset_patch_atomically() {
    let source = fixture("test-data/ooxml/xlsb/sample.xlsb");
    let decorated = rewrite_index(&source, insert_opaque_index_record);
    let before_index = package_part(&decorated, INDEX1);
    let before_sheet = package_part(&decorated, SHEET1);
    assert!(records(&before_index).iter().any(|record| {
        record.kind.get() == OPAQUE_KIND && record.payload == b"extension bytes"
    }));
    let mut workbook = Workbook::new(Cursor::new(decorated)).expect("workbook");
    let snapshot = workbook.cell_values(0).expect("snapshot");
    let mut edit = snapshot.edit();
    edit.remove(Reference::new(0, 1).expect("reference"))
        .expect("remove second cell");
    let commit = edit.commit().expect("commit");
    assert!(workbook.apply_cell_values(0, &commit).is_err());
    let after_package = save_workbook(&workbook);
    assert_eq!(before_index, package_part(&after_package, INDEX1));
    assert_eq!(before_sheet, package_part(&after_package, SHEET1));
}

#[test]
fn malformed_or_forged_index_is_refused_before_lookup() {
    let cases: [(&str, fn(&[u8]) -> Vec<u8>); 2] = [
        ("truncated", |_index| vec![42]),
        ("forged offset", forge_first_index_offset),
    ];
    for (name, transform) in cases {
        let source = rewrite_index(&fixture("test-data/ooxml/xlsb/sample.xlsb"), transform);
        let workbook = SourceBackedWorkbook::from_reader(Cursor::new(source)).expect("catalog");
        let sheet = workbook
            .worksheet_by_index(0)
            .expect("worksheet selector")
            .expect("worksheet");
        assert!(sheet.indexed_values().is_err(), "accepted {name} index");
    }
}

#[test]
fn malformed_index_structure_is_refused() {
    let cases: [(&str, &str, fn(&[u8]) -> Vec<u8>); 2] = [
        (
            "truncated offset array",
            "test-data/ooxml/xlsb/sample.xlsb",
            truncate_first_row_block,
        ),
        (
            "overlapping index ranges",
            "test-data/poi/test-data/spreadsheet/testVarious.xlsb",
            overlap_second_index_block,
        ),
    ];
    for (name, fixture_name, transform) in cases {
        let source = rewrite_index(&fixture(fixture_name), transform);
        let workbook = SourceBackedWorkbook::from_reader(Cursor::new(source)).expect("catalog");
        let sheet = workbook
            .worksheet_by_index(0)
            .expect("worksheet selector")
            .expect("worksheet");
        assert!(sheet.indexed_values().is_err(), "accepted {name}");
    }
}

#[test]
fn index_rejects_zero_blocks_and_more_than_two_end_markers() {
    let cases: [(&str, &str, fn(&[u8]) -> Vec<u8>); 2] = [
        (
            "zero blocks",
            "test-data/ooxml/xlsb/sample.xlsb",
            empty_index_with_end,
        ),
        (
            "three end markers",
            "test-data/ooxml/xlsb/sample.xlsb",
            append_two_index_part_ends,
        ),
    ];
    for (name, fixture_name, transform) in cases {
        let source = rewrite_index(&fixture(fixture_name), transform);
        let workbook = SourceBackedWorkbook::from_reader(Cursor::new(source)).expect("catalog");
        let sheet = workbook
            .worksheet_by_index(0)
            .expect("worksheet selector")
            .expect("worksheet");
        assert!(sheet.indexed_values().is_err(), "accepted {name}");
    }
}

#[derive(Debug)]
struct VersionedSource {
    bytes: Vec<u8>,
    revision: AtomicU64,
}

impl VersionedSource {
    fn new(bytes: Vec<u8>) -> Self {
        Self {
            bytes,
            revision: AtomicU64::new(0),
        }
    }

    fn change(&self) {
        self.revision.fetch_add(1, Ordering::SeqCst);
    }
}

impl ReadAt for VersionedSource {
    fn len(&self) -> io::Result<u64> {
        Ok(u64::try_from(self.bytes.len()).expect("source length"))
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        let offset = usize::try_from(offset)
            .map_err(|_| io::Error::new(io::ErrorKind::InvalidInput, "offset too large"))?;
        if offset >= self.bytes.len() {
            return Ok(0);
        }
        let end = offset.saturating_add(output.len()).min(self.bytes.len());
        output[..end - offset].copy_from_slice(&self.bytes[offset..end]);
        Ok(end - offset)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        Ok(SourceVersion::new(
            0x584c_5342,
            self.revision.load(Ordering::SeqCst),
        ))
    }
}

fn managed_context() -> (CancellationSource, ExecutionContext) {
    let budget = Budget::root(
        "xlsb-index-integration",
        BudgetLimits::new(u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX),
    );
    let (cancellation_source, cancellation) = CancellationSource::pair();
    let execution_limits = ExecutionLimits::new(
        std::num::NonZeroUsize::new(1).expect("concurrency"),
        std::num::NonZeroUsize::new(1).expect("recursion"),
        std::num::NonZeroU64::new(u64::MAX).expect("work"),
        0,
    )
    .expect("execution limits");
    (
        cancellation_source,
        ExecutionContext::new(budget, cancellation, execution_limits),
    )
}

#[test]
fn indexed_handle_refuses_stale_source_and_cancellation() {
    let source = Arc::new(VersionedSource::new(fixture(
        "test-data/ooxml/xlsb/sample.xlsb",
    )));
    let workbook = SourceBackedWorkbook::from_read_at(source.clone()).expect("source workbook");
    let indexed = workbook
        .worksheet_by_index(0)
        .expect("worksheet selector")
        .expect("worksheet")
        .indexed_values()
        .expect("indexed values")
        .expect("index");
    source.change();
    assert!(matches!(
        indexed.cached_value(0, 1),
        Err(PackageError::Opc(OpcError::SourceChanged { .. }))
    ));

    let (cancellation_source, context) = managed_context();
    let managed = SourceBackedWorkbook::from_read_at_with_execution_context(
        Arc::new(OwnedSource::new(fixture(
            "test-data/ooxml/xlsb/sample.xlsb",
        ))),
        ReadLimits::default(),
        context,
    )
    .expect("managed source workbook");
    let managed_index = managed
        .worksheet_by_index(0)
        .expect("worksheet selector")
        .expect("worksheet")
        .indexed_values()
        .expect("indexed values")
        .expect("index");
    cancellation_source.cancel();
    assert!(matches!(
        managed_index.cached_value(0, 1),
        Err(PackageError::Opc(OpcError::Cancelled))
    ));
}

#[derive(Debug)]
struct CountingSource {
    bytes: Vec<u8>,
    reads: AtomicUsize,
}

impl CountingSource {
    fn new(bytes: Vec<u8>) -> Self {
        Self {
            bytes,
            reads: AtomicUsize::new(0),
        }
    }

    fn reads(&self) -> usize {
        self.reads.load(Ordering::SeqCst)
    }
}

impl ReadAt for CountingSource {
    fn len(&self) -> io::Result<u64> {
        Ok(u64::try_from(self.bytes.len()).expect("source length"))
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        self.reads.fetch_add(1, Ordering::SeqCst);
        let offset = usize::try_from(offset)
            .map_err(|_| io::Error::new(io::ErrorKind::InvalidInput, "offset too large"))?;
        if offset >= self.bytes.len() {
            return Ok(0);
        }
        let end = offset.saturating_add(output.len()).min(self.bytes.len());
        output[..end - offset].copy_from_slice(&self.bytes[offset..end]);
        Ok(end - offset)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        Ok(SourceVersion::new(0x584c_5342, 0))
    }
}

#[test]
fn managed_index_handle_keeps_worksheet_part_data_alive() {
    let source = Arc::new(CountingSource::new(fixture(
        "test-data/ooxml/xlsb/62815.xlsb",
    )));
    let (_cancellation_source, context) = managed_context();
    let workbook = SourceBackedWorkbook::from_read_at_with_execution_context(
        source.clone(),
        ReadLimits::default(),
        context,
    )
    .expect("managed source workbook");
    let indexed = workbook
        .worksheet_by_index(0)
        .expect("worksheet selector")
        .expect("worksheet")
        .indexed_values()
        .expect("indexed values")
        .expect("index");
    let reads_after_open = source.reads();
    drop(workbook);

    assert!(matches!(
        indexed.cached_value(0, 0).expect("retained cell lookup"),
        Some(CellValue::Float(_))
    ));
    assert_eq!(source.reads(), reads_after_open);
}

#[test]
fn source_bytes_limit_rejects_index_before_payload_reads() {
    let source = Arc::new(CountingSource::new(fixture(
        "test-data/ooxml/xlsb/sample.xlsb",
    )));
    let workbook = SourceBackedWorkbook::from_read_at(source.clone()).expect("source workbook");
    let sheet = workbook
        .worksheet_by_index(0)
        .expect("worksheet selector")
        .expect("worksheet");
    let reads_before = source.reads();
    let limits = BinaryIndexLimits::new_with_source_bytes(
        RawLimits::DEFAULT,
        1,
        1_000_000,
        32_768,
        1_000_000,
    );
    assert!(sheet.indexed_values_with_limits(limits).is_err());
    assert_eq!(source.reads(), reads_before);
}
