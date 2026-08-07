//! Public-reader regression fixtures for the MS-XLSB sparkline wire format.
//!
//! The BIFF12 bytes below are assembled directly from the record layouts and
//! identifiers in MS-XLSB. They deliberately do not use `WorkbookWriter`, the
//! private sparkline encoder, or the crate's generic BIFF12 writer.

use litchi_opc::constants::relationship_type;
use litchi_opc::{BlobPart, OpcPackage, PackURI, Part};
use litchi_sheet::sparkline::{AxisType, EmptyCells, SparklineType};
use litchi_xlsb::Workbook;
use litchi_xlsb::sparkline::{ColorType, FormulaKind, Options};
use std::io::Cursor;

// MS-XLSB record identifiers used by this deliberately independent fixture.
const BRT_BEGIN_SHEET: u16 = 0x0081;
const BRT_END_SHEET: u16 = 0x0082;
const BRT_BEGIN_SHEET_DATA: u16 = 0x0091;
const BRT_END_SHEET_DATA: u16 = 0x0092;
const BRT_BUNDLE_SH: u16 = 0x009c;
const BRT_SUP_SELF: u16 = 0x0165;
const BRT_EXTERN_SHEET: u16 = 0x016a;
const BRT_BEGIN_SPARKLINE_GROUP: u16 = 0x0411;
const BRT_END_SPARKLINE_GROUP: u16 = 0x0412;
const BRT_SPARKLINE: u16 = 0x0413;
const BRT_BEGIN_SPARKLINES: u16 = 0x0420;
const BRT_END_SPARKLINES: u16 = 0x0421;
const BRT_BEGIN_SPARKLINE_GROUPS: u16 = 0x0422;
const BRT_END_SPARKLINE_GROUPS: u16 = 0x0423;

#[test]
fn public_reader_decodes_hand_authored_sparkline_records() {
    let workbook = Workbook::new(Cursor::new(package_bytes(false))).expect("open fixture package");
    let snapshot = workbook.sparklines(0).expect("read fixture worksheet");
    let groups = snapshot
        .groups()
        .expect("fixture contains sparkline groups");

    assert_eq!(groups.len(), 1);
    let group = &groups.as_slice()[0];
    assert_eq!(group.kind(), SparklineType::Column);
    assert_eq!(group.empty_cells(), EmptyCells::Gap);
    assert_eq!(
        group.options(),
        Options::MARKERS | Options::HIGH | Options::NEGATIVE | Options::AXIS
    );
    assert_eq!(group.minimum().kind(), AxisType::Custom);
    assert_eq!(group.maximum().kind(), AxisType::Group);
    assert_eq!(group.minimum().manual(), -4.25);
    assert_eq!(group.maximum().manual(), 0.0);
    assert_eq!(group.line_weight(), 1.75);
    assert!(!group.date_axis());
    assert!(group.date_formula().is_none());
    assert_eq!(group.colors().series().color_type(), ColorType::Rgb);
    assert_eq!(group.colors().series().rgba(), [0x11, 0x22, 0x33, 0xff]);

    let sparkline = &group.sparklines()[0];
    assert_eq!(sparkline.location().row(), 4);
    assert_eq!(sparkline.location().column(), 5);
    let state = sparkline.location().state();
    assert!(state.adjusted_deleted());
    assert!(state.adjusted_changed());
    assert!(state.edited());
    assert!(state.unused());

    let formula = sparkline.formula().expect("fixture has a source formula");
    assert_eq!(formula.kind(), FormulaKind::Reference3d);
    assert_eq!(formula.ixti(), Some(0));
    assert!(formula.ancillary().is_empty());
    let reference = formula.reference().expect("PtgRef3d coordinates");
    assert_eq!(reference.row(), 8);
    assert_eq!(reference.column(), 3);
    assert!(reference.row_relative());
    assert!(reference.column_relative());
}

#[test]
fn public_reader_rejects_nonempty_sparkline_delimiter() {
    let workbook = Workbook::new(Cursor::new(package_bytes(true))).expect("open fixture package");
    let error = workbook
        .sparklines(0)
        .expect_err("nonempty BrtEndSparklines must be rejected");

    assert!(
        error
            .to_string()
            .contains("delimiter payload must be empty"),
        "unexpected public read error: {error}"
    );
}

fn package_bytes(malformed_delimiter: bool) -> Vec<u8> {
    let mut workbook_payload = Vec::new();
    push_record(&mut workbook_payload, BRT_SUP_SELF, &[]);
    let mut extern_sheet = Vec::new();
    extern_sheet.extend_from_slice(&1u32.to_le_bytes());
    extern_sheet.extend_from_slice(&0u32.to_le_bytes());
    extern_sheet.extend_from_slice(&0u32.to_le_bytes());
    extern_sheet.extend_from_slice(&0u32.to_le_bytes());
    push_record(&mut workbook_payload, BRT_EXTERN_SHEET, &extern_sheet);

    let mut bundle_sheet = Vec::new();
    bundle_sheet.extend_from_slice(&0u32.to_le_bytes());
    bundle_sheet.extend_from_slice(&1u32.to_le_bytes());
    bundle_sheet.extend_from_slice(&wide_string("rIdSheet1"));
    bundle_sheet.extend_from_slice(&wide_string("Sheet1"));
    push_record(&mut workbook_payload, BRT_BUNDLE_SH, &bundle_sheet);

    let mut workbook_part = BlobPart::new(
        PackURI::new("/xl/workbook.bin").expect("workbook URI"),
        "application/vnd.ms-excel.sheet.binary.macroEnabled.main".to_owned(),
        workbook_payload,
    );
    workbook_part.rels_mut().add_relationship(
        relationship_type::WORKSHEET.to_owned(),
        "worksheets/sheet1.bin".to_owned(),
        "rIdSheet1".to_owned(),
        false,
    );

    let sheet_part = BlobPart::new(
        PackURI::new("/xl/worksheets/sheet1.bin").expect("worksheet URI"),
        "application/vnd.ms-excel.worksheet".to_owned(),
        worksheet_records(malformed_delimiter),
    );

    let mut package = OpcPackage::new();
    package.add_part(Box::new(workbook_part));
    package.add_part(Box::new(sheet_part));
    let mut bytes = Cursor::new(Vec::new());
    package
        .to_stream(&mut bytes)
        .expect("serialize fixture OPC");
    bytes.into_inner()
}

fn worksheet_records(malformed_delimiter: bool) -> Vec<u8> {
    let mut bytes = Vec::new();
    push_record(&mut bytes, BRT_BEGIN_SHEET, &[]);
    push_record(&mut bytes, BRT_BEGIN_SHEET_DATA, &[]);
    push_record(&mut bytes, BRT_END_SHEET_DATA, &[]);
    push_record(&mut bytes, BRT_BEGIN_SPARKLINE_GROUPS, &[]);
    push_record(
        &mut bytes,
        BRT_BEGIN_SPARKLINE_GROUP,
        &sparkline_group_payload(),
    );
    push_record(&mut bytes, BRT_BEGIN_SPARKLINES, &[]);
    push_record(&mut bytes, BRT_SPARKLINE, &sparkline_payload());
    push_record(
        &mut bytes,
        BRT_END_SPARKLINES,
        if malformed_delimiter { &[0x00] } else { &[] },
    );
    push_record(&mut bytes, BRT_END_SPARKLINE_GROUP, &[]);
    push_record(&mut bytes, BRT_END_SPARKLINE_GROUPS, &[]);
    push_record(&mut bytes, BRT_END_SHEET, &[]);
    bytes
}

fn sparkline_group_payload() -> Vec<u8> {
    let mut payload = Vec::new();
    payload.extend_from_slice(&0u32.to_le_bytes()); // FRTHeader: no date-axis formula.
    payload.extend_from_slice(&0x231au16.to_le_bytes());

    for rgba in [
        [0x11, 0x22, 0x33, 0xff],
        [0x44, 0x55, 0x66, 0xff],
        [0x77, 0x88, 0x99, 0xff],
        [0xaa, 0xbb, 0xcc, 0xff],
        [0x10, 0x20, 0x30, 0xff],
        [0x40, 0x50, 0x60, 0xff],
        [0x70, 0x80, 0x90, 0xff],
        [0xa0, 0xb0, 0xc0, 0xff],
    ] {
        payload.extend_from_slice(&[0x05, 0x00, 0x00, 0x00]); // RGB BrtColor, zero tint.
        payload.extend_from_slice(&rgba);
    }

    payload.extend_from_slice(&0.0f64.to_le_bytes());
    payload.extend_from_slice(&(-4.25f64).to_le_bytes());
    payload.extend_from_slice(&1.75f64.to_le_bytes());
    payload.extend_from_slice(&1u32.to_le_bytes()); // Column sparkline.
    payload
}

fn sparkline_payload() -> Vec<u8> {
    let mut payload = Vec::new();
    payload.extend_from_slice(&0x06u32.to_le_bytes()); // FRTSqrefs plus formula.
    payload.extend_from_slice(&1u32.to_le_bytes()); // One FRTSqref.
    payload.extend_from_slice(&0x0001_000fu32.to_le_bytes());
    payload.extend_from_slice(&1i32.to_le_bytes()); // One UncheckedRfX.
    for coordinate in [4u32, 4, 5, 5] {
        payload.extend_from_slice(&coordinate.to_le_bytes());
    }

    payload.extend_from_slice(&1u32.to_le_bytes()); // One FRTFormula.
    payload.extend_from_slice(&2u32.to_le_bytes()); // Required reserved flags.
    payload.extend_from_slice(&9u32.to_le_bytes()); // PtgRef3d byte count.
    payload.extend_from_slice(&0u32.to_le_bytes()); // No ancillary formula bytes.
    payload.push(0x3a); // PtgRef3d with reference-class PtgDataType.
    payload.extend_from_slice(&0u16.to_le_bytes()); // XTI 0: Sheet1 via BrtSupSelf.
    payload.extend_from_slice(&8u32.to_le_bytes());
    payload.extend_from_slice(&0xc003u16.to_le_bytes()); // Column 3; both coordinates relative.
    payload
}

fn push_record(output: &mut Vec<u8>, kind: u16, payload: &[u8]) {
    if kind < 0x80 {
        output.push(kind as u8);
    } else {
        output.push((kind as u8 & 0x7f) | 0x80);
        output.push((kind >> 7) as u8);
    }
    push_varint(output, payload.len());
    output.extend_from_slice(payload);
}

fn push_varint(output: &mut Vec<u8>, mut value: usize) {
    loop {
        let mut byte = (value & 0x7f) as u8;
        value >>= 7;
        if value != 0 {
            byte |= 0x80;
        }
        output.push(byte);
        if value == 0 {
            break;
        }
    }
}

fn wide_string(value: &str) -> Vec<u8> {
    let mut bytes = Vec::new();
    bytes.extend_from_slice(&(value.encode_utf16().count() as u32).to_le_bytes());
    for unit in value.encode_utf16() {
        bytes.extend_from_slice(&unit.to_le_bytes());
    }
    bytes
}
