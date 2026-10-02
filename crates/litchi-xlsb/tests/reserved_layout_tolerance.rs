#![allow(
    clippy::unwrap_used,
    reason = "test assertions use panic-on-failure extraction by design"
)]

//! Audit finding F3.11: reserved bits, strict layouts, and trailing payload
//! bytes are repaired the way Excel reads past them, and every repair is
//! recorded. Each fixture is a conformant workbook with one deviation
//! injected; the unmodified workbook is the control and repairs nothing.

use std::io::Cursor;

use litchi_core::sheet::traits::WorkbookTrait;
use litchi_opc::{OpcPackage, PackURI};
use litchi_xlsb::raw::{Kind, Records, Writer, kind};
use litchi_xlsb::{
    Defect, LayoutDefect, Package, ReservedFieldDefect, SourceBackedWorkbook, ToleranceReport,
    Workbook,
};

type Stream = Vec<(Kind, Vec<u8>)>;

const WORKBOOK: &str = "/xl/workbook.bin";
const STYLES: &str = "/xl/styles.bin";
const SHEET1: &str = "/xl/worksheets/sheet1.bin";

/// Columns, rows, RK, ISST, BLANK, and BrtFmlaNum cells on sheet 1.
fn fixture() -> Vec<u8> {
    std::fs::read(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../test-data/ooxml/xlsb/universal-content.xlsb"
    ))
    .unwrap()
}

fn rewrite(package_bytes: &[u8], part: &str, transform: impl FnOnce(&mut Stream)) -> Vec<u8> {
    let mut package = OpcPackage::from_reader(Cursor::new(package_bytes)).unwrap();
    let uri = PackURI::new(part).unwrap();
    let mut stream: Stream = Records::new(package.get_part(&uri).unwrap().blob())
        .map(|record| {
            let record = record.unwrap();
            (record.kind(), record.payload().to_vec())
        })
        .collect();
    transform(&mut stream);
    let mut blob = Vec::new();
    let mut writer = Writer::new(&mut blob);
    for (record_type, payload) in &stream {
        writer.write_record(*record_type, payload).unwrap();
    }
    package.get_part_mut(&uri).unwrap().set_blob(blob);
    let mut output = Cursor::new(Vec::new());
    package.to_stream(&mut output).unwrap();
    output.into_inner()
}

fn first_payload(stream: &mut Stream, record_type: Kind) -> &mut Vec<u8> {
    &mut stream
        .iter_mut()
        .find(|(kind, _)| *kind == record_type)
        .unwrap()
        .1
}

fn position_after_last(stream: &Stream, record_type: Kind) -> usize {
    stream
        .iter()
        .rposition(|(kind, _)| *kind == record_type)
        .unwrap()
        + 1
}

fn wide_string(value: &str) -> Vec<u8> {
    let units: Vec<u16> = value.encode_utf16().collect();
    let mut output = u32::try_from(units.len()).unwrap().to_le_bytes().to_vec();
    for unit in units {
        output.extend_from_slice(&unit.to_le_bytes());
    }
    output
}

/// A `BrtName` defining `name` as the constant 1.
fn name_record(flags: u32, name: &str) -> Vec<u8> {
    let mut data = flags.to_le_bytes().to_vec();
    data.push(0);
    data.extend_from_slice(&u32::MAX.to_le_bytes());
    data.extend_from_slice(&wide_string(name));
    data.extend_from_slice(
        &litchi_xlsb::ParsedFormula {
            rgce: vec![litchi_xlsb::ptg_types::PTG_INT, 1, 0],
            rgcb: Vec::new(),
        }
        .to_bytes()
        .unwrap(),
    );
    data.extend_from_slice(&u32::MAX.to_le_bytes());
    data
}

/// A `BrtFmt` payload with identifier `id` and format code `code`.
fn fmt_record(id: u16, code: &str) -> Vec<u8> {
    let mut data = id.to_le_bytes().to_vec();
    data.extend_from_slice(&wide_string(code));
    data
}

/// The single repair a fixture must report, as (class, record type, observed).
fn assert_single(
    report: &ToleranceReport,
    defect: impl Into<Defect>,
    record_type: Kind,
    observed: u32,
) {
    let defect = defect.into();
    assert_eq!(report.total(), 1, "{:?}", report.defects());
    let entry = report.defects()[0];
    assert_eq!(entry.defect(), defect);
    assert_eq!(entry.record_type(), record_type.get());
    assert_eq!(entry.observed(), observed);
}

/// Cell values and row/column metadata a repaired worksheet must keep.
fn sheet_projection(workbook: &Workbook) -> String {
    let worksheet = workbook.worksheet(0).unwrap();
    let mut cells = Vec::new();
    for row in 0..16 {
        for column in 0..16 {
            if let Some(cell) = worksheet.get_cell(row, column) {
                cells.push(format!("{cell:?}"));
            }
        }
    }
    format!(
        "{cells:?}\n{:?}\n{:?}",
        worksheet.column_infos(),
        worksheet.row_infos()
    )
}

fn control() -> Workbook {
    Workbook::new(Cursor::new(fixture())).unwrap()
}

#[test]
fn the_unmodified_control_repairs_nothing() {
    let workbook = control();
    assert!(workbook.tolerance_report().is_clean());
    for index in 0..workbook.worksheet_count() {
        assert!(
            workbook
                .worksheet(index)
                .unwrap()
                .tolerance_report()
                .is_clean()
        );
        assert!(
            workbook
                .cell_values(index)
                .unwrap()
                .tolerance_report()
                .is_clean()
        );
    }
    assert!(
        Package::from_bytes(fixture())
            .unwrap()
            .tolerance_report()
            .is_clean()
    );
}

/// Open a sheet-1 fixture on every path and return the eager worksheet report,
/// after checking the repair kept every value and the other paths agree.
fn open_sheet_fixture(bytes: &[u8], snapshot_sees_it: bool) -> ToleranceReport {
    let workbook = Workbook::new(Cursor::new(bytes.to_vec())).unwrap();
    assert!(workbook.tolerance_report().is_clean());
    assert_eq!(sheet_projection(&workbook), sheet_projection(&control()));
    let report = workbook.worksheet(0).unwrap().tolerance_report().clone();

    let snapshot = workbook.cell_values(0).unwrap();
    if snapshot_sees_it {
        assert_eq!(snapshot.tolerance_report(), &report);
    } else {
        assert!(snapshot.tolerance_report().is_clean());
    }

    let source = SourceBackedWorkbook::from_reader(Cursor::new(bytes.to_vec())).unwrap();
    let materialized = source
        .worksheet_by_index(0)
        .unwrap()
        .unwrap()
        .materialize()
        .unwrap();
    assert_eq!(materialized.tolerance_report(), &report);
    assert!(source.tolerance_report().is_clean());
    report
}

#[test]
fn reserved_cell_flag_bits_are_masked_and_recorded() {
    let bytes = rewrite(&fixture(), SHEET1, |stream| {
        first_payload(stream, kind::CELL_RK)[7] |= 0x80;
    });
    let report = open_sheet_fixture(&bytes, true);
    assert_single(&report, ReservedFieldDefect::CellFlags, kind::CELL_RK, 0x80);
}

#[test]
fn reserved_row_header_flag_bits_are_masked_and_recorded() {
    let bytes = rewrite(&fixture(), SHEET1, |stream| {
        first_payload(stream, kind::ROW_HDR)[10] |= 0x10;
    });
    let report = open_sheet_fixture(&bytes, false);
    assert_single(
        &report,
        ReservedFieldDefect::RowHeaderFlags,
        kind::ROW_HDR,
        0x10,
    );
}

#[test]
fn reserved_column_info_flag_bits_are_masked_and_recorded() {
    let bytes = rewrite(&fixture(), SHEET1, |stream| {
        first_payload(stream, kind::COL_INFO)[16] |= 0x40;
    });
    let report = open_sheet_fixture(&bytes, false);
    assert_single(
        &report,
        ReservedFieldDefect::ColumnInfoFlags,
        kind::COL_INFO,
        0x40,
    );
}

#[test]
fn reserved_formula_flag_bits_are_masked_and_recorded() {
    let bytes = rewrite(&fixture(), SHEET1, |stream| {
        first_payload(stream, kind::FMLA_NUM)[16] |= 0x01;
    });
    let report = open_sheet_fixture(&bytes, true);
    assert_single(
        &report,
        ReservedFieldDefect::FormulaFlags,
        kind::FMLA_NUM,
        0x01,
    );
}

#[test]
fn trailing_formula_cell_bytes_are_ignored_and_recorded() {
    let bytes = rewrite(&fixture(), SHEET1, |stream| {
        first_payload(stream, kind::FMLA_NUM).extend_from_slice(&[0xAB, 0xCD]);
    });
    let report = open_sheet_fixture(&bytes, true);
    assert_single(&report, LayoutDefect::TrailingBytes, kind::FMLA_NUM, 2);
}

/// Open a workbook-stream or styles fixture on every path and return the
/// workbook report, checking the paths agree.
fn open_workbook_fixture(bytes: &[u8]) -> ToleranceReport {
    let workbook = Workbook::new(Cursor::new(bytes.to_vec())).unwrap();
    let report = workbook.tolerance_report().clone();
    assert_eq!(
        Package::from_bytes(bytes.to_vec())
            .unwrap()
            .tolerance_report(),
        &report
    );
    for index in 0..workbook.worksheet_count() {
        assert!(
            workbook
                .worksheet(index)
                .unwrap()
                .tolerance_report()
                .is_clean()
        );
    }

    // The source-backed catalog parses the workbook stream at open and the
    // styles part once a worksheet materialization needs it.
    let source = SourceBackedWorkbook::from_reader(Cursor::new(bytes.to_vec())).unwrap();
    source
        .worksheet_by_index(0)
        .unwrap()
        .unwrap()
        .materialize()
        .unwrap();
    assert_eq!(source.tolerance_report(), report);
    // A second materialization reuses the cached styles: nothing is recorded twice.
    source
        .worksheet_by_index(0)
        .unwrap()
        .unwrap()
        .materialize()
        .unwrap();
    assert_eq!(source.tolerance_report(), report);
    report
}

#[test]
fn reserved_calc_option_bits_are_masked_and_recorded() {
    let bytes = rewrite(&fixture(), WORKBOOK, |stream| {
        first_payload(stream, kind::CALC_PROP)[25] |= 0x80;
    });
    let report = open_workbook_fixture(&bytes);
    assert_single(
        &report,
        ReservedFieldDefect::CalcOptions,
        kind::CALC_PROP,
        0x8000,
    );
    let workbook = Workbook::new(Cursor::new(bytes)).unwrap();
    assert_eq!(workbook.calc(), control().calc());
}

#[test]
fn trailing_calc_property_bytes_are_ignored_and_recorded() {
    let bytes = rewrite(&fixture(), WORKBOOK, |stream| {
        first_payload(stream, kind::CALC_PROP).extend_from_slice(&[0; 4]);
    });
    let report = open_workbook_fixture(&bytes);
    assert_single(&report, LayoutDefect::TrailingBytes, kind::CALC_PROP, 4);
}

#[test]
fn trailing_bundle_sheet_bytes_are_ignored_and_recorded() {
    let bytes = rewrite(&fixture(), WORKBOOK, |stream| {
        first_payload(stream, kind::BUNDLE_SH).push(0);
    });
    let report = open_workbook_fixture(&bytes);
    assert_single(&report, LayoutDefect::TrailingBytes, kind::BUNDLE_SH, 1);
    let workbook = Workbook::new(Cursor::new(bytes)).unwrap();
    assert_eq!(workbook.worksheet_names(), control().worksheet_names());
}

#[test]
fn reserved_name_flags_and_trailing_name_bytes_are_repaired_and_recorded() {
    let mut name = name_record(0x8000_0000, "Constant");
    name.extend_from_slice(&[1, 2, 3]);
    let bytes = rewrite(&fixture(), WORKBOOK, |stream| {
        let at = position_after_last(stream, kind::END_BUNDLE_SHS);
        stream.insert(at, (kind::NAME, name));
    });
    let report = open_workbook_fixture(&bytes);
    assert_eq!(report.total(), 2, "{:?}", report.defects());
    assert_eq!(
        report.defects()[0].defect(),
        Defect::ReservedField(ReservedFieldDefect::NameFlags)
    );
    assert_eq!(report.defects()[0].observed(), 0x8000_0000);
    assert_eq!(
        report.defects()[1].defect(),
        Defect::Layout(LayoutDefect::TrailingBytes)
    );
    assert_eq!(report.defects()[1].record_type(), kind::NAME.get());
    assert_eq!(report.defects()[1].observed(), 3);
    let workbook = Workbook::new(Cursor::new(bytes)).unwrap();
    assert_eq!(workbook.defined_names(), ["Constant".to_string()]);
}

#[test]
fn reserved_border_flag_bits_are_masked_and_recorded() {
    let bytes = rewrite(&fixture(), STYLES, |stream| {
        first_payload(stream, kind::BORDER)[0] |= 0x80;
    });
    let report = open_workbook_fixture(&bytes);
    assert_single(
        &report,
        ReservedFieldDefect::BorderFlags,
        kind::BORDER,
        0x80,
    );
    let workbook = Workbook::new(Cursor::new(bytes)).unwrap();
    let border = workbook.styles().get_border(0).unwrap();
    assert!(!border.diagonal_down && !border.diagonal_up);
}

#[test]
fn a_nonzero_border_side_reserved_byte_is_ignored_and_recorded() {
    let bytes = rewrite(&fixture(), STYLES, |stream| {
        // The top side's Blxf starts at offset 1; its reserved byte follows
        // the line style.
        first_payload(stream, kind::BORDER)[2] = 0x5A;
    });
    let report = open_workbook_fixture(&bytes);
    assert_single(
        &report,
        ReservedFieldDefect::BorderSideReserved,
        kind::BORDER,
        0x5A,
    );
}

#[test]
fn an_undefined_border_line_style_reads_as_no_border_and_is_recorded() {
    let bytes = rewrite(&fixture(), STYLES, |stream| {
        first_payload(stream, kind::BORDER)[1] = 0x20;
    });
    let report = open_workbook_fixture(&bytes);
    assert_single(&report, LayoutDefect::BorderStyle, kind::BORDER, 0x20);
    let workbook = Workbook::new(Cursor::new(bytes)).unwrap();
    assert!(workbook.styles().get_border(0).unwrap().top.is_none());
}

#[test]
fn a_direct_color_without_valid_rgb_reads_as_no_color_and_is_recorded() {
    let bytes = rewrite(&fixture(), STYLES, |stream| {
        // BrtFont's BrtColor begins at offset 12: xColorType = 2 (RGB) with
        // fValidRGB clear.
        first_payload(stream, kind::FONT)[12] = 0x04;
    });
    let report = open_workbook_fixture(&bytes);
    assert_single(
        &report,
        LayoutDefect::ColorWithoutValidRgb,
        kind::FONT,
        0x04,
    );
    let workbook = Workbook::new(Cursor::new(bytes)).unwrap();
    assert_eq!(workbook.styles().get_font(0).unwrap().color, None);
}

#[test]
fn out_of_spec_number_formats_are_kept_verbatim_and_recorded() {
    let mut trailing = fmt_record(170, "0.000");
    trailing.push(0);
    let bytes = rewrite(&fixture(), STYLES, |stream| {
        stream.insert(0, (kind::BEGIN_FMTS, 3u32.to_le_bytes().to_vec()));
        stream.insert(1, (kind::FMT, fmt_record(100, "0.0")));
        stream.insert(2, (kind::FMT, fmt_record(171, "")));
        stream.insert(3, (kind::FMT, trailing));
        stream.insert(4, (kind::END_FMTS, Vec::new()));
    });
    let report = open_workbook_fixture(&bytes);
    let entries: Vec<_> = report
        .iter()
        .map(|entry| (entry.defect(), entry.location(), entry.observed()))
        .collect();
    assert_eq!(
        entries,
        [
            (Defect::Layout(LayoutDefect::NumberFormatId), 0, 100),
            (Defect::Layout(LayoutDefect::NumberFormatCodeLength), 1, 0),
            (Defect::Layout(LayoutDefect::TrailingBytes), 2, 1),
        ]
    );
    let workbook = Workbook::new(Cursor::new(bytes)).unwrap();
    assert_eq!(workbook.styles().get_num_fmt(100), Some("0.0"));
    assert_eq!(workbook.styles().get_num_fmt(171), Some(""));
    assert_eq!(workbook.styles().get_num_fmt(170), Some("0.000"));
}

#[test]
fn standalone_codec_functions_still_refuse_what_they_cannot_report() {
    let mut calc = None;
    let mut border = None;
    let bytes = fixture();
    let package = OpcPackage::from_reader(Cursor::new(bytes)).unwrap();
    for record in Records::new(
        package
            .get_part(&PackURI::new(WORKBOOK).unwrap())
            .unwrap()
            .blob(),
    ) {
        let record = record.unwrap();
        if record.kind() == kind::CALC_PROP {
            calc = Some(record.payload().to_vec());
        }
    }
    for record in Records::new(
        package
            .get_part(&PackURI::new(STYLES).unwrap())
            .unwrap()
            .blob(),
    ) {
        let record = record.unwrap();
        if record.kind() == kind::BORDER {
            border = Some(record.payload().to_vec());
        }
    }
    let mut calc = calc.unwrap();
    litchi_xlsb::calc::read(&calc).unwrap();
    calc[25] |= 0x80;
    assert!(litchi_xlsb::calc::read(&calc).is_err());

    let mut border = border.unwrap();
    litchi_xlsb::styles::parse_border(&border).unwrap();
    border[0] |= 0x80;
    assert!(litchi_xlsb::styles::parse_border(&border).is_err());
}

#[test]
fn truncated_records_remain_fatal() {
    // Record framing is never repaired: a cell shorter than its Cell header
    // still fails the worksheet.
    let bytes = rewrite(&fixture(), SHEET1, |stream| {
        first_payload(stream, kind::CELL_BLANK).truncate(6);
    });
    let workbook = Workbook::new(Cursor::new(bytes)).unwrap();
    assert!(workbook.worksheet(0).is_err());

    let bytes = rewrite(&fixture(), STYLES, |stream| {
        first_payload(stream, kind::BORDER).truncate(50);
    });
    assert!(Workbook::new(Cursor::new(bytes)).is_err());
}

#[test]
fn many_defects_stay_bounded() {
    // Every cell and row header carries a reserved bit; the report stays
    // within its bound while counting every repair.
    let bytes = rewrite(&fixture(), SHEET1, |stream| {
        for (record_type, payload) in stream.iter_mut() {
            if (kind::CELL_BLANK.get()..=kind::FMLA_ERROR.get()).contains(&record_type.get()) {
                payload[7] |= 0x02;
            }
        }
    });
    let workbook = Workbook::new(Cursor::new(bytes)).unwrap();
    let worksheet = workbook.worksheet(0).unwrap();
    let report = worksheet.tolerance_report();
    assert!(report.total() > 1);
    assert_eq!(
        report.count(ReservedFieldDefect::CellFlags) as u64,
        report.total()
    );
    assert!(report.defects().len() <= ToleranceReport::RECORD_LIMIT);
}
