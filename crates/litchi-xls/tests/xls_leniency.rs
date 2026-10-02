//! End-to-end coverage for [`OpenOptions::leniency`].
//!
//! Every case opens a real third-party workbook twice: once strictly, proving
//! the defect is still rejected by default, and once leniently, proving the
//! workbook opens *and* that the repair is enumerable rather than silent.

use std::path::PathBuf;

use litchi_xls::{Defect, FormattingDefect, Leniency, OpenOptions, ToleranceReport, Workbook};

fn poi_fixture(name: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/poi/test-data/spreadsheet")
        .join(name)
}

fn libreoffice_fixture(name: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/libreoffice-core/sc/qa/unit/data/xls")
        .join(name)
}

fn lenient_options() -> OpenOptions<'static> {
    OpenOptions::new().with_leniency(Leniency::TolerateFormattingDefects)
}

fn open_lenient(path: PathBuf) -> Workbook<std::fs::File> {
    let file = std::fs::File::open(&path).expect("fixture is readable");
    Workbook::new_with_options(file, lenient_options())
        .expect("a lenient open tolerates cosmetic formatting defects")
}

fn strict_error(path: PathBuf) -> String {
    let file = std::fs::File::open(&path).expect("fixture is readable");
    match Workbook::new(file) {
        Ok(_) => panic!("a strict open must still reject {}", path.display()),
        Err(error) => error.to_string(),
    }
}

/// Assert that `path` is rejected by default, opens leniently, and reports
/// exactly `defect` as the reason. Returns the first recorded entry.
fn assert_repaired(
    path: PathBuf,
    defect: FormattingDefect,
    strict_needle: &str,
) -> litchi_xls::ToleratedDefect {
    let message = strict_error(path.clone());
    assert!(
        message.contains(strict_needle),
        "unexpected strict error for {}: {message}",
        path.display()
    );

    let workbook = open_lenient(path);
    let report = workbook.tolerance_report();
    assert!(!report.is_clean());
    assert_eq!(report.unrecorded(), 0);
    assert_eq!(report.total(), report.defects().len() as u64);
    assert_eq!(
        report.count(defect),
        report.defects().len(),
        "the fixture must isolate a single defect class"
    );
    // The workbook is genuinely usable, not merely constructible.
    assert!(!workbook.sheets().is_empty());

    let entry = report.defects()[0];
    assert_eq!(entry.defect(), Defect::Formatting(defect));
    assert_eq!(entry.record_type(), Some(defect.record_type()));
    entry
}

#[test]
fn strict_open_is_the_default_for_every_tolerable_defect() {
    // `Workbook::new` and a default-constructed options value must agree,
    // and both must reject what a lenient open would repair.
    let path = poi_fixture("29942.xls");
    let explicit = std::fs::File::open(&path).expect("fixture is readable");
    assert!(Workbook::new_with_options(explicit, OpenOptions::default()).is_err());
    assert!(strict_error(path).contains("Font family"));
}

#[test]
fn lenient_open_repairs_an_out_of_range_font_family() {
    let entry = assert_repaired(
        poi_fixture("29942.xls"),
        FormattingDefect::FontFamily,
        "Font family",
    );
    // The repair substitutes `NotApplicable`; the offending byte survives in
    // the report so a caller can tell exactly what the file claimed.
    assert!(entry.observed() > 5);
}

#[test]
fn lenient_open_repairs_an_empty_font_name() {
    let entry = assert_repaired(
        libreoffice_fixture("tdf170189.xls"),
        FormattingDefect::FontNameEmpty,
        "Font name has 0 characters",
    );
    assert_eq!(entry.observed(), 0);
}

#[test]
fn lenient_open_repairs_an_xfcrc_count_disagreement() {
    let entry = assert_repaired(
        poi_fixture("SharedFormulaTest.xls"),
        FormattingDefect::ExtendedFormatCountMismatch,
        "XFCRC declares",
    );
    // `ordinal` is the number of XF records actually parsed, `observed` the
    // count XFCRC claimed; the two must disagree or there was no defect.
    assert_ne!(entry.ordinal(), entry.observed());
}

#[test]
fn a_conforming_workbook_reports_nothing_under_either_policy() {
    for options in [OpenOptions::default(), lenient_options()] {
        let file = std::fs::File::open(poi_fixture("colwidth.xls")).expect("fixture is readable");
        let workbook =
            Workbook::new_with_options(file, options).expect("a conforming workbook opens");
        assert!(workbook.tolerance_report().is_clean());
        assert_eq!(workbook.tolerance_report().total(), 0);
    }
}

#[test]
fn a_defect_free_report_is_the_default_state() {
    let report = ToleranceReport::default();
    assert!(report.is_clean());
    assert_eq!(report.defects(), &[]);
    assert_eq!(report.count(FormattingDefect::FontNameEmpty), 0);
    assert_eq!(report.total(), 0);
}

// ---------------------------------------------------------------------------
// Excel-compatible repairs: applied and recorded under the *default* policy
// ---------------------------------------------------------------------------

use std::io::Cursor;

use litchi_cfb::{OleFile, OleWriter};
use litchi_core::sheet::{Cell as SheetCell, CellValue};
use litchi_xls::{CellFormatDefect, ExtendedFormat, Writer};

/// BIFF `Number` record type.
const NUMBER_RECORD: u16 = 0x0203;

/// A one-cell workbook built by the crate writer, with the cell's `ixfe`
/// rewritten to `ixfe` — one of the most common non-Microsoft-producer
/// defects in the wild.
fn workbook_with_cell_ixfe(ixfe: u16) -> Vec<u8> {
    let mut writer = Writer::new();
    let sheet = writer.add_worksheet("Sheet1").unwrap();
    writer.write_number(sheet, 0, 0, 42.5).unwrap();
    let mut bytes = Vec::new();
    writer.write_to(&mut Cursor::new(&mut bytes)).unwrap();

    let mut stream = OleFile::open(Cursor::new(&bytes))
        .unwrap()
        .open_stream(&["Workbook"])
        .unwrap();
    let mut cursor = 0;
    let mut patched = false;
    while cursor + 4 <= stream.len() {
        let kind = u16::from_le_bytes([stream[cursor], stream[cursor + 1]]);
        let length = usize::from(u16::from_le_bytes([stream[cursor + 2], stream[cursor + 3]]));
        if kind == NUMBER_RECORD {
            let payload = cursor + 4;
            stream[payload + 4..payload + 6].copy_from_slice(&ixfe.to_le_bytes());
            patched = true;
            break;
        }
        cursor += 4 + length;
    }
    assert!(patched, "the writer workbook must contain a Number record");

    let mut writer = OleWriter::new();
    writer.create_stream(&["Workbook"], &stream).unwrap();
    let mut package = Vec::new();
    writer.write_to(&mut Cursor::new(&mut package)).unwrap();
    package
}

/// A cell whose `ixfe` cannot format a cell opens under the default policy:
/// the value is untouched, the format falls back to the default cell XF, and
/// the repair is recorded rather than silent.
fn assert_default_open_repairs_cell_xf(ixfe: u16) {
    let workbook = Workbook::new(Cursor::new(workbook_with_cell_ixfe(ixfe)))
        .expect("the default policy applies the Excel-compatible repair");

    let worksheet = workbook.xls_worksheet(0).unwrap();
    let cell = worksheet.get_cell(0, 0).unwrap();
    // The value is never at risk...
    assert_eq!(cell.value(), &CellValue::Float(42.5));
    // ...the record's own claim is preserved raw...
    assert_eq!(cell.xf_index(), ixfe);
    // ...and the resolved format is the default cell XF — the very format the
    // writer gave this cell before the mutation, since it writes ixfe 15.
    let clean = Workbook::new(Cursor::new({
        let mut writer = Writer::new();
        let sheet = writer.add_worksheet("Sheet1").unwrap();
        writer.write_number(sheet, 0, 0, 42.5).unwrap();
        let mut bytes = Vec::new();
        writer.write_to(&mut Cursor::new(&mut bytes)).unwrap();
        bytes
    }))
    .unwrap();
    let format = worksheet.format_for_cell(0, 0).expect("substituted format");
    assert_eq!(
        format.number_format_id(),
        clean
            .xls_worksheet(0)
            .unwrap()
            .format_for_cell(0, 0)
            .map(ExtendedFormat::number_format_id)
            .unwrap(),
    );

    let report = workbook.tolerance_report();
    assert_eq!(report.unrecorded(), 0);
    assert_eq!(report.total(), 1);
    let entry = report.defects()[0];
    assert_eq!(
        entry.defect(),
        Defect::CellFormat(CellFormatDefect::XfReference)
    );
    assert_eq!(entry.ordinal(), 0, "the cell's row");
    assert_eq!(entry.observed(), u32::from(ixfe));
    assert_eq!(entry.record_type(), None);
}

#[test]
fn a_default_open_repairs_an_out_of_range_cell_xf_reference() {
    assert_default_open_repairs_cell_xf(4095);
}

#[test]
fn a_default_open_repairs_a_reserved_style_slot_cell_xf_reference() {
    assert_default_open_repairs_cell_xf(3);
}

#[test]
fn a_conforming_writer_workbook_reports_nothing() {
    let mut writer = Writer::new();
    let sheet = writer.add_worksheet("Sheet1").unwrap();
    writer.write_number(sheet, 0, 0, 42.5).unwrap();
    let mut bytes = Vec::new();
    writer.write_to(&mut Cursor::new(&mut bytes)).unwrap();
    let workbook = Workbook::new(Cursor::new(bytes)).unwrap();
    assert!(workbook.tolerance_report().is_clean());
}
