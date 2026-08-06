//! Regression coverage for the layered BIFF8 number-format owner.

use super::codec::{
    is_custom_date_time, parse_date_system,
    parse_number_format as parse_number_format_with_tolerance, parse_xf as parse_xf_with_tolerance,
    parse_xfcrc,
};
use super::model::*;
use super::{
    DATE1904_RECORD, FORMAT_RECORD, MIN_XF_RECORDS, XF_RECORD, XFCRC_RECORD,
    XL_UNICODE_STRING_HIGH_BYTE,
};
use crate::error::Result;
use crate::leniency::{FormattingDefect, ToleranceLog};

/// Strict-mode shim used by both test modules: the production reader threads a
/// tolerance log, but these tests exercise the default reject-everything policy.
fn parse_globals_strict(records: &[litchi_biff::Record]) -> Result<Formatting> {
    let refs = records
        .iter()
        .map(|record| record.as_ref())
        .collect::<Vec<_>>();
    Formatting::parse_globals(
        &refs,
        &mut ToleranceLog::new(crate::leniency::Leniency::Strict),
    )
}

use litchi_biff::{Encoder, Kind, Record as Frame};

/// Strict-mode shim: the production reader threads a tolerance log, but
/// these tests exercise the default (reject-everything) policy.
fn parse_xf(data: &[u8], index: u16) -> Result<ExtendedFormat> {
    parse_xf_with_tolerance(
        data,
        index,
        &mut ToleranceLog::new(crate::leniency::Leniency::Strict),
    )
}

/// Strict-mode shim for [`super::parse_number_format`].
fn parse_number_format(data: &[u8]) -> Result<NumberFormat> {
    parse_number_format_with_tolerance(
        data,
        0,
        &mut ToleranceLog::new(crate::leniency::Leniency::Strict),
    )
}

fn semantic_xf(style: bool, parent: u16, application_bits: u8) -> [u8; 20] {
    let mut data = [0; 20];
    let flags = (parent << 4) | if style { 0x0004 } else { 0 };
    data[4..6].copy_from_slice(&flags.to_le_bytes());
    data[9] = application_bits << 2;
    data
}

#[test]
fn decodes_cell_apply_polarity_and_style_local_semantics() {
    let cell = parse_xf(&semantic_xf(false, 0, 0b10_1010), 1).unwrap();
    let apply = cell.applications();
    assert!(apply.inherits_number_format());
    assert!(apply.applies_font());
    assert!(apply.inherits_alignment());
    assert!(apply.applies_border());
    assert!(apply.inherits_fill());
    assert!(apply.applies_protection());

    let style = parse_xf(&semantic_xf(true, 0x0fff, 0), 0).unwrap();
    let apply = style.applications();
    assert!(apply.applies_number_format());
    assert!(apply.applies_font());
    assert!(apply.applies_alignment());
    assert!(apply.applies_border());
    assert!(apply.applies_fill());
    assert!(apply.applies_protection());
}

#[test]
fn decodes_cell_special_flags_and_ignores_style_reserved_overlays() {
    let mut cell_data = semantic_xf(false, 0, 0);
    cell_data[4] |= 0x08;
    let mut border2 = 1u32 << 25;
    cell_data[14..18].copy_from_slice(&border2.to_le_bytes());
    cell_data[18..20].copy_from_slice(&(1u16 << 14).to_le_bytes());
    let cell = parse_xf(&cell_data, 1).unwrap();
    assert!(cell.quote_prefix());
    assert!(cell.has_xf_extension());
    assert!(cell.pivot_button());

    let mut style_data = semantic_xf(true, 0x0fff, 0x3f);
    border2 |= 1 << 25;
    style_data[14..18].copy_from_slice(&border2.to_le_bytes());
    style_data[18..20].copy_from_slice(&(3u16 << 14).to_le_bytes());
    let style = parse_xf(&style_data, 0).unwrap();
    assert!(!style.has_xf_extension());
    assert!(!style.pivot_button());
    assert!(style.applications().applies_fill());
}

#[test]
fn resolves_effective_components_by_borrowing_parent_or_cell() {
    let mut style_data = semantic_xf(true, 0x0fff, 0);
    style_data[2..4].copy_from_slice(&14u16.to_le_bytes());
    let style = parse_xf(&style_data, 0).unwrap();
    let mut cell_data = semantic_xf(false, 0, 0b00_0010);
    cell_data[0..2].copy_from_slice(&1u16.to_le_bytes());
    cell_data[2..4].copy_from_slice(&1u16.to_le_bytes());
    let cell = parse_xf(&cell_data, 1).unwrap();
    let formatting = Formatting {
        extended_formats: vec![style, cell],
        ..Formatting::default()
    };

    let effective = formatting.effective_extended_format(1).unwrap();
    assert_eq!(effective.parent_style().unwrap().index(), 0);
    assert_eq!(effective.number_format_id(), 14);
    assert_eq!(effective.number_format_source().index(), 0);
    assert_eq!(effective.font_index(), 1);
    assert_eq!(effective.font_source().index(), 1);
    assert!(std::ptr::eq(
        effective.alignment(),
        formatting.extended_formats()[0].alignment(),
    ));
}
use crate::cell::Cell;
use crate::records::CellRecord;
use crate::workbook::Workbook;
use litchi_core::sheet::{Cell as SheetCell, CellValue, Worksheet};
use std::fs::{self, File};
use std::io::Cursor;
use std::path::{Path, PathBuf};

fn record(record_type: u16, data: Vec<u8>) -> Frame {
    let mut encoder = Encoder::new();
    encoder
        .push(Kind::from_wire(record_type), &data)
        .expect("test frame fits the BIFF wire limit");
    Frame::open(encoder.finish()).expect("test frame is complete")
}

fn xf(style: bool, parent: u16, format_id: u16) -> Frame {
    let mut data = vec![0u8; 20];
    data[2..4].copy_from_slice(&format_id.to_le_bytes());
    let flags = (parent << 4) | u16::from(style) << 2 | 1;
    data[4..6].copy_from_slice(&flags.to_le_bytes());
    record(XF_RECORD, data)
}

fn format_record(id: u16, code: &str) -> Frame {
    let mut data = Vec::new();
    data.extend_from_slice(&id.to_le_bytes());
    data.extend_from_slice(&(code.len() as u16).to_le_bytes());
    data.push(0);
    data.extend_from_slice(code.as_bytes());
    record(FORMAT_RECORD, data)
}

fn fixture(relative: &str) -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../..")
        .join(relative)
}

#[test]
fn recognizes_date_formats_without_literal_false_positives() {
    for code in ["yyyy-mm-dd", "h:mm AM/PM", "[hh]:mm:ss", "dd mmmm"] {
        assert!(is_custom_date_time(code), "{code}");
    }
    for code in ["0.00E+00", "0.00 \"m\"", "0.0\\m", "[Red]0.00", "General"] {
        assert!(!is_custom_date_time(code), "{code}");
    }
}

#[test]
fn parses_strict_unicode_format_records() {
    let mut data = vec![164, 0, 4, 0, 1];
    for unit in "日期mm".encode_utf16() {
        data.extend_from_slice(&unit.to_le_bytes());
    }
    let format = parse_number_format(&data).unwrap();
    assert_eq!(format.id(), 164);
    assert_eq!(format.code(), "日期mm");
    assert!(format.is_date_time());

    data.push(0);
    assert!(parse_number_format(&data).is_err());
}

/// Build a `Format` payload whose `cch` claims `declared` characters while
/// only `present` characters follow.
fn overlong_format_record(declared: u16, present: &str) -> Vec<u8> {
    let mut data = vec![164, 0];
    data.extend_from_slice(&declared.to_le_bytes());
    data.push(XL_UNICODE_STRING_HIGH_BYTE);
    for unit in present.encode_utf16() {
        data.extend_from_slice(&unit.to_le_bytes());
    }
    data
}

#[test]
fn a_lenient_policy_truncates_a_format_string_that_overstates_its_payload() {
    const DECLARED: u16 = 40;
    const ORDINAL: u32 = 3;
    let data = overlong_format_record(DECLARED, "0.00");

    let mut tolerance = ToleranceLog::new(crate::leniency::Leniency::TolerateFormattingDefects);
    let format = parse_number_format_with_tolerance(&data, ORDINAL, &mut tolerance)
        .expect("a lenient policy decodes the characters that are present");
    assert_eq!(format.code(), "0.00");

    let report = tolerance.into_report();
    assert_eq!(report.count(FormattingDefect::FormatStringOverrun), 1);
    assert_eq!(report.defects()[0].ordinal(), ORDINAL);
    assert_eq!(report.defects()[0].observed(), u32::from(DECLARED));
    assert_eq!(report.defects()[0].record_type(), FORMAT_RECORD);
}

#[test]
fn a_truncated_format_string_drops_a_split_utf16_code_unit() {
    // One trailing byte cannot form a UTF-16 code unit; the repair must
    // discard it rather than decode half a character.
    let mut data = overlong_format_record(40, "0.00");
    data.push(0x30);

    let mut tolerance = ToleranceLog::new(crate::leniency::Leniency::TolerateFormattingDefects);
    let format = parse_number_format_with_tolerance(&data, 0, &mut tolerance)
        .expect("a lenient policy keeps only whole characters");
    assert_eq!(format.code(), "0.00");
    assert_eq!(
        tolerance
            .into_report()
            .count(FormattingDefect::FormatStringOverrun),
        1
    );
}

#[test]
fn a_lenient_policy_still_rejects_trailing_bytes_and_a_bad_format_identifier() {
    let mut tolerance = ToleranceLog::new(crate::leniency::Leniency::TolerateFormattingDefects);
    // Trailing bytes past a satisfied `cch` are a framing defect.
    let mut trailing = overlong_format_record(4, "0.00");
    trailing.push(0);
    trailing.push(0);
    assert!(parse_number_format_with_tolerance(&trailing, 0, &mut tolerance).is_err());
    // An identifier outside the permitted ranges is not a formatting defect.
    let mut bad_id = overlong_format_record(4, "0.00");
    bad_id[0..2].copy_from_slice(&0u16.to_le_bytes());
    assert!(parse_number_format_with_tolerance(&bad_id, 0, &mut tolerance).is_err());
    assert!(tolerance.into_report().is_clean());
}

#[test]
fn rejects_invalid_date_xf_and_crc_shapes() {
    assert!(parse_date_system(&[2, 0]).is_err());
    assert!(parse_date_system(&[0]).is_err());
    assert!(parse_xf(&[0; 19], 0).is_err());

    let mut crc = [0u8; 20];
    crc[0..2].copy_from_slice(&XFCRC_RECORD.to_le_bytes());
    crc[14..16].copy_from_slice(&16u16.to_le_bytes());
    assert_eq!(parse_xfcrc(&crc).unwrap(), 16);
    crc[2] = 1;
    assert!(parse_xfcrc(&crc).is_err());
}

#[test]
fn classifies_numeric_and_formula_caches_without_literal_false_positives() {
    let mut records = vec![record(DATE1904_RECORD, vec![0, 0])];
    records.push(format_record(164, "yyyy-mm-dd"));
    records.push(format_record(165, "0.00 \"m\""));
    for _ in 0..15 {
        records.push(xf(true, 0x0fff, 0));
    }
    records.push(xf(false, 0, 14));
    records.push(xf(false, 0, 164));
    records.push(xf(false, 0, 165));
    let formatting = parse_globals_strict(&records).unwrap();

    let builtin = CellRecord::Number {
        row: 0,
        col: 0,
        xf_index: 15,
        value: 39_304.0,
    };
    let custom = CellRecord::Formula {
        row: 0,
        col: 1,
        xf_index: 16,
        value: crate::records::FormulaValue::Number(45_000.5),
        metadata: crate::FormulaMetadata::default(),
        formula: vec![0x1e, 1, 0],
    };
    let literal = CellRecord::Number {
        row: 0,
        col: 2,
        xf_index: 17,
        value: 12.5,
    };

    let builtin =
        Cell::from_record_with_formula_context(&builtin, None, None, Some(&formatting)).unwrap();
    let custom =
        Cell::from_record_with_formula_context(&custom, None, None, Some(&formatting)).unwrap();
    let literal =
        Cell::from_record_with_formula_context(&literal, None, None, Some(&formatting)).unwrap();
    assert_eq!(builtin.value(), &CellValue::DateTime(39_304.0));
    assert_eq!(custom.value(), &CellValue::DateTime(45_000.5));
    assert_eq!(literal.value(), &CellValue::Float(12.5));
    assert_eq!(literal.xf_index(), 17);
}

#[test]
fn opens_poi_date_and_epoch_fixtures() {
    let dates = Workbook::new(
        File::open(fixture(
            "test-data/poi/test-data/spreadsheet/DateFormats.xls",
        ))
        .unwrap(),
    )
    .unwrap();
    let sheet = dates.xls_worksheet(0).unwrap();
    let mut found = false;
    for row in 0..sheet.row_count() as u32 {
        for col in 0..sheet.column_count() as u32 {
            if matches!(
                sheet.get_cell(row, col).map(SheetCell::value),
                Some(CellValue::DateTime(_))
            ) {
                found = true;
            }
        }
    }
    assert!(
        found,
        "POI DateFormats.xls contains no classified date cells"
    );

    for (path, expected) in [
        (
            "test-data/poi/test-data/spreadsheet/1900DateWindowing.xls",
            DateSystem::Excel1900,
        ),
        (
            "test-data/poi/test-data/spreadsheet/1904DateWindowing.xls",
            DateSystem::Excel1904,
        ),
    ] {
        let mut bytes = fs::read(fixture(path)).unwrap();
        assert!(Workbook::new(Cursor::new(bytes.clone())).is_err());

        // Both POI windowing fixtures declare sector 10 as their sole FAT
        // sector but mark FAT[10] ENDOFCHAIN rather than FATSECT. Normalize
        // that one proven MS-CFB defect in the test copy only; production
        // container validation remains strict.
        let sector_shift = u16::from_le_bytes([bytes[0x1e], bytes[0x1f]]);
        let sector_size = 1usize << sector_shift;
        let fat_sector = u32::from_le_bytes(bytes[0x4c..0x50].try_into().unwrap()) as usize;
        assert_eq!(u32::from_le_bytes(bytes[0x2c..0x30].try_into().unwrap()), 1);
        let marker_offset = (fat_sector + 1) * sector_size + fat_sector * 4;
        assert_eq!(
            u32::from_le_bytes(bytes[marker_offset..marker_offset + 4].try_into().unwrap()),
            0xffff_fffe
        );
        bytes[marker_offset..marker_offset + 4].copy_from_slice(&0xffff_fffdu32.to_le_bytes());

        assert_eq!(u32::from_le_bytes(bytes[0x30..0x34].try_into().unwrap()), 0);
        for sid in 5..=7usize {
            let entry = sector_size + sid * 128;
            assert_eq!(bytes[entry + 66], 0);
            assert_eq!(
                u16::from_le_bytes(bytes[entry + 64..entry + 66].try_into().unwrap()),
                2
            );
            assert_eq!(&bytes[entry..entry + 2], &[0, 0]);
            bytes[entry + 64..entry + 66].copy_from_slice(&0u16.to_le_bytes());
        }

        let workbook = Workbook::new(Cursor::new(bytes)).unwrap();
        assert_eq!(workbook.date_system(), expected, "{path}");
    }
}

#[test]
fn opens_libreoffice_format_fixtures_with_ordered_xfs() {
    for path in [
        "test-data/libreoffice-core/sc/qa/unit/data/xls/formats.xls",
        "test-data/libreoffice-core/sc/qa/unit/data/xls/cellformat.xls",
    ] {
        let workbook = Workbook::new(File::open(fixture(path)).unwrap()).unwrap();
        assert!(workbook.extended_formats().len() >= 16, "{path}");
        for (index, format) in workbook.extended_formats().iter().enumerate() {
            assert_eq!(format.index() as usize, index, "{path}");
        }
    }
}

mod real_world_tolerance_tests {
    use super::*;
    use litchi_biff::{Encoder, Kind, Record as Frame};

    fn record(record_type: u16, data: Vec<u8>) -> Frame {
        let mut encoder = Encoder::new();
        encoder
            .push(Kind::from_wire(record_type), &data)
            .expect("test frame fits the BIFF wire limit");
        Frame::open(encoder.finish()).expect("test frame is complete")
    }

    fn format_record(id: u16, code: &str) -> Frame {
        let mut data = id.to_le_bytes().to_vec();
        data.extend_from_slice(&(code.len() as u16).to_le_bytes());
        data.push(0); // compressed characters, no reserved bits
        data.extend_from_slice(code.as_bytes());
        record(FORMAT_RECORD, data)
    }

    /// `MIN_XF_RECORDS` style XFs plus one cell XF, the smallest set
    /// `parse_globals` accepts.
    fn minimal_xf_records() -> Vec<Frame> {
        fn xf(style: bool, parent: u16, format_id: u16) -> Frame {
            let mut data = vec![0u8; 20];
            data[2..4].copy_from_slice(&format_id.to_le_bytes());
            let flags = (parent << 4) | u16::from(style) << 2 | 1;
            data[4..6].copy_from_slice(&flags.to_le_bytes());
            record(XF_RECORD, data)
        }
        let mut records: Vec<Frame> = (0..MIN_XF_RECORDS - 1)
            .map(|_| xf(true, 0x0fff, 0))
            .collect();
        records.push(xf(false, 0, 164));
        records
    }

    /// MS-XLS 2.1 lists Date1904 in the globals grammar, but workbooks written
    /// by other producers omit it. The record selects between exactly two date
    /// systems, so the absent value is unambiguous and must not make the
    /// workbook unreadable.
    #[test]
    fn defaults_to_the_1900_date_system_when_date1904_is_absent() {
        let mut records = vec![format_record(164, "yyyy-mm-dd")];
        records.extend(minimal_xf_records());

        let formatting = parse_globals_strict(&records).expect("a missing Date1904 is not fatal");
        assert_eq!(formatting.date_system(), DateSystem::Excel1900);
    }

    /// An explicit record still wins over the fallback.
    #[test]
    fn honours_an_explicit_1904_date_system() {
        let mut records = vec![record(DATE1904_RECORD, vec![1, 0])];
        records.push(format_record(164, "yyyy-mm-dd"));
        records.extend(minimal_xf_records());

        let formatting = parse_globals_strict(&records).expect("explicit Date1904 parses");
        assert_eq!(formatting.date_system(), DateSystem::Excel1904);
    }

    /// MS-XLS 2.5.294 gives XLUnicodeString exactly one meaningful option bit;
    /// the rest are reserved and "MUST be zero, and MUST be ignored". A writer
    /// that leaves one set must not make the FORMAT record unreadable, and the
    /// bits must not be mistaken for the `fRichSt`/`fExtSt` of the unrelated
    /// XLUnicodeRichExtendedString, which would shift the character offset.
    #[test]
    fn ignores_reserved_option_bits_in_format_record_strings() {
        const CODE: &str = "0.00";
        for reserved in [0x02u8, 0x04, 0x08, 0x20, 0xf0] {
            let mut data = 164u16.to_le_bytes().to_vec();
            data.extend_from_slice(&(CODE.len() as u16).to_le_bytes());
            data.push(reserved); // fHighByte clear, reserved bits set
            data.extend_from_slice(CODE.as_bytes());

            let mut records = vec![
                record(DATE1904_RECORD, vec![0, 0]),
                record(FORMAT_RECORD, data),
            ];
            records.extend(minimal_xf_records());

            let formatting = parse_globals_strict(&records)
                .unwrap_or_else(|error| panic!("reserved bits {reserved:#04x} rejected: {error}"));
            assert_eq!(
                formatting.number_format(164).map(NumberFormat::code),
                Some(CODE),
                "reserved bits {reserved:#04x} must not shift the characters"
            );
        }
    }

    /// Tolerating reserved bits must not turn a genuinely truncated record into
    /// a silently mis-parsed one.
    #[test]
    fn still_rejects_a_truncated_format_record() {
        let mut data = 164u16.to_le_bytes().to_vec();
        data.extend_from_slice(&10u16.to_le_bytes()); // claims 10 characters
        data.push(0);
        data.extend_from_slice(b"0.00"); // supplies 4
        let mut records = vec![
            record(DATE1904_RECORD, vec![0, 0]),
            record(FORMAT_RECORD, data),
        ];
        records.extend(minimal_xf_records());

        assert!(parse_globals_strict(&records).is_err());
    }

    #[test]
    fn a_lenient_policy_repairs_an_xfcrc_count_disagreement_only() {
        let mut records = vec![format_record(164, "yyyy-mm-dd")];
        records.extend(minimal_xf_records());
        let mut crc = [0u8; 20];
        crc[0..2].copy_from_slice(&XFCRC_RECORD.to_le_bytes());
        crc[16..20].copy_from_slice(&99u32.to_le_bytes());
        crc[14..16].copy_from_slice(&(MIN_XF_RECORDS as u16 + 1).to_le_bytes());
        records.push(record(XFCRC_RECORD, crc.to_vec()));

        assert!(parse_globals_strict(&records).is_err());

        let mut tolerance = ToleranceLog::new(crate::leniency::Leniency::TolerateFormattingDefects);
        let refs = records
            .iter()
            .map(|record| record.as_ref())
            .collect::<Vec<_>>();
        let formatting = Formatting::parse_globals(&refs, &mut tolerance)
            .expect("a lenient policy trusts the XF records that were parsed");
        assert_eq!(formatting.extended_formats().len(), MIN_XF_RECORDS);

        let report = tolerance.into_report();
        assert_eq!(
            report.count(FormattingDefect::ExtendedFormatCountMismatch),
            1
        );
        let entry = report.defects()[0];
        assert_eq!(entry.ordinal(), MIN_XF_RECORDS as u32);
        assert_eq!(entry.observed(), MIN_XF_RECORDS as u32 + 1);
    }
}
