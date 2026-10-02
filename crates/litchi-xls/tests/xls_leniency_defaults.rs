//! End-to-end coverage for the Excel-compatible repairs a read applies by
//! default (and the opt-in ones behind named [`Leniency`] variants).
//!
//! Every case builds a conforming workbook with the crate writer, plants
//! exactly one cataloged non-conformity in its Workbook stream, and opens it
//! through both public read paths — the eager [`Workbook`] and the
//! source-backed [`SourceBackedWorkbook`] — asserting that the workbook
//! opens, that the repair is recorded with its location, and that the same
//! workbook without the deviation records nothing.

use std::io::Cursor;
use std::sync::Arc;

use litchi_cfb::{OleFile, OleWriter};
use litchi_core::OwnedSource;
use litchi_core::sheet::{Cell as _, CellValue};
use litchi_xls::{
    CellFormatDefect, Defect, EncodingDefect, FormattingDefect, Leniency, OpenOptions,
    PaletteDefect, SharedStringDefect, SheetDefect, SheetNameDefect, SourceBackedOptions,
    SourceBackedWorkbook, ToleranceReport, ToleratedDefect, Workbook, Writer,
};

const CODEPAGE: u16 = 0x0042;
const BOUND_SHEET: u16 = 0x0085;
const PALETTE: u16 = 0x0092;
const FONT: u16 = 0x0031;
const XF: u16 = 0x00E0;
const SST: u16 = 0x00FC;
const NUMBER: u16 = 0x0203;
const LABEL: u16 = 0x0204;
const INDEX: u16 = 0x020B;

/// One BIFF record of a Workbook stream, remembering where it sat in the
/// stream the writer produced so absolute offsets can be re-targeted after
/// an edit changes record sizes.
#[derive(Clone)]
struct Record {
    kind: u16,
    payload: Vec<u8>,
    original_offset: Option<usize>,
}

impl Record {
    fn new(kind: u16, payload: Vec<u8>) -> Self {
        Self {
            kind,
            payload,
            original_offset: None,
        }
    }
}

/// A writer-built workbook's Workbook stream, split into records.
struct Stream {
    records: Vec<Record>,
}

impl Stream {
    /// The Workbook stream of a workbook the crate writer built.
    fn from_writer(build: impl FnOnce(&mut Writer)) -> Self {
        let mut writer = Writer::new();
        build(&mut writer);
        let mut bytes = Vec::new();
        writer.write_to(&mut Cursor::new(&mut bytes)).unwrap();
        let stream = OleFile::open(Cursor::new(&bytes))
            .unwrap()
            .open_stream(&["Workbook"])
            .unwrap();
        let mut records = Vec::new();
        let mut cursor = 0;
        while cursor + 4 <= stream.len() {
            let kind = u16::from_le_bytes([stream[cursor], stream[cursor + 1]]);
            let length = usize::from(u16::from_le_bytes([stream[cursor + 2], stream[cursor + 3]]));
            records.push(Record {
                kind,
                payload: stream[cursor + 4..cursor + 4 + length].to_vec(),
                original_offset: Some(cursor),
            });
            cursor += 4 + length;
        }
        Self { records }
    }

    /// A one-sheet workbook holding one number.
    fn simple() -> Self {
        Self::from_writer(|writer| {
            let sheet = writer.add_worksheet("Sheet1").unwrap();
            writer.write_number(sheet, 0, 0, 42.5).unwrap();
        })
    }

    /// Index of the `nth` record of `kind`.
    fn position(&self, kind: u16, nth: usize) -> usize {
        self.records
            .iter()
            .enumerate()
            .filter(|(_, record)| record.kind == kind)
            .nth(nth)
            .unwrap_or_else(|| panic!("no record {kind:#06x} #{nth}"))
            .0
    }

    /// The payload of the `nth` record of `kind`.
    fn payload_mut(&mut self, kind: u16, nth: usize) -> &mut Vec<u8> {
        let index = self.position(kind, nth);
        &mut self.records[index].payload
    }

    /// Insert `record` immediately before the `nth` record of `kind`.
    fn insert_before(&mut self, kind: u16, nth: usize, record: Record) {
        let index = self.position(kind, nth);
        self.records.insert(index, record);
    }

    /// Byte offset the `index`th record will occupy in the serialized stream.
    fn offset_of(&self, index: usize) -> usize {
        self.records[..index]
            .iter()
            .map(|record| 4 + record.payload.len())
            .sum()
    }

    /// Byte offset of the `nth` record of `kind` in the serialized stream.
    fn offset_of_kind(&self, kind: u16, nth: usize) -> u32 {
        u32::try_from(self.offset_of(self.position(kind, nth))).unwrap()
    }

    /// Serialize the stream, re-targeting the absolute offsets
    /// (`BoundSheet8.lbPlyPos`, `INDEX.rgibRw`) the edit may have moved.
    fn serialize(&self) -> Vec<u8> {
        let relocate = |old: u32| -> u32 {
            let index = self
                .records
                .iter()
                .position(|record| record.original_offset == Some(old as usize))
                .unwrap_or_else(|| panic!("no record at original offset {old}"));
            u32::try_from(self.offset_of(index)).unwrap()
        };
        let mut out = Vec::new();
        for record in &self.records {
            let mut payload = record.payload.clone();
            if record.original_offset.is_some() {
                if record.kind == BOUND_SHEET {
                    let old = u32::from_le_bytes(payload[0..4].try_into().unwrap());
                    payload[0..4].copy_from_slice(&relocate(old).to_le_bytes());
                } else if record.kind == INDEX {
                    for start in (16..payload.len()).step_by(4) {
                        let old = u32::from_le_bytes(payload[start..start + 4].try_into().unwrap());
                        payload[start..start + 4].copy_from_slice(&relocate(old).to_le_bytes());
                    }
                }
            }
            out.extend_from_slice(&record.kind.to_le_bytes());
            out.extend_from_slice(&u16::try_from(payload.len()).unwrap().to_le_bytes());
            out.extend_from_slice(&payload);
        }
        out
    }

    /// The serialized stream, packaged as a compound file.
    fn package(&self) -> Vec<u8> {
        let mut writer = OleWriter::new();
        writer
            .create_stream(&["Workbook"], &self.serialize())
            .unwrap();
        let mut package = Vec::new();
        writer.write_to(&mut Cursor::new(&mut package)).unwrap();
        package
    }
}

/// Open `package` through both public read paths under `leniency`, returning
/// each path's outcome (the source-backed one reduced to its report).
fn try_open_both(
    package: &[u8],
    leniency: Leniency,
) -> (
    Result<Workbook<Cursor<Vec<u8>>>, String>,
    Result<ToleranceReport, String>,
) {
    let workbook = Workbook::new_with_options(
        Cursor::new(package.to_vec()),
        OpenOptions::new().with_leniency(leniency),
    )
    .map_err(|error| error.to_string());
    let owner = SourceBackedWorkbook::from_read_at_with_options(
        Arc::new(OwnedSource::new(package.to_vec())),
        SourceBackedOptions::new().with_leniency(leniency),
    )
    .map(|owner| owner.tolerance_report())
    .map_err(|error| error.to_string());
    (workbook, owner)
}

/// Open `package` through both public read paths under `leniency` and return
/// the eager workbook plus both reports.
fn open_both(package: &[u8], leniency: Leniency) -> (Workbook<Cursor<Vec<u8>>>, ToleranceReport) {
    let (workbook, owner) = try_open_both(package, leniency);
    (
        workbook.expect("the eager open applies the repair"),
        owner.expect("the source-backed open applies the same repair"),
    )
}

/// Both read paths open `package` under `leniency` and record exactly the
/// `expected` repairs, in order. Returns the eager workbook.
fn assert_repairs_under(
    package: &[u8],
    leniency: Leniency,
    expected: &[(Defect, u32, u32)],
) -> Workbook<Cursor<Vec<u8>>> {
    let (workbook, source_report) = open_both(package, leniency);
    for (path, report) in [
        ("eager", workbook.tolerance_report()),
        ("source-backed", &source_report),
    ] {
        assert_eq!(report.unrecorded(), 0, "{path}");
        let found = report
            .defects()
            .iter()
            .map(|entry: &ToleratedDefect| (entry.defect(), entry.ordinal(), entry.observed()))
            .collect::<Vec<_>>();
        assert_eq!(found, expected, "{path} report");
        for entry in report.defects() {
            assert!(!entry.defect().description().is_empty());
        }
    }
    workbook
}

/// Both read paths open `package` under the *default* policy and record
/// exactly the `expected` repairs, in order. Returns the eager workbook.
fn assert_repairs(package: &[u8], expected: &[(Defect, u32, u32)]) -> Workbook<Cursor<Vec<u8>>> {
    assert_repairs_under(package, Leniency::default(), expected)
}

/// The eager open records exactly `expected` under the default policy; the
/// source-backed open, which never reads the carrying record, opens it with
/// nothing to repair.
fn assert_eager_repairs(
    package: &[u8],
    expected: &[(Defect, u32, u32)],
) -> Workbook<Cursor<Vec<u8>>> {
    let (workbook, source_report) = open_both(package, Leniency::default());
    let found = workbook
        .tolerance_report()
        .defects()
        .iter()
        .map(|entry| (entry.defect(), entry.ordinal(), entry.observed()))
        .collect::<Vec<_>>();
    assert_eq!(found, expected);
    assert!(source_report.is_clean(), "{source_report:?}");
    workbook
}

/// Both read paths refuse `package` under `leniency`, each with an error
/// mentioning `needle`.
fn assert_refused(package: &[u8], leniency: Leniency, needle: &str) {
    let (workbook, owner) = try_open_both(package, leniency);
    let Err(eager) = workbook else {
        panic!("the eager open must refuse");
    };
    let Err(source) = owner else {
        panic!("the source-backed open must refuse");
    };
    assert!(eager.contains(needle), "eager: {eager}");
    assert!(source.contains(needle), "source-backed: {source}");
}

/// Both read paths open `package` under `leniency` and record nothing.
fn assert_clean_under(package: &[u8], leniency: Leniency) {
    let (workbook, source_report) = open_both(package, leniency);
    assert!(
        workbook.tolerance_report().is_clean(),
        "{:?}",
        workbook.tolerance_report()
    );
    assert!(source_report.is_clean(), "{source_report:?}");
}

/// Both read paths open `package` under the default policy and record nothing.
fn assert_clean(package: &[u8]) {
    assert_clean_under(package, Leniency::default());
}

#[test]
fn the_unmodified_fixture_is_clean_on_both_paths() {
    assert_clean(&Stream::simple().package());
}

// ---------------------------------------------------------------------------
// F3.2: CODEPAGE without an exact codec
// ---------------------------------------------------------------------------

/// The fixture with its `CODEPAGE` set to `codepage`, plus the record offset.
fn with_codepage(codepage: u16) -> (Vec<u8>, u32) {
    let mut stream = Stream::simple();
    *stream.payload_mut(CODEPAGE, 0) = codepage.to_le_bytes().to_vec();
    let offset = stream.offset_of_kind(CODEPAGE, 0);
    (stream.package(), offset)
}

#[test]
fn an_unsupported_code_page_falls_back_to_windows_1252() {
    // CP437 (DOS Latin US) is a valid [CODEPG] value Lotus-era converters
    // still emit, but this reader has no exact codec for it.
    for codepage in [437_u16, 850, 10029, 0xFFFF] {
        let (package, offset) = with_codepage(codepage);
        let workbook = assert_repairs(
            &package,
            &[(
                Defect::Encoding(EncodingDefect::UnsupportedCodePage),
                offset,
                u32::from(codepage),
            )],
        );
        // The workbook is genuinely usable: names and values survive.
        assert_eq!(workbook.sheets()[0].name(), "Sheet1");
        assert_eq!(
            workbook
                .xls_worksheet(0)
                .unwrap()
                .get_cell(0, 0)
                .unwrap()
                .value(),
            &CellValue::Float(42.5)
        );
        let entry = workbook.tolerance_report().defects()[0];
        assert_eq!(entry.record_type(), Some(CODEPAGE));
    }
}

#[test]
fn a_supported_code_page_records_nothing() {
    for codepage in [1200_u16, 1252, 1251, 932] {
        assert_clean(&with_codepage(codepage).0);
    }
}

// ---------------------------------------------------------------------------
// F3.3: worksheet-name rules (opt-in: Leniency::TolerateSheetNameDefects)
// ---------------------------------------------------------------------------

/// Rewrite the `nth` sheet's `BoundSheet8.stName` to `name`, compressed when
/// every character fits in one byte.
fn set_sheet_name(stream: &mut Stream, nth: usize, name: &str) {
    let payload = stream.payload_mut(BOUND_SHEET, nth);
    payload.truncate(6);
    let units = name.encode_utf16().collect::<Vec<_>>();
    payload.push(u8::try_from(units.len()).unwrap());
    if units.iter().all(|unit| *unit < 0x100) {
        payload.push(0);
        payload.extend(units.iter().map(|unit| *unit as u8));
    } else {
        payload.push(1);
        payload.extend(units.iter().flat_map(|unit| unit.to_le_bytes()));
    }
}

fn two_sheets() -> Stream {
    Stream::from_writer(|writer| {
        let first = writer.add_worksheet("Data").unwrap();
        writer.write_number(first, 0, 0, 1.0).unwrap();
        let second = writer.add_worksheet("Other").unwrap();
        writer.write_number(second, 0, 0, 2.0).unwrap();
    })
}

fn with_second_sheet_name(name: &str) -> Vec<u8> {
    let mut stream = two_sheets();
    set_sheet_name(&mut stream, 1, name);
    stream.package()
}

#[test]
fn a_rule_breaking_sheet_name_is_refused_by_default() {
    for name in ["", "a:b", "'quoted", &"x".repeat(32)] {
        let package = with_second_sheet_name(name);
        for leniency in [Leniency::Strict, Leniency::TolerateFormattingDefects] {
            assert_refused(&package, leniency, "Invalid BoundSheet8 sheet name");
        }
    }
    let package = with_second_sheet_name("DATA");
    for leniency in [Leniency::Strict, Leniency::TolerateFormattingDefects] {
        assert_refused(
            &package,
            leniency,
            "duplicate case-insensitive BoundSheet8 name",
        );
    }
}

#[test]
fn the_sheet_name_opt_in_keeps_each_rule_breaking_name_verbatim() {
    let long = "x".repeat(40);
    let cases: [(&str, SheetNameDefect, u32); 6] = [
        ("", SheetNameDefect::Empty, 0),
        (&long, SheetNameDefect::TooLong, 40),
        ("a:b", SheetNameDefect::ForbiddenCharacter, u32::from(':')),
        ("Q1/Q2", SheetNameDefect::ForbiddenCharacter, u32::from('/')),
        ("'quoted", SheetNameDefect::EdgeApostrophe, 0x27),
        ("DATA", SheetNameDefect::Duplicate, 0),
    ];
    for leniency in [
        Leniency::TolerateSheetNameDefects,
        Leniency::TolerateAllDefects,
    ] {
        for (name, defect, observed) in cases {
            let workbook = assert_repairs_under(
                &with_second_sheet_name(name),
                leniency,
                &[(Defect::SheetName(defect), 1, observed)],
            );
            // The name is kept exactly as stored, and both sheets survive.
            assert_eq!(workbook.sheets()[1].name(), name);
            assert_eq!(workbook.sheets().len(), 2);
            assert_eq!(
                workbook
                    .xls_worksheet(1)
                    .unwrap()
                    .get_cell(0, 0)
                    .unwrap()
                    .value(),
                &CellValue::Float(2.0)
            );
            assert_eq!(
                workbook.tolerance_report().defects()[0].record_type(),
                Some(BOUND_SHEET)
            );
        }
    }
}

#[test]
fn the_sheet_name_opt_in_records_nothing_for_conforming_names() {
    for leniency in [
        Leniency::TolerateSheetNameDefects,
        Leniency::TolerateAllDefects,
    ] {
        assert_clean_under(&two_sheets().package(), leniency);
        // A name that differs from another by more than case is no duplicate.
        assert_clean_under(&with_second_sheet_name("Data2"), leniency);
    }
}

// ---------------------------------------------------------------------------
// F3.4: Palette reserved byte, color count, duplicates
// ---------------------------------------------------------------------------

/// A `Palette` payload of `count` entries, entry `i` colored `(i, i, i)` with
/// reserved byte `reserved`.
fn palette_payload(count: i16, reserved: u8) -> Vec<u8> {
    let mut payload = count.to_le_bytes().to_vec();
    for index in 0..count {
        let level = u8::try_from(index).unwrap();
        payload.extend_from_slice(&[level, level, level, reserved]);
    }
    payload
}

/// The fixture with one `Palette` record per payload inserted before the
/// first `BoundSheet8`, plus each record's offset.
fn with_palettes(payloads: &[Vec<u8>]) -> (Vec<u8>, Vec<u32>) {
    let mut stream = Stream::simple();
    for payload in payloads {
        stream.insert_before(BOUND_SHEET, 0, Record::new(PALETTE, payload.clone()));
    }
    let offsets = (0..payloads.len())
        .map(|nth| stream.offset_of_kind(PALETTE, nth))
        .collect();
    (stream.package(), offsets)
}

#[test]
fn a_palette_reserved_byte_is_ignored_and_recorded() {
    let (package, offsets) = with_palettes(&[palette_payload(56, 0xFF)]);
    let workbook = assert_eager_repairs(
        &package,
        &[(Defect::Palette(PaletteDefect::ReservedByte), offsets[0], 56)],
    );
    let color = workbook.palette().color(0x08 + 9).unwrap();
    assert_eq!((color.red(), color.green(), color.blue()), (9, 9, 9));
}

#[test]
fn a_short_palette_keeps_default_colors_for_the_remainder() {
    let (package, offsets) = with_palettes(&[palette_payload(16, 0)]);
    let workbook = assert_eager_repairs(
        &package,
        &[(Defect::Palette(PaletteDefect::ColorCount), offsets[0], 16)],
    );
    let custom = workbook.palette().color(0x08 + 15).unwrap();
    assert_eq!(custom.red(), 15);
    // Entry 0x3F is past the record's 16 colors: the default (51, 51, 51).
    let default = workbook.palette().color(0x3F).unwrap();
    assert_eq!(
        (default.red(), default.green(), default.blue()),
        (51, 51, 51)
    );
}

#[test]
fn the_last_of_two_palettes_wins_and_is_recorded() {
    let mut second = palette_payload(56, 0);
    second[2..5].copy_from_slice(&[200, 100, 50]);
    let (package, offsets) = with_palettes(&[palette_payload(56, 0), second]);
    let workbook = assert_eager_repairs(
        &package,
        &[(Defect::Palette(PaletteDefect::Duplicate), offsets[1], 2)],
    );
    let color = workbook.palette().color(0x08).unwrap();
    assert_eq!((color.red(), color.green(), color.blue()), (200, 100, 50));
}

#[test]
fn a_conforming_palette_records_nothing() {
    let (package, _) = with_palettes(&[palette_payload(56, 0)]);
    assert_clean(&package);
}

#[test]
fn a_palette_whose_length_contradicts_its_count_is_still_refused() {
    let mut payload = palette_payload(56, 0);
    payload.truncate(payload.len() - 4);
    let (package, _) = with_palettes(&[payload]);
    let Err(error) = Workbook::new(Cursor::new(package)) else {
        panic!("record framing stays a hard error");
    };
    let error = error.to_string();
    assert!(error.contains("Palette"), "{error}");
}

// ---------------------------------------------------------------------------
// F3.5: Font scalar fields (opt-in: Leniency::TolerateFormattingDefects)
// ---------------------------------------------------------------------------

/// The fixture with the physical-index-`nth` font's weight set to `weight`
/// and its character set to `charset`.
fn with_font_fields(nth: usize, weight: u16, charset: u8) -> Vec<u8> {
    let mut stream = Stream::simple();
    let payload = stream.payload_mut(FONT, nth);
    payload[6..8].copy_from_slice(&weight.to_le_bytes());
    payload[12] = charset;
    stream.package()
}

fn open_eager(package: &[u8], leniency: Leniency) -> Result<Workbook<Cursor<Vec<u8>>>, String> {
    Workbook::new_with_options(
        Cursor::new(package.to_vec()),
        OpenOptions::new().with_leniency(leniency),
    )
    .map_err(|error| error.to_string())
}

#[test]
fn invalid_font_scalars_stay_refused_by_default() {
    let package = with_font_fields(3, 1200, 3);
    for leniency in [Leniency::Strict, Leniency::TolerateSheetNameDefects] {
        let Err(error) = open_eager(&package, leniency) else {
            panic!("{leniency:?} must refuse");
        };
        assert!(error.contains("Font weight"), "{error}");
    }
}

#[test]
fn the_formatting_opt_in_repairs_invalid_font_scalars() {
    let package = with_font_fields(3, 1200, 3);
    for leniency in [
        Leniency::TolerateFormattingDefects,
        Leniency::TolerateAllDefects,
    ] {
        let workbook = open_eager(&package, leniency).expect("the opt-in repairs the font");
        let found = workbook
            .tolerance_report()
            .defects()
            .iter()
            .map(|entry| (entry.defect(), entry.ordinal(), entry.observed()))
            .collect::<Vec<_>>();
        assert_eq!(
            found,
            [
                (Defect::Formatting(FormattingDefect::FontWeight), 3, 1200),
                (Defect::Formatting(FormattingDefect::FontCharset), 3, 3),
            ]
        );
        let font = workbook.font(3).expect("the repaired font is indexed");
        assert_eq!(font.weight(), 1000);
    }
    // The source-backed open never reads Font records.
    assert!(SourceBackedWorkbook::from_read_at(Arc::new(OwnedSource::new(package))).is_ok());
}

#[test]
fn conforming_fonts_record_nothing_under_the_opt_in() {
    assert_clean_under(
        &Stream::simple().package(),
        Leniency::TolerateFormattingDefects,
    );
}

// ---------------------------------------------------------------------------
// F3.6: BoundSheet8 reserved bits, visibility 3, unknown type, trailing bytes
// ---------------------------------------------------------------------------

/// The two-sheet fixture with the second `BoundSheet8` payload edited.
fn with_second_bound_sheet(edit: impl FnOnce(&mut Vec<u8>)) -> Vec<u8> {
    let mut stream = two_sheets();
    edit(stream.payload_mut(BOUND_SHEET, 1));
    stream.package()
}

#[test]
fn reserved_sheet_name_option_bits_are_ignored_and_recorded() {
    let package = with_second_bound_sheet(|payload| payload[7] |= 0xF0);
    let workbook = assert_repairs(
        &package,
        &[(Defect::Sheet(SheetDefect::ReservedNameFlags), 1, 0xF0)],
    );
    assert_eq!(workbook.sheets()[1].name(), "Other");
}

#[test]
fn visibility_three_reads_as_very_hidden_and_is_recorded() {
    let package = with_second_bound_sheet(|payload| payload[4] = 3);
    let workbook = assert_repairs(&package, &[(Defect::Sheet(SheetDefect::Visibility), 1, 3)]);
    assert_eq!(
        workbook.sheets()[1].visibility(),
        litchi_xls::SheetVisibility::VeryHidden
    );
}

#[test]
fn an_unknown_sheet_type_reads_as_a_worksheet_and_is_recorded() {
    let package = with_second_bound_sheet(|payload| payload[5] = 9);
    let workbook = assert_repairs(&package, &[(Defect::Sheet(SheetDefect::SheetType), 1, 9)]);
    assert_eq!(
        workbook
            .xls_worksheet(1)
            .unwrap()
            .get_cell(0, 0)
            .unwrap()
            .value(),
        &CellValue::Float(2.0)
    );
}

#[test]
fn trailing_bound_sheet_bytes_are_ignored_and_recorded() {
    let package = with_second_bound_sheet(|payload| payload.extend_from_slice(&[0xAA; 3]));
    let workbook = assert_repairs(
        &package,
        &[(Defect::Sheet(SheetDefect::TrailingBytes), 1, 3)],
    );
    assert_eq!(workbook.sheets()[1].name(), "Other");
}

#[test]
fn a_bound_sheet_too_short_for_its_name_is_still_refused() {
    let package = with_second_bound_sheet(|payload| {
        payload.pop();
    });
    assert_refused(&package, Leniency::default(), "BoundSheet8");
}

// ---------------------------------------------------------------------------
// F3.7: XF alignment / border / fill enumerations
// ---------------------------------------------------------------------------

/// The fixture with XF 15 — the default cell XF the number uses — edited.
fn with_cell_xf(edit: impl FnOnce(&mut Vec<u8>)) -> Vec<u8> {
    let mut stream = Stream::simple();
    edit(stream.payload_mut(XF, 15));
    stream.package()
}

#[test]
fn undefined_xf_enumerations_read_as_defaults_and_are_recorded() {
    let package = with_cell_xf(|payload| {
        payload[6] = (payload[6] & 0x8F) | (6 << 4); // alcV = 6
        payload[7] = 200; // trot
        payload[8] |= 0xE0; // reserved bit + iReadOrder = 3
        let mut border1 = u32::from_le_bytes(payload[10..14].try_into().unwrap());
        border1 = (border1 & !0x0F) | 14; // dgLeft = 14
        border1 = (border1 & !(0x7F << 16)) | (0x60 << 16); // icvLeft
        payload[10..14].copy_from_slice(&border1.to_le_bytes());
        let mut border2 = u32::from_le_bytes(payload[14..18].try_into().unwrap());
        border2 = (border2 & !(0x3F << 26)) | (20 << 26); // fls = 20
        payload[14..18].copy_from_slice(&border2.to_le_bytes());
    });
    let workbook = assert_repairs(
        &package,
        &[
            (
                Defect::CellFormat(CellFormatDefect::AlignmentVertical),
                15,
                6,
            ),
            (
                Defect::CellFormat(CellFormatDefect::AlignmentRotation),
                15,
                200,
            ),
            (
                Defect::CellFormat(CellFormatDefect::AlignmentReservedBit),
                15,
                0xE0,
            ),
            (
                Defect::CellFormat(CellFormatDefect::AlignmentReadingOrder),
                15,
                3,
            ),
            (Defect::CellFormat(CellFormatDefect::ColorIndex), 15, 0x60),
            (Defect::CellFormat(CellFormatDefect::BorderStyle), 15, 14),
            (Defect::CellFormat(CellFormatDefect::FillPattern), 15, 20),
        ],
    );
    let worksheet = workbook.xls_worksheet(0).unwrap();
    let format = worksheet
        .format_for_cell(0, 0)
        .expect("the repaired cell XF");
    assert_eq!(
        format.alignment().vertical(),
        litchi_xls::VerticalAlignment::Bottom
    );
    assert_eq!(
        format.borders().left().style(),
        litchi_xls::BorderStyle::None
    );
    assert_eq!(format.borders().left().color_index(), 0x7F);
    assert_eq!(format.fill().pattern(), litchi_xls::FillPattern::None);
    assert_eq!(
        worksheet.get_cell(0, 0).unwrap().value(),
        &CellValue::Float(42.5)
    );
    for entry in workbook.tolerance_report().defects() {
        assert_eq!(entry.record_type(), Some(XF));
    }
}

// ---------------------------------------------------------------------------
// F3.8: unpaired UTF-16 surrogates and out-of-order formatting runs
// ---------------------------------------------------------------------------

/// A one-string `SST` payload holding `units` as uncompressed UTF-16, with
/// the given rich-text runs.
fn sst_payload(units: &[u16], runs: &[(u16, u16)]) -> Vec<u8> {
    let mut payload = Vec::new();
    payload.extend_from_slice(&1u32.to_le_bytes());
    payload.extend_from_slice(&1u32.to_le_bytes());
    payload.extend_from_slice(&u16::try_from(units.len()).unwrap().to_le_bytes());
    payload.push(if runs.is_empty() { 0x01 } else { 0x09 });
    if !runs.is_empty() {
        payload.extend_from_slice(&u16::try_from(runs.len()).unwrap().to_le_bytes());
    }
    for unit in units {
        payload.extend_from_slice(&unit.to_le_bytes());
    }
    for (character, font) in runs {
        payload.extend_from_slice(&character.to_le_bytes());
        payload.extend_from_slice(&font.to_le_bytes());
    }
    payload
}

/// A one-sheet workbook whose only cell is shared string 0, with the SST
/// replaced by `payload`.
fn with_sst(payload: Vec<u8>) -> Vec<u8> {
    let mut stream = Stream::from_writer(|writer| {
        let sheet = writer.add_worksheet("Sheet1").unwrap();
        writer.write_string(sheet, 0, 0, "x").unwrap();
    });
    *stream.payload_mut(SST, 0) = payload;
    stream.package()
}

/// The text of cell (0, 0) through both read paths.
fn cell_text_both(package: &[u8]) -> (CellValue, CellValue) {
    let (workbook, _) = open_both(package, Leniency::default());
    let eager = workbook
        .xls_worksheet(0)
        .unwrap()
        .get_cell(0, 0)
        .unwrap()
        .value()
        .clone();
    let owner =
        SourceBackedWorkbook::from_read_at(Arc::new(OwnedSource::new(package.to_vec()))).unwrap();
    let source = owner
        .worksheet_by_index(0)
        .unwrap()
        .unwrap()
        .cell(0, 0)
        .unwrap()
        .unwrap()
        .into_value();
    (eager, source)
}

#[test]
fn an_unpaired_shared_string_surrogate_decodes_to_the_replacement_character() {
    let package = with_sst(sst_payload(&[0x0041, 0xD800, 0x0042], &[]));
    assert_repairs(
        &package,
        &[(
            Defect::SharedString(SharedStringDefect::UnpairedSurrogate),
            0,
            0xD800,
        )],
    );
    let expected = CellValue::String("A\u{FFFD}B".to_string());
    assert_eq!(cell_text_both(&package), (expected.clone(), expected));
}

#[test]
fn an_out_of_order_formatting_run_is_dropped_and_recorded() {
    let package = with_sst(sst_payload(&[0x0041, 0x0042, 0x0043], &[(2, 5), (1, 6)]));
    assert_repairs(
        &package,
        &[(
            Defect::SharedString(SharedStringDefect::FormattingRun),
            0,
            1,
        )],
    );
    let expected = CellValue::String("ABC".to_string());
    assert_eq!(cell_text_both(&package), (expected.clone(), expected));
}

#[test]
fn well_formed_shared_strings_record_nothing() {
    // A valid surrogate pair (U+1F600) and increasing runs.
    assert_clean(&with_sst(sst_payload(
        &[0xD83D, 0xDE00, 0x0041],
        &[(0, 5), (2, 6)],
    )));
}

/// A `Label` payload for (row, col) with XF 15 holding `units` as UTF-16.
fn label_payload(row: u16, col: u16, units: &[u16]) -> Vec<u8> {
    let mut payload = Vec::new();
    payload.extend_from_slice(&row.to_le_bytes());
    payload.extend_from_slice(&col.to_le_bytes());
    payload.extend_from_slice(&15u16.to_le_bytes());
    payload.extend_from_slice(&u16::try_from(units.len()).unwrap().to_le_bytes());
    payload.push(0x01);
    for unit in units {
        payload.extend_from_slice(&unit.to_le_bytes());
    }
    payload
}

/// The fixture with its number cell replaced by a `Label` holding `units`.
fn with_label(units: &[u16]) -> Vec<u8> {
    let mut stream = Stream::simple();
    let index = stream.position(NUMBER, 0);
    stream.records[index].kind = LABEL;
    stream.records[index].payload = label_payload(0, 0, units);
    stream.package()
}

#[test]
fn an_unpaired_label_surrogate_decodes_to_the_replacement_character() {
    let package = with_label(&[0xDC00, 0x0041]);
    let expected_repair = (Defect::Encoding(EncodingDefect::LabelSurrogate), 0, 0);
    // The eager open decodes every cell at open time. The source-backed open
    // reads worksheets lazily, so — like the cell-XF repair — it records the
    // repair when a query scans the cell, not at open.
    let workbook = assert_eager_repairs(&package, &[expected_repair]);
    assert_eq!(
        workbook.tolerance_report().defects()[0].record_type(),
        Some(LABEL)
    );
    let expected = CellValue::String("\u{FFFD}A".to_string());
    assert_eq!(
        cell_text_both(&package),
        (expected.clone(), expected.clone())
    );

    let owner =
        SourceBackedWorkbook::from_read_at(Arc::new(OwnedSource::new(package.clone()))).unwrap();
    let value = owner
        .worksheet_by_index(0)
        .unwrap()
        .unwrap()
        .cell(0, 0)
        .unwrap()
        .unwrap()
        .into_value();
    assert_eq!(value, expected);
    let found = owner
        .tolerance_report()
        .defects()
        .iter()
        .map(|entry| (entry.defect(), entry.ordinal(), entry.observed()))
        .collect::<Vec<_>>();
    assert_eq!(found, [expected_repair]);
}

#[test]
fn a_well_formed_label_records_nothing() {
    assert_clean(&with_label(&[0xD83D, 0xDE00, 0x0041]));
}

#[test]
fn an_unpaired_font_name_surrogate_decodes_to_the_replacement_character() {
    let mut stream = Stream::simple();
    let payload = stream.payload_mut(FONT, 0);
    payload.truncate(14);
    payload.extend_from_slice(&[2, 0x01]); // cch, fHighByte
    payload.extend_from_slice(&0x0041u16.to_le_bytes());
    payload.extend_from_slice(&0xDBFFu16.to_le_bytes());
    let package = stream.package();
    // The source-backed open never reads Font records.
    let workbook = assert_eager_repairs(
        &package,
        &[(
            Defect::Encoding(EncodingDefect::FontNameSurrogate),
            0,
            0xDBFF,
        )],
    );
    assert_eq!(workbook.font(0).unwrap().name(), "A\u{FFFD}");
}

// ---------------------------------------------------------------------------
// F3.9: number-format, font, and XFCRC cross-references
// ---------------------------------------------------------------------------

const XFCRC: u16 = 0x087C;

#[test]
fn an_unresolvable_number_format_and_reserved_font_fall_back_and_are_recorded() {
    let package = with_cell_xf(|payload| {
        payload[0..2].copy_from_slice(&4u16.to_le_bytes()); // ifnt 4 is reserved
        payload[2..4].copy_from_slice(&300u16.to_le_bytes()); // no Format 300
    });
    let workbook = assert_repairs(
        &package,
        &[
            (Defect::CellFormat(CellFormatDefect::FontReference), 15, 4),
            (
                Defect::CellFormat(CellFormatDefect::NumberFormatReference),
                15,
                300,
            ),
        ],
    );
    let worksheet = workbook.xls_worksheet(0).unwrap();
    let format = worksheet.format_for_cell(0, 0).unwrap();
    assert_eq!(format.number_format_id(), 0, "General");
    assert_eq!(format.font_index(), 0);
    assert_eq!(
        worksheet.get_cell(0, 0).unwrap().value(),
        &CellValue::Float(42.5)
    );
}

/// The fixture with an `XFCRC` record after its last `XF`, declaring the
/// true XF count, with the reserved bytes set to `reserved`.
fn with_xfcrc(reserved: u8) -> (Vec<u8>, u32) {
    let mut stream = Stream::simple();
    let xf_count = stream
        .records
        .iter()
        .filter(|record| record.kind == XF)
        .count();
    let mut payload = vec![0u8; 20];
    payload[0..2].copy_from_slice(&XFCRC.to_le_bytes());
    payload[4..12].fill(reserved);
    payload[14..16].copy_from_slice(&u16::try_from(xf_count).unwrap().to_le_bytes());
    let last_xf = stream.position(XF, xf_count - 1);
    stream
        .records
        .insert(last_xf + 1, Record::new(XFCRC, payload));
    let offset = stream.offset_of_kind(XFCRC, 0);
    (stream.package(), offset)
}

#[test]
fn xfcrc_reserved_bytes_are_ignored_and_recorded() {
    let (package, offset) = with_xfcrc(0xAB);
    assert_repairs(
        &package,
        &[(
            Defect::CellFormat(CellFormatDefect::XfcrcReserved),
            offset,
            8,
        )],
    );
}

#[test]
fn a_conforming_xfcrc_records_nothing() {
    assert_clean(&with_xfcrc(0).0);
}

// ---------------------------------------------------------------------------
// Compound-file (container) repairs join the workbook report
// ---------------------------------------------------------------------------

/// The simple fixture packaged with a 5000-byte `LeniencyProbe` stream beside
/// the Workbook stream.
fn package_with_probe() -> Vec<u8> {
    let mut writer = OleWriter::new();
    writer
        .create_stream(&["Workbook"], &Stream::simple().serialize())
        .unwrap();
    writer
        .create_stream(&["LeniencyProbe"], &[0xAB; 5000])
        .unwrap();
    let mut package = Vec::new();
    writer.write_to(&mut Cursor::new(&mut package)).unwrap();
    package
}

/// Shrink the probe's declared size from 5000 to 4096 bytes, so its
/// ten-sector chain runs two sectors past the eight the size needs.
fn shrink_probe(mut package: Vec<u8>) -> Vec<u8> {
    let name = "LeniencyProbe"
        .encode_utf16()
        .flat_map(u16::to_le_bytes)
        .collect::<Vec<_>>();
    let entry = package
        .windows(name.len())
        .position(|window| window == name.as_slice())
        .expect("probe directory entry");
    assert_eq!(
        u32::from_le_bytes(package[entry + 120..entry + 124].try_into().unwrap()),
        5000
    );
    package[entry + 120..entry + 124].copy_from_slice(&4096u32.to_le_bytes());
    package
}

/// A non-critical stream whose FAT chain runs past its declared size opens
/// (POI and LibreOffice read exactly the declared bytes) on both read paths,
/// and the compound-file repair is listed as a `Defect::Container` entry —
/// located by the probe's directory SID — ahead of every BIFF-level repair.
#[test]
fn container_repairs_join_the_workbook_report_on_both_paths() {
    let conformant = package_with_probe();
    assert_clean(&conformant);

    let package = shrink_probe(conformant);
    // The container report, in discovery order: the over-long chain, then
    // the two sectors it carries past the declared size.
    let container = OleFile::open(Cursor::new(package.clone()))
        .unwrap()
        .tolerance_report()
        .clone();
    assert!(container.contains(litchi_cfb::AllocationDefect::ChainLongerThanSize));
    let expected = container
        .defects()
        .iter()
        .map(|tolerated| (Defect::Container(tolerated.defect), tolerated.context, 0))
        .collect::<Vec<_>>();
    let workbook = assert_repairs(&package, &expected);
    assert_eq!(workbook.sheets()[0].name(), "Sheet1");
    // The eager path built from an already-opened container folds it too.
    let ole = OleFile::open(Cursor::new(package)).unwrap();
    let workbook = Workbook::from_ole_file(ole).unwrap();
    assert_eq!(workbook.tolerance_report().defects().len(), expected.len());
}

// ---------------------------------------------------------------------------
// Property-set repairs join the workbook report
// ---------------------------------------------------------------------------

/// A one-section `SummaryInformation` stream whose title is a `VT_LPWSTR`
/// holding an unpaired high surrogate (or, with `surrogate` false, a plain
/// `A`).
fn summary_information(surrogate: bool) -> Vec<u8> {
    const SUMMARY_INFORMATION_FMTID: [u8; 16] = [
        0xE0, 0x85, 0x9F, 0xF2, 0xF9, 0x4F, 0x68, 0x10, 0xAB, 0x91, 0x08, 0x00, 0x2B, 0x27, 0xB3,
        0xD9,
    ];
    let first_unit: u16 = if surrogate { 0xD800 } else { 0x41 };
    let code_page = [2u16.to_le_bytes(), [0, 0], 1200u16.to_le_bytes(), [0, 0]].concat();
    let title = [
        0x1Fu16.to_le_bytes(),
        [0, 0],
        3u16.to_le_bytes(),
        [0, 0],
        first_unit.to_le_bytes(),
        0x41u16.to_le_bytes(),
        [0, 0],
        [0, 0],
    ]
    .concat();
    let table_end = 8 + 2 * 8;
    let mut section = Vec::new();
    section.extend_from_slice(&((table_end + code_page.len() + title.len()) as u32).to_le_bytes());
    section.extend_from_slice(&2u32.to_le_bytes());
    section.extend_from_slice(&1u32.to_le_bytes());
    section.extend_from_slice(&(table_end as u32).to_le_bytes());
    section.extend_from_slice(&2u32.to_le_bytes());
    section.extend_from_slice(&((table_end + code_page.len()) as u32).to_le_bytes());
    section.extend_from_slice(&code_page);
    section.extend_from_slice(&title);
    let mut stream = Vec::new();
    stream.extend_from_slice(&0xFFFEu16.to_le_bytes());
    stream.extend_from_slice(&0u16.to_le_bytes());
    stream.extend_from_slice(&0u32.to_le_bytes());
    stream.extend_from_slice(&[0; 16]);
    stream.extend_from_slice(&1u32.to_le_bytes());
    stream.extend_from_slice(&SUMMARY_INFORMATION_FMTID);
    stream.extend_from_slice(&48u32.to_le_bytes());
    stream.extend_from_slice(&section);
    stream
}

/// The simple fixture packaged with a `\x05SummaryInformation` stream.
fn package_with_summary(surrogate: bool) -> Vec<u8> {
    let mut writer = OleWriter::new();
    writer
        .create_stream(&["Workbook"], &Stream::simple().serialize())
        .unwrap();
    writer
        .create_stream(
            &["\u{5}SummaryInformation"],
            &summary_information(surrogate),
        )
        .unwrap();
    let mut package = Vec::new();
    writer.write_to(&mut Cursor::new(&mut package)).unwrap();
    package
}

/// A title with an unpaired surrogate in `\x05SummaryInformation` decodes
/// lossily the way ole32 and POI decode it, and the eager open lists the
/// repair as a `Defect::PropertySet` entry located by the title's property
/// identifier; the source-backed owner, which reads no metadata stream, opens
/// it with nothing to repair. The same metadata without the surrogate records
/// nothing.
#[test]
fn property_set_repairs_join_the_eager_workbook_report() {
    assert_clean(&package_with_summary(false));
    let mut workbook = assert_eager_repairs(
        &package_with_summary(true),
        &[(
            Defect::PropertySet(litchi_ole_common::leniency::Defect::String(
                litchi_ole_common::leniency::StringDefect::LossyDecode,
            )),
            2,
            0,
        )],
    );
    // The metadata accessor still hands the lossy title back with its own
    // report, and reading it does not grow the workbook report.
    let summary = workbook.summary_information().unwrap().unwrap();
    assert!(!summary.tolerance_report().is_clean());
    assert_eq!(workbook.tolerance_report().defects().len(), 1);
}
