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
use litchi_core::sheet::Cell as _;
use litchi_xls::{
    Defect, EncodingDefect, Leniency, OpenOptions, PaletteDefect, SheetNameDefect,
    SourceBackedOptions, SourceBackedWorkbook, ToleranceReport, ToleratedDefect, Workbook, Writer,
};

const CODEPAGE: u16 = 0x0042;
const BOUND_SHEET: u16 = 0x0085;
const PALETTE: u16 = 0x0092;
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
            &litchi_core::sheet::CellValue::Float(42.5)
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
                &litchi_core::sheet::CellValue::Float(2.0)
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
