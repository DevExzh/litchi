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
    Defect, EncodingDefect, SourceBackedWorkbook, ToleranceReport, ToleratedDefect, Workbook,
    Writer,
};

const CODEPAGE: u16 = 0x0042;
const BOUND_SHEET: u16 = 0x0085;
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

/// Open `package` through both public read paths under default options and
/// return both reports.
fn open_both(package: &[u8]) -> (Workbook<Cursor<Vec<u8>>>, ToleranceReport) {
    let workbook = Workbook::new(Cursor::new(package.to_vec()))
        .expect("the default policy applies the Excel-compatible repair");
    let owner = SourceBackedWorkbook::from_read_at(Arc::new(OwnedSource::new(package.to_vec())))
        .expect("the source-backed default policy applies the same repair");
    let source_report = owner.tolerance_report();
    (workbook, source_report)
}

/// Both read paths open `package` and record exactly the `expected` repairs,
/// in order. Returns the eager workbook for further assertions.
fn assert_repairs(package: &[u8], expected: &[(Defect, u32, u32)]) -> Workbook<Cursor<Vec<u8>>> {
    let (workbook, source_report) = open_both(package);
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

/// Both read paths open `package` and record nothing.
fn assert_clean(package: &[u8]) {
    let (workbook, source_report) = open_both(package);
    assert!(
        workbook.tolerance_report().is_clean(),
        "{:?}",
        workbook.tolerance_report()
    );
    assert!(source_report.is_clean(), "{source_report:?}");
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
