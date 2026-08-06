use std::io::Cursor;

use litchi_biff::{Encoder, Kind};
use litchi_cfb::OleWriter;

use crate::chart;
use crate::{Error, Limits};

use super::codec::{COMP_OBJ, OLE, WORKBOOK};
use super::semantic::{Package, PackageRef, Workbook, WorkbookRef};
use super::validation::{
    BOF, BOF_BYTES, CHART_SHEET, EOF, GLOBALS, OGRAPH_VERSION, REQUIRED_PLATFORM_FLAGS,
};

fn bof(doc_type: u16) -> [u8; BOF_BYTES] {
    let mut payload = [0; BOF_BYTES];
    payload[0..2].copy_from_slice(&OGRAPH_VERSION.to_le_bytes());
    payload[2..4].copy_from_slice(&doc_type.to_le_bytes());
    payload[4..6].copy_from_slice(&0x0DBB_u16.to_le_bytes());
    payload[6..8].copy_from_slice(&0x07CD_u16.to_le_bytes());
    payload[8..12].copy_from_slice(&(REQUIRED_PLATFORM_FLAGS | (6 << 14)).to_le_bytes());
    payload[12..16].copy_from_slice(&(0x06_u32 | (6 << 8)).to_le_bytes());
    payload
}

fn workbook() -> Vec<u8> {
    workbook_with_bofs(bof(GLOBALS), bof(CHART_SHEET))
}

fn workbook_with_bofs(globals: [u8; BOF_BYTES], chart: [u8; BOF_BYTES]) -> Vec<u8> {
    let mut out = Encoder::with_limits(Limits::default().biff).expect("BIFF limits");
    out.push(BOF, &globals).expect("globals BOF");
    out.push(Kind::from_wire(0x7777), &[1, 2, 3])
        .expect("unknown record");
    out.push(EOF, &[]).expect("globals EOF");
    out.push(BOF, &chart).expect("chart BOF");
    out.push(Kind::from_wire(0x7778), &[4, 5])
        .expect("unknown record");
    out.push(EOF, &[]).expect("chart EOF");
    out.finish()
}

fn package(workbook: Option<&[u8]>, extras: &[(&str, &[u8])]) -> Vec<u8> {
    let mut writer = OleWriter::new();
    if let Some(workbook) = workbook {
        writer
            .create_stream(&[WORKBOOK], workbook)
            .expect("Workbook stream");
    }
    for (name, bytes) in extras {
        writer.create_stream(&[*name], bytes).expect("extra stream");
    }
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).expect("write CFB");
    output.into_inner()
}

#[test]
fn accepts_exact_standalone_topology_and_reads_workbook() {
    let workbook = workbook();
    let bytes = package(Some(&workbook), &[(COMP_OBJ, &[1, 2]), (OLE, &[3, 4, 5])]);
    let parsed = PackageRef::open(&bytes).expect("valid package");
    assert_eq!(parsed.topology().stream_count(), 3);
    assert_eq!(parsed.topology().workbook_bytes(), workbook.len() as u64);
    let opened = parsed.workbook().expect("read Workbook");
    assert_eq!(opened.as_bytes(), workbook);
    assert_eq!(opened.chart().kind(), chart::Kind::Graph);
    assert_eq!(opened.chart().records().count(), 3);
}

#[test]
fn owned_finish_reuses_the_input_allocation() {
    let bytes = package(Some(&workbook()), &[]);
    let pointer = bytes.as_ptr();
    let capacity = bytes.capacity();
    let payload = Package::open(bytes).expect("valid").finish();
    let bytes = payload.into_bytes();
    assert_eq!(bytes.as_ptr(), pointer);
    assert_eq!(bytes.capacity(), capacity);
}

#[test]
fn rejects_missing_unknown_and_nested_root_entries() {
    let missing = package(None, &[(COMP_OBJ, &[])]);
    assert!(matches!(
        PackageRef::open(&missing),
        Err(Error::MissingStream { name: WORKBOOK })
    ));

    let unknown = package(Some(&workbook()), &[("Other", &[])]);
    assert!(matches!(
        PackageRef::open(&unknown),
        Err(Error::UnexpectedEntry { name, .. }) if name == "Other"
    ));

    let mut writer = OleWriter::new();
    writer.create_storage(&["Nested"]).expect("storage");
    writer
        .create_stream(&[WORKBOOK], &workbook())
        .expect("Workbook");
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).expect("write CFB");
    assert!(matches!(
        PackageRef::open(output.get_ref()),
        Err(Error::UnexpectedEntry { name, .. }) if name == "Nested"
    ));
}

#[test]
fn rejects_wrong_substream_shape_and_trailing_records() {
    let mut wrong = Encoder::with_limits(Limits::default().biff).expect("BIFF limits");
    wrong.push(BOF, &bof(GLOBALS)).expect("BOF");
    wrong.push(EOF, &[]).expect("EOF");
    wrong.push(BOF, &bof(GLOBALS)).expect("wrong BOF");
    wrong.push(EOF, &[]).expect("EOF");
    let bytes = package(Some(&wrong.finish()), &[]);
    assert!(matches!(
        PackageRef::open(&bytes),
        Err(Error::InvalidWorkbook { .. })
    ));

    let mut trailing = workbook();
    trailing.extend_from_slice(&[0x0A, 0, 0, 0]);
    let bytes = package(Some(&trailing), &[]);
    assert!(matches!(
        PackageRef::open(&bytes),
        Err(Error::InvalidWorkbook { .. })
    ));
}

#[test]
fn checks_package_and_workbook_limits_before_exposing_data() {
    let bytes = package(Some(&workbook()), &[]);
    let package_limits = Limits {
        max_package_bytes: bytes.len() - 1,
        ..Limits::default()
    };
    assert!(matches!(
        PackageRef::with_limits(&bytes, package_limits),
        Err(Error::LimitExceeded {
            resource: "package bytes",
            ..
        })
    ));

    let workbook_limits = Limits {
        max_workbook_bytes: 1,
        ..Limits::default()
    };
    assert!(matches!(
        PackageRef::with_limits(&bytes, workbook_limits),
        Err(Error::LimitExceeded {
            resource: "Workbook bytes",
            ..
        })
    ));
}

#[test]
fn strict_bof_checks_must_fields_but_preserves_undefined_bits() {
    let mut globals = bof(GLOBALS);
    let ignored = (0b11 << 6) | (1 << 9) | (1 << 10) | (0b11 << 11) | (1 << 13) | (1 << 18);
    globals[8..12].copy_from_slice(&(REQUIRED_PLATFORM_FLAGS | ignored | (6 << 14)).to_le_bytes());
    let bytes = workbook_with_bofs(globals, bof(CHART_SHEET));
    let opened = Workbook::open(bytes).expect("undefined BOF bits are preserved");
    assert_eq!(opened.as_bytes(), opened.as_ref().as_bytes());

    let mut invalid = Vec::new();

    let mut year = bof(GLOBALS);
    year[6..8].copy_from_slice(&0x07CB_u16.to_le_bytes());
    invalid.push(year);

    let mut platform = bof(GLOBALS);
    platform[8..12].copy_from_slice(&(6_u32 << 14).to_le_bytes());
    invalid.push(platform);

    let mut forbidden = bof(GLOBALS);
    forbidden[8..12]
        .copy_from_slice(&(REQUIRED_PLATFORM_FLAGS | (1 << 1) | (6 << 14)).to_le_bytes());
    invalid.push(forbidden);

    let mut reserved1 = bof(GLOBALS);
    reserved1[8..12]
        .copy_from_slice(&(REQUIRED_PLATFORM_FLAGS | (6 << 14) | (1 << 19)).to_le_bytes());
    invalid.push(reserved1);

    let mut highest = bof(GLOBALS);
    highest[8..12].copy_from_slice(&(REQUIRED_PLATFORM_FLAGS | (5 << 14)).to_le_bytes());
    invalid.push(highest);

    let mut lowest = bof(GLOBALS);
    lowest[12..16].copy_from_slice(&(0x05_u32 | (4 << 8)).to_le_bytes());
    invalid.push(lowest);

    let mut last = bof(GLOBALS);
    last[8..12].copy_from_slice(&(REQUIRED_PLATFORM_FLAGS | (4 << 14)).to_le_bytes());
    last[12..16].copy_from_slice(&(0x06_u32 | (6 << 8)).to_le_bytes());
    invalid.push(last);

    let mut reserved2 = bof(GLOBALS);
    reserved2[12..16].copy_from_slice(&(0x06_u32 | (6 << 8) | (1 << 12)).to_le_bytes());
    invalid.push(reserved2);

    for globals in invalid {
        assert!(matches!(
            WorkbookRef::open(&workbook_with_bofs(globals, bof(CHART_SHEET))),
            Err(Error::InvalidWorkbook { .. })
        ));
    }
}
