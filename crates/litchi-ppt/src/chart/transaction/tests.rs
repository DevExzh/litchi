use std::io::Cursor;

use litchi_cfb::{OleFile, OleWriter};
use litchi_ograph::PackageRef;

use super::PackageEditor;

const BOF: u16 = 0x0809;
const EOF: u16 = 0x000A;
const WORKBOOK: &str = "Workbook";
const COMP_OBJ: &str = "\u{1}CompObj";
const OLE: &str = "\u{1}Ole";

fn bof(doc_type: u16) -> [u8; 16] {
    let mut payload = [0; 16];
    payload[0..2].copy_from_slice(&0x0680_u16.to_le_bytes());
    payload[2..4].copy_from_slice(&doc_type.to_le_bytes());
    payload[4..6].copy_from_slice(&0x0DBB_u16.to_le_bytes());
    payload[6..8].copy_from_slice(&0x07CD_u16.to_le_bytes());
    payload[8..12].copy_from_slice(&(0x0000_0009_u32 | (6 << 14)).to_le_bytes());
    payload[12..16].copy_from_slice(&(0x06_u32 | (6 << 8)).to_le_bytes());
    payload
}

fn push_record(output: &mut Vec<u8>, kind: u16, payload: &[u8]) {
    output.extend_from_slice(&kind.to_le_bytes());
    output.extend_from_slice(
        &(u16::try_from(payload.len()).expect("small test record")).to_le_bytes(),
    );
    output.extend_from_slice(payload);
}

fn chart_stream(marker: u8, tail: &[u8]) -> Vec<u8> {
    let mut output = Vec::new();
    push_record(&mut output, BOF, &bof(0x8000));
    push_record(&mut output, 0x7777, &[marker]);
    push_record(&mut output, 0x7778, tail);
    push_record(&mut output, EOF, &[]);
    output
}

fn excel_chart_stream() -> Vec<u8> {
    let mut chart_bof = bof(0x0020);
    chart_bof[0..2].copy_from_slice(&0x0600_u16.to_le_bytes());
    let mut output = Vec::new();
    push_record(&mut output, BOF, &chart_bof);
    push_record(&mut output, EOF, &[]);
    output
}

fn workbook(chart: &[u8]) -> Vec<u8> {
    let mut bytes = Vec::new();
    push_record(&mut bytes, BOF, &bof(0x0005));
    push_record(&mut bytes, 0x7776, &[0xA0]);
    push_record(&mut bytes, EOF, &[]);
    bytes.extend_from_slice(chart);
    bytes
}

fn package(chart: &[u8]) -> Vec<u8> {
    let workbook = workbook(chart);
    let mut writer = OleWriter::new();
    writer
        .create_stream(&[WORKBOOK], &workbook)
        .expect("Workbook stream");
    writer
        .create_stream(&[COMP_OBJ], b"component object")
        .expect("CompObj stream");
    writer
        .create_stream(&[OLE], b"ole metadata")
        .expect("Ole stream");
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).expect("CFB package");
    output.into_inner()
}

#[test]
fn clean_commit_preserves_the_original_package_bytes() {
    let chart = chart_stream(1, &[2]);
    let bytes = package(&chart);
    let original = bytes.clone();
    let editor = PackageEditor::open(bytes).expect("open Graph package");

    let snapshot = editor.snapshot().expect("snapshot");
    assert!(!snapshot.is_dirty());
    assert_eq!(snapshot.chart().kind(), litchi_ograph::chart::Kind::Graph);
    assert_eq!(snapshot.chart().as_bytes(), chart.as_slice());

    assert_eq!(editor.commit().expect("clean commit"), original);
}

#[test]
fn replacement_is_atomic_and_retains_ograph_host_streams() {
    let original_chart = chart_stream(1, &[2]);
    let replacement = chart_stream(3, &[4, 5, 6, 7]);
    let original_package = package(&original_chart);
    let mut editor = PackageEditor::open(original_package.clone()).expect("open package");

    editor
        .replace_chart(
            litchi_ograph::chart::Stream::open(replacement.clone()).expect("typed chart stream"),
        )
        .expect("replace chart");
    let snapshot = editor.snapshot().expect("edited snapshot");
    assert!(snapshot.is_dirty());
    assert_eq!(snapshot.chart().as_bytes(), replacement.as_slice());
    let original_workbook = PackageRef::open(&original_package)
        .expect("original package")
        .workbook()
        .expect("original Workbook");
    assert_ne!(snapshot.workbook(), original_workbook.as_bytes());

    let output = editor.commit().expect("commit replacement");
    let parsed = PackageRef::open(&output).expect("reopen replacement");
    let output_workbook = parsed.workbook().expect("Workbook");
    assert_eq!(output_workbook.chart().as_bytes(), replacement);

    let mut cfb = OleFile::open(Cursor::new(&output)).expect("open replacement CFB");
    assert_eq!(
        cfb.open_stream(&[COMP_OBJ]).expect("CompObj"),
        b"component object"
    );
    assert_eq!(cfb.open_stream(&[OLE]).expect("Ole"), b"ole metadata");
}

#[test]
fn rejecting_an_excel_stream_leaves_the_transaction_untouched() {
    let chart = chart_stream(9, &[8]);
    let original = package(&chart);
    let mut editor = PackageEditor::open(original.clone()).expect("open package");
    let before = editor.snapshot().expect("before snapshot");
    let before_workbook = before.workbook().to_vec();

    let error = editor
        .replace_chart(
            litchi_ograph::chart::Stream::open(excel_chart_stream())
                .expect("typed Excel chart stream"),
        )
        .expect_err("Excel chart must not enter a Graph package");
    assert!(matches!(error, crate::package::Error::InvalidFormat(_)));
    assert!(!editor.is_dirty());
    let after = editor.snapshot().expect("after snapshot");
    assert_eq!(after.workbook(), before_workbook.as_slice());
    assert_eq!(editor.commit().expect("clean commit"), original);
}
