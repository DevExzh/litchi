//! Public exact-output guard for source-backed multi-cell edits.
//!
//! Include this module from `source_backed_cell_values.rs` after copying it to
//! `source_backed_cell_values/public_exact_output.rs`.

use super::*;

#[test]
fn public_multi_edit_matches_baseline_whole_worksheet_output() {
    let source_xml = format!(
        r#"<worksheet xmlns="{SML}"><sheetData>
  <row r="1" spans="1:3"><c r="A1"><v>1</v></c><!--keep-row-one--><c r='B1' t='b'><v>1</v></c></row>
  <row r="2" spans="1:3"><c r="A2"><v>2</v></c><!--keep-row-two--><c r="C2"><v>3</v></c></row>
</sheetData></worksheet>"#
    );
    // Preserve a noncompact external source; the authored-package writer
    // deliberately requires compact XML and is not this test's input route.
    let base = two_sheets();
    let archive = ArchiveReader::new(&base).unwrap();
    let mut writer = soapberry_zip::office::StreamingArchiveWriter::new();
    for member in archive.file_names() {
        let content = if member == "xl/worksheets/sheet1.xml" {
            source_xml.as_bytes().to_vec()
        } else {
            archive.read(member).unwrap()
        };
        writer.write_deflated(member, &content).unwrap();
    }
    let bytes = writer.finish_to_bytes().unwrap();
    let source = Arc::new(VersionedSource::new(bytes.clone()));
    let editor = SourceBackedEditor::from_read_at(source.clone()).unwrap();

    let mut edit = editor.edit_sheets(["Sheet1".into()]).unwrap();
    edit.set("Sheet1", address("A1"), 10u32).unwrap();
    edit.set("Sheet1", address("C2"), 30u32).unwrap();
    let commit = edit.commit().unwrap();
    assert!(commit.changed());
    assert_eq!(commit.diagnostics().changed_cells(), 2);
    assert_eq!(commit.diagnostics().touched_worksheets(), 1);

    let mut published = Vec::new();
    editor
        .publish_multi_commit_to_stream(&mut published, &commit)
        .unwrap();

    // This is the explicit baseline writer output: only the two selected
    // primary value spans change; row tags, comments, whitespace, and the
    // untouched B1/A2 cell bytes remain exactly as authored above.
    let expected_sheet = format!(
        r#"<worksheet xmlns="{SML}"><sheetData>
  <row r="1" spans="1:3"><c r="A1"><v>10</v></c><!--keep-row-one--><c r='B1' t='b'><v>1</v></c></row>
  <row r="2" spans="1:3"><c r="A2"><v>2</v></c><!--keep-row-two--><c r="C2"><v>30</v></c></row>
</sheetData></worksheet>"#
    );
    assert_eq!(
        zip_member(&published, "xl/worksheets/sheet1.xml"),
        expected_sheet.as_bytes()
    );
    assert_eq!(source.bytes.as_slice(), bytes.as_slice());

    // The public commit snapshot must agree with the exact worksheet oracle.
    assert_eq!(
        commit.snapshot().value(0, address("A1")),
        Some(&Value::Number(Number::new("10").unwrap()))
    );
    assert_eq!(
        commit.snapshot().value(0, address("C2")),
        Some(&Value::Number(Number::new("30").unwrap()))
    );
}
