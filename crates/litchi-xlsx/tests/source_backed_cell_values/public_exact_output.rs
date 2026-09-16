//! Public exact-output guard for source-backed multi-cell edits.
//!
//! Include this module from `source_backed_cell_values.rs` after copying it to
//! `source_backed_cell_values/public_exact_output.rs`.

use super::*;

#[test]
fn public_multi_edit_matches_baseline_whole_worksheet_output() {
    let source_xml = format!(
        r#"<worksheet xmlns="{SML}"><sheetData><row r="1" spans="1:3"><c r="A1"><v>1</v></c><!--keep-row-one--><c r='B1' t='b'><v>1</v></c></row><row r="2" spans="1:3"><c r="A2"><v> 2 </v></c><!--keep-row-two--><c r="C2"><v>3</v></c></row></sheetData></worksheet>"#
    );
    let bytes = source_with_worksheet_xml(&source_xml);
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
    // primary value spans change; row tags, comments, value whitespace, and the
    // untouched B1/A2 cell bytes remain exactly as authored above.
    let expected_sheet = format!(
        r#"<worksheet xmlns="{SML}"><sheetData><row r="1" spans="1:3"><c r="A1"><v>10</v></c><!--keep-row-one--><c r='B1' t='b'><v>1</v></c></row><row r="2" spans="1:3"><c r="A2"><v> 2 </v></c><!--keep-row-two--><c r="C2"><v>30</v></c></row></sheetData></worksheet>"#
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

fn source_with_worksheet_xml(source_xml: &str) -> Vec<u8> {
    // Build a ZIP source fixture so the public source-backed path owns
    // worksheet validation. Changed publication requires compact element spacing.
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
    writer.finish_to_bytes().unwrap()
}

#[test]
fn public_multi_edit_publishes_a_producer_formatted_worksheet_verbatim() {
    // Change 0654 loosened the audit of the *original* bytes and recorded
    // that this test still refused, because the value editor's replacement
    // carries the source's own formatting and the replacement audit then
    // raised the byte-identical refusal. Change 0657 declares that
    // replacement source-derived, so the producer's indentation is published
    // rather than refused, and only the edited cell changes.
    let xml = format!(
        r#"<worksheet xmlns="{SML}"><sheetData>
  <row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c></row>
</sheetData></worksheet>"#
    );
    let bytes = source_with_worksheet_xml(&xml);
    let source = Arc::new(VersionedSource::new(bytes.clone()));
    for _ in 0..2 {
        let editor = SourceBackedEditor::from_read_at(source.clone()).unwrap();
        let mut edit = editor.edit_sheets(["Sheet1".into()]).unwrap();
        edit.set("Sheet1", address("A1"), 10u32).unwrap();
        let commit = edit.commit().unwrap();
        assert!(commit.changed());
        let mut published = Vec::new();
        editor
            .publish_multi_commit_to_stream(&mut published, &commit)
            .expect("a source-derived replacement publishes");
        let package = OpcPackage::from_bytes(&published).expect("published package");
        let worksheet = package
            .get_part(&PackURI::new("/xl/worksheets/sheet1.xml").unwrap())
            .expect("published worksheet");
        let text = String::from_utf8(worksheet.blob().to_vec()).expect("UTF-8 worksheet");
        assert!(
            text.contains("<sheetData>\n  <row r=\"1\">"),
            "the producer's indentation is preserved verbatim: {text}"
        );
        assert!(
            text.contains(r#"<c r="B1"><v>2</v></c>"#),
            "an unedited neighbour is preserved verbatim: {text}"
        );
        assert!(text.contains("<v>10</v>"), "the edited value is published");
        assert_eq!(source.bytes.as_slice(), bytes.as_slice());
    }
}

/// Every non-compactness check the publication audit makes still refuses.
#[test]
fn a_source_backed_replacement_still_fails_closed_on_malformed_xml() {
    use litchi_opc::{SourceBackedPackage, SourceTopologyPlan};

    let compact = format!(
        r#"<worksheet xmlns="{SML}"><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>"#
    );
    let bytes = source_with_worksheet_xml(&compact);
    let target = PackURI::new("/xl/worksheets/sheet1.xml").unwrap();
    for (label, replacement) in [
        ("unterminated element", "<worksheet>".to_string()),
        (
            "two document elements",
            format!("<worksheet xmlns=\"{SML}\"/><worksheet xmlns=\"{SML}\"/>"),
        ),
        (
            "document type declaration",
            format!("<!DOCTYPE worksheet><worksheet xmlns=\"{SML}\"/>"),
        ),
    ] {
        let package =
            SourceBackedPackage::from_read_at(Arc::new(VersionedSource::new(bytes.clone())))
                .unwrap();
        let mut plan = SourceTopologyPlan::new();
        plan.try_replace_part(target.clone(), replacement.into_bytes())
            .unwrap();
        let mut published = Vec::new();
        let error = package
            .write_topology_to_stream(&mut published, plan)
            .unwrap_err();
        assert!(
            matches!(error, OpcError::XmlPublication { .. }),
            "{label} must still be refused: {error:?}"
        );
        assert!(published.is_empty(), "{label} must emit no archive");
    }
}
