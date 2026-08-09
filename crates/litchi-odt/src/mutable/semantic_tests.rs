//! Regression coverage for contextual mutable-document views and inline text edits.

use super::MutableDocument;
use crate::Document;
use crate::core::PackageWriter;

#[test]
fn contextual_views_keep_legacy_operations_and_separate_content_from_styles() {
    let mut document = MutableDocument::new();
    document.content_mut().add_paragraph("body").unwrap();
    document.content_mut().add_heading("title", 1).unwrap();
    document
        .styles_mut()
        .add_master_page("Standard", "pm1")
        .unwrap();

    assert_eq!(document.content().paragraphs().len(), 1);
    assert_eq!(document.content().headings().len(), 1);
    assert_eq!(document.styles().master_pages().unwrap().len(), 1);
    assert_eq!(document.styles().page_layouts().unwrap().len(), 1);

    document
        .content_mut()
        .replace_paragraph_at(litchi_core::Position::new(0), "changed")
        .unwrap();
    assert_eq!(document.paragraphs()[0].text().unwrap(), "changed");
}

#[test]
fn contextual_content_editor_can_reborrow_for_reads() {
    let mut document = MutableDocument::new();
    document.add_paragraph("one").unwrap();

    let mut content = document.content_mut();
    assert_eq!(content.read().paragraphs().len(), 1);
    content.add_paragraph("two").unwrap();
    assert_eq!(content.read().paragraphs().len(), 2);
}

#[test]
fn appending_line_break_uses_odf_inline_markup_and_preserves_rich_content() {
    let mut document = MutableDocument::new();
    document.add_paragraph("before").unwrap();
    document
        .append_line_break_at(litchi_core::Position::new(0))
        .unwrap();

    let output = Document::from_bytes(document.to_bytes().unwrap()).unwrap();
    assert_eq!(output.paragraphs().unwrap()[0].text().unwrap(), "before\n");
    let content = String::from_utf8(output.get_file("content.xml").unwrap()).unwrap();
    assert!(content.contains("<text:line-break/>"));

    let rich_xml = r#"<?xml version="1.0"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><office:body><office:text><t:p><t:span t:style-name="Emphasis">rich</t:span></t:p></office:text></office:body></office:document-content>"#;
    let mut writer = PackageWriter::new();
    writer
        .set_mimetype("application/vnd.oasis.opendocument.text")
        .unwrap();
    writer.add_file("content.xml", rich_xml.as_bytes()).unwrap();
    let source = Document::from_bytes(writer.finish_to_bytes().unwrap()).unwrap();
    let mut mutable = MutableDocument::from_document(source).unwrap();
    mutable
        .append_line_break_at(litchi_core::Position::new(0))
        .unwrap();

    let round_trip = Document::from_bytes(mutable.to_bytes().unwrap()).unwrap();
    assert_eq!(
        round_trip.paragraphs().unwrap()[0].text().unwrap(),
        "rich\n"
    );
    let content = String::from_utf8(round_trip.get_file("content.xml").unwrap()).unwrap();
    assert!(content.contains("<t:span t:style-name=\"Emphasis\">rich</t:span>"));
    assert!(content.contains("<t:line-break/>"));
}

#[test]
fn appending_line_break_reports_a_missing_document_paragraph() {
    let mut document = MutableDocument::new();
    document.add_paragraph("only").unwrap();
    assert!(
        document
            .append_line_break_at(litchi_core::Position::new(1))
            .is_err()
    );
}
