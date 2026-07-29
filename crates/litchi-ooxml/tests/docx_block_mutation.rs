use litchi_ooxml::docx::{
    MutableDocument, Package, SectionProperties, WdSectionStart,
};
use litchi_ooxml::docx::writer::TableOfContents;
use litchi_opc::packuri::PackURI;
use std::io::Cursor;

const RAW: &str = r#"<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>alpha</w:t></w:r></w:p><w:p><w:r><w:t>beta</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>"#;

#[test]
fn inserts_and_removes_paragraphs_and_tables_then_reparses() {
    let mut package = Package::new().unwrap();
    {
        let doc = package.document_mut().unwrap();
        doc.add_paragraph_with_text("one");
        doc.add_paragraph_with_text("two");
        doc.add_paragraph_with_text("three");

        doc.insert_paragraph(1)
            .unwrap()
            .add_run()
            .set_text("inserted");
        assert_eq!(doc.paragraph_count(), 4);
        doc.insert_table(0, 2, 3).unwrap();
        assert_eq!(doc.table_count(), 1);

        let xml = doc.to_xml().unwrap();
        // With no existing tables, the table anchors at the end of the body.
        assert!(xml.find("three").unwrap() < xml.find("<w:tbl").unwrap());
        assert!(xml.find("one").unwrap() < xml.find("inserted").unwrap());
        assert!(xml.find("inserted").unwrap() < xml.find("two").unwrap());

        doc.remove_paragraph(1).unwrap();
        assert_eq!(doc.paragraph_count(), 3);
        doc.remove_table(0).unwrap();
        assert_eq!(doc.table_count(), 0);
        assert!(!doc.to_xml().unwrap().contains("inserted"));
    }

    let mut bytes = Cursor::new(Vec::new());
    package.to_stream(&mut bytes).unwrap();
    bytes.set_position(0);
    let reopened = Package::from_reader(bytes).unwrap();
    let text = reopened.document().unwrap().text().unwrap();
    assert!(text.find("one").unwrap() < text.find("two").unwrap());
    assert!(text.find("two").unwrap() < text.find("three").unwrap());
    assert!(!text.contains("inserted"));
}

#[test]
fn rejects_out_of_range_insertion_and_removal() {
    let mut doc = MutableDocument::new();
    doc.add_paragraph_with_text("only");
    assert!(doc.insert_paragraph(2).is_err());
    assert!(doc.remove_paragraph(1).is_err());
    assert!(doc.insert_table(1, 1, 1).is_err());
    assert!(doc.remove_table(0).is_err());
    assert_eq!(doc.paragraph_count(), 1);
    assert!(!doc.to_xml().unwrap().is_empty());
}

#[test]
fn mutates_blocks_around_preserved_paragraphs_and_final_section() {
    let mut doc = MutableDocument::from_xml(RAW).unwrap();
    assert_eq!(doc.paragraph_count(), 2);

    // Append lands before the body-final sectPr.
    doc.insert_paragraph(2).unwrap().add_run().set_text("omega");
    // Insert between two preserved paragraphs.
    doc.insert_paragraph(1).unwrap().add_run().set_text("middle");
    // Remove a preserved paragraph.
    doc.remove_paragraph(0).unwrap();
    assert_eq!(doc.paragraph_count(), 3);

    let xml = doc.to_xml().unwrap();
    assert!(!xml.contains("alpha"));
    assert!(xml.find("middle").unwrap() < xml.find("beta").unwrap());
    assert!(xml.find("beta").unwrap() < xml.find("omega").unwrap());
    assert!(xml.find("omega").unwrap() < xml.find("<w:sectPr>").unwrap());
    assert!(xml.trim_end().ends_with("</w:document>"));

    // The result still parses and validates section placement.
    let reparsed = MutableDocument::from_xml(&xml).unwrap();
    assert_eq!(reparsed.paragraph_count(), 3);
    assert_eq!(reparsed.section().page_width, 12240);
}

#[test]
fn removing_a_section_ending_paragraph_merges_sections() {
    let mut doc = MutableDocument::new();
    doc.add_paragraph_with_text("first");
    doc.add_paragraph_with_text("second");
    doc.insert_section_break(
        0,
        SectionProperties::default().with_start_type(WdSectionStart::NewPage),
    )
    .unwrap();
    assert_eq!(doc.section_break_count().unwrap(), 1);

    doc.remove_paragraph(0).unwrap();
    assert_eq!(doc.section_break_count().unwrap(), 0);
    assert_eq!(doc.paragraph_count(), 1);
    assert!(!doc.to_xml().unwrap().contains("first"));
}

#[test]
fn structural_edits_keep_the_pending_toc_anchored_after_existing_content() {
    let mut package = Package::new().unwrap();
    {
        let doc = package.document_mut().unwrap();
        doc.add_paragraph_with_text("body-text");
        doc.add_toc(TableOfContents::new()).unwrap();
        doc.insert_paragraph(0)
            .unwrap()
            .add_run()
            .set_text("front-text");
    }

    let mut bytes = Cursor::new(Vec::new());
    package.to_stream(&mut bytes).unwrap();
    bytes.set_position(0);
    let reopened = Package::from_reader(bytes).unwrap();
    let part = reopened
        .opc_package()
        .get_part(&PackURI::new("/word/document.xml").unwrap())
        .unwrap();
    let xml = std::str::from_utf8(part.blob()).unwrap();
    let front = xml.find("front-text").unwrap();
    let body = xml.find("body-text").unwrap();
    let toc = xml.find("TOC ").unwrap();
    assert!(front < body, "inserted paragraph must lead the body");
    assert!(body < toc, "TOC must stay anchored after existing content");
}
