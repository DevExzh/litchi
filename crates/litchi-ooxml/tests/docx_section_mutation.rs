use litchi_ooxml::docx::{
    Package, SectionProperties, WdHeaderFooter, WdSectionStart,
};
use std::io::Cursor;

const TRANSITIONAL: &str = r#"<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:test"><w:body><w:p><w:pPr><w:sectPr><w:type w:val="continuous"/><w:pgSz w:w="12240" w:h="15840"/><mc:AlternateContent><mc:Fallback><x:opaque x:value="keep"/></mc:Fallback></mc:AlternateContent></w:sectPr></w:pPr><w:r><w:t>one</w:t></w:r></w:p><w:p><w:r><w:t>two</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>"#;

#[test]
fn mutates_moves_and_removes_paragraph_section_breaks_without_losing_mce() {
    let mut document = litchi_ooxml::docx::MutableDocument::from_xml(TRANSITIONAL).unwrap();
    assert_eq!(document.section_break_count().unwrap(), 1);
    assert_eq!(
        document.section_break(0).unwrap().start_type,
        Some(WdSectionStart::Continuous)
    );

    document
        .update_section_break(0, |section| section.margin_left = 900)
        .unwrap();
    document.move_section_break(0, 1).unwrap();
    let output = document.to_xml().unwrap();
    assert!(output.contains("w:left=\"900\""));
    assert!(output.contains("<x:opaque x:value=\"keep\"/>"));
    assert_eq!(document.section_break_count().unwrap(), 1);

    document.remove_section_break(0).unwrap();
    assert_eq!(document.section_break_count().unwrap(), 0);
}

#[test]
fn accepts_strict_sections_and_rejects_malformed_body_final_placement() {
    let strict = r#"<s:document xmlns:s="http://purl.oclc.org/ooxml/wordprocessingml/main"><s:body><s:p/><s:sectPr><s:type s:val="oddPage"/></s:sectPr></s:body></s:document>"#;
    let document = litchi_ooxml::docx::MutableDocument::from_xml(strict).unwrap();
    assert_eq!(document.section().start_type, Some(WdSectionStart::OddPage));
    assert!(document.to_xml().unwrap().contains("<s:type s:val=\"oddPage\"/>"));

    let nonfinal = r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:sectPr/><w:p/></w:body></w:document>"#;
    assert!(litchi_ooxml::docx::MutableDocument::from_xml(nonfinal).is_err());
    let duplicate = r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:sectPr/><w:sectPr/></w:body></w:document>"#;
    assert!(litchi_ooxml::docx::MutableDocument::from_xml(duplicate).is_err());
    let out_of_order = r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:sectPr><w:pgMar/><w:pgSz/></w:sectPr></w:body></w:document>"#;
    assert!(litchi_ooxml::docx::MutableDocument::from_xml(out_of_order).is_err());
}

#[test]
fn package_roundtrip_keeps_distinct_section_header_footer_parts() {
    let mut package = Package::new().unwrap();
    {
        let document = package.document_mut().unwrap();
        document.add_paragraph_with_text("first");
        document.add_paragraph_with_text("second");
        let mut first = SectionProperties::default().with_start_type(WdSectionStart::NewPage);
        first
            .set_header_part(
                WdHeaderFooter::Primary,
                "first-header",
                r#"<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>First header</w:t></w:r></w:p></w:hdr>"#,
            )
            .unwrap();
        document.insert_section_break(0, first).unwrap();
        document
            .section_mut()
            .set_header_part(
                WdHeaderFooter::Primary,
                "last-header",
                r#"<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>Last header</w:t></w:r></w:p></w:hdr>"#,
            )
            .unwrap();
        document
            .section_mut()
            .set_footer_part(
                WdHeaderFooter::Primary,
                "last-footer",
                r#"<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>Last footer</w:t></w:r></w:p></w:ftr>"#,
            )
            .unwrap();
    }

    let mut bytes = Cursor::new(Vec::new());
    package.to_stream(&mut bytes).unwrap();
    bytes.set_position(0);
    let reopened = Package::from_reader(bytes).unwrap();
    assert_eq!(reopened.document().unwrap().sections().unwrap().len(), 2);
}
