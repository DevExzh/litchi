use litchi_ooxml::docx::enums::WdHeaderFooter;
use litchi_ooxml::docx::header_footer::HeaderFooter;
use litchi_ooxml::xlsx::SharedStrings;

#[test]
fn docx_header_uses_fallback_without_mutating_raw_xml() {
    let raw = br#"<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:unsupported"><mc:AlternateContent><mc:Choice Requires="x"><w:p><w:r><w:t>choice</w:t></w:r></w:p></mc:Choice><mc:Fallback><w:p><w:r><w:t>fallback</w:t></w:r></w:p></mc:Fallback></mc:AlternateContent></w:hdr>"#;
    let header = HeaderFooter::from_xml_bytes(raw.to_vec(), WdHeaderFooter::Primary);
    assert_eq!(header.xml_bytes(), raw);
    assert_eq!(header.text().expect("fallback text"), "fallback");
    assert_eq!(header.paragraph_count().expect("fallback paragraph"), 1);
}

#[test]
fn xlsx_shared_strings_uses_fallback() {
    let xml = r#"<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:unsupported" count="1" uniqueCount="1"><mc:AlternateContent><mc:Choice Requires="x"><si><t>choice</t></si></mc:Choice><mc:Fallback><si><t>fallback</t></si></mc:Fallback></mc:AlternateContent></sst>"#;
    let strings = SharedStrings::parse(xml).expect("valid shared strings");
    assert_eq!(strings.get(0), Some("fallback"));
}

#[test]
fn generic_chart_reader_uses_fallback() {
    let xml = br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:unsupported"><mc:AlternateContent><mc:Choice Requires="x"><x:chart/></mc:Choice><mc:Fallback><c:chart/></mc:Fallback></mc:AlternateContent></c:chartSpace>"#;
    litchi_ooxml::charts::reader::parse_chart(xml.as_slice()).expect("fallback chart");
}

#[test]
fn alternate_content_picture_selects_fallback() {
    let xml = br#"<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:unsupported"><mc:AlternateContent><mc:Choice Requires="x"><x:picture/></mc:Choice><mc:Fallback><w:pict><w:t>fallback-picture</w:t></w:pict></mc:Fallback></mc:AlternateContent></w:r>"#;
    let output = litchi_ooxml_common::process_ooxml(xml).expect("valid fallback");
    let semantic = std::str::from_utf8(output.as_ref()).expect("UTF-8 fallback");
    assert!(semantic.contains("w:pict"));
    assert!(!semantic.contains("x:picture"));
    assert!(!semantic.contains("AlternateContent"));
}
