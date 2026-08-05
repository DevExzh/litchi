use litchi_docx::header_footer::HeaderFooter;
use litchi_docx::header_footer::Kind;

#[test]
fn header_uses_fallback_without_mutating_raw_xml() {
    let raw = br#"<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:unsupported"><mc:AlternateContent><mc:Choice Requires="x"><w:p><w:r><w:t>choice</w:t></w:r></w:p></mc:Choice><mc:Fallback><w:p><w:r><w:t>fallback</w:t></w:r></w:p></mc:Fallback></mc:AlternateContent></w:hdr>"#;
    let header = HeaderFooter::from_xml_bytes(raw.to_vec(), Kind::Primary);
    assert_eq!(header.xml_bytes(), raw);
    assert_eq!(header.text().unwrap(), "fallback");
    assert_eq!(header.paragraph_count().unwrap(), 1);
}
