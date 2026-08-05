#[test]
fn alternate_content_selects_the_fallback_branch() {
    let raw = br#"<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:unsupported"><mc:AlternateContent><mc:Choice Requires="x"><x:picture/></mc:Choice><mc:Fallback><w:pict><w:t>fallback-picture</w:t></w:pict></mc:Fallback></mc:AlternateContent></w:r>"#;
    let output = litchi_ooxml_common::mce::process_ooxml(raw).unwrap();
    let semantic = std::str::from_utf8(output.as_ref()).unwrap();
    assert!(semantic.contains("w:pict"));
    assert!(!semantic.contains("x:picture"));
    assert!(!semantic.contains("AlternateContent"));
}
