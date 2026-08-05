use super::{Content, Meta, Styles, XmlPart};

#[test]
fn part_from_bytes() {
    let xml = b"<?xml version=\"1.0\"?><root><child>text</child></root>";
    let part = XmlPart::from_bytes(xml).unwrap();
    assert_eq!(
        part.content(),
        "<?xml version=\"1.0\"?><root><child>text</child></root>"
    );
}

#[test]
fn part_rejects_invalid_utf8() {
    assert!(XmlPart::from_bytes(&[0x80, 0x81, 0x82]).is_err());
}

#[test]
fn part_exposes_bytes_without_copying() {
    let part = XmlPart::from_bytes(b"<root/>").unwrap();
    assert_eq!(part.as_bytes(), b"<root/>");
}

#[test]
fn content_preserves_xml() {
    let xml = b"<?xml version=\"1.0\"?><office:document-content xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\"><office:body><office:text><text:p>Hello World</text:p></office:text></office:body></office:document-content>";
    let content = Content::from_bytes(xml).unwrap();
    assert!(content.xml_content().contains("Hello World"));
}

#[test]
fn styles_preserve_xml() {
    let xml = b"<?xml version=\"1.0\"?><office:document-styles xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\"></office:document-styles>";
    let styles = Styles::from_bytes(xml).unwrap();
    assert!(styles.xml_content().contains("document-styles"));
}

#[test]
fn meta_preserves_xml() {
    let xml = b"<?xml version=\"1.0\"?><office:document-meta xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\"></office:document-meta>";
    let meta = Meta::from_bytes(xml).unwrap();
    assert!(meta.xml_content().contains("document-meta"));
}

#[test]
fn meta_extracts_empty_common_metadata() {
    let xml = br#"<?xml version="1.0"?>
        <office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
                              xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0"
                              xmlns:dc="http://purl.org/dc/elements/1.1/">
        </office:document-meta>"#;
    let metadata = Meta::from_bytes(xml)
        .unwrap()
        .try_extract_metadata()
        .unwrap();
    assert!(metadata.title.is_none());
}
