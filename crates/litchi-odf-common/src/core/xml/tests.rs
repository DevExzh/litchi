use super::{Content, Meta, Styles, XmlPart};
use litchi_core::Result;

#[test]
fn part_from_bytes() -> Result<()> {
    let xml = b"<?xml version=\"1.0\"?><root><child>text</child></root>";
    let part = XmlPart::from_bytes(xml)?;
    assert_eq!(
        part.content(),
        "<?xml version=\"1.0\"?><root><child>text</child></root>"
    );
    Ok(())
}

#[test]
fn part_rejects_invalid_utf8() {
    assert!(XmlPart::from_bytes(&[0x80, 0x81, 0x82]).is_err());
}

#[test]
fn part_exposes_bytes_without_copying() -> Result<()> {
    let part = XmlPart::from_bytes(b"<root/>")?;
    assert_eq!(part.as_bytes(), b"<root/>");
    Ok(())
}

#[test]
fn content_preserves_xml() -> Result<()> {
    let xml = b"<?xml version=\"1.0\"?><office:document-content xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\"><office:body><office:text><text:p>Hello World</text:p></office:text></office:body></office:document-content>";
    let content = Content::from_bytes(xml)?;
    assert!(content.xml_content().contains("Hello World"));
    Ok(())
}

#[test]
fn styles_preserve_xml() -> Result<()> {
    let xml = b"<?xml version=\"1.0\"?><office:document-styles xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\"></office:document-styles>";
    let styles = Styles::from_bytes(xml)?;
    assert!(styles.xml_content().contains("document-styles"));
    Ok(())
}

#[test]
fn meta_preserves_xml() -> Result<()> {
    let xml = b"<?xml version=\"1.0\"?><office:document-meta xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\"></office:document-meta>";
    let meta = Meta::from_bytes(xml)?;
    assert!(meta.xml_content().contains("document-meta"));
    Ok(())
}

#[test]
fn meta_extracts_empty_common_metadata() -> Result<()> {
    let xml = br#"<?xml version="1.0"?>
        <office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
                              xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0"
                              xmlns:dc="http://purl.org/dc/elements/1.1/">
        </office:document-meta>"#;
    let metadata = Meta::from_bytes(xml)?.try_extract_metadata()?;
    assert!(metadata.title.is_none());
    Ok(())
}
