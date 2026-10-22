use super::{Content, Meta, Styles, XmlPart};
use crate::{SemanticLimitKind, SemanticLimits, SemanticLimitsBuilder, XmlStreamLimits};
use litchi_core::Result;

#[test]
fn part_from_bytes() -> Result<()> {
    let xml = b"<?xml version=\"1.0\"?><root><child>text</child></root>";
    let part = XmlPart::from_bytes(xml, XmlStreamLimits::default())?;
    assert_eq!(
        part.content(),
        "<?xml version=\"1.0\"?><root><child>text</child></root>"
    );
    Ok(())
}

#[test]
fn part_rejects_invalid_utf8() {
    assert!(XmlPart::from_bytes(&[0x80, 0x81, 0x82], XmlStreamLimits::default()).is_err());
}

#[test]
fn part_exposes_bytes_without_copying() -> Result<()> {
    let part = XmlPart::from_bytes(b"<root/>", XmlStreamLimits::default())?;
    assert_eq!(part.as_bytes(), b"<root/>");
    Ok(())
}

#[test]
fn content_preserves_xml() -> Result<()> {
    let xml = b"<?xml version=\"1.0\"?><office:document-content xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\"><office:body><office:text><text:p>Hello World</text:p></office:text></office:body></office:document-content>";
    let content = Content::from_bytes(xml, SemanticLimits::DEFAULT)?;
    assert!(content.xml_content().contains("Hello World"));
    Ok(())
}

#[test]
fn styles_preserve_xml() -> Result<()> {
    let xml = b"<?xml version=\"1.0\"?><office:document-styles xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\"></office:document-styles>";
    let styles = Styles::from_bytes(xml, SemanticLimits::DEFAULT)?;
    assert!(styles.xml_content().contains("document-styles"));
    Ok(())
}

#[test]
fn meta_preserves_xml() -> Result<()> {
    let xml = b"<?xml version=\"1.0\"?><office:document-meta xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\"></office:document-meta>";
    let meta = Meta::from_bytes(xml, XmlStreamLimits::default())?;
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
    let metadata = Meta::from_bytes(xml, XmlStreamLimits::default())?.try_extract_metadata()?;
    assert!(metadata.title.is_none());
    Ok(())
}

#[test]
fn part_admission_uses_the_caller_bound_at_its_origin() {
    let xml = b"<root/>";
    let exact = SemanticLimits::builder()
        .max_content_bytes(xml.len() as u64)
        .and_then(SemanticLimitsBuilder::build)
        .expect("exact content bound");
    assert!(Content::from_vec(xml.to_vec(), exact).is_ok());
    let short = SemanticLimits::builder()
        .max_content_bytes(xml.len() as u64 - 1)
        .and_then(SemanticLimitsBuilder::build)
        .expect("short content bound");
    match Content::from_vec(xml.to_vec(), short) {
        Err(crate::Error::Limit {
            kind: crate::LimitKind::Semantic(SemanticLimitKind::ContentBytes),
            observed: 7,
            maximum: 6,
        }) => {},
        other => panic!("unexpected content admission: {other:?}"),
    }
    // A styles bound does not govern content and vice versa.
    assert!(Styles::from_vec(xml.to_vec(), short).is_ok());
    let raw = XmlStreamLimits::new(6, 8, 8, 8, 8).expect("raw bound");
    match Meta::from_bytes(xml, raw) {
        Err(crate::Error::Limit {
            kind: crate::LimitKind::RawXml(crate::RawXmlLimitKind::InputBytes),
            observed: 7,
            maximum: 6,
        }) => {},
        other => panic!("unexpected meta admission: {other:?}"),
    }
    match XmlPart::from_vec(xml.to_vec(), raw) {
        Err(crate::Error::Limit { observed: 7, .. }) => {},
        other => panic!("unexpected part admission: {other:?}"),
    }
}
