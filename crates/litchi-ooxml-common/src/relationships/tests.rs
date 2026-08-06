use super::{Id, STRICT_NAMESPACE, TRANSITIONAL_NAMESPACE, attribute_id, attribute_value};
use quick_xml::reader::NsReader;

#[test]
fn accepts_transitional_and_strict_relationship_attributes() {
    for (xml, expected) in [
        (
            format!(
                r#"<p:item xmlns:p="urn:test" xmlns:r="{}" r:id="rId1"/>"#,
                String::from_utf8_lossy(TRANSITIONAL_NAMESPACE)
            ),
            "rId1",
        ),
        (
            format!(
                r#"<p:item xmlns:p="urn:test" xmlns:r="{}" r:id="rId2"/>"#,
                String::from_utf8_lossy(STRICT_NAMESPACE)
            ),
            "rId2",
        ),
    ] {
        let mut reader = NsReader::from_reader(xml.as_bytes());
        let (_, event) = reader.read_resolved_event().expect("item");
        let quick_xml::events::Event::Empty(element) = event else {
            panic!("expected empty item");
        };
        assert_eq!(
            attribute_value(&element, b"id", reader.decoder(), reader.resolver())
                .expect("relationship attribute")
                .as_deref(),
            Some(expected)
        );
        assert_eq!(
            attribute_id(&element, b"id", reader.decoder(), reader.resolver())
                .expect("typed relationship attribute")
                .as_ref()
                .map(Id::as_str),
            Some(expected)
        );
    }
}

#[test]
fn rejects_duplicate_relationship_attributes() {
    let xml = format!(
        r#"<p:item xmlns:p="urn:test" xmlns:r="{}" xmlns:q="{}" r:id="one" q:id="two"/>"#,
        String::from_utf8_lossy(TRANSITIONAL_NAMESPACE),
        String::from_utf8_lossy(TRANSITIONAL_NAMESPACE)
    );
    let mut reader = NsReader::from_reader(xml.as_bytes());
    let (_, event) = reader.read_resolved_event().expect("item");
    let quick_xml::events::Event::Empty(element) = event else {
        panic!("expected empty item");
    };
    assert!(matches!(
        attribute_value(&element, b"id", reader.decoder(), reader.resolver()),
        Err(crate::XmlError::Invalid(message)) if message.contains("duplicate")
    ));
}

#[test]
fn relationship_ids_use_the_xml_ncname_domain() {
    let valid = Id::new("rId42").expect("valid ID");
    assert_eq!(valid.as_str(), "rId42");
    assert!(Id::new("").is_err());
    assert!(Id::new("42-id").is_err());
    assert!(Id::new("relationship id").is_err());
}
