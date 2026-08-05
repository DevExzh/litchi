//! Focused tests for the semantic annotation model and bounded XML codec.

use super::*;

#[test]
fn builds_rich_annotation_and_escapes_values() {
    let mut annotation = Annotation::new("first & <line>");
    annotation.set_creator(Some("A&B"));
    annotation.set_display(Some(true));
    annotation.set_attribute("svg:width", "3\" & 4cm").unwrap();

    let mut span = Element::new("text:span").unwrap();
    span.set_attribute("text:style-name", "Strong").unwrap();
    span.push_text("bold");
    let mut paragraph = Element::new("text:p").unwrap();
    paragraph.push_element(span);
    paragraph.push_element(Element::new("text:line-break").unwrap());
    paragraph.push_text("after");
    annotation.push_element(paragraph);

    assert_eq!(annotation.creator().as_deref(), Some("A&B"));
    assert_eq!(annotation.display(), Some(true));
    assert_eq!(annotation.text(), "first & <line>\nbold\nafter");

    let mut xml = String::new();
    annotation.write_xml(&mut xml);
    assert!(xml.contains("office:display=\"true\""));
    assert!(
        xml.contains("svg:width=\"3&amp;quot; &amp; 4cm\"")
            || xml.contains("svg:width=\"3&quot; &amp; 4cm\"")
    );
    assert!(xml.contains("first &amp; &lt;line&gt;"));
}

#[test]
fn rejects_names_that_could_inject_xml() {
    assert!(Element::new("text:p><evil").is_err());
    let mut annotation = Annotation::default();
    assert!(annotation.set_attribute("x\" y", "value").is_err());
    assert!(annotation.set_attribute("xmlns:evil", "urn:evil").is_err());
}

#[test]
fn validates_custom_extension_namespaces() {
    let mut annotation = Annotation::new("root");
    annotation.push_element(Element::new("vendor:thread").unwrap());
    assert!(annotation.validate().is_err());

    annotation
        .set_namespace("vendor", "urn:example:annotation")
        .unwrap();
    annotation.validate().unwrap();
    let mut xml = String::new();
    annotation.write_xml(&mut xml);
    assert!(xml.contains("xmlns:vendor=\"urn:example:annotation\""));
    assert!(xml.contains("<vendor:thread/>"));
}

#[test]
fn reads_legacy_annotation_metadata_attributes() {
    let mut annotation = Annotation::default();
    annotation
        .set_attribute("office:author", "Legacy Author")
        .unwrap();
    annotation
        .set_attribute("office:create-date", "2002-01-01T00:00:00")
        .unwrap();
    annotation
        .set_attribute("office:create-date-string", "January 1, 2002")
        .unwrap();

    assert_eq!(annotation.creator().as_deref(), Some("Legacy Author"));
    assert_eq!(annotation.date().as_deref(), Some("2002-01-01T00:00:00"));
    assert_eq!(annotation.date_string().as_deref(), Some("January 1, 2002"));
}

#[test]
fn builder_enforces_annotation_nesting_limit() {
    let root = quick_xml::events::BytesStart::new("office:annotation");
    let reader = quick_xml::reader::NsReader::from_str("");
    let decoder = reader.decoder();
    let mut builder = Builder::new(&root, decoder, std::collections::BTreeMap::new()).unwrap();

    for _ in 0..package::MAX_ANNOTATION_NESTING {
        let start = quick_xml::events::BytesStart::new("text:p");
        builder.start(&start, decoder).unwrap();
    }

    let start = quick_xml::events::BytesStart::new("text:p");
    assert!(builder.start(&start, decoder).is_err());
}
