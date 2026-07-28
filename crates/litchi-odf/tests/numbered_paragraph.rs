//! Tests for `text:numbered-paragraph` extraction and its inert numbering
//! attributes.

use litchi_odf::DocumentParser;
use litchi_odf::elements::element::ElementBase;
use litchi_odf::elements::parser::DocumentOrderElement;

const CONTENT: &str = concat!(
    r#"<?xml version="1.0"?><office:document-content "#,
    r#"xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" "#,
    r#"xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" "#,
    r#"office:version="1.3"><office:body><office:text>"#,
    r#"<text:p>Before</text:p>"#,
    r#"<text:numbered-paragraph text:style-name="P1" text:level="2" text:list-id="list1" text:start-value="5">"#,
    r#"Numbered <text:span text:style-name="T1">item</text:span></text:numbered-paragraph>"#,
    r#"<text:p>After</text:p>"#,
    r#"</office:text></office:body></office:document-content>"#,
);

#[test]
fn numbered_paragraphs_are_extracted_with_numbering_attributes() {
    let elements = DocumentParser::parse_elements_in_order(CONTENT).unwrap();
    assert_eq!(elements.len(), 3);
    let DocumentOrderElement::NumberedParagraph(para) = &elements[1] else {
        panic!("expected a numbered paragraph, got {:?}", elements.len())
    };
    assert_eq!(para.text().unwrap(), "Numbered item");
    assert_eq!(para.style_name(), Some("P1"));
    assert_eq!(para.level().unwrap().unwrap(), 2);
    assert_eq!(para.list_id(), Some("list1"));
    assert_eq!(para.start_value().unwrap().unwrap(), 5);
    assert!(para.element().tag_name() == "text:numbered-paragraph");
}

#[test]
fn numbered_paragraphs_convert_to_plain_paragraphs() {
    let elements = DocumentParser::parse_elements_in_order(CONTENT).unwrap();
    let DocumentOrderElement::NumberedParagraph(para) = elements.into_iter().nth(1).unwrap()
    else {
        panic!()
    };
    let plain = para.into_paragraph();
    assert_eq!(plain.text().unwrap(), "Numbered item");
}

#[test]
fn surrounding_paragraphs_are_unaffected() {
    let elements = DocumentParser::parse_elements_in_order(CONTENT).unwrap();
    let DocumentOrderElement::Paragraph(before) = &elements[0] else {
        panic!()
    };
    let DocumentOrderElement::Paragraph(after) = &elements[2] else {
        panic!()
    };
    assert_eq!(before.text().unwrap(), "Before");
    assert_eq!(after.text().unwrap(), "After");
}

#[test]
fn invalid_numbering_attributes_are_reported() {
    let bad = CONTENT.replace("text:level=\"2\"", "text:level=\"two\"");
    let elements = DocumentParser::parse_elements_in_order(&bad).unwrap();
    let DocumentOrderElement::NumberedParagraph(para) = &elements[1] else {
        panic!()
    };
    assert!(para.level().unwrap().is_err());
}
