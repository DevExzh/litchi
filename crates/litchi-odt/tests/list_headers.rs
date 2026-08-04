//! Unnumbered list headers (`text:list-header`).
//!
//! ODF lets a list open with a single `text:list-header` whose paragraphs get
//! no bullet or number. It is authored content exactly like a list item, so it
//! must be reachable from the structured list model and not only through the
//! flattened text.

use litchi_odt::elements::parser::{OrderElement, Parser};
use litchi_odt::elements::text::TextElements;
use litchi_odt::elements::text::{ListHeader, Paragraph};

/// A list whose first block is an unnumbered header followed by two items.
const LIST_XML: &str = concat!(
    r#"<office:document-content"#,
    r#" xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0""#,
    r#" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">"#,
    "<office:body><office:text><text:list>",
    "<text:list-header><text:p>lead in</text:p><text:p>still lead in</text:p></text:list-header>",
    "<text:list-item><text:p>one</text:p></text:list-item>",
    "<text:list-item><text:p>two</text:p></text:list-item>",
    "</text:list></office:text></office:body></office:document-content>",
);

fn parsed_list() -> litchi_odt::elements::text::List {
    let elements = Parser::parse_elements_in_order(LIST_XML).unwrap();
    elements
        .into_iter()
        .find_map(|element| match element {
            OrderElement::List(list) => Some(list),
            _ => None,
        })
        .expect("document contains one list")
}

#[test]
fn a_list_header_is_reported_separately_from_the_items() {
    let list = parsed_list();

    let items = list.items().unwrap();
    assert_eq!(items.len(), 2);
    assert_eq!(items[0].text().unwrap(), "one");

    let header = list.header().unwrap().expect("list has a header");
    let paragraphs = header.paragraphs().unwrap();
    assert_eq!(paragraphs.len(), 2);
    assert_eq!(paragraphs[0].text().unwrap(), "lead in");
    assert_eq!(paragraphs[1].text().unwrap(), "still lead in");
}

#[test]
fn header_text_also_reaches_the_flattened_document_text() {
    let text = TextElements::extract_text(LIST_XML).unwrap();
    assert_eq!(text, "lead in\nstill lead in\none\ntwo");
}

#[test]
fn a_list_without_a_header_reports_none() {
    let xml = LIST_XML.replace(
        "<text:list-header><text:p>lead in</text:p><text:p>still lead in</text:p></text:list-header>",
        "",
    );
    let elements = Parser::parse_elements_in_order(&xml).unwrap();
    let list = elements
        .into_iter()
        .find_map(|element| match element {
            OrderElement::List(list) => Some(list),
            _ => None,
        })
        .expect("document contains one list");

    assert!(list.header().unwrap().is_none());
}

#[test]
fn setting_a_header_replaces_the_existing_one_and_keeps_it_first() {
    let mut list = parsed_list();

    let mut header = ListHeader::new();
    let mut paragraph = Paragraph::new();
    paragraph.set_text("replaced");
    header.add_paragraph(paragraph);
    list.set_header(header);

    let stored = list.header().unwrap().expect("header was set");
    assert_eq!(stored.paragraphs().unwrap().len(), 1);
    assert_eq!(stored.text().unwrap(), "replaced");
    // Replacing the header must not disturb the items.
    assert_eq!(list.items().unwrap().len(), 2);
}

#[test]
fn a_foreign_element_cannot_be_wrapped_as_a_list_header() {
    let paragraph: litchi_odt::elements::element::Element = Paragraph::new().into();
    assert!(ListHeader::from_element(paragraph).is_err());
}
