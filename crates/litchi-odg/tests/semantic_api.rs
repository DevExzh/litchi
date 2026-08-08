#![allow(
    clippy::unwrap_used,
    reason = "test code panics on failure; unwrap keeps assertions concise"
)]

use litchi_odg::{Builder, Drawing, layer::Layer, page::Page};

const COMPACT_CONTENT: &str = concat!(
    r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content "#,
    r#"xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" "#,
    r#"xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" office:version="1.3">"#,
    r#"<office:body><office:drawing><text:p>compact</text:p></office:drawing></office:body>"#,
    r#"</office:document-content>"#,
);

const NONCOMPACT_CONTENT: &str = concat!(
    r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content "#,
    r#"xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" office:version="1.3">"#,
    "\n<office:body><office:drawing/></office:body></office:document-content>",
);

const SEMANTIC_WHITESPACE_CONTENT: &str = concat!(
    r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content "#,
    r#"xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" "#,
    r#"xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" office:version="1.3">"#,
    "<office:body><office:drawing><text:p>line one\n  line two</text:p></office:drawing>",
    "</office:body></office:document-content>",
);

#[test]
fn focused_modules_are_the_canonical_semantic_api() {
    assert_eq!(Page::new("Page 1").name(), "Page 1");
    assert_eq!(Layer::new("Foreground").name(), "Foreground");

    let bytes = Builder::new().build().unwrap();
    let drawing = Drawing::from_bytes(bytes).unwrap();
    assert!(drawing.content_xml().contains("<office:drawing"));
}

#[test]
fn compact_authored_content_is_published_without_rewriting() {
    let drawing =
        Drawing::from_bytes(Builder::new().content_xml(COMPACT_CONTENT).build().unwrap()).unwrap();
    assert_eq!(drawing.content_xml(), COMPACT_CONTENT);
}

#[test]
fn noncompact_authored_content_returns_a_typed_error() {
    assert!(matches!(
        Builder::new()
            .content_xml(NONCOMPACT_CONTENT)
            .build()
            .unwrap_err(),
        litchi_core::Error::XmlCompactness {
            kind: litchi_core::xml::CompactnessKind::FormattingWhitespace,
            ..
        }
    ));
}

#[test]
fn semantic_whitespace_is_preserved_exactly() {
    let drawing = Drawing::from_bytes(
        Builder::new()
            .content_xml(SEMANTIC_WHITESPACE_CONTENT)
            .build()
            .unwrap(),
    )
    .unwrap();
    assert_eq!(drawing.content_xml(), SEMANTIC_WHITESPACE_CONTENT);
}
