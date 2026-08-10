#![allow(
    clippy::unwrap_used,
    reason = "test code panics on failure; unwrap keeps assertions concise"
)]

use litchi_odm::{
    Builder, Master,
    section::Section,
    structure::{IndexKind, Kind},
    subdocument::Subdocument,
    transaction::BodyItemSpec,
};

const COMPACT_CONTENT: &str = concat!(
    r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content "#,
    r#"xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" "#,
    r#"xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" office:version="1.3">"#,
    r#"<office:body><office:text><text:p>compact</text:p></office:text></office:body>"#,
    r#"</office:document-content>"#,
);

const NONCOMPACT_CONTENT: &str = concat!(
    r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content "#,
    r#"xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" office:version="1.3">"#,
    "\n<office:body><office:text/></office:body></office:document-content>",
);

const SEMANTIC_WHITESPACE_CONTENT: &str = concat!(
    r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content "#,
    r#"xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" "#,
    r#"xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" office:version="1.3">"#,
    "<office:body><office:text><text:p>line one\n  line two</text:p></office:text>",
    "</office:body></office:document-content>",
);

#[test]
fn focused_modules_are_the_canonical_semantic_api() {
    let mut section = Section::new("Introduction");
    section.push(Subdocument::new("chapter-1.odt"));
    assert_eq!(section.children()[0].href(), "chapter-1.odt");

    let bytes = Builder::new().build().unwrap();
    let master = Master::from_bytes(bytes).unwrap();
    assert!(master.content_xml().contains("<office:text"));
}

#[test]
fn compact_authored_content_is_published_without_rewriting() {
    let master =
        Master::from_bytes(Builder::new().content_xml(COMPACT_CONTENT).build().unwrap()).unwrap();
    assert_eq!(master.content_xml(), COMPACT_CONTENT);
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
    let master = Master::from_bytes(
        Builder::new()
            .content_xml(SEMANTIC_WHITESPACE_CONTENT)
            .build()
            .unwrap(),
    )
    .unwrap();
    assert_eq!(master.content_xml(), SEMANTIC_WHITESPACE_CONTENT);
}

#[test]
fn typed_builder_authors_common_master_body_content() {
    let bytes = Builder::new()
        .body_item(BodyItemSpec::paragraph("Introduction & scope").unwrap())
        .body_item(BodyItemSpec::heading(1, "Part one").unwrap())
        .body_item(BodyItemSpec::list(vec!["First".to_string(), "Second".to_string()]).unwrap())
        .body_item(
            BodyItemSpec::table(
                "Summary",
                vec![
                    vec!["Key".to_string(), "Value".to_string()],
                    vec!["Mode".to_string(), "Exact".to_string()],
                ],
            )
            .unwrap(),
        )
        .body_item(BodyItemSpec::generated_index(IndexKind::TableOfContents, "Contents").unwrap())
        .build()
        .unwrap();
    let master = Master::from_bytes(bytes).unwrap();
    assert_eq!(
        master.structure().items(),
        &[
            Kind::Paragraph,
            Kind::Heading,
            Kind::List,
            Kind::Table,
            Kind::GeneratedIndex(IndexKind::TableOfContents),
        ]
    );
    assert!(master.content_xml().contains("Introduction &amp; scope"));
    assert!(!master.content_xml().contains('\n'));
    assert!(BodyItemSpec::heading(0, "invalid").is_err());
    assert!(BodyItemSpec::list(Vec::new()).is_err());
    assert!(
        BodyItemSpec::table(
            "Ragged",
            vec![
                vec!["one".to_string()],
                vec!["two".to_string(), "three".to_string()]
            ],
        )
        .is_err()
    );
}
