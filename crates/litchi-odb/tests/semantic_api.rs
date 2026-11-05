#![allow(
    clippy::unwrap_used,
    reason = "test code panics on failure; unwrap keeps assertions concise"
)]

use litchi_odb::{Builder, connection::Connection, query::Query};

const COMPACT_CONTENT: &str = concat!(
    r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content "#,
    r#"xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" "#,
    r#"xmlns:db="urn:oasis:names:tc:opendocument:xmlns:database:1.0" "#,
    r#"xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" office:version="1.3">"#,
    r#"<office:body><office:database><db:data-source/><text:p>compact</text:p></office:database></office:body>"#,
    r#"</office:document-content>"#,
);

const NONCOMPACT_CONTENT: &str = concat!(
    r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content "#,
    r#"xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" office:version="1.3">"#,
    "\n<office:body><office:database/></office:body></office:document-content>",
);

const SEMANTIC_WHITESPACE_CONTENT: &str = concat!(
    r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content "#,
    r#"xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" "#,
    r#"xmlns:db="urn:oasis:names:tc:opendocument:xmlns:database:1.0" "#,
    r#"xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" office:version="1.3">"#,
    "<office:body><office:database><db:data-source/><text:p>line one\n  line two</text:p></office:database>",
    "</office:body></office:document-content>",
);

#[test]
fn focused_modules_are_the_canonical_semantic_api() {
    let query = Query::new("recent", "SELECT * FROM records");
    assert_eq!(query.name(), "recent");
    assert_eq!(query.command(), "SELECT * FROM records");
    assert!(matches!(
        Connection::File("records.db".into()),
        Connection::File(_)
    ));

    let database = Builder::new().build().unwrap();
    assert!(
        database
            .database()
            .content_xml()
            .contains("<office:database")
    );
}

#[test]
fn compact_authored_content_is_published_without_rewriting() {
    let database = Builder::new().content_xml(COMPACT_CONTENT).build().unwrap();
    assert_eq!(database.database().content_xml(), COMPACT_CONTENT);
}

#[test]
fn noncompact_authored_content_returns_a_typed_error() {
    assert!(matches!(
        Builder::new()
            .content_xml(NONCOMPACT_CONTENT)
            .build()
            .unwrap_err(),
        litchi_odb::Error::Core(litchi_core::Error::XmlCompactness {
            kind: litchi_core::xml::CompactnessKind::FormattingWhitespace,
            ..
        })
    ));
}

#[test]
fn semantic_whitespace_is_preserved_exactly() {
    let database = Builder::new()
        .content_xml(SEMANTIC_WHITESPACE_CONTENT)
        .build()
        .unwrap();
    assert_eq!(
        database.database().content_xml(),
        SEMANTIC_WHITESPACE_CONTENT
    );
}
