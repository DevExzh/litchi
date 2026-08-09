#![allow(
    clippy::unwrap_used,
    reason = "test code panics on failure; unwrap keeps assertions concise"
)]

use litchi_odi::{Builder, Image, frame::Frame, source::Source};

const COMPACT_CONTENT: &str = concat!(
    r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content "#,
    r#"xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" "#,
    r#"xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" office:version="1.3">"#,
    r#"<office:body><office:image><text:p>compact</text:p></office:image></office:body>"#,
    r#"</office:document-content>"#,
);

const NONCOMPACT_CONTENT: &str = concat!(
    r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content "#,
    r#"xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" office:version="1.3">"#,
    "\n<office:body><office:image/></office:body></office:document-content>",
);

const SEMANTIC_WHITESPACE_CONTENT: &str = concat!(
    r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content "#,
    r#"xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" "#,
    r#"xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" office:version="1.3">"#,
    "<office:body><office:image><text:p>line one\n  line two</text:p></office:image>",
    "</office:body></office:document-content>",
);

#[test]
fn focused_modules_are_the_canonical_semantic_api() {
    let frame = Frame::new(Source::Linked("Pictures/photo.png".into())).with_name("Photo");
    assert_eq!(frame.name(), Some("Photo"));
    assert!(matches!(frame.source(), Source::Linked(_)));

    let bytes = Builder::new().build().unwrap();
    let image = Image::from_bytes(bytes).unwrap();
    assert!(image.content_xml().contains("<office:image"));
}

#[test]
fn compact_authored_content_is_published_without_rewriting() {
    let image =
        Image::from_bytes(Builder::new().content_xml(COMPACT_CONTENT).build().unwrap()).unwrap();
    assert_eq!(image.content_xml(), COMPACT_CONTENT);
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
    let image = Image::from_bytes(
        Builder::new()
            .content_xml(SEMANTIC_WHITESPACE_CONTENT)
            .build()
            .unwrap(),
    )
    .unwrap();
    assert_eq!(image.content_xml(), SEMANTIC_WHITESPACE_CONTENT);
}

#[test]
fn package_edit_is_source_checked_reversible_and_preserves_unknown_content() {
    const CONTENT: &str = concat!(
        r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content "#,
        r#"xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" "#,
        r#"xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" "#,
        r#"xmlns:xlink="http://www.w3.org/1999/xlink" office:version="1.4">"#,
        r#"<office:body><office:image><draw:frame draw:name="before"><draw:image xlink:href="Pictures/before.png"/></draw:frame><foreign:keep xmlns:foreign="urn:example">opaque</foreign:keep></office:image></office:body></office:document-content>"#,
    );
    let source = Image::from_bytes(Builder::new().content_xml(CONTENT).build().unwrap()).unwrap();
    let mut edit = source.edit();
    edit.set_frame_name(0, Some("after".to_string())).unwrap();
    edit.set_source(0, Source::Linked("Pictures/after.png".to_string()))
        .unwrap();
    let commit = edit.commit().unwrap();
    assert!(commit.changed());
    assert_eq!(commit.image().frames()[0].name(), Some("after"));
    assert_eq!(
        commit.image().frames()[0].source(),
        &Source::Linked("Pictures/after.png".to_string())
    );
    assert!(
        commit
            .image()
            .content_xml()
            .contains("<foreign:keep xmlns:foreign=\"urn:example\">opaque</foreign:keep>")
    );
    assert_eq!(commit.patch().changes().len(), 1);
    assert_eq!(
        commit
            .patch()
            .inverse()
            .apply(commit.image())
            .unwrap()
            .as_bytes(),
        source.as_bytes()
    );
}
