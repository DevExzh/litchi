#![allow(
    clippy::unwrap_used,
    reason = "test code panics on failure; unwrap keeps assertions concise"
)]

use litchi_odi::{Builder, Image, frame::Frame, source::Source};

const COMPACT_CONTENT: &str = concat!(
    r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content "#,
    r#"xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" "#,
    r#"xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" "#,
    r#"xmlns:xlink="http://www.w3.org/1999/xlink" office:version="1.4">"#,
    r#"<office:body><office:image><draw:frame><draw:image xlink:href="Pictures/photo.png"/></draw:frame></office:image></office:body>"#,
    r#"</office:document-content>"#,
);

const NONCOMPACT_CONTENT: &str = concat!(
    r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content "#,
    r#"xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" "#,
    r#"xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" "#,
    r#"xmlns:xlink="http://www.w3.org/1999/xlink" office:version="1.4">"#,
    "\n<office:body><office:image><draw:frame><draw:image xlink:href=\"x\"/></draw:frame></office:image></office:body></office:document-content>",
);

const SEMANTIC_WHITESPACE_CONTENT: &str = concat!(
    r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content "#,
    r#"xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" "#,
    r#"xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" "#,
    r#"xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" "#,
    r#"xmlns:xlink="http://www.w3.org/1999/xlink" office:version="1.4">"#,
    "<office:body><office:image><draw:frame><draw:image xlink:href=\"x\"><text:p>line one\n  line two</text:p></draw:image></draw:frame></office:image>",
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
    assert_eq!(image.frames().len(), 1);
    assert!(matches!(image.frames()[0].source(), Source::Embedded(_)));
    assert_eq!(image.frames()[0].media_type(), Some("image/png"));
}

#[test]
fn semantic_frame_authoring_is_minimal_and_preserves_accessibility() {
    let frame = Frame::new(Source::Linked("Pictures/photo&one.png".into()))
        .with_name("Photo & one")
        .with_title("Short & useful")
        .with_description("Longer <description>");
    let image = Image::from_bytes(Builder::new().frame(&frame).build().unwrap()).unwrap();
    assert!(!image.content_xml().contains("> <"));
    assert!(!image.content_xml().contains(">\n<"));
    let actual_frame = &image.frames()[0];
    assert_eq!(actual_frame.name(), Some("Photo & one"));
    assert_eq!(actual_frame.title(), Some("Short & useful"));
    assert_eq!(actual_frame.description(), Some("Longer <description>"));
    assert_eq!(
        actual_frame.source(),
        &Source::Linked("Pictures/photo&one.png".into())
    );
}

#[test]
fn package_resources_are_inert_typed_and_preserved_by_edits() {
    let source = Image::from_bytes(
        Builder::new()
            .frame(&Frame::new(Source::Linked("Pictures/photo.png".into())).with_name("before"))
            .resource("Pictures/photo.png", "image/png", b"png bytes".to_vec())
            .build()
            .unwrap(),
    )
    .unwrap();
    assert_eq!(source.resources().len(), 1);
    let resource = &source.resources()[0];
    assert_eq!(resource.frame(), 0);
    assert_eq!(resource.path(), "Pictures/photo.png");
    assert_eq!(resource.media_type(), Some("image/png"));
    assert!(resource.is_present());
    assert_eq!(
        source.resource_bytes(0).unwrap(),
        Some(b"png bytes".to_vec())
    );

    let mut unchanged = source.edit();
    unchanged
        .set_resource(0, "image/png".into(), b"png bytes".to_vec())
        .unwrap();
    let unchanged_commit = unchanged.commit().unwrap();
    assert!(!unchanged_commit.changed());
    assert_eq!(unchanged_commit.image().as_bytes(), source.as_bytes());

    let mut edit = source.edit();
    edit.set_name(Some("after".into())).unwrap();
    edit.set_resource(0, "image/webp".into(), b"webp bytes".to_vec())
        .unwrap();
    let commit = edit.commit().unwrap();
    assert_eq!(commit.patch().resource_changes().len(), 1);
    assert_eq!(
        commit.patch().resource_changes()[0].path(),
        "Pictures/photo.png"
    );
    assert_eq!(
        commit.patch().resource_changes()[0].before_size(),
        Some(b"png bytes".len())
    );
    let target = commit.image();
    assert_eq!(
        target.resource_bytes(0).unwrap(),
        Some(b"webp bytes".to_vec())
    );
    assert_eq!(target.resources()[0].media_type(), Some("image/webp"));
    assert_eq!(
        commit.patch().inverse().apply(target).unwrap().as_bytes(),
        source.as_bytes()
    );

    let mut removal = target.edit();
    removal.remove_resource(0).unwrap();
    let removed = removal.commit().unwrap();
    assert!(!removed.image().resources()[0].is_present());
    assert_eq!(removed.image().resource_bytes(0).unwrap(), None);
}

#[test]
fn missing_package_resource_remains_visible_without_dereference() {
    let image = Image::from_bytes(
        Builder::new()
            .frame(&Frame::new(Source::Linked("Pictures/missing.png".into())))
            .build()
            .unwrap(),
    )
    .unwrap();
    assert_eq!(image.resources().len(), 1);
    assert!(!image.resources()[0].is_present());
    assert_eq!(image.resource_bytes(0).unwrap(), None);

    let mut edit = image.edit();
    edit.set_resource(0, "image/png".into(), b"created".to_vec())
        .unwrap();
    let created = edit.commit().unwrap();
    assert!(created.image().resources()[0].is_present());
    assert_eq!(
        created.image().resource_bytes(0).unwrap(),
        Some(b"created".to_vec())
    );
}

#[test]
fn package_edit_refuses_every_odf_signature_filename() {
    let source = Image::from_bytes(
        Builder::new()
            .frame(&Frame::new(Source::Linked("photo.png".into())).with_name("before"))
            .resource("photo.png", "image/png", b"png".to_vec())
            .resource(
                "META-INF/vendor-signatures-v2.xml",
                "text/xml",
                br"<signature/>".to_vec(),
            )
            .build()
            .unwrap(),
    )
    .unwrap();
    let mut edit = source.edit();
    edit.set_frame_name(0, Some("after".into())).unwrap();
    assert!(matches!(
        edit.commit(),
        Err(litchi_core::Error::InvalidFormat(message))
            if message.contains("signed packages")
    ));
}

#[test]
fn compact_authored_content_is_published_without_rewriting() {
    let image =
        Image::from_bytes(Builder::new().content_xml(COMPACT_CONTENT).build().unwrap()).unwrap();
    assert_eq!(image.content_xml(), COMPACT_CONTENT);
}

#[test]
fn package_content_accepts_namespace_aliases_without_prefix_guessing() {
    const ALIASED: &str = r#"<?xml version="1.0" encoding="UTF-8"?><o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:d="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:l="http://www.w3.org/1999/xlink" o:version="1.4"><o:body><o:image><d:frame><d:image l:href="photo.png"/></d:frame></o:image></o:body></o:document-content>"#;
    let image = Image::from_bytes(Builder::new().content_xml(ALIASED).build().unwrap()).unwrap();
    assert_eq!(image.content_xml(), ALIASED);
    assert_eq!(
        image.frames()[0].source(),
        &Source::Linked("photo.png".into())
    );
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
