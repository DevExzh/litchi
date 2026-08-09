use litchi_core::Error;
use litchi_odp::FlatPresentation;

const FLAT: &str = r#"<?xml version="1.0" encoding="UTF-8"?><office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" office:mimetype="application/vnd.oasis.opendocument.presentation" office:version="1.3"><office:automatic-styles><style:style style:name="dp1" style:family="drawing-page"><style:drawing-page-properties draw:fill="bitmap"><style:background-image xlink:href="https://example.invalid/background.png"/></style:drawing-page-properties></style:style></office:automatic-styles><office:body><office:presentation><draw:page draw:name="page1" draw:style-name="dp1"><draw:frame presentation:class="object" svg:width="10cm" svg:height="5cm"><draw:text-box><text:p>Slide with remote background image.</text:p></draw:text-box></draw:frame></draw:page></office:presentation></office:body></office:document>"#;

const NOTES: &str = r#"<?xml version="1.0"?><office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" office:mimetype="application/vnd.oasis.opendocument.presentation"><office:body><office:presentation><draw:page draw:name="one"><draw:frame presentation:class="object"><draw:text-box><text:p>First slide</text:p></draw:text-box></draw:frame></draw:page><draw:page draw:name="two"><draw:frame presentation:class="object"><draw:text-box><text:p>Second slide</text:p></draw:text-box></draw:frame><presentation:notes><draw:frame><draw:text-box><text:p>Remember to breathe</text:p></draw:text-box></draw:frame></presentation:notes></draw:page></office:presentation></office:body></office:document>"#;

fn flat_with_body(body: &str) -> String {
    format!(
        r#"<?xml version="1.0"?><office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" office:mimetype="application/vnd.oasis.opendocument.presentation"><office:body>{body}</office:body></office:document>"#
    )
}

fn flat_with_pages(page: &str, count: usize) -> String {
    flat_with_body(&format!(
        "<office:presentation>{}</office:presentation>",
        page.repeat(count)
    ))
}

#[test]
fn flat_read_slides_notes_and_byte_exact_noop() {
    let flat = FlatPresentation::from_bytes(FLAT.as_bytes().to_vec()).unwrap();
    assert_eq!(flat.slide_count(), 1);
    assert!(
        flat.slides()[0]
            .text()
            .unwrap()
            .contains("remote background")
    );
    assert_eq!(flat.as_bytes(), FLAT.as_bytes());
    assert_eq!(flat.to_bytes(), FLAT.as_bytes());

    let notes = FlatPresentation::from_bytes(NOTES.as_bytes().to_vec()).unwrap();
    assert_eq!(notes.slide_count(), 2);
    assert_eq!(notes.slides()[0].notes().unwrap(), None);
    assert_eq!(
        notes.slides()[1].notes().unwrap(),
        Some("Remember to breathe")
    );
}

#[test]
fn flat_transaction_edits_one_slide_and_is_reversible() {
    let flat = FlatPresentation::from_bytes(FLAT.as_bytes().to_vec()).unwrap();
    let source = flat.snapshot();
    let mut transaction = source.transaction();
    transaction
        .replace(0, "New Title", "New Body")
        .unwrap()
        .unwrap();
    let commit = transaction.commit().unwrap();
    assert!(commit.changed());
    assert!(!commit.patch().is_noop());
    let reopened = commit.snapshot().to_presentation();
    assert_eq!(reopened.slides()[0].title().unwrap(), Some("New Title"));
    assert_eq!(reopened.slides()[0].text().unwrap(), "New Body");
    let output = std::str::from_utf8(reopened.as_bytes()).unwrap();
    assert!(output.contains("background.png"));
    let restored = commit.patch().inverse().apply(commit.snapshot()).unwrap();
    assert_eq!(restored.bytes(), FLAT.as_bytes());

    let noop = source.transaction().commit().unwrap();
    assert!(!noop.changed());
    assert!(noop.patch().is_noop());
    assert_eq!(noop.snapshot().bytes(), FLAT.as_bytes());
}

#[test]
fn flat_transaction_preserves_self_closing_direct_notes() {
    let document = NOTES.replace(
        r"<presentation:notes><draw:frame><draw:text-box><text:p>Remember to breathe</text:p></draw:text-box></draw:frame></presentation:notes>",
        r"<presentation:notes/>",
    );
    let source = FlatPresentation::from_bytes(document.into_bytes())
        .unwrap()
        .snapshot();
    let mut transaction = source.transaction();
    transaction.replace(1, "", "Updated body").unwrap().unwrap();
    let commit = transaction.commit().unwrap();
    let output = std::str::from_utf8(commit.snapshot().bytes()).unwrap();
    assert!(output.contains("<presentation:notes/>"));
}

#[test]
fn flat_changed_publication_refuses_formatted_source_xml() {
    let formatted = FLAT.replace("<office:body>", "<office:body>\n");
    let source = FlatPresentation::from_bytes(formatted.into_bytes())
        .unwrap()
        .snapshot();
    let mut transaction = source.transaction();
    transaction
        .replace(0, "New Title", "New Body")
        .unwrap()
        .unwrap();
    assert!(matches!(
        transaction.commit(),
        Err(Error::Unsupported(message)) if message.contains("not compact")
    ));
}

#[test]
fn flat_rejects_wrong_family_packaged_garbage_and_unsafe_rewrite() {
    let text = FLAT.replace(
        "application/vnd.oasis.opendocument.presentation",
        "application/vnd.oasis.opendocument.text",
    );
    assert!(FlatPresentation::from_bytes(text.into_bytes()).is_err());
    assert!(FlatPresentation::from_bytes(b"PK\x03\x04mimetype".to_vec()).is_err());
    assert!(FlatPresentation::from_bytes(Vec::new()).is_err());

    let unsupported = FLAT.replace(
        "</draw:page>",
        "<presentation:event-listeners/></draw:page>",
    );
    let source = FlatPresentation::from_bytes(unsupported.into_bytes())
        .unwrap()
        .snapshot();
    let mut transaction = source.transaction();
    assert!(matches!(
        transaction.replace(0, "title", "body"),
        Err(Error::Unsupported(_))
    ));

    let self_closing_unmodeled = FLAT.replace("</draw:page>", "<draw:image/></draw:page>");
    let source = FlatPresentation::from_bytes(self_closing_unmodeled.into_bytes())
        .unwrap()
        .snapshot();
    let mut transaction = source.transaction();
    assert!(matches!(
        transaction.replace(0, "title", "body"),
        Err(Error::Unsupported(_))
    ));
}

#[test]
fn flat_rejects_misparented_and_excessive_empty_or_nonempty_pages() {
    for page in [
        r#"<draw:page draw:name="bad"/>"#,
        r#"<draw:page draw:name="bad"></draw:page>"#,
    ] {
        let document = flat_with_body(&format!(
            "<office:presentation></office:presentation><office:other>{page}</office:other>"
        ));
        assert!(matches!(
            FlatPresentation::from_bytes(document.into_bytes()),
            Err(Error::InvalidFormat(message)) if message.contains("outside office:presentation")
        ));
    }

    for page in [
        r#"<draw:page draw:name="p"/>"#,
        r#"<draw:page draw:name="p"></draw:page>"#,
    ] {
        assert!(matches!(
            FlatPresentation::from_bytes(flat_with_pages(page, 65_537).into_bytes()),
            Err(Error::InvalidFormat(message)) if message.contains("slide-count limit")
        ));
    }
}

#[test]
fn flat_changed_publication_never_materializes_an_oversized_output() {
    let mut document = FLAT.to_owned();
    let filler = "x".repeat(50 * 1024 * 1024);
    document.insert_str(
        document.len() - "</office:document>".len(),
        &format!("<!--{filler}-->"),
    );
    let source = FlatPresentation::from_bytes(document.into_bytes())
        .unwrap()
        .snapshot();
    let text = "x".repeat(16 * 1024 * 1024);
    let mut transaction = source.transaction();
    transaction.replace(0, "", &text).unwrap().unwrap();
    assert!(matches!(
        transaction.commit(),
        Err(Error::InvalidFormat(message)) if message.contains("output limit")
    ));
}
