#![allow(
    clippy::shadow_unrelated,
    clippy::unwrap_used,
    reason = "integration-test assertions panic on failure by design"
)]

use litchi_core::Error;
use litchi_odp::flat::Presentation as FlatPresentation;

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
    let flat = FlatPresentation::from_bytes(FLAT.as_bytes()).unwrap();
    assert_eq!(flat.slide_count(), 1);
    assert!(
        flat.slides()[0]
            .text()
            .unwrap()
            .contains("remote background")
    );
    assert_eq!(flat.as_bytes(), FLAT.as_bytes());
    assert_eq!(flat.to_bytes(), FLAT.as_bytes());

    let notes = FlatPresentation::from_bytes(NOTES.as_bytes()).unwrap();
    assert_eq!(notes.slide_count(), 2);
    assert_eq!(notes.slides()[0].notes().unwrap(), None);
    assert_eq!(
        notes.slides()[1].notes().unwrap(),
        Some("Remember to breathe")
    );
}

#[test]
fn flat_transaction_edits_one_slide_and_is_reversible() {
    let source = FlatPresentation::from_bytes(FLAT.as_bytes()).unwrap();
    let mut transaction = source.edit();
    transaction
        .replace(0, "New Title", "New Body")
        .unwrap()
        .unwrap();
    let commit = transaction.commit().unwrap();
    assert!(commit.diagnostics().changed());
    assert!(!commit.patch().is_noop());
    let reopened = commit.package();
    assert_eq!(reopened.slides()[0].title().unwrap(), Some("New Title"));
    assert_eq!(reopened.slides()[0].text().unwrap(), "New Body");
    let output = std::str::from_utf8(reopened.as_bytes()).unwrap();
    assert!(output.contains("background.png"));
    let restored = commit.package().apply(&commit.patch().inverse()).unwrap();
    assert_eq!(restored.package().as_bytes(), FLAT.as_bytes());
    assert!(
        reopened
            .apply(commit.patch())
            .is_err_and(|error| matches!(error, litchi_odp::flat::Error::PatchConflict { .. }))
    );

    let noop = source.edit().commit().unwrap();
    assert!(!noop.diagnostics().changed());
    assert!(noop.patch().is_noop());
    assert_eq!(noop.package().as_bytes(), FLAT.as_bytes());
}

#[test]
fn flat_transaction_preserves_self_closing_direct_notes() {
    let document = NOTES.replace(
        r"<presentation:notes><draw:frame><draw:text-box><text:p>Remember to breathe</text:p></draw:text-box></draw:frame></presentation:notes>",
        r"<presentation:notes/>",
    );
    let source = FlatPresentation::from_bytes(document.as_bytes()).unwrap();
    let mut transaction = source.edit();
    transaction.replace(1, "", "Updated body").unwrap().unwrap();
    let commit = transaction.commit().unwrap();
    let output = std::str::from_utf8(commit.package().as_bytes()).unwrap();
    assert!(output.contains("<presentation:notes/>"));
}

#[test]
fn flat_changed_publication_refuses_formatted_source_xml() {
    let formatted = FLAT.replace("<office:body>", "<office:body>\n");
    let source = FlatPresentation::from_bytes(formatted.as_bytes()).unwrap();
    let mut transaction = source.edit();
    transaction
        .replace(0, "New Title", "New Body")
        .unwrap()
        .unwrap();
    assert!(matches!(
        transaction.commit(),
        Err(litchi_odp::flat::Error::Refused(refusal)) if refusal.reason().contains("not compact")
    ));
}

#[test]
fn flat_rejects_wrong_family_packaged_garbage_and_unsafe_rewrite() {
    let text = FLAT.replace(
        "application/vnd.oasis.opendocument.presentation",
        "application/vnd.oasis.opendocument.text",
    );
    assert!(FlatPresentation::from_bytes(text.as_bytes()).is_err());
    assert!(FlatPresentation::from_bytes(b"PK\x03\x04mimetype").is_err());
    assert!(FlatPresentation::from_bytes(&[]).is_err());

    let unsupported = FLAT.replace(
        "</draw:page>",
        "<presentation:event-listeners/></draw:page>",
    );
    let source = FlatPresentation::from_bytes(unsupported.as_bytes()).unwrap();
    let mut transaction = source.edit();
    assert!(matches!(
        transaction.replace(0, "title", "body"),
        Err(Error::Unsupported(_) | Error::Refused(_))
    ));

    let self_closing_unmodeled = FLAT.replace("</draw:page>", "<draw:image/></draw:page>");
    let source = FlatPresentation::from_bytes(self_closing_unmodeled.as_bytes()).unwrap();
    let mut transaction = source.edit();
    assert!(matches!(
        transaction.replace(0, "title", "body"),
        Err(Error::Unsupported(_) | Error::Refused(_))
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
            FlatPresentation::from_bytes(document.as_bytes()),
            Err(litchi_odp::flat::Error::Core(Error::InvalidFormat(message))) if message.contains("outside office:presentation")
        ));
    }

    for page in [
        r#"<draw:page draw:name="p"/>"#,
        r#"<draw:page draw:name="p"></draw:page>"#,
    ] {
        assert!(matches!(
            FlatPresentation::from_bytes_with_options(
                flat_with_pages(page, 65_537).as_bytes(),
                slide_limit(65_536)
            ),
            Err(litchi_odp::flat::Error::Limit {
                kind: litchi_odp::LimitKind::Semantic(litchi_odp::SemanticLimitKind::Units),
                ..
            })
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
    let source =
        FlatPresentation::from_bytes_with_options(document.as_bytes(), content_limit(64 << 20))
            .unwrap();
    let text = "x".repeat(16 * 1024 * 1024);
    let mut transaction = source.edit();
    transaction.replace(0, "", &text).unwrap().unwrap();
    assert!(matches!(
        transaction.commit(),
        Err(litchi_odp::flat::Error::Limit { .. })
    ));
}

/// Record 0341: XML's predefined entities and character references need no
/// document type, so a flat presentation holding them opens (LibreOffice's
/// `theme.fodp` and `alternativeText.fodp` write `&apos;`, `&lt;` and
/// `&gt;`); an undeclared custom entity is still refused.
#[test]
fn flat_predefined_entities_and_character_references_open() {
    let page = r#"<draw:page draw:name="one"><draw:frame><draw:text-box><text:p>Tom&apos;s &lt;a&gt; &amp; &quot;b&quot; &#169; &#x41;</text:p></draw:text-box></draw:frame></draw:page>"#;
    let document = NOTES.replace(
        r#"<draw:page draw:name="one"><draw:frame presentation:class="object"><draw:text-box><text:p>First slide</text:p></draw:text-box></draw:frame></draw:page>"#,
        page,
    );
    assert_ne!(document, NOTES);
    let flat = FlatPresentation::from_bytes(document.as_bytes()).unwrap();
    assert_eq!(flat.slide_count(), 2);
    assert_eq!(flat.as_bytes(), document.as_bytes());
    assert_eq!(flat.slides()[0].text().unwrap(), "Tom's <a> & \"b\" © A");

    let control = NOTES.replace("First slide", "First &#1; slide");
    assert!(FlatPresentation::from_bytes(control.as_bytes()).is_err());
    let custom = NOTES.replace("First slide", "First &nasty; slide");
    assert!(matches!(
        FlatPresentation::from_bytes(custom.as_bytes()),
        Err(litchi_odp::flat::Error::Core(Error::InvalidFormat(message))) if message.contains("custom entities")
    ));
}

fn slide_limit(slides: u64) -> litchi_odp::OpenOptions {
    litchi_odp::OpenOptions::new().semantic_limits(
        litchi_odp::SemanticLimits::builder()
            .max_units(slides)
            .unwrap()
            .build()
            .unwrap(),
    )
}

fn content_limit(bytes: u64) -> litchi_odp::OpenOptions {
    litchi_odp::OpenOptions::new().semantic_limits(
        litchi_odp::SemanticLimits::builder()
            .max_content_bytes(bytes)
            .unwrap()
            .build()
            .unwrap(),
    )
}
