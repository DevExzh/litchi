#![allow(
    clippy::shadow_unrelated,
    clippy::unwrap_used,
    reason = "integration-test assertions panic on failure by design"
)]

use litchi_odp::core::{OwnedPackage, PackageWriter};
use litchi_odp::{Builder, Presentation, edit};

fn fixture(name: &str) -> Vec<u8> {
    std::fs::read(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/odf/odp")
            .join(name),
    )
    .unwrap()
}

fn page_fragments(xml: &str) -> Vec<&str> {
    let mut remaining = xml;
    let mut pages = Vec::new();
    while let Some(start) = remaining.find("<draw:page ") {
        remaining = &remaining[start..];
        let end = remaining.find("</draw:page>").unwrap() + "</draw:page>".len();
        pages.push(&remaining[..end]);
        remaining = &remaining[end..];
    }
    pages
}

fn semantic_slide(mut slide: litchi_odp::Slide) -> litchi_odp::Slide {
    slide.index = 0;
    slide
}

#[test]
fn unsupported_real_producer_pages_are_refused_before_staging() {
    let source_bytes = fixture("tdf169979.odp");
    let source = edit::Snapshot::from_bytes(source_bytes.clone()).unwrap();
    let mut transaction = source.transaction().unwrap();
    let before = transaction.slides().to_vec();
    let error = transaction
        .move_slide(0, edit::SlidePosition::Last)
        .unwrap_err();
    assert!(
        error
            .to_string()
            .contains("declarations or settings cannot yet be reordered losslessly")
    );
    assert_eq!(transaction.slides(), before);
    let commit = transaction.commit().unwrap();
    assert!(!commit.changed());
    assert_eq!(commit.snapshot().bytes(), source_bytes);
}

#[test]
fn noncompact_source_without_settings_or_declarations_is_refused() {
    const CONTENT: &[u8] = br#"<?xml version="1.0"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"><office:body><office:presentation><draw:page draw:name="page-a" /><draw:page draw:name="page-b" /></office:presentation></office:body></office:document-content>"#;
    let mut archive = soapberry_zip::office::StreamingArchiveWriter::new();
    archive
        .write_stored(
            "mimetype",
            b"application/vnd.oasis.opendocument.presentation",
        )
        .unwrap();
    archive.write_deflated("content.xml", CONTENT).unwrap();
    archive
        .write_deflated(
            "META-INF/manifest.xml",
            br#"<?xml version="1.0"?><manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.presentation"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/></manifest:manifest>"#,
        )
        .unwrap();
    let source = edit::Snapshot::from_bytes(archive.finish_to_bytes().unwrap()).unwrap();
    let mut transaction = source.transaction().unwrap();

    let error = transaction
        .move_slide(0, edit::SlidePosition::Last)
        .unwrap_err();
    assert!(error.to_string().contains("current writer"));
    let commit = transaction.commit().unwrap();
    assert!(!commit.changed());
    assert_eq!(commit.snapshot().bytes(), source.bytes());
}

#[test]
fn retained_declarations_and_settings_are_conservatively_refused() {
    const CONTENT: &[u8] = br#"<?xml version="1.0"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0"><office:body><office:presentation><presentation:header-decl presentation:name="header">Header</presentation:header-decl><draw:page draw:name="page-a" presentation:use-header-name="header"/><draw:page draw:name="page-b"/><presentation:settings presentation:start-page="page-a"/></office:presentation></office:body></office:document-content>"#;
    let mut writer = PackageWriter::new();
    writer
        .set_mimetype("application/vnd.oasis.opendocument.presentation")
        .unwrap();
    writer.add_file("content.xml", CONTENT).unwrap();
    let source = edit::Snapshot::from_bytes(writer.finish_to_bytes().unwrap()).unwrap();
    let mut transaction = source.transaction().unwrap();
    let before = transaction.slides().to_vec();

    let error = transaction
        .move_slide(0, edit::SlidePosition::Last)
        .unwrap_err();
    assert!(error.to_string().contains("declarations or settings"));
    assert_eq!(transaction.slides(), before);
    let commit = transaction.commit().unwrap();
    assert!(!commit.changed());
    assert_eq!(commit.snapshot().bytes(), source.bytes());
}

fn compact_producer_deck() -> Vec<u8> {
    const CONTENT: &str = r#"<?xml version="1.0"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:vendor="urn:example:producer"><office:body><office:presentation><vendor:deck-token vendor:value="keep"/><draw:page draw:name="page-a" draw:id="id-a"><draw:frame draw:name="shape-a" draw:protect="content"><draw:image xlink:href="Pictures/pixel.png"/><draw:text-box><text:p>A</text:p></draw:text-box></draw:frame><office:annotation office:name="comment-a"><text:p>annotation-a</text:p></office:annotation><vendor:page-token vendor:value="alpha"/></draw:page><draw:page draw:name="page-b" draw:id="id-b"><draw:frame draw:name="shape-b"><draw:text-box><text:p>B</text:p></draw:text-box></draw:frame><presentation:notes><text:p>notes-b</text:p></presentation:notes><vendor:page-token vendor:value="beta"/></draw:page><draw:page draw:name="page-c" draw:id="id-c"><draw:frame draw:name="shape-c"><draw:text-box><text:p>C</text:p></draw:text-box></draw:frame><vendor:transition vendor:type="fade"/><vendor:page-token vendor:value="gamma"/></draw:page></office:presentation></office:body></office:document-content>"#;
    let mut writer = PackageWriter::new();
    writer
        .set_mimetype("application/vnd.oasis.opendocument.presentation")
        .unwrap();
    writer.add_file("content.xml", CONTENT.as_bytes()).unwrap();
    writer
        .add_file_with_media_type("Pictures/pixel.png", b"producer-pixel", "image/png")
        .unwrap();
    writer
        .add_file_with_media_type(
            "Producer/opaque.bin",
            b"producer-opaque",
            "application/octet-stream",
        )
        .unwrap();
    writer.finish_to_bytes().unwrap()
}

#[test]
fn compact_source_pages_move_as_exact_fragments_with_durable_inverse() {
    let source_bytes = compact_producer_deck();
    let source_package = OwnedPackage::from_bytes(source_bytes.clone()).unwrap();
    let source_content =
        String::from_utf8(source_package.get_file("content.xml").unwrap()).unwrap();
    let source_pages = page_fragments(&source_content);
    assert_eq!(source_pages.len(), 3);

    let source = edit::Snapshot::from_bytes(source_bytes.clone()).unwrap();
    let original_slides = source.slides().to_vec();
    let mut transaction = source.transaction().unwrap();
    transaction
        .move_slide(0, edit::SlidePosition::Last)
        .unwrap()
        .unwrap();
    let commit = transaction.commit().unwrap();
    assert!(commit.changed());

    let moved = commit.snapshot().to_presentation().unwrap();
    let moved_slides = moved.slides().unwrap();
    for (actual, expected) in moved_slides.iter().cloned().map(semantic_slide).zip(
        [1usize, 2, 0]
            .into_iter()
            .map(|index| semantic_slide(original_slides[index].clone())),
    ) {
        assert_eq!(actual, expected);
    }
    let moved_pages = moved.pages().unwrap();
    assert_eq!(
        moved_pages
            .pages()
            .iter()
            .map(|page| page.name.as_deref().unwrap())
            .collect::<Vec<_>>(),
        ["page-b", "page-c", "page-a"]
    );
    let annotations = moved.annotations().unwrap();
    assert_eq!(annotations.len(), 1);
    assert_eq!(annotations[0].annotation.text(), "annotation-a");
    assert_eq!(
        annotations[0].anchor.position(),
        &litchi_odp::annotation::Position::Page { index: 2 }
    );

    let moved_package = OwnedPackage::from_bytes(commit.snapshot().bytes().to_vec()).unwrap();
    let moved_content = String::from_utf8(moved_package.get_file("content.xml").unwrap()).unwrap();
    let moved_fragments = page_fragments(&moved_content);
    assert_eq!(
        moved_fragments,
        [source_pages[1], source_pages[2], source_pages[0]]
    );
    assert!(moved_content.contains(r#"<vendor:deck-token vendor:value="keep"/>"#));

    for path in source_package.files().unwrap() {
        if matches!(path.as_str(), "content.xml" | "META-INF/manifest.xml") {
            continue;
        }
        assert_eq!(
            moved_package.get_file(&path).unwrap(),
            source_package.get_file(&path).unwrap(),
            "raw package member changed: {path}"
        );
    }

    let durable =
        edit::Patch::from_durable_bytes(&commit.patch().to_durable_bytes().unwrap()).unwrap();
    let replayed = durable.apply(&source).unwrap();
    assert_eq!(replayed.bytes(), commit.snapshot().bytes());
    assert_eq!(
        durable.inverse().apply(&replayed).unwrap().bytes(),
        source_bytes
    );
    let mut other = Builder::new();
    other.add_slide_with_title("unrelated", "source").unwrap();
    let other = edit::Snapshot::from_bytes(other.build().unwrap()).unwrap();
    assert!(durable.apply(&other).is_err());
}

#[test]
fn staged_coordinates_cover_middle_first_last_noop_and_atomic_failures() {
    let mut builder = Builder::new();
    for title in ["A", "B", "C", "D"] {
        builder
            .add_slide_with_title(title, &format!("body-{title}"))
            .unwrap();
    }
    let source = edit::Snapshot::from_bytes(builder.build().unwrap()).unwrap();

    let mut transaction = source.transaction().unwrap();
    transaction
        .move_slide("B", edit::SlidePosition::Last)
        .unwrap()
        .unwrap();
    assert_eq!(
        transaction
            .slides()
            .iter()
            .map(|slide| slide.title.as_deref().unwrap())
            .collect::<Vec<_>>(),
        ["A", "C", "D", "B"]
    );
    transaction
        .move_slide(2, edit::SlidePosition::First)
        .unwrap()
        .unwrap();
    assert_eq!(
        transaction
            .slides()
            .iter()
            .map(|slide| slide.title.as_deref().unwrap())
            .collect::<Vec<_>>(),
        ["D", "A", "C", "B"]
    );
    let before_failure = transaction.slides().to_vec();
    assert!(
        transaction
            .move_slide(1, edit::SlidePosition::Index(4))
            .is_err()
    );
    assert_eq!(transaction.slides(), before_failure);
    assert!(
        transaction
            .move_slide("missing", edit::SlidePosition::First)
            .unwrap()
            .is_none()
    );
    assert_eq!(transaction.slides(), before_failure);
    transaction
        .move_slide(0, edit::SlidePosition::Index(2))
        .unwrap()
        .unwrap();
    assert_eq!(
        transaction
            .slides()
            .iter()
            .map(|slide| slide.title.as_deref().unwrap())
            .collect::<Vec<_>>(),
        ["A", "C", "D", "B"]
    );
    let before_mixed_edit = transaction.slides().to_vec();
    assert!(transaction.remove(0).is_err());
    assert_eq!(transaction.slides(), before_mixed_edit);
    let commit = transaction.commit().unwrap();
    Presentation::from_bytes(commit.snapshot().bytes().to_vec()).unwrap();

    let mut noop = source.transaction().unwrap();
    noop.move_slide(0, edit::SlidePosition::First)
        .unwrap()
        .unwrap();
    let noop = noop.commit().unwrap();
    assert!(!noop.changed());
    assert_eq!(noop.snapshot().bytes(), source.bytes());
}

#[test]
fn page_indexed_operation_conflict_is_refused_without_moving_slides() {
    let source = edit::Snapshot::from_bytes(compact_producer_deck()).unwrap();
    let mut transaction = source.transaction().unwrap();
    let annotation = litchi_odp::annotation::Annotation::new("staged comment");
    transaction
        .add_annotation(&litchi_odp::annotation::Anchor::page(1), &annotation)
        .unwrap();
    let before = transaction.slides().to_vec();

    let error = transaction
        .move_slide(0, edit::SlidePosition::Last)
        .unwrap_err();
    assert!(error.to_string().contains("page-indexed operations"));
    assert_eq!(transaction.slides(), before);

    let commit = transaction.commit().unwrap();
    let reopened = commit.snapshot().to_presentation().unwrap();
    assert_eq!(
        reopened
            .pages()
            .unwrap()
            .pages()
            .iter()
            .map(|page| page.name.as_deref().unwrap())
            .collect::<Vec<_>>(),
        ["page-a", "page-b", "page-c"]
    );
    assert_eq!(reopened.annotations().unwrap().len(), 2);
}
