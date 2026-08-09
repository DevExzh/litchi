#![allow(clippy::unwrap_used, reason = "test assertions use unwrap for clarity")]

use litchi_core::Position;
use litchi_odf_common::{compact_xml, core::PackageWriter};
use litchi_odm::{Master, subdocument::Target};

const MIME: &str = "application/vnd.oasis.opendocument.text-master";
const CONTENT: &str = concat!(
    r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content "#,
    r#"xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" "#,
    r#"xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" "#,
    r#"xmlns:xlink="http://www.w3.org/1999/xlink" office:version="1.4">"#,
    r#"<office:body><office:text><text:section text:name="A"><text:section-source "#,
    r#"xlink:type="simple" xlink:href="Chapters/a.odt" xlink:show="embed" "#,
    r#"text:section-name="Body" text:filter-name="writer8"/><text:p>Cached</text:p>"#,
    r#"</text:section><text:section text:name="B"><text:section-source "#,
    r#"xlink:href="https://example.test/b.odt"/></text:section></office:text>"#,
    r#"</office:body></office:document-content>"#,
);

fn package(content: &str, extra_path: Option<&str>) -> Vec<u8> {
    let mut writer = PackageWriter::new();
    writer.set_mimetype(MIME).unwrap();
    writer.add_file("content.xml", content.as_bytes()).unwrap();
    if let Some(path) = extra_path {
        writer
            .add_file_with_media_type(path, b"opaque", "application/octet-stream")
            .unwrap();
    }
    writer.finish_to_bytes().unwrap()
}

fn document(inner: &str) -> String {
    format!(
        r#"<?xml version="1.0"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:xlink="http://www.w3.org/1999/xlink"><office:body><office:text>{inner}</office:text></office:body></office:document-content>"#
    )
}

#[test]
fn projects_odf_section_source_attributes_without_activating_targets() {
    let master = Master::from_bytes(package(CONTENT, None)).unwrap();
    let links = master.subdocuments();
    assert_eq!(links.len(), 2);
    assert_eq!(links[0].source_section(), Some("Body"));
    assert_eq!(links[0].filter_name(), Some("writer8"));
    assert!(matches!(links[0].target(), Target::Package(_)));
    assert!(links[1].target().is_external());

    let local_only = document(
        r#"<text:section text:name="Local"><text:section-source text:section-name="Other"/><text:p>Cached</text:p></text:section>"#,
    );
    assert!(
        Master::from_bytes(package(&local_only, None))
            .unwrap()
            .subdocuments()
            .is_empty()
    );
}

#[test]
fn enforces_odf_linked_section_structure_and_attribute_domains() {
    let invalid = [
        document(
            r#"<text:section text:name="late"><text:p>Cached</text:p><text:section-source xlink:href="a.odt"/></text:section>"#,
        ),
        document(
            r#"<text:section text:name="twice"><text:section-source xlink:href="a.odt"/><text:section-source xlink:href="b.odt"/></text:section>"#,
        ),
        document(
            r#"<text:section text:name="child"><text:section-source xlink:href="a.odt"><text:p>invalid</text:p></text:section-source></text:section>"#,
        ),
        document(
            r#"<text:section text:name="type"><text:section-source xlink:type="extended" xlink:href="a.odt"/></text:section>"#,
        ),
        document(
            r#"<text:section text:name="show"><text:section-source xlink:href="a.odt" xlink:show="new"/></text:section>"#,
        ),
        document(
            r#"<text:section text:name="orphan"><text:section-source xlink:type="simple"/></text:section>"#,
        ),
    ];
    for content in invalid {
        assert!(Master::from_bytes(package(&content, None)).is_err());
    }
}

#[test]
fn linked_section_edit_is_compact_reversible_and_source_checked() {
    let source = Master::from_bytes(package(CONTENT, Some("custom/opaque.bin"))).unwrap();
    let mut edit = source.edit_link(Position::new(0)).unwrap();
    assert_eq!(edit.reference(), Position::new(0));
    edit.set_href("Chapters/revised & final.odt").unwrap();
    let commit = edit.commit().unwrap();
    let changed = commit.snapshot();
    assert_eq!(
        changed.subdocuments()[0].href(),
        "Chapters/revised & final.odt"
    );
    assert!(matches!(
        changed.subdocuments()[0].target(),
        Target::Package(_)
    ));
    assert_eq!(changed.subdocuments()[0].source_section(), Some("Body"));
    assert_eq!(changed.subdocuments()[0].filter_name(), Some("writer8"));
    assert!(changed.content_xml().contains("revised &amp; final.odt"));
    assert!(
        changed
            .files()
            .unwrap()
            .contains(&"custom/opaque.bin".to_string())
    );
    compact_xml::validate(changed.content_xml().as_bytes()).unwrap();
    assert!(!changed.content_xml().contains('\n'));
    assert!(!changed.content_xml().contains('\t'));
    assert!(!changed.content_xml().contains("  "));
    assert_eq!(commit.patch().change().reference(), Position::new(0));
    assert_eq!(commit.patch().change().before(), "Chapters/a.odt");
    assert_eq!(
        commit.patch().change().after(),
        "Chapters/revised & final.odt"
    );
    assert!(commit.patch().is_applicable_to(&source));
    assert_eq!(
        source.apply_link_patch(commit.patch()).unwrap().as_bytes(),
        changed.as_bytes()
    );
    assert_eq!(
        commit.patch().inverse().apply(changed).unwrap().as_bytes(),
        source.as_bytes()
    );

    let stale = Master::from_bytes(package(&CONTENT.replace("Cached", "Different"), None)).unwrap();
    assert!(!commit.patch().is_applicable_to(&stale));
    assert!(commit.patch().apply(&stale).is_err());
}

#[test]
fn linked_section_noop_and_refusals_are_typed() {
    let source = Master::from_bytes(package(CONTENT, None)).unwrap();
    let noop = source.edit_link("A").unwrap().commit().unwrap();
    assert!(noop.patch().is_noop());
    assert_eq!(noop.snapshot().as_bytes(), source.as_bytes());
    assert!(source.edit_link(Position::new(2)).is_err());
    assert!(source.edit_link("missing").is_err());

    let mut invalid = source.edit_link(Position::new(0)).unwrap();
    assert!(invalid.set_href("bad\0target").is_err());

    let signed =
        Master::from_bytes(package(CONTENT, Some("META-INF/documentsignatures.xml"))).unwrap();
    let mut edit = signed.edit_link(Position::new(0)).unwrap();
    edit.set_href("changed.odt").unwrap();
    assert!(edit.commit().is_err());
}

#[test]
fn opening_requires_compact_content_xml() {
    let noncompact = CONTENT.replace("<office:body>", "<office:body>\n");
    assert!(Master::from_bytes(package(&noncompact, None)).is_err());
}

#[test]
fn namespace_aliases_are_not_mistaken_for_document_semantics() {
    let aliased = r#"<?xml version="1.0"?><o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:l="http://www.w3.org/1999/xlink"><o:body><o:text><t:section t:name="Alias"><t:section-source l:href="alias.odt"/></t:section></o:text></o:body></o:document-content>"#;
    let master = Master::from_bytes(package(aliased, None)).unwrap();
    assert_eq!(master.subdocuments()[0].href(), "alias.odt");
}
