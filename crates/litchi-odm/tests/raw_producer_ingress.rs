#![allow(
    clippy::unwrap_used,
    reason = "fixed in-memory raw ZIP fixtures keep boundary assertions concise"
)]

use litchi_core::Position;
use litchi_odf_common::compact_xml;
use litchi_odm::{Master, transaction::SectionSpec};
use std::io::{Cursor, Write as _};

const MIME: &str = "application/vnd.oasis.opendocument.text-master";
const COMPACT_CONTENT: &str = concat!(
    r#"<?xml version="1.0"?><office:document-content "#,
    r#"xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" "#,
    r#"xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">"#,
    r#"<office:body><office:text><text:section text:name="A"/>"#,
    r#"</office:text></office:body></office:document-content>"#,
);

fn raw_package(content: &str) -> Vec<u8> {
    let manifest = format!(
        concat!(
            r#"<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0">"#,
            r#"<manifest:file-entry manifest:full-path="/" manifest:media-type="{MIME}"/>"#,
            r#"<manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>"#,
            r#"</manifest:manifest>"#,
        ),
        MIME = MIME,
    );
    let mut output = Cursor::new(Vec::new());
    {
        let mut zip = zip::ZipWriter::new(&mut output);
        let stored = zip::write::SimpleFileOptions::default()
            .compression_method(zip::CompressionMethod::Stored);
        let deflated = zip::write::SimpleFileOptions::default()
            .compression_method(zip::CompressionMethod::Deflated);
        zip.start_file("mimetype", stored).unwrap();
        zip.write_all(MIME.as_bytes()).unwrap();
        zip.start_file("content.xml", deflated).unwrap();
        zip.write_all(content.as_bytes()).unwrap();
        zip.start_file("META-INF/manifest.xml", deflated).unwrap();
        zip.write_all(manifest.as_bytes()).unwrap();
        zip.finish().unwrap();
    }
    output.into_inner()
}

#[test]
fn raw_pretty_producer_package_is_ingested_exactly_then_compacted_on_change() {
    let pretty = COMPACT_CONTENT
        .replace("<office:body>", "<office:body>\n  ")
        .replace("<office:text>", "<office:text>\n    ")
        .replace("</office:text>", "\n  </office:text>");
    let original = raw_package(&pretty);
    let source = Master::from_bytes(original.clone()).unwrap();
    assert_eq!(source.as_bytes(), original);

    let mut edit = source.edit();
    edit.add_section(SectionSpec::new("B").unwrap()).unwrap();
    let changed = edit.commit().unwrap().into_snapshot();
    compact_xml::validate(changed.content_xml().as_bytes()).unwrap();
    assert!(!changed.content_xml().contains('\n'));
    assert!(Master::from_bytes(changed.as_bytes().to_vec()).is_ok());
}

#[test]
fn raw_unsafe_xml_is_rejected_at_ingress_or_changed_publication() {
    let dtd = COMPACT_CONTENT.replacen(
        "?>",
        "?><!DOCTYPE office:document-content [<!ENTITY x 'bad'>]>",
        1,
    );
    assert!(Master::from_bytes(raw_package(&dtd)).is_err());

    let ambiguous = COMPACT_CONTENT.replace(
        "<office:text><text:section",
        "<office:text>   <text:section",
    );
    let source = Master::from_bytes(raw_package(&ambiguous)).unwrap();
    let mut edit = source.edit();
    edit.add_section(SectionSpec::new("B").unwrap()).unwrap();
    assert!(edit.commit().is_err());
}

#[test]
fn subtree_removal_checks_incoming_references_to_descendants() {
    let content = COMPACT_CONTENT.replace(
        r#"<text:section text:name="A"/>"#,
        concat!(
            r#"<text:section text:name="Parent"><text:section text:name="Child"/>"#,
            r#"</text:section><text:section text:name="Outside">"#,
            r#"<text:section-source text:section-name="Child"/></text:section>"#,
        ),
    );
    let source = Master::from_bytes(raw_package(&content)).unwrap();
    let mut edit = source.edit();
    edit.remove_section(Position::new(0)).unwrap();
    assert!(edit.commit().is_err());
}

#[test]
fn generated_indexes_require_unique_addressable_names() {
    let missing = COMPACT_CONTENT.replace(
        r#"<text:section text:name="A"/>"#,
        r#"<text:table-of-content/>"#,
    );
    assert!(Master::from_bytes(raw_package(&missing)).is_err());

    let duplicate = COMPACT_CONTENT.replace(
        r#"<text:section text:name="A"/>"#,
        concat!(
            r#"<text:table-of-content text:name="Index"/>"#,
            r#"<text:alphabetical-index text:name="Index"/>"#,
        ),
    );
    assert!(Master::from_bytes(raw_package(&duplicate)).is_err());
}

#[test]
fn common_body_children_are_schema_checked_at_raw_ingress() {
    let invalid_list = COMPACT_CONTENT.replace(
        r#"<text:section text:name="A"/>"#,
        r#"<text:list><text:p>not-an-item</text:p></text:list>"#,
    );
    assert!(Master::from_bytes(raw_package(&invalid_list)).is_err());
    let header_only_list = COMPACT_CONTENT.replace(
        r#"<text:section text:name="A"/>"#,
        r#"<text:list><text:list-header><text:p>header</text:p></text:list-header></text:list>"#,
    );
    assert!(Master::from_bytes(raw_package(&header_only_list)).is_err());

    let invalid_table = COMPACT_CONTENT
        .replace(
            r#"xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0""#,
            concat!(
                r#"xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" "#,
                r#"xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0""#,
            ),
        )
        .replace(
            r#"<text:section text:name="A"/>"#,
            r#"<table:table><text:p>not-a-row</text:p></table:table>"#,
        );
    assert!(Master::from_bytes(raw_package(&invalid_table)).is_err());
    let no_row_table = invalid_table.replace(
        r#"<table:table><text:p>not-a-row</text:p></table:table>"#,
        r#"<table:table><table:table-column/></table:table>"#,
    );
    assert!(Master::from_bytes(raw_package(&no_row_table)).is_err());

    let invalid_index = COMPACT_CONTENT.replace(
        r#"<text:section text:name="A"/>"#,
        concat!(
            r#"<text:table-of-content text:name="Contents">"#,
            r#"<text:index-body/><text:table-of-content-source/>"#,
            r#"</text:table-of-content>"#,
        ),
    );
    assert!(Master::from_bytes(raw_package(&invalid_index)).is_err());
}
