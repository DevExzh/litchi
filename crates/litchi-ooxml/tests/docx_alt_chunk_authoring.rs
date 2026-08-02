use litchi_docx::alt::{Data, Import, Kind, MAX_CHUNKS, Target};
use litchi_ooxml::docx::{DocumentBlock, Package};
use litchi_opc::OpcPackage;
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::packuri::PackURI;
use litchi_opc::part::{BlobPart, Part};

const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const HTML_FIXTURE: &[u8] = include_bytes!(
    "../../../test-data/libreoffice-core/sw/qa/writerfilter/dmapper/data/alt-chunk-html.docx"
);
const DOCX_FIXTURE: &[u8] = include_bytes!(
    "../../../test-data/libreoffice-core/sw/qa/writerfilter/dmapper/data/alt-chunk.docx"
);
const HEADER_FIXTURE: &[u8] = include_bytes!(
    "../../../test-data/libreoffice-core/sw/qa/writerfilter/dmapper/data/alt-chunk-header.docx"
);

fn package_with_document(
    xml: String,
    relationships: &[(&str, &str, &str, bool)],
    payloads: &[(&str, &str, &[u8])],
) -> Package {
    let document_uri = PackURI::new("/word/document.xml").unwrap();
    let mut document = BlobPart::new(
        document_uri,
        ct::WML_DOCUMENT_MAIN.to_string(),
        xml.into_bytes(),
    );
    for (id, reltype, target, external) in relationships {
        document.rels_mut().add_relationship(
            (*reltype).to_string(),
            (*target).to_string(),
            (*id).to_string(),
            *external,
        );
    }
    let mut opc = OpcPackage::new();
    opc.add_part(Box::new(document));
    for (name, content_type, bytes) in payloads {
        opc.add_part(Box::new(BlobPart::new(
            PackURI::new(*name).unwrap(),
            (*content_type).to_string(),
            bytes.to_vec(),
        )));
    }
    opc.rels_mut().add_relationship(
        rt::OFFICE_DOCUMENT.to_string(),
        "word/document.xml".to_string(),
        "rId1".to_string(),
        false,
    );
    Package::from_opc_package(opc).unwrap()
}

#[test]
fn libreoffice_payloads_are_ordered_borrowed_and_opaque() {
    let html = Package::from_reader(std::io::Cursor::new(HTML_FIXTURE)).unwrap();
    let document = html.document().unwrap();
    let blocks = document.blocks().unwrap();
    assert_eq!(blocks.len(), 3);
    assert!(matches!(blocks[0], DocumentBlock::Paragraph(_)));
    let DocumentBlock::Alt(chunk) = &blocks[1] else {
        panic!("missing ordered alternative-format anchor")
    };
    assert!(matches!(blocks[2], DocumentBlock::Paragraph(_)));
    let payload = document.resolve_alt(chunk).unwrap();
    assert_eq!(payload.kind(), Kind::Html);
    assert_eq!(payload.media_type(), "text/html");
    assert_eq!(
        payload.bytes(),
        b"<html><body><p>HTML AltChunk</p></body></html>"
    );

    let docx = Package::from_reader(std::io::Cursor::new(DOCX_FIXTURE)).unwrap();
    let document = docx.document().unwrap();
    let chunk = document.alts().unwrap().remove(0);
    let payload = document.resolve_alt(&chunk).unwrap();
    assert_eq!(payload.kind(), Kind::Docx);
    assert!(payload.bytes().starts_with(b"PK"));

    let header = Package::from_reader(std::io::Cursor::new(HEADER_FIXTURE)).unwrap();
    let document = header.document().unwrap();
    let chunk = document.alts().unwrap().remove(0);
    let payload = document.resolve_alt(&chunk).unwrap();
    assert_eq!(payload.name().as_str(), "/word/afchunk2.docx");
    assert_eq!(payload.kind(), Kind::Docx);
}

#[test]
fn markup_compatibility_selectors_and_package_indexes_share_the_active_branch() {
    let xml = format!(
        r#"<w:document xmlns:w="{W}" xmlns:r="{R}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:unsupported"><w:body><mc:AlternateContent><mc:Choice Requires="x"><w:p><w:r><w:t>inactive</w:t></w:r></w:p><w:altChunk r:id="inactive"/></mc:Choice><mc:Fallback><w:p><w:r><w:t>active</w:t></w:r></w:p><w:altChunk r:id="fallback"/></mc:Fallback></mc:AlternateContent></w:body></w:document>"#
    );
    let mut package = package_with_document(
        xml,
        &[
            (
                "inactive",
                rt::MS_ALTERNATIVE_FORMAT_IMPORT,
                "inactive.html",
                false,
            ),
            (
                "fallback",
                rt::MS_ALTERNATIVE_FORMAT_IMPORT,
                "fallback.html",
                false,
            ),
        ],
        &[
            ("/word/inactive.html", "text/html", b"inactive"),
            ("/word/fallback.html", "text/html", b"fallback"),
        ],
    );
    {
        let document = package.document().unwrap();
        let chunks = document.alts().unwrap();
        assert_eq!(chunks.len(), 1);
        assert_eq!(chunks[0].relationship().as_str(), "fallback");
        let blocks = document.blocks().unwrap();
        assert_eq!(blocks.len(), 2);
        let DocumentBlock::Paragraph(paragraph) = &blocks[0] else {
            panic!("missing active fallback paragraph")
        };
        assert_eq!(paragraph.text().unwrap(), "active");
        assert!(matches!(blocks[1], DocumentBlock::Alt(_)));
    }
    assert_eq!(package.document_mut().unwrap().alts().len(), 1);
    assert!(package.remove_alt(1).is_err());
    let old = package
        .replace_alt(
            0,
            Import::data(Data::Text(b"replacement".to_vec())),
            Some(false),
        )
        .unwrap();
    assert_eq!(old.relationship().as_str(), "fallback");
    let output = package.document_mut().unwrap().to_xml().unwrap();
    assert!(output.contains(r#"<w:altChunk r:id="inactive"/>"#));
    assert!(output.contains("<w:t>inactive</w:t>"));
    assert!(!output.contains(r#"r:id="fallback""#));

    let directory = tempfile::tempdir().unwrap();
    let path = directory.path().join("active-mce.docx");
    package.save(&path).unwrap();
    let reopened = Package::open(&path).unwrap();
    let document = reopened.document().unwrap();
    let chunks = document.alts().unwrap();
    assert_eq!(chunks.len(), 1);
    assert_eq!(chunks[0].match_source(), Some(false));
    assert_eq!(document.resolve_alt(&chunks[0]).unwrap().kind(), Kind::Text);
}

#[test]
fn inherited_transitional_and_strict_aliases_support_package_crud() {
    const STRICT_W: &str = "http://purl.oclc.org/ooxml/wordprocessingml/main";
    const STRICT_R: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships";
    const STRICT_ALT: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships/afChunk";
    let directory = tempfile::tempdir().unwrap();

    for (case, word, relationships, relationship_type) in [
        ("transitional", W, R, rt::MS_ALTERNATIVE_FORMAT_IMPORT),
        ("strict", STRICT_W, STRICT_R, STRICT_ALT),
    ] {
        let xml = format!(
            r#"<?xml version="1.0"?><d:document xmlns:d="{word}" xmlns:link="{relationships}"><d:body><d:p><d:r><d:t>preserved</d:t></d:r></d:p><d:altChunk link:id="old"/></d:body></d:document>"#
        );
        let mut package = package_with_document(
            xml,
            &[("old", relationship_type, "old.html", false)],
            &[("/word/old.html", "text/html", b"old")],
        );
        assert_eq!(
            package.document().unwrap().alts().unwrap()[0]
                .relationship()
                .as_str(),
            "old"
        );
        assert_eq!(
            package.document_mut().unwrap().alts()[0]
                .relationship()
                .as_str(),
            "old"
        );

        let inserted = package
            .insert_alt(
                0,
                Import::data(Data::Html(b"inserted".to_vec())),
                Some(true),
            )
            .unwrap();
        package.move_alt(1, 0).unwrap();
        assert_eq!(
            package.document_mut().unwrap().alts()[0]
                .relationship()
                .as_str(),
            "old"
        );
        let replaced = package
            .replace_alt(
                0,
                Import::data(Data::Rtf(br"{\rtf1 replacement}".to_vec())),
                Some(false),
            )
            .unwrap();
        assert_eq!(replaced.relationship().as_str(), "old");
        let removed = package.remove_alt(1).unwrap();
        assert_eq!(removed, inserted);
        let output = package.document_mut().unwrap().to_xml().unwrap();
        assert!(output.contains("<d:document"));
        assert!(output.contains("<d:t>preserved</d:t>"));

        let path = directory.path().join(format!("aliased-{case}.docx"));
        package.save(&path).unwrap();
        let reopened = Package::open(&path).unwrap();
        let document = reopened.document().unwrap();
        let chunks = document.alts().unwrap();
        assert_eq!(chunks.len(), 1);
        assert_eq!(chunks[0].match_source(), Some(false));
        assert_eq!(document.resolve_alt(&chunks[0]).unwrap().kind(), Kind::Rtf);
    }
}

#[test]
fn missing_internal_payload_is_rejected_without_fetching_or_importing() {
    let xml = format!(
        r#"<w:document xmlns:w="{W}" xmlns:r="{R}"><w:body><w:altChunk r:id="missing"/></w:body></w:document>"#
    );
    let package = package_with_document(
        xml,
        &[(
            "missing",
            rt::MS_ALTERNATIVE_FORMAT_IMPORT,
            "missing.html",
            false,
        )],
        &[],
    );
    let document = package.document().unwrap();
    let chunk = document.alts().unwrap().remove(0);
    assert!(document.resolve_alt(&chunk).is_err());
    assert!(document.alt_target(&chunk).is_err());
}

#[test]
fn generated_internal_formats_have_canonical_content_types_and_reorder() {
    let directory = tempfile::tempdir().unwrap();
    let path = directory.path().join("formats.docx");
    let mut package = Package::new().unwrap();
    let imports = [
        Data::Docm(b"opaque macro document".to_vec()),
        Data::Dotx(b"opaque template".to_vec()),
        Data::Dotm(b"opaque macro template".to_vec()),
        Data::Mime(b"From: sender@example.invalid\r\n\r\nbody".to_vec()),
        Data::Html(b"<p>html</p>".to_vec()),
        Data::Xhtml(b"<html xmlns=\"http://www.w3.org/1999/xhtml\"/>".to_vec()),
        Data::Rtf(br"{\rtf1 rtf}".to_vec()),
        Data::Text(b"plain".to_vec()),
        Data::Xml(b"<root/>".to_vec()),
        Data::Docx(b"opaque nested package".to_vec()),
    ];
    for (index, data) in imports.into_iter().enumerate() {
        package
            .add_alt(Import::data(data), (index == 0).then_some(true))
            .unwrap();
    }
    package.move_alt(9, 0).unwrap();
    package.save(&path).unwrap();

    let reopened = Package::open(&path).unwrap();
    let document = reopened.document().unwrap();
    let chunks = document.alts().unwrap();
    assert_eq!(chunks.len(), 10);
    let actual = chunks
        .iter()
        .map(|chunk| {
            let part = document.resolve_alt(chunk).unwrap();
            (part.kind(), part.media_type().to_string())
        })
        .collect::<Vec<_>>();
    assert_eq!(
        actual,
        vec![
            (
                Kind::Docx,
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
                    .into(),
            ),
            (
                Kind::Docm,
                "application/vnd.ms-word.document.macroEnabled.main+xml".into(),
            ),
            (
                Kind::Dotx,
                "application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml"
                    .into(),
            ),
            (
                Kind::Dotm,
                "application/vnd.ms-word.template.macroEnabledTemplate.main+xml".into(),
            ),
            (Kind::Mime, "message/rfc822".into()),
            (Kind::Html, "text/html".into()),
            (Kind::Xhtml, "application/xhtml+xml".into()),
            (Kind::Rtf, "application/rtf".into()),
            (Kind::Text, "text/plain".into()),
            (Kind::Xml, "application/xml".into()),
        ]
    );
    assert_eq!(chunks[1].match_source(), Some(true));
}

#[test]
fn external_target_is_returned_without_access_and_strict_anchor_is_emitted() {
    let directory = tempfile::tempdir().unwrap();
    let path = directory.path().join("strict-external.docx");
    let strict_w = "http://purl.oclc.org/ooxml/wordprocessingml/main";
    let strict_r = "http://purl.oclc.org/ooxml/officeDocument/relationships";
    let xml = format!(
        r#"<?xml version="1.0"?><w:document xmlns:w="{strict_w}" xmlns:r="{strict_r}"><w:body><w:p/></w:body></w:document>"#
    );
    let mut package = package_with_document(xml, &[], &[]);
    let chunk = package
        .add_alt(
            Import::link("https://example.invalid/import.html").unwrap(),
            Some(false),
        )
        .unwrap();
    let document_uri = PackURI::new("/word/document.xml").unwrap();
    let relationship = package
        .opc_package()
        .get_part(&document_uri)
        .unwrap()
        .rels()
        .get(chunk.relationship().as_str())
        .unwrap();
    assert_eq!(
        relationship.reltype(),
        "http://purl.oclc.org/ooxml/officeDocument/relationships/afChunk"
    );
    assert!(relationship.is_external());
    package.save(&path).unwrap();

    let reopened = Package::open(&path).unwrap();
    let document = reopened.document().unwrap();
    let chunk = document.alts().unwrap().pop().unwrap();
    match document.alt_target(&chunk).unwrap() {
        Target::Link(uri) => {
            assert_eq!(uri, "https://example.invalid/import.html")
        },
        Target::Part(_) => panic!("expected external target"),
    }
    assert!(document.resolve_alt(&chunk).is_err());
}

#[test]
fn removal_preserves_a_target_still_referenced_by_another_relationship() {
    let xml = format!(
        r#"<w:document xmlns:w="{W}" xmlns:r="{R}"><w:body><w:altChunk r:id="rIdA"/><w:altChunk r:id="rIdB"/></w:body></w:document>"#
    );
    let mut package = package_with_document(
        xml,
        &[
            (
                "rIdA",
                rt::MS_ALTERNATIVE_FORMAT_IMPORT,
                "shared.html",
                false,
            ),
            (
                "rIdB",
                rt::MS_ALTERNATIVE_FORMAT_IMPORT,
                "shared.html",
                false,
            ),
        ],
        &[("/word/shared.html", "text/html", b"shared")],
    );
    package.remove_alt(0).unwrap();
    let target = PackURI::new("/word/shared.html").unwrap();
    assert!(package.opc_package().get_part(&target).is_ok());
    assert_eq!(package.document_mut().unwrap().alts().len(), 1);
}

#[test]
fn replace_moves_in_new_payload_and_reclaims_the_old_graph() {
    let mut package = Package::new().unwrap();
    let old = package
        .add_alt(Import::data(Data::Html(b"old".to_vec())), None)
        .unwrap();
    package
        .add_alt(Import::data(Data::Text(b"keep".to_vec())), None)
        .unwrap();

    let document_uri = PackURI::new("/word/document.xml").unwrap();
    let old_target = package
        .opc_package()
        .get_part(&document_uri)
        .unwrap()
        .rels()
        .get(old.relationship().as_str())
        .unwrap()
        .target_partname()
        .unwrap();

    let replaced = package
        .replace_alt(
            0,
            Import::data(Data::Rtf(br"{\rtf1 replacement}".to_vec())),
            Some(false),
        )
        .unwrap();
    assert_eq!(replaced, old);
    assert!(package.opc_package().get_part(&old_target).is_err());

    let directory = tempfile::tempdir().unwrap();
    let path = directory.path().join("replaced.docx");
    package.save(&path).unwrap();
    let reopened = Package::open(&path).unwrap();
    let document = reopened.document().unwrap();
    let chunks = document.alts().unwrap();
    assert_eq!(chunks.len(), 2);
    assert_eq!(chunks[0].match_source(), Some(false));
    assert_eq!(document.resolve_alt(&chunks[0]).unwrap().kind(), Kind::Rtf);
    assert_eq!(document.resolve_alt(&chunks[1]).unwrap().kind(), Kind::Text);
}

#[test]
fn malformed_relationship_and_invalid_index_leave_the_package_unchanged() {
    let xml = format!(
        r#"<w:document xmlns:w="{W}" xmlns:r="{R}"><w:body><w:altChunk r:id="rIdBad"/></w:body></w:document>"#
    );
    let mut package = package_with_document(
        xml,
        &[("rIdBad", rt::HYPERLINK, "https://example.invalid", true)],
        &[],
    );
    let part_count = package.opc_package().part_count();
    assert!(package.remove_alt(0).is_err());
    assert_eq!(package.document_mut().unwrap().alts().len(), 1);
    assert!(
        package
            .insert_alt(2, Import::data(Data::Html(b"new".to_vec())), None,)
            .is_err()
    );
    assert_eq!(package.opc_package().part_count(), part_count);
    assert_eq!(package.document_mut().unwrap().alts().len(), 1);
}

#[test]
fn authoring_refuses_to_exceed_the_parser_anchor_limit() {
    let anchors = r#"<w:altChunk r:id="shared"/>"#.repeat(MAX_CHUNKS);
    let xml = format!(
        r#"<w:document xmlns:w="{W}" xmlns:r="{R}"><w:body>{anchors}</w:body></w:document>"#
    );
    let mut package = package_with_document(
        xml,
        &[(
            "shared",
            rt::MS_ALTERNATIVE_FORMAT_IMPORT,
            "shared.html",
            false,
        )],
        &[("/word/shared.html", "text/html", b"shared")],
    );
    let part_count = package.opc_package().part_count();

    assert!(
        package
            .add_alt(Import::data(Data::Html(b"too many".to_vec())), None)
            .is_err()
    );
    assert_eq!(package.opc_package().part_count(), part_count);
    assert_eq!(package.document_mut().unwrap().alts().len(), MAX_CHUNKS);
}

#[test]
fn semantic_insert_preserves_unrelated_mce_and_unknown_xml() {
    let source = format!(
        r#"<?xml version="1.0"?><w:document xmlns:w="{W}" xmlns:r="{R}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:test"><w:body><mc:AlternateContent><mc:Choice Requires="x"><x:sentinel value="untouched"/></mc:Choice><mc:Fallback><w:p/></mc:Fallback></mc:AlternateContent><w:altChunk r:id="existing"><x:opaque keep="yes">foreign</x:opaque></w:altChunk><x:tail keep="yes"/></w:body></w:document>"#
    );
    let mut package = package_with_document(
        source,
        &[(
            "existing",
            rt::MS_ALTERNATIVE_FORMAT_IMPORT,
            "existing.html",
            false,
        )],
        &[("/word/existing.html", "text/html", b"existing")],
    );
    package
        .insert_alt(
            0,
            Import::data(Data::Html(b"<p>opaque</p>".to_vec())),
            Some(true),
        )
        .unwrap();
    let output = package.document_mut().unwrap().to_xml().unwrap();
    assert!(output.contains(r#"<x:sentinel value="untouched"/>"#));
    assert!(output.contains(r#"<x:opaque keep="yes">foreign</x:opaque>"#));
    assert!(output.contains(r#"<x:tail keep="yes"/>"#));
    assert!(output.contains("<w:altChunkPr><w:matchSrc w:val=\"1\"/>"));
}
