use litchi_docx::alt::{Data, Import, Kind, MAX_CHUNKS, Target};
use litchi_docx::{Block, Builder, Package};
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

/// The package as a reader sees it after publication: the alt-chunk
/// mutators stage their edits until the package is written.
fn published(package: &Package) -> Package {
    Package::from_bytes(&package.clone().to_bytes().unwrap()).unwrap()
}

fn published_main_xml(package: &Package) -> String {
    let published = published(package);
    String::from_utf8(
        published
            .opc_package()
            .main_document_part()
            .unwrap()
            .blob()
            .to_vec(),
    )
    .unwrap()
}

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
    assert!(matches!(blocks[0], Block::Paragraph(_)));
    let Block::Alt(chunk) = &blocks[1] else {
        panic!("missing ordered alternative-format anchor")
    };
    assert!(matches!(blocks[2], Block::Paragraph(_)));
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
    let package = package_with_document(
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
        let Block::Paragraph(paragraph) = &blocks[0] else {
            panic!("missing active fallback paragraph")
        };
        assert_eq!(paragraph.text().unwrap(), "active");
        assert!(matches!(blocks[1], Block::Alt(_)));
    }
    assert_eq!(
        published(&package)
            .document()
            .unwrap()
            .alts()
            .unwrap()
            .len(),
        1
    );
    // Alt-chunk CRUD on an opened package was removed with the writer
    // hydration of opened packages; only the shared active-branch reads remain.
    let output = published_main_xml(&package);
    assert!(output.contains(r#"<w:altChunk r:id="inactive"/>"#));
    assert!(output.contains(r#"<w:altChunk r:id="fallback"/>"#));
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
    let mut builder = Builder::new().unwrap();
    let wordprocessingml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body/></w:document>"#;
    let imports = [
        Data::Docm(wordprocessingml.to_vec()),
        Data::Dotx(wordprocessingml.to_vec()),
        Data::Dotm(wordprocessingml.to_vec()),
        Data::Mime(b"From: sender@example.invalid\r\n\r\nbody".to_vec()),
        Data::Html(b"<p>html</p>".to_vec()),
        Data::Xhtml(b"<html xmlns=\"http://www.w3.org/1999/xhtml\"/>".to_vec()),
        Data::Rtf(br"{\rtf1 rtf}".to_vec()),
        Data::Text(b"plain".to_vec()),
        Data::Xml(b"<root/>".to_vec()),
        Data::Docx(wordprocessingml.to_vec()),
    ];
    for (index, data) in imports.into_iter().enumerate() {
        builder
            .add_alt(Import::data(data), (index == 0).then_some(true))
            .unwrap();
    }
    builder.move_alt(9, 0).unwrap();
    let package = builder.build().unwrap();
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
fn external_target_is_returned_without_access() {
    let directory = tempfile::tempdir().unwrap();
    let path = directory.path().join("external.docx");
    let mut builder = Builder::new().unwrap();
    builder.document().unwrap().add_paragraph();
    let chunk = builder
        .add_alt(
            Import::link("https://example.invalid/import.html").unwrap(),
            Some(false),
        )
        .unwrap();
    let package = builder.build().unwrap();
    let document_uri = PackURI::new("/word/document.xml").unwrap();
    let relationship = package
        .opc_package()
        .get_part(&document_uri)
        .unwrap()
        .rels()
        .get(chunk.relationship().as_str())
        .unwrap();
    assert_eq!(relationship.reltype(), rt::MS_ALTERNATIVE_FORMAT_IMPORT);
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
fn replace_moves_in_new_payload_and_reclaims_the_old_graph() {
    fn authored() -> (Builder, litchi_docx::alt::Chunk) {
        let mut builder = Builder::new().unwrap();
        let old = builder
            .add_alt(Import::data(Data::Html(b"old".to_vec())), None)
            .unwrap();
        builder
            .add_alt(Import::data(Data::Text(b"keep".to_vec())), None)
            .unwrap();
        (builder, old)
    }

    // An identical draft built without the replacement names the old target.
    let (unreplaced, old) = authored();
    let unreplaced = unreplaced.build().unwrap();
    let document_uri = PackURI::new("/word/document.xml").unwrap();
    let old_target = unreplaced
        .opc_package()
        .get_part(&document_uri)
        .unwrap()
        .rels()
        .get(old.relationship().as_str())
        .unwrap()
        .target_partname()
        .unwrap();
    assert!(unreplaced.opc_package().get_part(&old_target).is_ok());

    let (mut builder, authored_old) = authored();
    assert_eq!(authored_old, old);
    let replaced = builder
        .replace_alt(
            0,
            Import::data(Data::Rtf(br"{\rtf1 replacement}".to_vec())),
            Some(false),
        )
        .unwrap();
    assert_eq!(replaced, old);
    let package = builder.build().unwrap();
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
fn authoring_refuses_to_exceed_the_parser_anchor_limit() {
    let mut builder = Builder::new().unwrap();
    for _ in 0..MAX_CHUNKS {
        builder
            .add_alt(
                Import::link("https://example.invalid/shared.html").unwrap(),
                None,
            )
            .unwrap();
    }

    assert!(
        builder
            .add_alt(Import::data(Data::Html(b"too many".to_vec())), None)
            .is_err()
    );
    let package = builder.build().unwrap();
    assert!(
        package
            .opc_package()
            .iter_parts()
            .all(|part| part.content_type() != "text/html")
    );
    assert_eq!(
        published(&package)
            .document()
            .unwrap()
            .alts()
            .unwrap()
            .len(),
        MAX_CHUNKS
    );
}
