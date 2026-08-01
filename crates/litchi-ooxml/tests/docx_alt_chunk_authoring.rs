use litchi_ooxml::docx::writer::MutableDocument;
use litchi_ooxml::docx::{
    AltChunk, AltChunkNamespace, AlternativeFormatData, AlternativeFormatImport,
    AlternativeFormatKind, AlternativeFormatTarget, Package,
};
use litchi_opc::OpcPackage;
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::packuri::PackURI;
use litchi_opc::part::{BlobPart, Part};

const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

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
fn generated_internal_formats_have_canonical_content_types_and_reorder() {
    let directory = tempfile::tempdir().unwrap();
    let path = directory.path().join("formats.docx");
    let mut package = Package::new().unwrap();
    let imports = [
        AlternativeFormatData::Html(b"<p>html</p>".to_vec()),
        AlternativeFormatData::Rtf(br"{\rtf1 rtf}".to_vec()),
        AlternativeFormatData::PlainText(b"plain".to_vec()),
        AlternativeFormatData::Xml(b"<root/>".to_vec()),
        AlternativeFormatData::WordprocessingMl(b"opaque nested package".to_vec()),
    ];
    for (index, data) in imports.into_iter().enumerate() {
        package
            .add_alt_chunk(
                AlternativeFormatImport::Internal(data),
                (index == 0).then_some(true),
            )
            .unwrap();
    }
    package.move_alt_chunk(4, 0).unwrap();
    package.save(&path).unwrap();

    let reopened = Package::open(&path).unwrap();
    let document = reopened.document().unwrap();
    let chunks = document.alt_chunks().unwrap();
    assert_eq!(chunks.len(), 5);
    let actual = chunks
        .iter()
        .map(|chunk| {
            let part = document.resolve_alt_chunk(chunk).unwrap();
            (part.kind(), part.content_type().to_string())
        })
        .collect::<Vec<_>>();
    assert_eq!(
        actual,
        vec![
            (
                AlternativeFormatKind::WordprocessingMl,
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
                    .into(),
            ),
            (AlternativeFormatKind::Html, "text/html".into()),
            (AlternativeFormatKind::Rtf, "application/rtf".into()),
            (AlternativeFormatKind::PlainText, "text/plain".into()),
            (AlternativeFormatKind::Xml, "application/xml".into()),
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
        .add_alt_chunk(
            AlternativeFormatImport::External("https://example.invalid/import.html".into()),
            Some(false),
        )
        .unwrap();
    let document_uri = PackURI::new("/word/document.xml").unwrap();
    let relationship = package
        .opc_package()
        .get_part(&document_uri)
        .unwrap()
        .rels()
        .get(chunk.relationship_id())
        .unwrap();
    assert_eq!(
        relationship.reltype(),
        "http://purl.oclc.org/ooxml/officeDocument/relationships/afChunk"
    );
    assert!(relationship.is_external());
    package.save(&path).unwrap();

    let reopened = Package::open(&path).unwrap();
    let document = reopened.document().unwrap();
    let chunk = document.alt_chunks().unwrap().pop().unwrap();
    match document.resolve_alt_chunk_target(&chunk).unwrap() {
        AlternativeFormatTarget::External(uri) => {
            assert_eq!(uri, "https://example.invalid/import.html")
        },
        AlternativeFormatTarget::Internal(_) => panic!("expected external target"),
    }
    assert!(document.resolve_alt_chunk(&chunk).is_err());
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
    package.remove_alt_chunk(0).unwrap();
    let target = PackURI::new("/word/shared.html").unwrap();
    assert!(package.opc_package().get_part(&target).is_ok());
    assert_eq!(package.document_mut().unwrap().alt_chunks().len(), 1);
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
    assert!(package.remove_alt_chunk(0).is_err());
    assert_eq!(package.document_mut().unwrap().alt_chunks().len(), 1);
    assert!(
        package
            .insert_alt_chunk(
                2,
                AlternativeFormatImport::Internal(AlternativeFormatData::Html(b"new".to_vec())),
                None,
            )
            .is_err()
    );
    assert_eq!(package.opc_package().part_count(), part_count);
    assert_eq!(package.document_mut().unwrap().alt_chunks().len(), 1);
}

#[test]
fn direct_body_mutation_preserves_unrelated_mce_and_unknown_xml() {
    let source = format!(
        r#"<?xml version="1.0"?><w:document xmlns:w="{W}" xmlns:r="{R}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:test"><w:body><mc:AlternateContent><mc:Choice Requires="x"><x:sentinel value="untouched"/></mc:Choice><mc:Fallback><w:p/></mc:Fallback></mc:AlternateContent><x:tail keep="yes"/></w:body></w:document>"#
    );
    let mut document = MutableDocument::from_xml(&source).unwrap();
    document
        .insert_alt_chunk(
            0,
            AltChunk::new("rIdAltChunk1", Some(true)).unwrap(),
            AltChunkNamespace::Transitional,
        )
        .unwrap();
    let output = document.to_xml().unwrap();
    assert!(output.contains(r#"<x:sentinel value="untouched"/>"#));
    assert!(output.contains(r#"<x:tail keep="yes"/>"#));
    assert!(output.contains("<w:altChunkPr><w:matchSrc w:val=\"1\"/>"));
}
