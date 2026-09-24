#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    clippy::panic,
    reason = "focused publication-audit assertions intentionally panic on fixture errors"
)]

//! Caller-supplied XML that a package records as source provenance.
//!
//! The writer publishes retained source bytes without its audit (change
//! 0665), while ADR 0006 requires publication to audit every XML member it
//! writes. So every route that installs caller bytes as provenance must audit
//! them before it mutates the package, as `try_replace_owned_xml_part`
//! already does.
//!
//! The malformed input is an `xml:space` value other than `default` or
//! `preserve` on the document element. The bounded source validator and the
//! typed manifest parsers accept it, so a source package carrying it opens and
//! yields tokens; the publication audit refuses it.

use std::sync::Arc;

use litchi_opc::{OpcError, OpcPackage, PackURI, PackageWriter, TargetMode};
use soapberry_zip::office::{ArchiveReader, StreamingArchiveWriter};

const BOGUS_SPACE: &str = r#" xml:space="bogus""#;
const ITEM: &str = "/custom/item.xml";

fn source(content_types_attribute: &str, relationships_attribute: &str, item: &str) -> Vec<u8> {
    let mut writer = StreamingArchiveWriter::new();
    writer
        .write_deflated(
            "[Content_Types].xml",
            format!(
                r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"{content_types_attribute}><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/></Types>"#
            )
            .as_bytes(),
        )
        .unwrap();
    writer
        .write_deflated(
            "_rels/.rels",
            format!(
                r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"{relationships_attribute}><Relationship Id="rId1" Type="urn:test:item" Target="custom/item.xml"/></Relationships>"#
            )
            .as_bytes(),
        )
        .unwrap();
    writer
        .write_deflated("custom/item.xml", item.as_bytes())
        .unwrap();
    writer.finish_to_bytes().unwrap()
}

fn uri(value: &str) -> PackURI {
    PackURI::new(value).unwrap()
}

fn assert_audit_refusal<T: std::fmt::Debug>(result: litchi_opc::Result<T>, member: &str) {
    match result {
        Err(OpcError::XmlPublication { part, .. }) => assert_eq!(part, member),
        other => panic!("the publication audit must refuse {member}, got {other:?}"),
    }
}

#[test]
fn added_owned_xml_parts_are_audited_before_insertion() {
    let name = uri(ITEM);
    let donor = OpcPackage::from_vec(source("", "", r#"<item xml:space="bogus"/>"#))
        .expect("the source validator admits the donor");
    let token = donor
        .source_xml_part(&name)
        .expect("source provenance yields a token");
    let bytes = Arc::new(token.bytes().to_vec());

    let mut receiver = OpcPackage::new();
    let before = PackageWriter::to_bytes(&receiver).unwrap();
    assert_audit_refusal(receiver.try_add_owned_xml_part(token), ITEM);
    assert!(receiver.part_metadata(&name).is_none());
    assert_audit_refusal(
        receiver.try_add_owned_xml_part_bytes(name.clone(), "application/xml".to_owned(), bytes),
        ITEM,
    );
    assert!(receiver.part_metadata(&name).is_none());
    assert_eq!(PackageWriter::to_bytes(&receiver).unwrap(), before);

    // Well-formed caller XML is still accepted and published unchanged.
    let donor = OpcPackage::from_vec(source("", "", "<item>kept</item>")).unwrap();
    receiver
        .try_add_owned_xml_part(donor.source_xml_part(&name).unwrap())
        .expect("benign caller XML");
    let output = PackageWriter::to_bytes(&receiver).unwrap();
    assert_eq!(
        ArchiveReader::new(&output)
            .unwrap()
            .read("custom/item.xml")
            .unwrap(),
        b"<item>kept</item>"
    );
}

#[test]
fn replaced_content_types_tokens_are_audited_before_installation() {
    let source = source(BOGUS_SPACE, "", "<item/>");
    let mut package = OpcPackage::from_vec(source.clone()).expect("the manifest parser admits it");
    let current = package.source_content_types().unwrap();
    let replacement = current
        .with_part_overrides(&[(&uri(ITEM), "application/xml")], 1 << 20)
        .unwrap();
    assert_ne!(replacement.bytes(), current.bytes());

    assert_audit_refusal(
        package.try_replace_content_types(current.bytes(), &replacement),
        "/[Content_Types].xml",
    );
    assert_eq!(
        package.source_content_types().unwrap().bytes(),
        current.bytes()
    );
    // The untouched source is still published exactly, its own bytes
    // included: source provenance stays exempt.
    assert_eq!(PackageWriter::to_bytes(&package).unwrap(), source);
}

#[test]
fn manifest_tokens_the_read_policy_admits_keep_their_size_under_the_audit() {
    // One lexical token (a comment) larger than the audit's default 4 MiB
    // token budget, in a manifest within the read policy's 8 MiB limit.
    let comment = format!("<!--{}-->", "p".repeat(5 * 1024 * 1024));
    let manifest = |root_attribute: &str, item_override: bool| {
        let item_override = if item_override {
            r#"<Override PartName="/custom/item.xml" ContentType="application/xml"/>"#
        } else {
            ""
        };
        let mut writer = StreamingArchiveWriter::new();
        writer
            .write_deflated(
                "[Content_Types].xml",
                format!(
                    r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"{root_attribute}>{comment}<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>{item_override}</Types>"#
                )
                .as_bytes(),
            )
            .unwrap();
        writer
            .write_deflated(
                "_rels/.rels",
                br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="urn:test:item" Target="custom/item.xml"/></Relationships>"#,
            )
            .unwrap();
        writer
            .write_deflated("custom/item.xml", b"<item/>")
            .unwrap();
        writer.finish_to_bytes().unwrap()
    };

    for (root_attribute, admitted) in [("", true), (BOGUS_SPACE, false)] {
        let donor = OpcPackage::from_vec(manifest(root_attribute, true)).unwrap();
        let token = donor.source_content_types().unwrap();
        let mut package = OpcPackage::from_vec(manifest(root_attribute, false)).unwrap();
        let current = package.source_content_types().unwrap();
        let result = package.try_replace_content_types(current.bytes(), &token);
        if admitted {
            assert!(result.expect("the admitted size is audited, not refused"));
            let output = PackageWriter::to_bytes(&package).unwrap();
            assert_eq!(
                ArchiveReader::new(&output)
                    .unwrap()
                    .read("[Content_Types].xml")
                    .unwrap(),
                token.bytes()
            );
        } else {
            assert_audit_refusal(result, "/[Content_Types].xml");
        }
    }
}

#[test]
fn replaced_relationship_tokens_are_audited_before_installation() {
    let source = source("", BOGUS_SPACE, "<item/>");
    let mut package =
        OpcPackage::from_vec(source.clone()).expect("the relationships parser admits it");
    let root = uri("/");
    let expected = package.source_relationships(&root).unwrap();
    let replacement = expected
        .with_relationship(
            "urn:test:again",
            "custom/item.xml",
            "rId2",
            TargetMode::Internal,
            1 << 20,
        )
        .unwrap();

    assert_audit_refusal(
        package.try_replace_relationships(&expected, &replacement),
        "/_rels/.rels",
    );
    assert_eq!(package.source_relationships(&root).unwrap(), expected);
    assert_eq!(package.rels().len(), 1);
    assert_eq!(PackageWriter::to_bytes(&package).unwrap(), source);
}

#[test]
fn batch_source_tokens_are_audited_before_any_mutation() {
    let root = uri("/");
    let item = uri(ITEM);

    // A malformed relationships token with a clean manifest.
    let source_with_bad_rels = source("", BOGUS_SPACE, "<item/>");
    let mut package = OpcPackage::from_vec(source_with_bad_rels.clone()).unwrap();
    let content_types = package.source_content_types().unwrap();
    let expected = package.source_relationships(&root).unwrap();
    let replacement = expected
        .with_relationship(
            "urn:test:again",
            "custom/item.xml",
            "rId2",
            TargetMode::Internal,
            1 << 20,
        )
        .unwrap();
    assert_audit_refusal(
        package.try_add_parts_with_source_tokens(
            content_types.bytes(),
            &content_types,
            &expected,
            &replacement,
            Vec::new(),
        ),
        "/_rels/.rels",
    );
    assert_eq!(package.source_relationships(&root).unwrap(), expected);
    assert_eq!(
        PackageWriter::to_bytes(&package).unwrap(),
        source_with_bad_rels
    );

    // A malformed manifest token with an unchanged relationships token.
    let source_with_bad_manifest = source(BOGUS_SPACE, "", "<item/>");
    let mut package = OpcPackage::from_vec(source_with_bad_manifest.clone()).unwrap();
    let content_types = package.source_content_types().unwrap();
    let replacement = content_types
        .with_part_overrides(&[(&item, "application/xml")], 1 << 20)
        .unwrap();
    let relationships = package.source_relationships(&root).unwrap();
    assert_audit_refusal(
        package.try_add_parts_with_source_tokens(
            content_types.bytes(),
            &replacement,
            &relationships,
            &relationships,
            Vec::new(),
        ),
        "/[Content_Types].xml",
    );
    assert_eq!(
        package.source_content_types().unwrap().bytes(),
        content_types.bytes()
    );
    assert_eq!(
        PackageWriter::to_bytes(&package).unwrap(),
        source_with_bad_manifest
    );

    // Tokens equal to the retained source are that source, not caller XML,
    // and stay exempt: an unchanged batch on the same package succeeds.
    package
        .try_add_parts_with_source_tokens(
            content_types.bytes(),
            &content_types,
            &relationships,
            &relationships,
            Vec::new(),
        )
        .expect("unchanged source tokens are not re-audited");
}
