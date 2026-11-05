//! Untrusted durable metadata restoration must retain exact source guards.
#![allow(clippy::unwrap_used, clippy::expect_used, reason = "test assertions")]

use std::sync::Arc;

use litchi_opc::{
    BlobPart, OpcError, OpcPackage, PackURI, PackageWriter, SourceLimitKind, SourceLimits,
};
use soapberry_zip::office::{ArchiveReader, StreamingArchiveWriter};

const TYPES: &[u8] = br#"<t:Types xmlns:t='http://schemas.openxmlformats.org/package/2006/content-types'><!--types--><t:Default Extension='rels' ContentType='application/vnd.openxmlformats-package.relationships+xml'/><t:Default Extension='bin' ContentType='application/octet-stream'/></t:Types>"#;
const RELS: &[u8] = br#"<r:Relationships xmlns:r='http://schemas.openxmlformats.org/package/2006/relationships'><!--rels--><r:Relationship TargetMode='External' Target='https://example.test/old' Type='urn:test' Id='rId1'/></r:Relationships>"#;

fn package(rels: bool) -> OpcPackage {
    let mut writer =
        StreamingArchiveWriter::new(soapberry_zip::office::StreamingArchiveLimits::default());
    writer.write_deflated("[Content_Types].xml", TYPES).unwrap();
    writer.write_deflated("custom/item.bin", b"opaque").unwrap();
    writer.write_deflated("_rels/.rels", br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>"#).unwrap();
    if rels {
        writer
            .write_deflated("custom/_rels/item.bin.rels", RELS)
            .unwrap();
    }
    OpcPackage::from_vec(writer.finish_to_bytes().unwrap(), SourceLimits::default()).unwrap()
}

fn owner() -> PackURI {
    PackURI::new("/custom/item.bin").unwrap()
}

fn changed(bytes: &[u8], before: &str, after: &str) -> Arc<Vec<u8>> {
    Arc::new(
        std::str::from_utf8(bytes)
            .unwrap()
            .replace(before, after)
            .into_bytes(),
    )
}

#[test]
fn wire_metadata_restores_exact_bytes_after_publication_and_reopen() {
    let mut package = package(true);
    let limits = SourceLimits::default();
    let types = package
        .source_content_types(SourceLimits::default())
        .unwrap();
    let rels = package
        .source_relationships(&owner(), SourceLimits::default())
        .unwrap();
    let next_types = changed(TYPES, "<!--types-->", "<!--types--><?overlay retained?>");
    let next_rels = changed(RELS, "/old", "/new");
    assert!(
        package
            .try_replace_content_types_bytes(types.bytes(), Arc::clone(&next_types), limits)
            .unwrap()
    );
    assert!(
        package
            .try_replace_relationships_bytes(&rels, Arc::clone(&next_rels), true, limits)
            .unwrap()
    );
    let saved = PackageWriter::to_bytes(&package).unwrap();
    let archive =
        ArchiveReader::new(&saved, soapberry_zip::office::ArchiveLimits::default()).unwrap();
    assert_eq!(archive.read("[Content_Types].xml").unwrap(), *next_types);
    assert_eq!(
        archive.read("custom/_rels/item.bin.rels").unwrap(),
        *next_rels
    );
    let mut reopened = OpcPackage::from_bytes(&saved, SourceLimits::default()).unwrap();
    let expected = reopened
        .source_relationships(&owner(), SourceLimits::default())
        .unwrap();
    assert!(
        reopened
            .try_replace_content_types_bytes(&next_types, Arc::new(TYPES.to_vec()), limits)
            .unwrap()
    );
    assert!(
        reopened
            .try_replace_relationships_bytes(&expected, Arc::new(RELS.to_vec()), true, limits)
            .unwrap()
    );
    assert_eq!(
        reopened
            .source_content_types(SourceLimits::default())
            .unwrap()
            .bytes(),
        TYPES
    );
    assert_eq!(
        reopened
            .source_relationships(&owner(), SourceLimits::default())
            .unwrap()
            .bytes(),
        RELS
    );
}

#[test]
fn stale_guards_refuse_wire_images_without_overwriting_destination_overlays() {
    let mut package = package(true);
    let limits = SourceLimits::default();
    let types = package
        .source_content_types(SourceLimits::default())
        .unwrap();
    let rels = package
        .source_relationships(&owner(), SourceLimits::default())
        .unwrap();
    package
        .try_replace_content_types_bytes(
            types.bytes(),
            changed(TYPES, "<!--types-->", "<!--destination-->"),
            limits,
        )
        .unwrap();
    package
        .try_replace_relationships_bytes(
            &rels,
            changed(RELS, "<!--rels-->", "<!--destination-->"),
            true,
            limits,
        )
        .unwrap();
    let before = PackageWriter::to_bytes(&package).unwrap();
    // Staleness wins even over a malformed replacement payload.
    assert!(
        matches!(package.try_replace_content_types_bytes(types.bytes(), Arc::new(b"<".to_vec()), limits), Err(OpcError::InvalidContentTypesManifest(message)) if message.contains("stale"))
    );
    let error = package
        .try_replace_relationships_bytes(&rels, Arc::new(b"<".to_vec()), true, limits)
        .unwrap_err();
    assert!(error.to_string().contains("stale"));
    assert_eq!(PackageWriter::to_bytes(&package).unwrap(), before);
}

#[test]
fn invalid_and_oversized_wire_metadata_leave_the_package_unchanged() {
    let mut package = package(true);
    let limits = SourceLimits::default();
    let types = package
        .source_content_types(SourceLimits::default())
        .unwrap();
    let rels = package
        .source_relationships(&owner(), SourceLimits::default())
        .unwrap();
    let before = PackageWriter::to_bytes(&package).unwrap();
    for bytes in [b"<".as_slice(), b"<wrong/>", b"<!DOCTYPE x><x/>"] {
        assert!(
            package
                .try_replace_content_types_bytes(types.bytes(), Arc::new(bytes.to_vec()), limits)
                .is_err()
        );
        assert!(
            package
                .try_replace_relationships_bytes(&rels, Arc::new(bytes.to_vec()), true, limits)
                .is_err()
        );
        assert_eq!(PackageWriter::to_bytes(&package).unwrap(), before);
    }
    // An XML-looking metadata root is insufficient: attributes and token
    // content still pass the owner scanner and publication audit.
    for bytes in [
        changed(TYPES, "Extension='bin'", "Extension='bin' Extension='bin'"),
        changed(TYPES, "<!--types-->", "<!--bad\u{1}-->"),
    ] {
        assert!(
            package
                .try_replace_content_types_bytes(types.bytes(), bytes, limits)
                .is_err()
        );
    }
    for bytes in [
        changed(RELS, "Id='rId1'", "Id='rId1' Id='rId1'"),
        changed(RELS, "<!--rels-->", "<!--bad\u{1}-->"),
    ] {
        assert!(
            package
                .try_replace_relationships_bytes(&rels, bytes, true, limits)
                .is_err()
        );
    }
    let too_many_types = changed(
        TYPES,
        "<!--types-->",
        &format!("<!--{}-->", "x".repeat(256)),
    );
    let types_limit = SourceLimits::builder()
        .max_content_types_bytes(TYPES.len())
        .unwrap()
        .build()
        .unwrap();
    assert!(matches!(
        package.try_replace_content_types_bytes(types.bytes(), too_many_types, types_limit),
        Err(OpcError::Limit {
            kind: SourceLimitKind::ContentTypesBytes,
            ..
        })
    ));
    let too_many_rels = changed(RELS, "<!--rels-->", &format!("<!--{}-->", "x".repeat(256)));
    let rels_limit = SourceLimits::builder()
        .max_relationship_xml_bytes(RELS.len())
        .unwrap()
        .build()
        .unwrap();
    assert!(matches!(
        package.try_replace_relationships_bytes(&rels, too_many_rels, true, rels_limit),
        Err(OpcError::Limit {
            kind: SourceLimitKind::RelationshipXmlBytes,
            ..
        })
    ));
    assert_eq!(PackageWriter::to_bytes(&package).unwrap(), before);
}

#[test]
fn wire_member_presence_is_explicit_and_cannot_discard_lexical_content() {
    let mut package = package(false);
    let limits = SourceLimits::default();
    let absent = package
        .source_relationships(&owner(), SourceLimits::default())
        .unwrap();
    assert!(!absent.member_present());
    let empty = Arc::new(br#"<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><!--keep--></Relationships>"#.to_vec());
    assert!(
        package
            .try_replace_relationships_bytes(&absent, Arc::clone(&empty), false, limits)
            .is_err()
    );
    assert!(
        package
            .try_replace_relationships_bytes(&absent, empty, true, limits)
            .unwrap()
    );
    let present = package
        .source_relationships(&owner(), SourceLimits::default())
        .unwrap();
    assert!(present.member_present());
    let saved = PackageWriter::to_bytes(&package).unwrap();
    assert!(
        ArchiveReader::new(&saved, soapberry_zip::office::ArchiveLimits::default())
            .unwrap()
            .file_names()
            .any(|name| name == "custom/_rels/item.bin.rels")
    );
    assert!(
        package
            .try_replace_relationships_bytes(
                &present,
                Arc::new(absent.bytes().to_vec()),
                false,
                limits
            )
            .unwrap()
    );
    assert_eq!(
        package
            .source_relationships(&owner(), SourceLimits::default())
            .unwrap(),
        absent
    );
    let saved = PackageWriter::to_bytes(&package).unwrap();
    assert!(
        !ArchiveReader::new(&saved, soapberry_zip::office::ArchiveLimits::default())
            .unwrap()
            .file_names()
            .any(|name| name == "custom/_rels/item.bin.rels")
    );
    let root = package
        .source_relationships(&PackURI::new("/").unwrap(), SourceLimits::default())
        .unwrap();
    assert!(
        package
            .try_replace_relationships_bytes(&root, Arc::new(root.bytes().to_vec()), false, limits)
            .is_err()
    );
    assert_eq!(PackageWriter::to_bytes(&package).unwrap(), saved);
}

#[test]
fn signed_wire_replacements_refuse_changes_but_allow_exact_noops() {
    let mut package = package(true);
    package
        .try_add_part(Box::new(BlobPart::new(
            PackURI::new("/_xmlsignatures/origin.sigs").unwrap(),
            "application/vnd.openxmlformats-package.digital-signature-origin".into(),
            Vec::new(),
        )))
        .unwrap();
    let limits = SourceLimits::default();
    let types = package
        .source_content_types(SourceLimits::default())
        .unwrap();
    let rels = package
        .source_relationships(&owner(), SourceLimits::default())
        .unwrap();
    assert!(
        !package
            .try_replace_content_types_bytes(
                types.bytes(),
                Arc::new(types.bytes().to_vec()),
                limits
            )
            .unwrap()
    );
    assert!(
        !package
            .try_replace_relationships_bytes(&rels, Arc::new(rels.bytes().to_vec()), true, limits)
            .unwrap()
    );
    assert!(matches!(
        package.try_replace_content_types_bytes(types.bytes(), Arc::new(b"<".to_vec()), limits),
        Err(OpcError::SignedSourceRequiresExplicitPolicy)
    ));
    assert!(matches!(
        package.try_replace_relationships_bytes(&rels, Arc::new(b"<".to_vec()), true, limits),
        Err(OpcError::SignedSourceRequiresExplicitPolicy)
    ));
    assert_eq!(
        package
            .source_content_types(SourceLimits::default())
            .unwrap(),
        types
    );
    assert_eq!(
        package
            .source_relationships(&owner(), SourceLimits::default())
            .unwrap(),
        rels
    );
}

#[test]
fn entry_ceiling_applies_to_retained_and_authored_metadata_even_for_noops() {
    let mut authored = OpcPackage::new();
    authored
        .try_add_part(Box::new(BlobPart::new(
            owner(),
            "application/octet-stream".into(),
            b"opaque".to_vec(),
        )))
        .unwrap();
    for mut package in [package(true), authored] {
        let types = package
            .source_content_types(SourceLimits::default())
            .unwrap();
        let tight = SourceLimits::builder()
            .max_entry_bytes((types.bytes().len() - 1) as u64)
            .unwrap()
            .build()
            .unwrap();
        assert!(matches!(
            package.source_content_types(tight),
            Err(OpcError::Limit {
                kind: SourceLimitKind::EntryBytes,
                ..
            })
        ));
        assert!(matches!(
            package.try_replace_content_types_bytes(
                types.bytes(),
                Arc::new(types.bytes().to_vec()),
                tight
            ),
            Err(OpcError::Limit {
                kind: SourceLimitKind::EntryBytes,
                ..
            })
        ));
        let exact = SourceLimits::builder()
            .max_entry_bytes(types.bytes().len() as u64)
            .unwrap()
            .build()
            .unwrap();
        assert!(
            !package
                .try_replace_content_types_bytes(
                    types.bytes(),
                    Arc::new(types.bytes().to_vec()),
                    exact
                )
                .unwrap()
        );

        let rels = package
            .source_relationships(&owner(), SourceLimits::default())
            .unwrap();
        let tight = SourceLimits::builder()
            .max_entry_bytes((rels.bytes().len() - 1) as u64)
            .unwrap()
            .build()
            .unwrap();
        assert!(matches!(
            package.source_relationships(&owner(), tight),
            Err(OpcError::Limit {
                kind: SourceLimitKind::EntryBytes,
                ..
            })
        ));
        assert!(matches!(
            package.plan_source_relationships(&owner(), tight),
            Err(OpcError::Limit {
                kind: SourceLimitKind::EntryBytes,
                ..
            })
        ));
        assert!(matches!(
            package.try_replace_relationships_bytes(
                &rels,
                Arc::new(rels.bytes().to_vec()),
                rels.member_present(),
                tight
            ),
            Err(OpcError::Limit {
                kind: SourceLimitKind::EntryBytes,
                ..
            })
        ));
        let exact = SourceLimits::builder()
            .max_entry_bytes(rels.bytes().len() as u64)
            .unwrap()
            .build()
            .unwrap();
        assert!(
            !package
                .try_replace_relationships_bytes(
                    &rels,
                    Arc::new(rels.bytes().to_vec()),
                    rels.member_present(),
                    exact
                )
                .unwrap()
        );
    }
}
