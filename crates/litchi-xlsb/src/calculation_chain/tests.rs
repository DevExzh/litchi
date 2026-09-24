#![allow(
    clippy::expect_used,
    reason = "focused owner tests construct small exact package graphs and panic on fixture setup"
)]

use litchi_opc::constants::relationship_type as rt;
use litchi_opc::{BlobPart, PackURI, Part, TargetMode};

use super::*;
use crate::Package;
use crate::raw::{Kind, Writer};

fn opaque_records() -> Vec<u8> {
    let mut bytes = Vec::new();
    let mut writer = Writer::new(&mut bytes);
    writer
        .write_record(Kind::new(0x123).expect("kind"), &[0x01, 0x02])
        .expect("first record");
    writer
        .write_record(Kind::new(0x234).expect("kind"), &[])
        .expect("second record");
    writer
        .write_record(Kind::new(0x345).expect("kind"), &[0xFE, 0xED, 0xFA, 0xCE])
        .expect("third record");
    bytes
}

fn package_with_chain(bytes: Vec<u8>) -> Package {
    package_with_chain_at(bytes, "/xl/calcChain.bin", "calcChain.bin")
}

fn package_with_chain_at(bytes: Vec<u8>, physical_name: &str, target_ref: &str) -> Package {
    let mut package = Package::create().expect("base package").into_opc();
    let workbook_name = package
        .main_document_part()
        .expect("workbook")
        .partname()
        .clone();
    let chain_name = PackURI::new(physical_name).expect("chain URI");
    package.add_part(Box::new(BlobPart::new(
        chain_name,
        CONTENT_TYPE.to_string(),
        bytes,
    )));
    package
        .get_part_mut(&workbook_name)
        .expect("workbook")
        .rels_mut()
        .try_add_relationship(
            RELATIONSHIP_TYPE.to_string(),
            target_ref.to_string(),
            "rIdCalcChain".to_string(),
            TargetMode::Internal,
        )
        .expect("chain relationship");
    Package::from_opc(package).expect("validated package")
}

fn opened_package_with_explicit_empty_chain_relationships(bytes: Vec<u8>) -> Package {
    const EMPTY_RELATIONSHIPS: &[u8] =
        br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>"#;

    let package = package_with_chain(bytes).into_opc();
    let root = PackURI::new("/").expect("package URI");
    let content_types = package.source_content_types().expect("content types");
    let root_relationships = package
        .source_relationships(&root)
        .expect("root relationships");
    let chain = PackURI::new("/xl/calcChain.bin").expect("chain URI");

    let mut writer = soapberry_zip::office::StreamingArchiveWriter::new();
    writer
        .write_stored("[Content_Types].xml", content_types.bytes())
        .expect("content types member");
    writer
        .write_stored("_rels/.rels", root_relationships.bytes())
        .expect("root relationships member");

    let mut parts = package
        .try_iter_parts()
        .map(|part| part.expect("part payload decodes"))
        .collect::<Vec<_>>();
    parts.sort_unstable_by_key(|part| part.partname().as_str());
    for part in parts {
        writer
            .write_stored(part.partname().membername(), part.blob())
            .expect("part member");
        let relationships = package
            .source_relationships(part.partname())
            .expect("part relationships");
        if part.partname().is_equivalent_to(&chain) {
            let member = part
                .partname()
                .rels_uri()
                .expect("chain relationships URI")
                .membername()
                .to_owned();
            writer
                .write_stored(&member, EMPTY_RELATIONSHIPS)
                .expect("explicit empty chain relationships member");
        } else if relationships.member_present() || !part.rels().is_empty() {
            let member = part
                .partname()
                .rels_uri()
                .expect("relationships URI")
                .membername()
                .to_owned();
            writer
                .write_stored(&member, relationships.bytes())
                .expect("relationships member");
        }
    }

    Package::from_bytes(writer.finish_to_bytes().expect("ZIP bytes")).expect("opened package")
}

fn package_with_unrelated_root_relationship(id: String) -> Package {
    let mut package = package_with_chain(opaque_records()).into_opc();
    package
        .rels_mut()
        .try_add_relationship(
            "urn:litchi:unrelated-metadata".to_string(),
            "https://example.test/unrelated".to_string(),
            id,
            TargetMode::External,
        )
        .expect("unrelated root relationship");
    Package::from_opc(package).expect("validated package")
}

#[test]
fn opaque_snapshot_can_request_generic_record_diagnostic() {
    let bytes = opaque_records();
    let package = package_with_chain(bytes.clone());
    let snapshot = package.calculation_chain().expect("snapshot");
    let part = snapshot.part().expect("present");
    assert_eq!(part.bytes(), bytes.as_slice());
    assert_eq!(part.record_count().expect("diagnostic inventory"), 3);
    assert_eq!(part.content_type(), CONTENT_TYPE);
    assert_eq!(part.part_name(), "/xl/calcChain.bin");
}

#[test]
fn empty_opaque_stream_is_preserved_without_inferred_record_grammar() {
    let package = package_with_chain(Vec::new());
    let snapshot = package.calculation_chain().expect("snapshot");
    assert_eq!(
        snapshot
            .part()
            .expect("present")
            .record_count()
            .expect("diagnostic inventory"),
        0
    );
    assert!(snapshot.part().expect("present").bytes().is_empty());
}

#[test]
fn removal_and_inverse_restore_case_equivalent_physical_source() {
    let bytes = opaque_records();
    let package = package_with_chain_at(bytes.clone(), "/xl/CALCCHAIN.BIN", "calcChain.bin");
    let before = package.calculation_chain().expect("before");
    assert_eq!(
        before.part().expect("part").part_name(),
        "/xl/CALCCHAIN.BIN"
    );

    let mut transaction = before.edit();
    assert!(transaction.remove().expect("remove"));
    let commit = transaction.commit().expect("commit");
    assert!(commit.changed());

    let removed = package
        .apply_calculation_chain(&commit)
        .expect("remove publication");
    assert!(!removed.calculation_chain().expect("removed").is_present());

    let restored = removed
        .apply_calculation_chain_patch(&commit.patch().inverse())
        .expect("inverse publication");
    let after = restored.calculation_chain().expect("restored");
    assert_eq!(after, before);
    assert_eq!(after.part().expect("part").bytes(), bytes.as_slice());
    assert_eq!(after.part().expect("part").part_name(), "/xl/CALCCHAIN.BIN");
}

#[test]
fn exact_noop_does_not_publish_a_changed_patch() {
    let package = package_with_chain(opaque_records());
    let before = package.calculation_chain().expect("before");
    let commit = before.edit().commit().expect("no-op");
    assert!(!commit.changed());
    assert!(commit.patch().is_empty());
    let after = package
        .apply_calculation_chain(&commit)
        .expect("no-op publication");
    assert_eq!(after.calculation_chain().expect("after"), before);
}

#[test]
fn empty_owner_relationships_are_captured_and_restored() {
    const EMPTY_RELATIONSHIPS: &[u8] =
        br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>"#;

    let package = opened_package_with_explicit_empty_chain_relationships(opaque_records());
    let before = package.calculation_chain().expect("before");
    let owner = before.source().owner.as_ref().expect("owner source");
    assert!(owner.relationships.member_present());
    assert_eq!(owner.relationships.bytes(), EMPTY_RELATIONSHIPS);

    let mut transaction = before.edit();
    transaction.remove().expect("remove");
    let commit = transaction.commit().expect("commit");
    let removed = package
        .apply_calculation_chain(&commit)
        .expect("remove publication");
    let restored = removed
        .apply_calculation_chain_patch(&commit.patch().inverse())
        .expect("inverse publication");
    let restored_snapshot = restored.calculation_chain().expect("restored");
    assert_eq!(restored_snapshot, before);
    let restored_owner = restored_snapshot
        .source()
        .owner
        .as_ref()
        .expect("owner source");
    assert!(restored_owner.relationships.member_present());
    assert_eq!(restored_owner.relationships.bytes(), EMPTY_RELATIONSHIPS);

    let reopened = Package::from_bytes(restored.to_bytes().expect("restored ZIP")).expect("reopen");
    let chain = PackURI::new("/xl/calcChain.bin").expect("chain URI");
    let restored_member = reopened
        .opc_package()
        .source_relationships(&chain)
        .expect("reopened chain relationships");
    assert!(restored_member.member_present());
    assert_eq!(restored_member.bytes(), EMPTY_RELATIONSHIPS);
}

#[test]
fn duplicate_workbook_calculation_chain_edges_are_rejected() {
    let mut package = package_with_chain(opaque_records()).into_opc();
    let workbook_name = package
        .main_document_part()
        .expect("workbook")
        .partname()
        .clone();
    package
        .get_part_mut(&workbook_name)
        .expect("workbook")
        .rels_mut()
        .try_add_relationship(
            RELATIONSHIP_TYPE.to_string(),
            "calcChain.bin".to_string(),
            "rIdDuplicateCalcChain".to_string(),
            TargetMode::Internal,
        )
        .expect("duplicate edge");
    let package = Package::from_opc(package).expect("validated package");
    assert!(package.calculation_chain().is_err());
}

#[test]
fn foreign_inbound_edge_is_rejected_while_owner_exists() {
    let mut package = package_with_chain(opaque_records()).into_opc();
    let foreign_name = PackURI::new("/xl/foreign-owner.bin").expect("foreign URI");
    package.add_part(Box::new(BlobPart::new(
        foreign_name.clone(),
        "application/octet-stream".to_string(),
        Vec::new(),
    )));
    package
        .get_part_mut(&foreign_name)
        .expect("foreign part")
        .rels_mut()
        .try_add_relationship(
            "urn:litchi:foreign-inbound".to_string(),
            "calcChain.bin".to_string(),
            "rIdForeignInbound".to_string(),
            TargetMode::Internal,
        )
        .expect("foreign inbound edge");
    let package = Package::from_opc(package).expect("validated package");
    assert!(package.calculation_chain().is_err());
}

#[test]
fn workbook_bytes_changed_with_owner_and_relationships_unchanged_is_stale() {
    let package = package_with_chain(opaque_records());
    let before = package.calculation_chain().expect("before");
    let mut transaction = before.edit();
    transaction.remove().expect("remove");
    let commit = transaction.commit().expect("commit");

    let mut changed = package.clone().into_opc();
    let workbook_name = changed
        .main_document_part()
        .expect("workbook")
        .partname()
        .clone();
    let mut bytes = changed
        .get_part(&workbook_name)
        .expect("workbook")
        .blob()
        .to_vec();
    let payload_offset = {
        let raw_limits = crate::raw::Limits::new(bytes.len(), 0);
        let mut records = crate::raw::Records::try_with_limits(&bytes, raw_limits)
            .expect("workbook record stream");
        let record = loop {
            let record = records
                .next()
                .expect("workbook record")
                .expect("valid workbook record");
            if !record.payload().is_empty() {
                break record;
            }
        };
        let (_, header_bytes) = crate::raw::Header::parse(&bytes[record.offset()..], raw_limits)
            .expect("workbook record header");
        record.offset() + header_bytes
    };
    bytes[payload_offset] ^= 1;
    changed
        .get_part_mut(&workbook_name)
        .expect("workbook")
        .set_blob(bytes);
    let changed = Package::from_opc(changed).expect("changed package");
    assert!(changed.apply_calculation_chain(&commit).is_err());
}

#[test]
fn authored_relationship_metadata_exact_opc_attribute_limit_is_accepted() {
    let maximum = litchi_opc::ReadLimits::default().max_xml_attribute_bytes();
    let id = "x".repeat(maximum - "Id".len());
    let package = package_with_unrelated_root_relationship(id);
    assert!(package.calculation_chain().is_ok());
}

#[test]
fn huge_unrelated_authored_relationship_metadata_is_rejected_before_capture() {
    let maximum = litchi_opc::ReadLimits::default().max_xml_attribute_bytes();
    let id = "x".repeat(maximum - "Id".len() + 1);
    let package = package_with_unrelated_root_relationship(id);
    assert!(matches!(
        package.calculation_chain(),
        Err(Error::Opc(litchi_opc::OpcError::ReadLimit {
            resource: litchi_opc::ReadResource::XmlAttributeBytes,
            ..
        }))
    ));
}

#[test]
fn public_package_facade_preflights_oversized_stale_metadata_before_clone() {
    let package = package_with_chain(opaque_records());
    let before = package.calculation_chain().expect("before");
    let mut transaction = before.edit();
    transaction.remove().expect("remove");
    let commit = transaction.commit().expect("commit");

    let maximum = litchi_opc::ReadLimits::default().max_xml_attribute_bytes();
    let stale = package_with_unrelated_root_relationship("x".repeat(maximum - "Id".len() + 1));
    assert!(matches!(
        stale.apply_calculation_chain(&commit),
        Err(Error::Opc(litchi_opc::OpcError::ReadLimit {
            resource: litchi_opc::ReadResource::XmlAttributeBytes,
            ..
        }))
    ));
}

#[test]
fn checked_publication_keeps_signature_policy_after_source_preflight() {
    let package = package_with_chain(opaque_records());
    let before = package.calculation_chain().expect("before");
    let mut transaction = before.edit();
    transaction.remove().expect("remove");
    let commit = transaction.commit().expect("commit");
    let current = commit
        .patch()
        .check_source(package.opc_package())
        .expect("source preflight");

    let mut signed_candidate = signed_package(package.clone()).into_opc();
    let error = commit
        .patch()
        .apply_checked(&mut signed_candidate, current)
        .expect_err("changed signed candidate must require an explicit policy");
    assert!(matches!(
        error,
        Error::Opc(litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
    ));
}

#[test]
fn checked_publication_failure_does_not_publish_the_original_package() {
    let package = package_with_chain(opaque_records());
    let before = package.calculation_chain().expect("before");
    let mut transaction = before.edit();
    transaction.remove().expect("remove");
    let commit = transaction.commit().expect("commit");
    let current = commit
        .patch()
        .check_source(package.opc_package())
        .expect("source preflight");

    let mut candidate = package.clone().into_opc();
    let workbook_name = candidate
        .main_document_part()
        .expect("workbook")
        .partname()
        .clone();
    let mut workbook_bytes = candidate
        .get_part(&workbook_name)
        .expect("workbook")
        .blob()
        .to_vec();
    workbook_bytes[0] ^= 1;
    candidate
        .get_part_mut(&workbook_name)
        .expect("workbook")
        .set_blob(workbook_bytes);
    assert!(
        commit
            .patch()
            .apply_checked(&mut candidate, current)
            .is_err()
    );
    assert_eq!(package.calculation_chain().expect("original"), before);
}

#[test]
fn long_relationship_owner_and_many_short_edges_are_bounded_before_clone() {
    let mut package = package_with_chain(opaque_records()).into_opc();
    let owner_name = format!("/xl/{}.bin", "o".repeat(4_090));
    let owner_uri = PackURI::new(owner_name).expect("long owner URI");
    let mut owner = BlobPart::new(
        owner_uri.clone(),
        "application/octet-stream".to_string(),
        Vec::new(),
    );
    for index in 0..3_000 {
        owner
            .rels_mut()
            .try_add_relationship(
                "urn:litchi:short".to_string(),
                "https://example.test/short".to_string(),
                format!("rIdLongOwner{index}"),
                TargetMode::External,
            )
            .expect("short edge");
    }
    package.add_part(Box::new(owner));
    let package = Package::from_opc(package).expect("authored package");
    assert!(matches!(
        package.calculation_chain(),
        Err(Error::Opc(litchi_opc::OpcError::ReadLimit {
            resource: litchi_opc::ReadResource::TotalRelationshipXmlBytes,
            ..
        }))
    ));
}

fn signed_package(package: Package) -> Package {
    let mut package = package.into_opc();
    let origin = PackURI::new("/_xmlsignatures/origin.sigs").expect("signature origin URI");
    package.add_part(Box::new(BlobPart::new(
        origin,
        litchi_opc::constants::content_type::OPC_DIGITAL_SIGNATURE_ORIGIN.to_string(),
        Vec::new(),
    )));
    package.rels_mut().add_relationship(
        rt::DIGITAL_SIGNATURE_ORIGIN.to_string(),
        "_xmlsignatures/origin.sigs".to_string(),
        "rIdSignatureOrigin".to_string(),
        false,
    );
    Package::from_opc(package).expect("signed package")
}

#[test]
fn signed_source_allows_exact_noop_but_refuses_changed_publication() {
    let package = package_with_chain(opaque_records());
    let signed = signed_package(package.clone());
    let before = signed.calculation_chain().expect("signed before");
    let noop = before.edit().commit().expect("signed no-op");
    assert!(!noop.changed());
    assert!(signed.apply_calculation_chain(&noop).is_ok());

    let mut transaction = before.edit();
    transaction.remove().expect("signed remove");
    assert!(matches!(
        transaction.commit(),
        Err(Error::Opc(
            litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy
        ))
    ));
}

#[test]
fn owner_bytes_changed_with_workbook_bytes_unchanged_is_stale() {
    let package = package_with_chain(opaque_records());
    let before = package.calculation_chain().expect("before");
    let mut transaction = before.edit();
    transaction.remove().expect("remove");
    let commit = transaction.commit().expect("commit");

    let mut changed = package.clone().into_opc();
    let chain = PackURI::new("/xl/calcChain.bin").expect("chain URI");
    changed
        .get_part_mut(&chain)
        .expect("chain")
        .set_blob(vec![0x99, 0x88]);
    let changed = Package::from_opc(changed).expect("changed package");
    assert!(changed.apply_calculation_chain(&commit).is_err());
}

#[test]
fn relationship_source_changed_with_owner_and_workbook_bytes_unchanged_is_stale() {
    let package = package_with_chain(opaque_records());
    let before = package.calculation_chain().expect("before");
    let mut transaction = before.edit();
    transaction.remove().expect("remove");
    let commit = transaction.commit().expect("commit");

    let mut changed = package.clone().into_opc();
    changed.rels_mut().add_relationship(
        "urn:litchi:test-unrelated".to_string(),
        "https://example.test/unrelated".to_string(),
        "rIdUnrelated".to_string(),
        true,
    );
    let changed = Package::from_opc(changed).expect("changed package");
    assert!(changed.apply_calculation_chain(&commit).is_err());
}

#[test]
fn inverse_refuses_a_foreign_edge_to_the_removed_physical_target() {
    let package = package_with_chain(opaque_records());
    let before = package.calculation_chain().expect("before");
    let mut transaction = before.edit();
    transaction.remove().expect("remove");
    let commit = transaction.commit().expect("commit");
    let removed = package
        .apply_calculation_chain(&commit)
        .expect("remove publication");

    let mut changed = removed.into_opc();
    let target = PackURI::new("/xl/calcChain.bin").expect("target URI");
    changed.add_part(Box::new(BlobPart::new(
        target,
        "application/octet-stream".to_string(),
        opaque_records(),
    )));
    changed.rels_mut().add_relationship(
        "urn:litchi:foreign-chain-edge".to_string(),
        "xl/calcChain.bin".to_string(),
        "rIdForeignChain".to_string(),
        false,
    );
    let changed = Package::from_opc(changed).expect("changed package");
    assert!(
        changed
            .apply_calculation_chain_patch(&commit.patch().inverse())
            .is_err()
    );
}

#[test]
fn content_types_source_changed_with_owner_and_workbook_bytes_unchanged_is_stale() {
    let package = package_with_chain(opaque_records());
    let before = package.calculation_chain().expect("before");
    let mut transaction = before.edit();
    transaction.remove().expect("remove");
    let commit = transaction.commit().expect("commit");

    let mut changed = package.clone().into_opc();
    let stale_part = PackURI::new("/xl/stale-chain-neighbor.bin").expect("stale part URI");
    changed.add_part(Box::new(BlobPart::new(
        stale_part.clone(),
        "application/octet-stream".to_string(),
        Vec::new(),
    )));
    let current = changed.source_content_types().expect("content types");
    let replacement = current
        .without_parts(
            std::slice::from_ref(&stale_part),
            ReadLimits::DEFAULT.max_part_bytes,
        )
        .expect("replacement content types");
    changed
        .try_replace_content_types(current.bytes(), &replacement)
        .expect("replace content types");
    let changed = Package::from_opc(changed).expect("changed package");
    assert!(changed.apply_calculation_chain(&commit).is_err());
}

#[test]
fn finite_byte_graph_limits_and_explicit_record_diagnostics_are_enforced() {
    let bytes = opaque_records();
    let package = package_with_chain(bytes.clone());
    let mut large_bytes = Vec::new();
    let mut writer = Writer::new(&mut large_bytes);
    writer
        .write_record(Kind::new(0x123).expect("kind"), &[0_u8; 1024])
        .expect("large record");
    let large_package = package_with_chain(large_bytes.clone());
    let small_bytes = ReadLimits {
        max_part_bytes: large_bytes.len() - 1,
        ..ReadLimits::DEFAULT
    };
    assert!(matches!(
        large_package.calculation_chain_with_limits(small_bytes),
        Err(Error::LimitExceeded {
            resource: "Calculation Chain part bytes",
            ..
        })
    ));

    let small_records = ReadLimits {
        max_records: 2,
        ..ReadLimits::DEFAULT
    };
    let snapshot = package
        .calculation_chain_with_limits(small_records)
        .expect("record quota does not reject opaque admission");
    let before_diagnostic = snapshot.part().expect("part").bytes().to_vec();
    assert!(matches!(
        snapshot.part().expect("part").record_count(),
        Err(Error::LimitExceeded {
            resource: "Calculation Chain records",
            ..
        })
    ));
    assert_eq!(snapshot.part().expect("part").bytes(), before_diagnostic);

    let no_relationships = ReadLimits {
        max_relationships: 0,
        ..ReadLimits::DEFAULT
    };
    assert!(matches!(
        package.calculation_chain_with_limits(no_relationships),
        Err(Error::LimitExceeded {
            resource: "Calculation Chain relationships",
            ..
        })
    ));

    let no_parts = ReadLimits {
        max_graph_parts: 0,
        ..ReadLimits::DEFAULT
    };
    assert!(matches!(
        package.calculation_chain_with_limits(no_parts),
        Err(Error::LimitExceeded {
            resource: "Calculation Chain graph parts",
            ..
        })
    ));
}

#[test]
fn orphan_and_foreign_owner_edges_are_rejected_when_requested() {
    let mut orphan = Package::create().expect("base package").into_opc();
    orphan.add_part(Box::new(BlobPart::new(
        PackURI::new("/xl/orphanCalcChain.bin").expect("URI"),
        CONTENT_TYPE.to_string(),
        opaque_records(),
    )));
    let orphan = Package::from_opc(orphan).expect("package");
    assert!(orphan.calculation_chain().is_err());

    let mut foreign = Package::create().expect("base package").into_opc();
    let workbook_name = foreign
        .main_document_part()
        .expect("workbook")
        .partname()
        .clone();
    let chain_name = PackURI::new("/xl/foreignCalcChain.bin").expect("URI");
    foreign.add_part(Box::new(BlobPart::new(
        chain_name.clone(),
        CONTENT_TYPE.to_string(),
        opaque_records(),
    )));
    foreign.rels_mut().add_relationship(
        RELATIONSHIP_TYPE.to_string(),
        "xl/foreignCalcChain.bin".to_string(),
        "rIdRootCalcChain".to_string(),
        false,
    );
    foreign
        .get_part_mut(&workbook_name)
        .expect("workbook")
        .rels_mut()
        .add_relationship(
            RELATIONSHIP_TYPE.to_string(),
            "xl/foreignCalcChain.bin".to_string(),
            "rIdCalcChain".to_string(),
            false,
        );
    let foreign = Package::from_opc(foreign).expect("package");
    assert!(foreign.calculation_chain().is_err());
}

#[test]
fn malformed_or_non_biff_opaque_payload_is_admitted_and_diagnosed_on_request() {
    let bytes = vec![0x80];
    let package = package_with_chain(bytes.clone());
    let snapshot = package
        .calculation_chain()
        .expect("opaque payload admission");
    assert_eq!(snapshot.part().expect("part").bytes(), bytes.as_slice());
    let before_diagnostic = snapshot.part().expect("part").bytes().to_vec();
    assert!(matches!(
        snapshot.part().expect("part").record_count(),
        Err(Error::Wire(_))
    ));
    assert_eq!(snapshot.part().expect("part").bytes(), before_diagnostic);
}

#[test]
fn malformed_or_non_biff_opaque_payload_removes_and_restores_exactly() {
    let bytes = vec![0x80];
    let package = package_with_chain(bytes.clone());
    let before = package.calculation_chain().expect("before");
    let mut transaction = before.edit();
    transaction.remove().expect("remove");
    let commit = transaction.commit().expect("commit");
    let removed = package
        .apply_calculation_chain(&commit)
        .expect("remove publication");
    assert!(!removed.calculation_chain().expect("removed").is_present());
    let restored = removed
        .apply_calculation_chain_patch(&commit.patch().inverse())
        .expect("inverse publication");
    let after = restored.calculation_chain().expect("restored");
    assert_eq!(after.part().expect("part").bytes(), bytes.as_slice());
    assert!(after.part().expect("part").record_count().is_err());
}

#[test]
#[ignore = "requires the vendored native XLSB corpus and exercises producer-specific streams"]
fn native_calculation_chain_fixtures_are_opaque_inventory_cases() {
    let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../3rdparty");
    let fixtures = [
        "poi/test-data/spreadsheet/sample.xlsb",
        "poi/test-data/spreadsheet/testVarious.xlsb",
        "poi/test-data/spreadsheet/bug66682.xlsb",
        "poi/test-data/spreadsheet/62815.xlsb",
        "libreoffice-core/sc/qa/unit/data/xlsb/tdf94627.xlsb",
        "libreoffice-core/sc/qa/unit/data/xlsb/pivottable_error_item_filter.xlsb",
        "libreoffice-core/sc/qa/unit/data/xlsb/universal-content.xlsb",
        "libreoffice-core/sc/qa/unit/data/xlsb/shared_formula.xlsb",
    ];
    let mut found = 0usize;
    for fixture in fixtures {
        let path = root.join(fixture);
        if !path.is_file() {
            continue;
        }
        found += 1;
        let package = Package::open(&path).expect("native XLSB package");
        let snapshot = package.calculation_chain().expect("Calculation Chain");
        assert!(snapshot.is_present());
        assert!(!snapshot.part().expect("part").bytes().is_empty());
        // Native streams are inventory-only coverage. Their bytes are
        // accepted regardless of whether the optional diagnostic framing
        // parser recognizes them.
        let _ = snapshot.part().expect("part").record_count();
    }
    assert!(found > 0, "native fixture corpus is unavailable");
}
