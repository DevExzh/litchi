#![allow(
    clippy::expect_used,
    clippy::pedantic,
    clippy::unwrap_used,
    reason = "the focused integration fixtures fail fast while asserting the public signature policy"
)]

use litchi_opc::{BlobPart, OpcPackage, PackURI, Part, TargetMode};
use litchi_xlsb::Package;
use litchi_xlsb::raw::{Records, Writer, kind};
use litchi_xlsb::volatile_dependencies::{
    CachedValue, CellReference, DEFAULT_PART_NAME, Dependencies, DependencyKind, MainTopic, Topic,
    VolatileType,
};

fn dependencies(value: &'static str) -> Dependencies {
    Dependencies::new(vec![VolatileType {
        kind: DependencyKind::Rtd,
        mains: vec![MainTopic {
            first: value.to_owned(),
            topics: vec![Topic {
                subtopics: vec!["topic".to_owned()],
                value: CachedValue::Number(1.5),
                references: vec![CellReference::new(0, 0, 0).expect("valid reference")],
            }],
        }],
    }])
}

fn package_with_owner() -> Package {
    let package = Package::create().expect("workbook");
    let mut edit = package
        .edit_volatile_dependencies()
        .expect("volatile snapshot");
    edit.set_dependencies(dependencies("Prog.ID"))
        .expect("stage owner");
    let commit = edit.commit().expect("commit owner");
    package
        .apply_volatile_dependencies(&commit)
        .expect("publish owner")
}

fn authored_opc(source: OpcPackage) -> OpcPackage {
    let mut authored = OpcPackage::new();
    for relationship in source.rels().iter() {
        authored
            .rels_mut()
            .try_add_relationship(
                relationship.reltype().to_owned(),
                relationship.target_ref().to_owned(),
                relationship.r_id().to_owned(),
                relationship.target_mode(),
            )
            .expect("package relationship");
    }
    for part in source
        .try_iter_parts()
        .map(|part| part.expect("part payload decodes"))
    {
        let mut copy = BlobPart::new(
            part.partname().clone(),
            part.content_type().to_owned(),
            part.blob().to_vec(),
        );
        for relationship in part.rels().iter() {
            copy.rels_mut()
                .try_add_relationship(
                    relationship.reltype().to_owned(),
                    relationship.target_ref().to_owned(),
                    relationship.r_id().to_owned(),
                    relationship.target_mode(),
                )
                .expect("part relationship");
        }
        authored
            .try_add_part(Box::new(copy))
            .expect("authored part");
    }
    authored
}

fn add_signature_marker(package: &mut OpcPackage) {
    let origin = PackURI::new("/_xmlsignatures/origin.sigs").expect("origin URI");
    let signature = PackURI::new("/_xmlsignatures/sig1.xml").expect("signature URI");
    let mut origin_part = BlobPart::new(
        origin.clone(),
        litchi_opc::constants::content_type::OPC_DIGITAL_SIGNATURE_ORIGIN.to_owned(),
        Vec::new(),
    );
    origin_part
        .rels_mut()
        .try_add_relationship(
            "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/signature".to_owned(),
            signature.relative_ref(origin.base_uri()),
            "rIdSignature".to_owned(),
            TargetMode::Internal,
        )
        .expect("signature relationship");
    package.add_part(Box::new(origin_part));
    package.add_part(Box::new(BlobPart::new(
        signature,
        litchi_opc::constants::content_type::OPC_DIGITAL_SIGNATURE_XMLSIGNATURE.to_owned(),
        b"<Signature/>".to_vec(),
    )));
    package
        .rels_mut()
        .try_add_relationship(
            litchi_opc::constants::relationship_type::DIGITAL_SIGNATURE_ORIGIN.to_owned(),
            origin.as_str().to_owned(),
            "rIdSignatureOrigin".to_owned(),
            TargetMode::Internal,
        )
        .expect("signature origin relationship");
}

fn signed_package() -> Package {
    // Build the fixture as a newly authored OPC graph. This permits the
    // initial signed serialization while still making every later changed
    // publication pass through the explicit policy guard under test.
    let mut opc = authored_opc(package_with_owner().into_opc());
    add_signature_marker(&mut opc);
    Package::from_opc(opc).expect("signed workbook")
}

fn reserved_type_bits_package() -> Package {
    let package = package_with_owner();
    let owner = PackURI::new(DEFAULT_PART_NAME).expect("owner URI");
    let source = package
        .opc_package()
        .get_part(&owner)
        .expect("owner part")
        .blob();
    let mut records = Vec::new();
    for record in Records::new(source) {
        let record = record.expect("owner record");
        let payload = if record.kind() == kind::BEGIN_VOL_TYPE {
            let mut payload = record.payload().to_vec();
            let flags = u32::from_le_bytes(payload[..4].try_into().expect("type flags"));
            payload[..4].copy_from_slice(&(flags | 2).to_le_bytes());
            payload
        } else {
            record.payload().to_vec()
        };
        records.push((record.kind(), payload));
    }
    let mut replacement = Vec::new();
    let mut writer = Writer::new(&mut replacement);
    for (record_kind, payload) in records {
        writer
            .write_record(record_kind, &payload)
            .expect("replacement record");
    }
    let mut opc = package.into_opc();
    opc.get_part_mut(&owner)
        .expect("owner part")
        .set_blob(replacement);
    Package::from_opc(opc).expect("reserved-bit workbook")
}

fn case_mismatched_owner_package() -> Package {
    let package = package_with_owner();
    let mut opc = package.into_opc();
    let old_name = PackURI::new(DEFAULT_PART_NAME).expect("owner URI");
    let stored_name = PackURI::new("/xl/VolatileDependencies.BIN").expect("case variant URI");
    let workbook = opc
        .main_document_part()
        .expect("workbook part")
        .partname()
        .clone();
    let (content_type, payload, relationship_id, relationship_type) = {
        let owner = opc.get_part(&old_name).expect("owner part");
        let relationship = opc
            .get_part(&workbook)
            .expect("workbook part")
            .rels()
            .iter()
            .find(|relationship| {
                relationship.reltype() == litchi_xlsb::volatile_dependencies::RELATIONSHIP_TYPE
            })
            .expect("volatile relationship");
        (
            owner.content_type().to_owned(),
            owner.blob().to_vec(),
            relationship.r_id().to_owned(),
            relationship.reltype().to_owned(),
        )
    };
    assert!(opc.remove_part(&old_name));
    opc.get_part_mut(&workbook)
        .expect("workbook part")
        .rels_mut()
        .remove(&relationship_id);
    opc.validate_new_part_name(&stored_name)
        .expect("case variant owner URI is available after removal");
    opc.try_add_part(Box::new(BlobPart::new(stored_name, content_type, payload)))
        .expect("case variant owner");

    // The relationship target intentionally retains the old spelling. OPC
    // part identity is ASCII case-insensitive, while this lexical spelling
    // belongs to the source relationship token and must survive an inverse.
    let target = old_name.relative_ref(workbook.base_uri());
    opc.get_part_mut(&workbook)
        .expect("workbook part")
        .rels_mut()
        .add_relationship(relationship_type, target, relationship_id, false);
    Package::from_opc(opc).expect("case-equivalent relationship target")
}

#[test]
fn signed_exact_noop_preserves_signature_and_bytes() {
    let package = signed_package();
    assert!(package.opc_package().is_signed());
    let before = package.to_bytes().expect("signed bytes");
    let snapshot = package.volatile_dependencies().expect("volatile snapshot");
    let commit = snapshot.edit().commit().expect("no-op commit");
    assert!(!commit.changed());
    let applied = package
        .apply_volatile_dependencies(&commit)
        .expect("apply no-op");
    assert!(applied.opc_package().is_signed());
    assert_eq!(applied.to_bytes().expect("applied bytes"), before);
}

#[test]
fn signed_changed_edit_requires_explicit_policy() {
    let package = signed_package();
    let snapshot = package.volatile_dependencies().expect("volatile snapshot");
    let mut edit = snapshot.edit();
    edit.set_dependencies(dependencies("Changed.ID"))
        .expect("stage changed owner");
    let error = edit.commit().expect_err("signed change must refuse");
    assert!(matches!(
        error,
        litchi_xlsb::volatile_dependencies::Error::Opc(
            litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy
        )
    ));
}

#[test]
fn explicit_unsigning_requires_a_fresh_plan_and_allows_reversible_edits() {
    let signed = signed_package();
    let signed_noop = signed
        .edit_volatile_dependencies()
        .expect("signed snapshot")
        .commit()
        .expect("signed no-op");
    let unsigned = signed.without_signatures().expect("explicitly unsigned");
    assert!(!unsigned.opc_package().is_signed());
    let error = unsigned
        .apply_volatile_dependencies(&signed_noop)
        .expect_err("signed source plan is stale after unsigning");
    assert!(matches!(
        error,
        litchi_xlsb::volatile_dependencies::Error::InvalidFormat(_)
    ));

    let before = unsigned.to_bytes().expect("unsigned source bytes");
    let expected = dependencies("Changed.ID");
    let mut edit = unsigned
        .edit_volatile_dependencies()
        .expect("fresh unsigned snapshot");
    edit.set_dependencies(expected.clone()).expect("stage edit");
    let commit = edit.commit().expect("unsigned edit commit");
    let changed = unsigned
        .apply_volatile_dependencies(&commit)
        .expect("unsigned edit publication");
    let reopened = Package::from_bytes(changed.to_bytes().expect("changed bytes"))
        .expect("reopen changed package");
    assert!(!reopened.opc_package().is_signed());
    assert_eq!(
        reopened
            .volatile_dependencies()
            .expect("changed snapshot")
            .dependencies(),
        Some(&expected)
    );
    let restored = reopened
        .apply_volatile_dependencies_patch(&commit.patch().inverse())
        .expect("inverse after reopen");
    assert_eq!(restored.to_bytes().expect("restored bytes"), before);
}

#[test]
fn nonzero_begin_type_reserved_bits_are_opaque_and_refuse_rewrite() {
    let package = reserved_type_bits_package();
    let before = package.to_bytes().expect("source bytes");
    let snapshot = package
        .volatile_dependencies()
        .expect("reserved bits remain readable");
    assert!(!snapshot.can_edit());
    let noop = snapshot.edit().commit().expect("opaque no-op");
    let applied = package
        .apply_volatile_dependencies(&noop)
        .expect("apply opaque no-op");
    assert_eq!(applied.to_bytes().expect("applied bytes"), before);

    let mut edit = snapshot.edit();
    let error = edit
        .set_dependencies(dependencies("Changed.ID"))
        .expect_err("opaque source rewrite");
    assert!(matches!(
        error,
        litchi_xlsb::volatile_dependencies::Error::UnsupportedFeature(_)
    ));
}

#[test]
fn case_equivalent_target_is_owned_and_inverse_restores_lexical_source() {
    let package = case_mismatched_owner_package();
    let before = package.to_bytes().expect("source bytes");
    let snapshot = package
        .volatile_dependencies()
        .expect("case-equivalent owner snapshot");
    let mut remove = snapshot.edit();
    assert!(remove.remove().expect("stage owner removal"));
    let removal = remove.commit().expect("commit owner removal");
    let removed = package
        .apply_volatile_dependencies(&removal)
        .expect("remove case-equivalent owner");
    let restored = removed
        .apply_volatile_dependencies_patch(&removal.patch().inverse())
        .expect("inverse case-equivalent owner removal");
    assert_eq!(restored.to_bytes().expect("restored bytes"), before);
}
