#![allow(
    clippy::expect_used,
    clippy::pedantic,
    clippy::unwrap_used,
    reason = "integration fixtures deliberately fail fast while asserting public package contracts"
)]

//! Public, source-preserving Volatile Dependencies package contracts.
//!
//! The owner is described by [MS-XLSB] 2.1.7.60 and the hierarchy in
//! [MS-XLSB] 2.2.13.  These tests deliberately enter through `Package` and
//! `volatile_dependencies`; OPC is used only to prepare source-graph fixtures
//! that ordinary callers can receive from a saved workbook.

use litchi_opc::{BlobPart, PackURI, TargetMode};
use litchi_xlsb::Package;
use litchi_xlsb::raw::{Kind, Records, Writer};
use litchi_xlsb::volatile_dependencies::{
    CachedValue, CellReference, Commit, DEFAULT_PART_NAME, Dependencies, DependencyKind, MainTopic,
    RELATIONSHIP_TYPE, ReadLimits, Topic, VolatileType,
};

fn sample_dependencies() -> Dependencies {
    Dependencies::new(vec![
        VolatileType {
            kind: DependencyKind::Rtd,
            mains: vec![MainTopic {
                first: "Prog.ID".to_owned(),
                topics: vec![Topic {
                    subtopics: vec!["Topic".to_owned(), "Region".to_owned()],
                    value: CachedValue::Number(42.5),
                    references: vec![
                        CellReference::new(0, 1, 0).expect("valid first reference"),
                        CellReference::new(9, 2, 0).expect("valid second reference"),
                    ],
                }],
            }],
        },
        VolatileType {
            kind: DependencyKind::Cube,
            mains: vec![MainTopic {
                first: "SalesCube".to_owned(),
                topics: vec![Topic {
                    subtopics: vec!["[Measures].[Sales]".to_owned()],
                    value: CachedValue::String("ready".to_owned()),
                    references: vec![CellReference::new(4, 4, 0).expect("valid cube reference")],
                }],
            }],
        },
    ])
}

fn changed_dependencies() -> Dependencies {
    let mut dependencies = sample_dependencies();
    dependencies.types[0].mains[0].topics[0].value = CachedValue::Bool(true);
    dependencies.types[0].mains[0].topics[0].subtopics[1] = "Europe".to_owned();
    dependencies
}

fn package_with_dependencies(dependencies: Dependencies) -> Package {
    let package = Package::create().expect("create workbook");
    let mut transaction = package
        .edit_volatile_dependencies()
        .expect("start volatile dependency transaction");
    assert!(
        transaction
            .set_dependencies(dependencies)
            .expect("stage volatile dependencies")
    );
    let commit = transaction.commit().expect("commit volatile dependencies");
    package
        .apply_volatile_dependencies(&commit)
        .expect("publish volatile dependencies")
}

fn package_bytes(package: &Package) -> Vec<u8> {
    package.to_bytes().expect("serialize workbook")
}

fn owner_uri() -> PackURI {
    PackURI::new(DEFAULT_PART_NAME).expect("default owner URI")
}

fn owner_bytes(package: &Package) -> Vec<u8> {
    package
        .opc_package()
        .get_part(&owner_uri())
        .expect("volatile owner")
        .blob()
        .to_vec()
}

fn package_with_owner_bytes(package: &Package, bytes: Vec<u8>) -> Package {
    let mut opc = package.clone().into_opc();
    opc.get_part_mut(&owner_uri())
        .expect("volatile owner")
        .set_blob(bytes);
    Package::from_opc(opc).expect("reopen owner fixture")
}

fn rewrite_records(bytes: &[u8], insert_at: usize, kind: Kind, payload: &[u8]) -> Vec<u8> {
    let mut records = Vec::new();
    for record in Records::new(bytes) {
        let record = record.expect("valid owner record");
        records.push((record.kind(), record.payload().to_vec()));
    }
    records.insert(insert_at.min(records.len()), (kind, payload.to_vec()));
    let mut output = Vec::new();
    let mut writer = Writer::new(&mut output);
    for (record_kind, record_payload) in records {
        writer
            .write_record(record_kind, &record_payload)
            .expect("write owner record");
    }
    output
}

fn changed_workbook_relationship_source(package: &Package) -> Package {
    let mut opc = package.clone().into_opc();
    let workbook = opc
        .main_document_part()
        .expect("workbook part")
        .partname()
        .clone();
    let source = opc
        .source_relationships(&workbook)
        .expect("workbook relationship source");
    let relationship = opc
        .get_part(&workbook)
        .expect("workbook part")
        .rels()
        .iter()
        .next()
        .expect("workbook relationship");
    let id = relationship.r_id().to_owned();
    let reltype = relationship.reltype().to_owned();
    let target = relationship.target_ref().to_owned();
    let mode = relationship.target_mode();
    let mut replacement = source
        .without_relationship(&id, usize::MAX)
        .expect("remove relationship source token");
    replacement = replacement
        .with_relationship(&reltype, &target, &id, mode, usize::MAX)
        .expect("reinsert relationship source token");
    if replacement.bytes() == source.bytes() {
        replacement = replacement
            .with_relationship(
                "urn:litchi:test:stale-volatile",
                &target,
                "rIdStaleVolatile",
                TargetMode::Internal,
                usize::MAX,
            )
            .expect("add lexical stale relationship token");
    }
    assert_ne!(replacement.bytes(), source.bytes());
    opc.try_replace_relationships(&source, &replacement)
        .expect("replace workbook relationship source");
    Package::from_opc(opc).expect("reopen relationship stale fixture")
}

fn changed_content_types_source(package: &Package) -> Package {
    let original_bytes = package_bytes(package);
    let opc = package.clone().into_opc();
    let source = opc.source_content_types().expect("content-types source");
    let candidates = opc
        .iter_parts()
        .map(|part| (part.partname().clone(), part.content_type().to_owned()))
        .collect::<Vec<_>>();

    for (part_name, content_type) in candidates {
        let Ok(replacement) =
            source.with_part_overrides(&[(&part_name, content_type.as_str())], usize::MAX)
        else {
            continue;
        };
        if replacement.bytes() == source.bytes() {
            continue;
        }
        let mut candidate = opc.clone();
        if candidate
            .try_replace_content_types(source.bytes(), &replacement)
            .is_err()
        {
            continue;
        }
        if let Ok(package) = Package::from_opc(candidate) {
            assert_ne!(package_bytes(&package), original_bytes);
            return package;
        }
    }
    panic!("no content-types source-preserving lexical fixture was admitted")
}

fn relocate_owner(package: &Package) -> Package {
    let mut opc = package.clone().into_opc();
    let old_name = owner_uri();
    let workbook = opc
        .main_document_part()
        .expect("workbook part")
        .partname()
        .clone();
    let (content_type, payload, relationship_id) = {
        let owner = opc.get_part(&old_name).expect("default owner");
        let relationship = opc
            .get_part(&workbook)
            .expect("workbook part")
            .rels()
            .iter()
            .find(|relationship| relationship.reltype() == RELATIONSHIP_TYPE)
            .expect("volatile relationship");
        (
            owner.content_type().to_owned(),
            owner.blob().to_vec(),
            relationship.r_id().to_owned(),
        )
    };
    let custom_name =
        PackURI::new("/xl/custom/volatileDependencies.custom.bin").expect("custom owner URI");
    assert!(opc.remove_part(&old_name));
    opc.get_part_mut(&workbook)
        .expect("workbook part")
        .rels_mut()
        .remove(&relationship_id);
    opc.validate_new_part_name(&custom_name)
        .expect("custom owner name is available");
    opc.try_add_part(Box::new(BlobPart::new(
        custom_name.clone(),
        content_type,
        payload,
    )))
    .expect("add custom owner");
    let target = custom_name.relative_ref(workbook.base_uri());
    opc.get_part_mut(&workbook)
        .expect("workbook part")
        .rels_mut()
        .add_relationship(
            RELATIONSHIP_TYPE.to_owned(),
            target,
            "rIdVolatileCustom".to_owned(),
            false,
        );
    Package::from_opc(opc).expect("reopen custom owner fixture")
}

fn changed_commit(package: &Package) -> Commit {
    let snapshot = package
        .volatile_dependencies()
        .expect("read volatile dependencies");
    let mut transaction = snapshot.edit();
    transaction
        .set_dependencies(changed_dependencies())
        .expect("stage changed dependencies");
    transaction.commit().expect("commit changed dependencies")
}

#[test]
fn public_create_change_remove_save_reopen_and_exact_inverse() {
    let empty = Package::create().expect("create empty workbook");
    assert!(
        !empty
            .volatile_dependencies()
            .expect("read absent owner")
            .is_present()
    );

    let package = package_with_dependencies(sample_dependencies());
    let snapshot = package.volatile_dependencies().expect("read created owner");
    assert_eq!(snapshot.dependencies(), Some(&sample_dependencies()));

    let serialized = package_bytes(&package);
    let reopened = Package::from_bytes(serialized.clone()).expect("save/reopen package");
    assert_eq!(package_bytes(&reopened), serialized);
    assert_eq!(
        reopened
            .volatile_dependencies()
            .expect("read reopened owner")
            .dependencies(),
        Some(&sample_dependencies())
    );

    let changed = package
        .apply_volatile_dependencies(&changed_commit(&package))
        .expect("apply changed dependency owner");
    assert_eq!(
        changed
            .volatile_dependencies()
            .expect("read changed owner")
            .dependencies(),
        Some(&changed_dependencies())
    );
    let changed_commit = changed_commit(&package);
    let restored = changed
        .apply_volatile_dependencies_patch(&changed_commit.patch().inverse())
        .expect("inverse changed owner");
    assert_eq!(package_bytes(&restored), serialized);

    let mut remove = package
        .volatile_dependencies()
        .expect("read owner for removal")
        .edit();
    assert!(remove.remove().expect("stage owner removal"));
    let removal = remove.commit().expect("commit owner removal");
    let removed = package
        .apply_volatile_dependencies(&removal)
        .expect("apply owner removal");
    assert!(
        !removed
            .volatile_dependencies()
            .expect("read removed owner")
            .is_present()
    );
    let restored_empty = removed
        .apply_volatile_dependencies_patch(&removal.patch().inverse())
        .expect("inverse owner removal");
    assert_eq!(package_bytes(&restored_empty), serialized);
}

#[test]
fn exact_noop_is_source_preserving_and_stale_workbook_relationship_is_atomic() {
    let package = package_with_dependencies(sample_dependencies());
    let before = package_bytes(&package);
    let snapshot = package
        .volatile_dependencies()
        .expect("read owner for no-op");
    let noop = snapshot.edit().commit().expect("commit no-op");
    assert!(!noop.changed());
    assert!(noop.patch().is_empty());
    let no_op_package = package
        .apply_volatile_dependencies(&noop)
        .expect("apply no-op");
    assert_eq!(package_bytes(&no_op_package), before);

    let stale = changed_workbook_relationship_source(&package);
    let stale_before = package_bytes(&stale);
    assert!(stale.apply_volatile_dependencies(&noop).is_err());
    assert_eq!(package_bytes(&stale), stale_before);

    let changed = changed_commit(&package);
    assert!(stale.apply_volatile_dependencies(&changed).is_err());
    assert_eq!(package_bytes(&stale), stale_before);
}

#[test]
fn stale_content_types_source_rejects_noop_without_mutating_package() {
    let package = package_with_dependencies(sample_dependencies());
    let snapshot = package
        .volatile_dependencies()
        .expect("read owner for no-op");
    let noop = snapshot.edit().commit().expect("commit no-op");
    let stale = changed_content_types_source(&package);
    let before = package_bytes(&stale);
    assert!(stale.apply_volatile_dependencies(&noop).is_err());
    assert_eq!(package_bytes(&stale), before);
}

#[test]
fn custom_owner_name_and_relationship_id_survive_remove_inverse_and_reopen() {
    let package = relocate_owner(&package_with_dependencies(sample_dependencies()));
    let before = package_bytes(&package);
    let snapshot = package.volatile_dependencies().expect("read custom owner");
    assert!(snapshot.is_present());
    let mut transaction = snapshot.edit();
    assert!(transaction.remove().expect("stage custom owner removal"));
    let commit = transaction.commit().expect("commit custom owner removal");
    let removed = package
        .apply_volatile_dependencies(&commit)
        .expect("remove custom owner");
    assert!(
        !removed
            .volatile_dependencies()
            .expect("read removed custom owner")
            .is_present()
    );
    let restored = removed
        .apply_volatile_dependencies_patch(&commit.patch().inverse())
        .expect("restore custom owner");
    assert_eq!(package_bytes(&restored), before);
    let reopened = Package::from_bytes(package_bytes(&restored)).expect("reopen restored owner");
    assert_eq!(package_bytes(&reopened), before);
    assert_eq!(
        reopened
            .volatile_dependencies()
            .expect("read restored custom owner")
            .dependencies(),
        Some(&sample_dependencies())
    );
}

#[test]
fn unknown_and_reserved_record_kinds_are_exact_noops_but_changed_edits_refuse() {
    let package = package_with_dependencies(sample_dependencies());
    let owner = owner_bytes(&package);
    let positions = [1, usize::MAX];
    for (index, kind_value) in positions.into_iter().zip([0x03fe_u16, 0x03ff_u16]) {
        let kind = Kind::new(kind_value).expect("reserved record kind");
        let bytes = rewrite_records(&owner, index, kind, &[9, 8, 7]);
        let unsupported = package_with_owner_bytes(&package, bytes);
        let before = package_bytes(&unsupported);
        let snapshot = unsupported
            .volatile_dependencies()
            .expect("read unsupported owner");
        assert!(
            snapshot
                .dependencies()
                .expect("unsupported dependency model")
                .has_unsupported_records()
        );
        assert!(!snapshot.can_edit());

        let noop = snapshot.edit().commit().expect("commit unsupported no-op");
        assert!(!noop.changed());
        let no_op = unsupported
            .apply_volatile_dependencies(&noop)
            .expect("apply unsupported no-op");
        assert_eq!(package_bytes(&no_op), before);

        let mut changed = snapshot.edit();
        let result = changed.set_dependencies(changed_dependencies());
        assert!(matches!(
            result,
            Err(litchi_xlsb::volatile_dependencies::Error::UnsupportedFeature(_))
        ));
        assert_eq!(package_bytes(&unsupported), before);
    }
}

#[test]
fn caller_limits_refuse_before_publication_and_leave_source_unchanged() {
    let package = package_with_dependencies(sample_dependencies());
    let before = package_bytes(&package);

    let mut part_limit = ReadLimits::DEFAULT;
    part_limit.max_part_bytes = owner_bytes(&package).len().saturating_sub(1);
    let part_error = package
        .volatile_dependencies_with_limits(part_limit)
        .expect_err("part-byte limit must reject owner");
    assert!(part_error.to_string().contains("limit"));
    assert_eq!(package_bytes(&package), before);

    let mut record_limit = ReadLimits::DEFAULT;
    record_limit.max_records = 1;
    let record_error = package
        .volatile_dependencies_with_limits(record_limit)
        .expect_err("record limit must reject owner");
    assert!(record_error.to_string().contains("limit"));
    assert_eq!(package_bytes(&package), before);

    let mut relationship_limit = ReadLimits::DEFAULT;
    relationship_limit.max_relationships = 0;
    let relationship_error = package
        .volatile_dependencies_with_limits(relationship_limit)
        .expect_err("relationship limit must reject source closure");
    assert!(relationship_error.to_string().contains("limit"));
    assert_eq!(package_bytes(&package), before);
}
