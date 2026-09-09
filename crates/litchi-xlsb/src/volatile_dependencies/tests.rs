#![allow(
    clippy::expect_used,
    reason = "synthetic package fixtures use panic-on-failure extraction for asserted valid input"
)]

use super::*;
use crate::package::error::Error;
use crate::raw::{Writer, kind};
use litchi_opc::{BlobPart, PackURI};

fn sample_dependencies() -> Dependencies {
    Dependencies::new(vec![VolatileType {
        kind: DependencyKind::Rtd,
        mains: vec![MainTopic {
            first: "Prog.ID".to_string(),
            topics: vec![Topic {
                subtopics: vec!["Topic".to_string()],
                value: CachedValue::Number(42.5),
                references: vec![CellReference::new(0, 1, 0).expect("cell")],
            }],
        }],
    }])
}

fn package_with_payload(payload: Vec<u8>) -> crate::Package {
    let mut package = crate::Package::create()
        .expect("generated package")
        .into_opc();
    let workbook = package
        .main_document_part()
        .expect("workbook")
        .partname()
        .clone();
    let owner = PackURI::new(DEFAULT_PART_NAME).expect("owner URI");
    package
        .try_add_part(Box::new(BlobPart::new(
            owner.clone(),
            CONTENT_TYPE.to_string(),
            payload,
        )))
        .expect("owner part");
    let target = owner.relative_ref(workbook.base_uri());
    package
        .get_part_mut(&workbook)
        .expect("workbook part")
        .rels_mut()
        .get_or_add(RELATIONSHIP_TYPE, &target);
    crate::Package::from_opc(package).expect("valid owner package")
}

#[test]
fn typed_stream_round_trips_without_execution() {
    let dependencies = sample_dependencies();
    let bytes = codec::write(&dependencies, ReadLimits::DEFAULT, 1).expect("write");
    let parsed = codec::read(&bytes, ReadLimits::DEFAULT, 1).expect("read");
    assert_eq!(parsed.dependencies, dependencies);
    assert!(parsed.rewrite_safe);
}

#[test]
fn reserved_bits_are_opaque_and_scalar_domains_are_rejected() {
    let mut bytes = Vec::new();
    let mut writer = Writer::new(&mut bytes);
    writer
        .write_record(kind::BEGIN_VOL_DEPS, &[])
        .expect("begin");
    writer
        .write_record(kind::BEGIN_VOL_TYPE, &2u32.to_le_bytes())
        .expect("reserved");
    writer
        .write_record(kind::END_VOL_TYPE, &[])
        .expect("end type");
    writer.write_record(kind::END_VOL_DEPS, &[]).expect("end");
    let parsed = codec::read(&bytes, ReadLimits::DEFAULT, 1).expect("reserved bits ignored");
    assert!(!parsed.rewrite_safe);
    assert!(parsed.dependencies.has_unsupported_records());

    let mut invalid = sample_dependencies();
    if let CachedValue::Number(value) = &mut invalid.types[0].mains[0].topics[0].value {
        *value = f64::NAN;
    }
    assert!(codec::write(&invalid, ReadLimits::DEFAULT, 1).is_err());
}

#[test]
fn sheet_ordinals_bind_to_workbook_bundle_sheet_count() {
    let mut dependencies = sample_dependencies();
    dependencies.types[0].mains[0].topics[0].references[0] =
        CellReference::new(0, 1, 1).expect("reference");
    assert!(codec::write(&dependencies, ReadLimits::DEFAULT, 1).is_err());

    let bytes =
        codec::write(&sample_dependencies(), ReadLimits::DEFAULT, 1).expect("valid payload");
    assert!(codec::read(&bytes, ReadLimits::DEFAULT, 0).is_err());
}

#[test]
fn absent_owner_can_be_created_and_removed_reversibly() {
    let package = crate::Package::create().expect("package");
    let before_bytes = package.to_bytes().expect("before bytes");
    assert!(!package.volatile_dependencies().expect("read").is_present());

    let mut transaction = package.edit_volatile_dependencies().expect("transaction");
    assert!(
        transaction
            .set_dependencies(sample_dependencies())
            .expect("stage")
    );
    let commit = transaction.commit().expect("commit");
    let package = package
        .apply_volatile_dependencies(&commit)
        .expect("publish");
    assert_eq!(
        package
            .volatile_dependencies()
            .expect("read published")
            .dependencies(),
        Some(&sample_dependencies())
    );

    let inverse = commit.patch().inverse();
    let restored = package
        .apply_volatile_dependencies_patch(&inverse)
        .expect("inverse");
    assert_eq!(restored.to_bytes().expect("restored bytes"), before_bytes);
    assert!(
        !restored
            .volatile_dependencies()
            .expect("read restored")
            .is_present()
    );
}

#[test]
fn foreign_feature_relationship_is_not_treated_as_absence() {
    let mut opc = crate::Package::create().expect("package").into_opc();
    opc.rels_mut()
        .get_or_add(RELATIONSHIP_TYPE, "/xl/foreign-volatile.bin");
    let package = crate::Package::from_opc(opc).expect("generic package validation");
    assert!(package.volatile_dependencies().is_err());
}

#[test]
fn exact_noop_shares_source_and_stale_owner_is_rejected() {
    let package = package_with_payload(
        codec::write(&sample_dependencies(), ReadLimits::DEFAULT, 1).expect("payload"),
    );
    let before_bytes = package.to_bytes().expect("before bytes");
    let snapshot = package.volatile_dependencies().expect("snapshot");
    let commit = snapshot.edit().commit().expect("noop commit");
    assert!(!commit.changed());
    let no_op = package
        .apply_volatile_dependencies(&commit)
        .expect("noop apply");
    assert_eq!(no_op.to_bytes().expect("noop bytes"), before_bytes);

    let mut stale_opc = package.clone().into_opc();
    let owner = PackURI::new(DEFAULT_PART_NAME).expect("owner URI");
    stale_opc.get_part_mut(&owner).expect("owner").set_blob({
        let mut changed = sample_dependencies();
        changed.types[0].mains[0].first = "Stale".to_string();
        codec::write(&changed, ReadLimits::DEFAULT, 1).expect("changed payload")
    });
    let stale = crate::Package::from_opc(stale_opc).expect("stale package");
    let mut changed = snapshot.edit();
    let mut replacement = sample_dependencies();
    replacement.types[0].mains[0].first = "Replacement".to_string();
    changed.set_dependencies(replacement).expect("stage change");
    let changed = changed.commit().expect("changed commit");
    assert!(stale.apply_volatile_dependencies(&changed).is_err());
}

#[test]
fn unsupported_records_are_read_but_typed_rewrite_is_refused() {
    let mut bytes = codec::write(&sample_dependencies(), ReadLimits::DEFAULT, 1).expect("base");
    let mut records = Vec::new();
    for record in crate::raw::Records::new(&bytes) {
        let record = record.expect("record");
        records.push((record.kind(), record.payload().to_vec()));
    }
    records.insert(
        1,
        (crate::raw::Kind::new(0x3ff).expect("kind"), vec![9, 8, 7]),
    );
    bytes.clear();
    let mut writer = Writer::new(&mut bytes);
    for (record_kind, payload) in records {
        writer.write_record(record_kind, &payload).expect("record");
    }
    let package = package_with_payload(bytes);
    let snapshot = package.volatile_dependencies().expect("snapshot");
    assert!(!snapshot.can_edit());
    let mut edit = snapshot.edit();
    let error = edit
        .set_dependencies(sample_dependencies())
        .expect_err("refusal");
    assert!(matches!(error, Error::UnsupportedFeature(_)));
}
