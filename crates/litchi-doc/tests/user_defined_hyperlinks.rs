//! Public DOC transaction coverage for inert `_PID_HLINKS` metadata.

use litchi_doc::{
    FieldType, FieldsTable, HyperlinkAssociation, Package, UserDefinedHyperlinks,
    package::Snapshot,
    user_defined_hyperlinks::{Hyperlink, Hyperlinks, Limits},
};
use litchi_ole_common::property_set::user_defined::{Edit, LinkBase, Properties};
use litchi_ole_common::property_set::{CodePage, Value};
use std::io::Cursor;
use std::path::PathBuf;

fn fixture() -> Snapshot {
    let root = PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../..");
    let bytes = std::fs::read(root.join("test-data/ole/doc/documentProperties.doc")).unwrap();
    Snapshot::from_bytes(bytes).unwrap()
}

fn seeded() -> Snapshot {
    let source = fixture();
    let mut transaction = source.transaction().unwrap();
    assert!(
        transaction
            .edit_user_defined_properties(|section| {
                let mut edit = Edit::new(section)?;
                edit.set_link_base(LinkBase::new("https://base.example/")?)?;
                edit.set_hyperlinks(Hyperlinks::new(vec![
                    Hyperlink::new(-1, 41, 0, "shape-target", "shape-location")?,
                    Hyperlink::new(-2, 0, 0, "opaque-target", "opaque-location")?,
                    Hyperlink::new(7, 0, 0, "stored-target", "stored-location")?,
                ]))?;
                Ok(())
            })
            .unwrap()
    );
    transaction.commit().unwrap().into_parts().0
}

fn hyperlink_fixture() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../test-data/ole/doc/hyperlink.doc")
}

fn snapshot_with_real_hyperlink_field() -> (Snapshot, FieldsTable) {
    let path = hyperlink_fixture();
    let bytes = std::fs::read(&path).unwrap();
    let mut package = Package::open(path).unwrap();
    let fields = package
        .document()
        .unwrap()
        .fields_table()
        .expect("fixture must retain the HYPERLINK field PLCF")
        .clone();
    let hyperlink = fields
        .stories()
        .iter()
        .flat_map(|story| story.fields())
        .find(|field| field.field_type == FieldType::Hyperlink)
        .expect("fixture must contain a HYPERLINK field")
        .clone();
    let index = fields
        .story(hyperlink.story)
        .unwrap()
        .markers()
        .iter()
        .position(|marker| marker.position == hyperlink.start_cp)
        .unwrap();

    let source = Snapshot::from_bytes(bytes).unwrap();
    let mut transaction = source.transaction().unwrap();
    assert!(
        transaction
            .edit_user_defined_properties(|section| {
                Edit::new(section)?.set_hyperlinks(Hyperlinks::new(vec![Hyperlink::new(
                    i32::try_from(index).unwrap(),
                    0,
                    0,
                    "https://example.invalid/target",
                    "stored-location",
                )?]))
            })
            .unwrap()
    );
    (transaction.commit().unwrap().into_parts().0, fields)
}

fn resolve_all(hyperlinks: &mut UserDefinedHyperlinks) {
    for hyperlink in hyperlinks.entries_mut() {
        let (story, index) = match hyperlink.association() {
            HyperlinkAssociation::FieldCandidates(candidates) => {
                (candidates[0].story, candidates[0].plcfld_index)
            },
            other => panic!("expected field candidates, got {other:?}"),
        };
        hyperlink.resolve_field(story, index).unwrap();
    }
}

#[test]
fn immutable_authoring_collection_recontexts_replaces_and_round_trips() {
    let (snapshot, fields) = snapshot_with_real_hyperlink_field();
    let index = match snapshot
        .user_defined_hyperlinks(Some(&fields))
        .unwrap()
        .unwrap()
        .entries()[0]
        .association()
    {
        HyperlinkAssociation::FieldCandidates(candidates) => candidates[0].plcfld_index,
        other => panic!("expected field candidates, got {other:?}"),
    };
    let first = Hyperlink::new(index as i32, 0, 0, "first-target", "first-location").unwrap();
    let second = Hyperlink::new(index as i32, 0, 0, "second-target", "second-location").unwrap();
    let mut hyperlinks =
        UserDefinedHyperlinks::from_hyperlinks(Hyperlinks::new(vec![first, second]), Some(&fields));
    let count = hyperlinks.entries().len();
    assert!(
        hyperlinks
            .replace_hyperlink(
                count,
                Hyperlink::new(index as i32, 0, 0, "ignored", "ignored").unwrap(),
                Some(&fields),
            )
            .is_none()
    );
    assert!(hyperlinks.remove_hyperlink(count).is_none());
    assert_eq!(hyperlinks.entries().len(), count);
    resolve_all(&mut hyperlinks);

    let mut transaction = snapshot.transaction().unwrap();
    assert!(
        transaction
            .put_user_defined_hyperlinks(&hyperlinks)
            .unwrap()
    );
    let committed = transaction.commit().unwrap().into_parts().0;
    let mut readback = committed
        .user_defined_hyperlinks(Some(&fields))
        .unwrap()
        .unwrap();
    assert_eq!(
        readback
            .entries()
            .iter()
            .map(|entry| (entry.target(), entry.location()))
            .collect::<Vec<_>>(),
        vec![
            ("first-target", "first-location"),
            ("second-target", "second-location")
        ]
    );

    assert!(
        readback
            .replace_hyperlink(
                0,
                Hyperlink::new(
                    index as i32,
                    0,
                    0,
                    "replacement-target",
                    "replacement-location"
                )
                .unwrap(),
                Some(&fields),
            )
            .is_some()
    );
    resolve_all(&mut readback);
    let mut transaction = committed.transaction().unwrap();
    assert!(transaction.put_user_defined_hyperlinks(&readback).unwrap());
    let committed = transaction.commit().unwrap().into_parts().0;
    assert_eq!(
        committed
            .user_defined_hyperlinks(Some(&fields))
            .unwrap()
            .unwrap()
            .entries()[0]
            .target(),
        "replacement-target"
    );
}

#[test]
fn contextualized_reads_use_real_plcf_fields_and_explicit_limits() {
    let (snapshot, fields) = snapshot_with_real_hyperlink_field();
    let limits = Limits::builder().max_links(1).build().unwrap();

    let default = snapshot
        .user_defined_hyperlinks(Some(&fields))
        .unwrap()
        .unwrap();
    let explicit = snapshot
        .user_defined_hyperlinks_with_limits(Some(&fields), limits)
        .unwrap()
        .unwrap();
    assert_eq!(default.entries().len(), 1);
    assert_eq!(explicit.entries().len(), 1);

    let entry = &explicit.entries()[0];
    assert_eq!(entry.target(), "https://example.invalid/target");
    assert_eq!(entry.location(), "stored-location");
    assert!(matches!(
        entry.association(),
        HyperlinkAssociation::FieldCandidates(candidates)
            if candidates.iter().any(|candidate| candidate.field.field_type == FieldType::Hyperlink)
    ));

    let mut package = Package::open(hyperlink_fixture()).unwrap();
    assert!(package.user_defined_hyperlinks().unwrap().is_none());
    let mut package = Package::open(hyperlink_fixture()).unwrap();
    assert!(
        package
            .user_defined_hyperlinks_with_limits(limits)
            .unwrap()
            .is_none()
    );
}

#[test]
fn unresolved_candidates_refuse_changed_mutation_without_staging_or_commit_side_effects() {
    let (snapshot, fields) = snapshot_with_real_hyperlink_field();
    let unresolved = snapshot
        .user_defined_hyperlinks(Some(&fields))
        .unwrap()
        .unwrap();
    assert!(matches!(
        unresolved.entries()[0].association(),
        HyperlinkAssociation::FieldCandidates(_)
    ));

    let mut transaction = snapshot.transaction().unwrap();
    assert!(
        !transaction
            .put_user_defined_hyperlinks(&unresolved)
            .unwrap()
    );
    assert!(!transaction.is_changed());
    let unchanged = transaction.commit().unwrap().into_parts().0;

    let mut remove = unchanged.transaction().unwrap();
    assert!(remove.remove_user_defined_properties().unwrap());
    let without_property = remove.commit().unwrap().into_parts().0;
    let mut transaction = without_property.transaction().unwrap();
    assert!(
        transaction
            .put_user_defined_hyperlinks_with_limits(&unresolved, Limits::default())
            .is_err()
    );
    assert!(!transaction.is_changed());
    assert_eq!(transaction.rollback().bytes(), without_property.bytes());

    let mut transaction = without_property.transaction().unwrap();
    assert!(
        transaction
            .put_user_defined_hyperlinks(&unresolved)
            .is_err()
    );
    assert!(!transaction.is_changed());
    let commit = transaction.commit().unwrap();
    assert!(!commit.changed());
    assert_eq!(commit.snapshot().bytes(), without_property.bytes());
}

#[test]
fn overlay_limits_reject_reads_and_resolved_puts_without_mutation() {
    let (snapshot, fields) = snapshot_with_real_hyperlink_field();
    let too_small = Limits::builder().max_string_units(1).build().unwrap();

    assert!(
        snapshot
            .user_defined_hyperlinks_with_limits(Some(&fields), too_small)
            .is_err()
    );
    assert_eq!(snapshot.bytes(), snapshot.clone().into_bytes());

    let mut package = Package::from_reader(Cursor::new(snapshot.bytes().to_vec())).unwrap();
    assert!(
        package
            .user_defined_hyperlinks_with_limits(too_small)
            .is_err()
    );

    let mut hyperlinks = snapshot
        .user_defined_hyperlinks(Some(&fields))
        .unwrap()
        .unwrap();
    let (story, index) = match hyperlinks.entries()[0].association() {
        HyperlinkAssociation::FieldCandidates(candidates) => {
            (candidates[0].story, candidates[0].plcfld_index)
        },
        other => panic!("expected field candidates, got {other:?}"),
    };
    hyperlinks.entries_mut()[0]
        .resolve_field(story, index)
        .unwrap();

    let mut transaction = snapshot.transaction().unwrap();
    assert!(
        transaction
            .put_user_defined_hyperlinks_with_limits(&hyperlinks, too_small)
            .is_err()
    );
    assert!(!transaction.is_changed());
    assert_eq!(transaction.rollback().bytes(), snapshot.bytes());
}

#[test]
fn resolved_metadata_creates_a_codepage_for_a_missing_user_defined_section() {
    let (contextualized, fields) = snapshot_with_real_hyperlink_field();
    let mut hyperlinks = contextualized
        .user_defined_hyperlinks(Some(&fields))
        .unwrap()
        .unwrap();
    let (story, index) = match hyperlinks.entries()[0].association() {
        HyperlinkAssociation::FieldCandidates(candidates) => {
            (candidates[0].story, candidates[0].plcfld_index)
        },
        other => panic!("expected field candidates, got {other:?}"),
    };
    hyperlinks.entries_mut()[0]
        .resolve_field(story, index)
        .unwrap();

    let mut transaction = contextualized.transaction().unwrap();
    assert!(transaction.remove_user_defined_properties().unwrap());
    let source = transaction.commit().unwrap().into_parts().0;
    assert!(source.user_defined_properties().unwrap().is_none());

    let mut transaction = source.transaction().unwrap();
    assert!(
        transaction
            .put_user_defined_hyperlinks_with_limits(&hyperlinks, Limits::default())
            .unwrap()
    );
    let committed = transaction.commit().unwrap().into_parts().0;
    let section = committed.user_defined_properties().unwrap().unwrap();
    assert_eq!(section.page(), Some(CodePage::WINDOWS_1252));
    assert_eq!(
        committed
            .user_defined_hyperlinks_with_limits(Some(&fields), Limits::default())
            .unwrap()
            .unwrap()
            .entries()
            .len(),
        1
    );
}

#[test]
fn stored_hash_mismatch_is_inert_and_an_exact_noop_preserves_it() {
    let (snapshot, fields) = snapshot_with_real_hyperlink_field();
    let mut transaction = snapshot.transaction().unwrap();
    assert!(
        transaction
            .edit_user_defined_properties(|section| {
                let (identifier, Value::Blob(value)) = section.find_named("_PID_HLINKS").unwrap()
                else {
                    panic!("fixture must contain the typed hyperlink property");
                };
                let mut value = value.clone();
                // `cbData` + `cElements` + VT_I4 tag/padding precede dwHash.
                value[12] = value[12].wrapping_add(1);
                section.update(identifier, Value::Blob(value))?;
                Ok(())
            })
            .unwrap()
    );
    let corrupted = transaction.commit().unwrap().into_parts().0;

    let hyperlinks = corrupted
        .user_defined_hyperlinks(Some(&fields))
        .unwrap()
        .unwrap();
    let entry = &hyperlinks.entries()[0];
    assert!(!entry.hash_matches());
    assert_ne!(entry.stored_hash(), entry.calculated_hash());

    let mut transaction = corrupted.transaction().unwrap();
    assert!(
        !transaction
            .put_user_defined_hyperlinks(&hyperlinks)
            .unwrap()
    );
    assert!(!transaction.is_changed());
    let commit = transaction.commit().unwrap();
    assert!(!commit.changed());
    assert_eq!(commit.snapshot().bytes(), corrupted.bytes());
}

#[test]
fn inert_hyperlinks_keep_source_order_and_no_op_commits_exact_bytes() {
    let snapshot = seeded();
    let hyperlinks = snapshot.user_defined_hyperlinks(None).unwrap().unwrap();
    let entries = hyperlinks.entries();

    assert_eq!(
        entries
            .iter()
            .map(|entry| entry.target())
            .collect::<Vec<_>>(),
        ["shape-target", "opaque-target", "stored-target"]
    );
    assert!(matches!(
        entries[0].association(),
        HyperlinkAssociation::OfficeArtShape
    ));
    assert!(matches!(
        entries[1].association(),
        HyperlinkAssociation::UnassociatedApplicationData
    ));
    assert!(matches!(
        entries[2].association(),
        HyperlinkAssociation::UnassociatedApplicationData
    ));

    let mut transaction = snapshot.transaction().unwrap();
    assert!(
        !transaction
            .put_user_defined_hyperlinks(&hyperlinks)
            .unwrap()
    );
    assert!(!transaction.is_changed());
    let commit = transaction.commit().unwrap();
    assert!(!commit.changed());
    assert_eq!(commit.snapshot().bytes(), snapshot.bytes());
}

#[test]
fn removing_hyperlinks_is_targeted_and_rollback_recovers_the_exact_source() {
    let snapshot = seeded();

    let mut rollback_transaction = snapshot.transaction().unwrap();
    assert!(
        rollback_transaction
            .remove_user_defined_hyperlinks()
            .unwrap()
    );
    assert!(rollback_transaction.is_changed());
    let rolled_back = rollback_transaction.rollback();
    assert_eq!(rolled_back.bytes(), snapshot.bytes());
    assert!(rolled_back.user_defined_hyperlinks(None).unwrap().is_some());

    let mut transaction = snapshot.transaction().unwrap();
    assert!(transaction.remove_user_defined_hyperlinks().unwrap());
    let (committed, _) = transaction.commit().unwrap().into_parts();
    assert!(committed.user_defined_hyperlinks(None).unwrap().is_none());

    let section = committed.user_defined_properties().unwrap().unwrap();
    assert_eq!(
        Properties::new(&section)
            .unwrap()
            .link_base()
            .unwrap()
            .unwrap()
            .value(),
        "https://base.example/"
    );

    let mut no_op = committed.transaction().unwrap();
    assert!(!no_op.remove_user_defined_hyperlinks().unwrap());
    assert!(!no_op.is_changed());
}
