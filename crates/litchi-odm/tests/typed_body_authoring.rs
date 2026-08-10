#![allow(
    clippy::unwrap_used,
    reason = "fixed corpus assertions use unwrap for clarity"
)]

use litchi_core::{BlobId, HistoryLimits, Position};
use litchi_odf_common::compact_xml;
use litchi_odm::{
    Builder, Master,
    structure::{IndexKind, Kind},
    transaction::{BodyItemChange, BodyItemSpec, Conflict},
};

const WRITER_ODM: &[u8] =
    include_bytes!("../../../3rdparty/libreoffice-core/sw/qa/extras/odfexport/data/tdf121119.odm");

fn authored_items() -> Vec<BodyItemSpec> {
    vec![
        BodyItemSpec::paragraph("Typed paragraph").unwrap(),
        BodyItemSpec::heading(2, "Typed heading").unwrap(),
        BodyItemSpec::list(vec!["Alpha".to_string(), "Beta".to_string()]).unwrap(),
        BodyItemSpec::table(
            "TypedTable",
            vec![
                vec!["A".to_string(), "B".to_string()],
                vec!["1".to_string(), "2".to_string()],
            ],
        )
        .unwrap(),
        BodyItemSpec::generated_index(IndexKind::Alphabetical, "Typed Index").unwrap(),
    ]
}

#[test]
fn genuine_writer_master_accepts_typed_body_authoring_with_full_patch_lifecycle() {
    let master = Master::from_bytes(WRITER_ODM.to_vec()).unwrap();
    let original_count = master.structure().items().len();
    let mut edit = master.edit();
    for item in authored_items() {
        edit.add_body_item(item).unwrap();
    }
    let commit = edit.commit().unwrap();
    let changed = commit.snapshot();
    assert_eq!(master.as_bytes(), WRITER_ODM);
    assert_eq!(changed.structure().items().len(), original_count + 5);
    assert_eq!(
        &changed.structure().items()[original_count..],
        &[
            Kind::Paragraph,
            Kind::Heading,
            Kind::List,
            Kind::Table,
            Kind::GeneratedIndex(IndexKind::Alphabetical),
        ]
    );
    assert_eq!(
        changed
            .structure()
            .generated_indexes()
            .last()
            .unwrap()
            .name(),
        Some("Typed Index")
    );
    compact_xml::validate(changed.content_xml().as_bytes()).unwrap();
    assert!(Master::from_bytes(changed.as_bytes().to_vec()).is_ok());
    assert_eq!(
        commit.patch().inverse().apply(changed).unwrap().as_bytes(),
        WRITER_ODM
    );
    assert_eq!(
        commit
            .patch()
            .durable()
            .unwrap()
            .apply(&master)
            .unwrap()
            .as_bytes(),
        changed.as_bytes()
    );

    let mut history = master.history(HistoryLimits::new(2, u64::MAX));
    history.record(&commit).unwrap();
    assert!(history.undo());
    assert_eq!(history.current().as_bytes(), WRITER_ODM);
    assert!(history.redo());
    assert_eq!(history.current().as_bytes(), changed.as_bytes());

    let mut title_edit = master.edit();
    title_edit.set_title("Typed body merge").unwrap();
    let title_commit = title_edit.commit().unwrap();
    let merged = commit
        .patch()
        .merge(title_commit.patch())
        .unwrap()
        .apply(&master)
        .unwrap();
    assert_eq!(merged.title(), Some("Typed body merge"));
    assert_eq!(merged.structure().items().len(), original_count + 5);
}

#[test]
fn body_transfer_retains_source_provenance_and_refuses_unclosed_dependencies() {
    let writer = Master::from_bytes(WRITER_ODM.to_vec()).unwrap();
    let original_count = writer.structure().items().len();
    let mut add = writer.edit();
    add.add_body_item(BodyItemSpec::paragraph("Portable body item").unwrap())
        .unwrap();
    let source = add.commit().unwrap().into_snapshot();
    let source_before = source.as_bytes().to_vec();
    let destination = Master::from_bytes(
        Builder::new()
            .body_item(BodyItemSpec::paragraph("Destination").unwrap())
            .build()
            .unwrap(),
    )
    .unwrap();

    let mut transfer_edit = destination.edit();
    transfer_edit
        .transfer_body_item(&source, Position::new(original_count))
        .unwrap();
    let transfer_commit = transfer_edit.commit().unwrap();
    let provenance = transfer_commit
        .patch()
        .changes()
        .body_items()
        .iter()
        .find_map(|change| match change {
            BodyItemChange::Add(spec) => spec.provenance(),
            BodyItemChange::Remove { .. } | _ => None,
        })
        .unwrap()
        .clone();
    assert_eq!(source.as_bytes(), source_before);
    assert_eq!(transfer_commit.snapshot().structure().items().len(), 2);
    assert!(Master::from_bytes(transfer_commit.snapshot().as_bytes().to_vec()).is_ok());
    assert_eq!(
        transfer_commit
            .patch()
            .inverse()
            .apply(transfer_commit.snapshot())
            .unwrap()
            .as_bytes(),
        destination.as_bytes()
    );
    assert_eq!(provenance.item(), Position::new(original_count));
    assert_eq!(provenance.kind(), Kind::Paragraph);
    assert_eq!(
        provenance.package_sha256(),
        BlobId::of(source.as_bytes()).as_hex()
    );

    let styled_writer_paragraph = writer
        .structure()
        .items()
        .iter()
        .position(|kind| *kind == Kind::Paragraph)
        .map(Position::new)
        .unwrap();
    let mut refused = destination.edit();
    assert!(
        refused
            .transfer_body_item(&writer, styled_writer_paragraph)
            .is_err()
    );
}

#[test]
fn concurrent_generated_index_identity_creation_is_typed_conflict() {
    let master = Master::from_bytes(WRITER_ODM.to_vec()).unwrap();
    let mut left_edit = master.edit();
    left_edit
        .add_body_item(
            BodyItemSpec::generated_index(IndexKind::Alphabetical, "Shared Index").unwrap(),
        )
        .unwrap();
    let left_commit = left_edit.commit().unwrap();
    let mut right_edit = master.edit();
    right_edit
        .add_body_item(BodyItemSpec::generated_index(IndexKind::User, "Shared Index").unwrap())
        .unwrap();
    let right_commit = right_edit.commit().unwrap();
    let plan = left_commit
        .patch()
        .plan_three_way(right_commit.patch())
        .unwrap();
    assert!(plan.conflicts().conflicts().iter().any(
        |conflict| matches!(conflict, Conflict::BodyItemName(name) if name == "Shared Index")
    ));
}
