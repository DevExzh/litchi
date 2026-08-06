//! Focused snapshot, semantic-edit, and package-atomicity coverage.

use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{BlobPart, OpcPackage, PackURI};

use super::{Snapshot, Transaction, validate_graph};

const COMMENTS_XML: &[u8] = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <authors><author>Alice</author><author>Bob</author></authors>
  <commentList><comment ref="A1" authorId="0" shapeId="1026"><text><t>legacy note</t></text></comment></commentList>
</comments>"#;
const VML_XML: &[u8] = br##"<xml><shape id="_x0000_s1026" type="#_x0000_t202"/></xml>"##;

fn fixture() -> (OpcPackage, PackURI) {
    let mut package = OpcPackage::new();
    let worksheet = PackURI::new("/xl/worksheets/sheet1.xml").unwrap();
    let comments = PackURI::new("/xl/comments1.xml").unwrap();
    let vml = PackURI::new("/xl/drawings/vmlDrawing1.vml").unwrap();
    package.add_part(Box::new(BlobPart::new(
        worksheet.clone(),
        ct::SML_WORKSHEET.into(),
        br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>"#
            .to_vec(),
    )));
    package.add_part(Box::new(BlobPart::new(
        comments,
        ct::SML_COMMENTS.into(),
        COMMENTS_XML.to_vec(),
    )));
    package.add_part(Box::new(BlobPart::new(
        vml,
        ct::OFC_VML_DRAWING.into(),
        VML_XML.to_vec(),
    )));
    let worksheet_part = package.get_part_mut(&worksheet).unwrap();
    worksheet_part.rels_mut().add_relationship(
        rt::COMMENTS.to_owned(),
        "../comments1.xml".to_owned(),
        "rIdComments".to_owned(),
        false,
    );
    worksheet_part.rels_mut().add_relationship(
        rt::VML_DRAWING.to_owned(),
        "../drawings/vmlDrawing1.vml".to_owned(),
        "rIdVml".to_owned(),
        false,
    );
    validate_graph(&package).unwrap();
    (package, worksheet)
}

fn comments_source(package: &OpcPackage) -> Vec<u8> {
    package
        .get_part(&PackURI::new("/xl/comments1.xml").unwrap())
        .unwrap()
        .blob()
        .to_vec()
}

fn vml_source(package: &OpcPackage) -> Vec<u8> {
    package
        .get_part(&PackURI::new("/xl/drawings/vmlDrawing1.vml").unwrap())
        .unwrap()
        .blob()
        .to_vec()
}

#[test]
fn snapshot_retains_exact_source_and_inert_shape_identity() {
    let (package, worksheet) = fixture();
    let snapshot = Snapshot::load(&package, &worksheet).unwrap();
    assert_eq!(snapshot.source_xml(), Some(COMMENTS_XML));
    let comment = snapshot.comments().unwrap().get("A1").unwrap();
    assert_eq!(comment.author, "Alice");
    assert_eq!(comment.text, "legacy note");
    assert_eq!(comment.shape_id, Some(1026));
    assert_eq!(vml_source(&package), VML_XML);
}

#[test]
fn semantic_no_op_keeps_exact_bytes_and_relationships() {
    let (mut package, worksheet) = fixture();
    let source = comments_source(&package);
    let relationship_ids: Vec<_> = package
        .get_part(&worksheet)
        .unwrap()
        .rels()
        .iter()
        .map(|relationship| relationship.r_id().to_owned())
        .collect();

    let mut transaction = Transaction::new(&mut package, &worksheet).unwrap();
    assert!(!transaction.set("A1", "Alice", "legacy note").unwrap());
    let commit = transaction.commit().unwrap();
    assert!(!commit.changed());
    assert!(commit.patch().is_empty());
    assert_eq!(comments_source(&package), source);
    let current_ids: Vec<_> = package
        .get_part(&worksheet)
        .unwrap()
        .rels()
        .iter()
        .map(|relationship| relationship.r_id().to_owned())
        .collect();
    assert_eq!(current_ids, relationship_ids);
}

#[test]
fn checked_edits_are_atomic_and_preserve_vml_shape_ids() {
    let (mut package, worksheet) = fixture();
    let vml_before = vml_source(&package);
    let mut transaction = Transaction::new(&mut package, &worksheet).unwrap();
    let before = transaction.comments().cloned();
    assert!(transaction.set("A0", "Alice", "invalid").is_err());
    assert_eq!(transaction.comments().cloned(), before);
    assert!(transaction.set_author(99, "Nobody").is_err());
    assert_eq!(transaction.comments().cloned(), before);

    assert!(transaction.set("A1", "Bob", "updated").unwrap());
    assert!(transaction.rename_author("Bob", "Robert").unwrap());
    let commit = transaction.commit().unwrap();
    let comment = commit.snapshot().comments().unwrap().get("A1").unwrap();
    assert_eq!(comment.author, "Robert");
    assert_eq!(comment.text, "updated");
    assert_eq!(comment.author_id, 1);
    assert_eq!(comment.shape_id, Some(1026));
    assert_eq!(vml_source(&package), vml_before);
}

#[test]
fn patch_replays_exact_source_and_inverse_atomically() {
    let (mut package, worksheet) = fixture();
    let original = comments_source(&package);
    let mut transaction = Transaction::new(&mut package, &worksheet).unwrap();
    assert!(transaction.set("A1", "Alice", "changed").unwrap());
    let commit = transaction.commit().unwrap();
    let patch = commit.patch().clone();
    let changed = comments_source(&package);

    let (mut replay, replay_worksheet) = fixture();
    assert_eq!(worksheet, replay_worksheet);
    patch.apply(&mut replay).unwrap();
    assert_eq!(comments_source(&replay), changed);
    patch.inverse().apply(&mut replay).unwrap();
    assert_eq!(comments_source(&replay), original);
    assert_eq!(vml_source(&replay), VML_XML);
}

#[test]
fn patch_conflict_does_not_publish_partial_relationship_changes() {
    let (mut source, worksheet) = fixture();
    let mut transaction = Transaction::new(&mut source, &worksheet).unwrap();
    assert!(transaction.set("A1", "Alice", "changed").unwrap());
    let patch = transaction.commit().unwrap().patch().clone();

    let (mut target, target_worksheet) = fixture();
    let conflicting = String::from_utf8(COMMENTS_XML.to_vec())
        .unwrap()
        .replace("legacy note", "different source");
    target
        .get_part_mut(&PackURI::new("/xl/comments1.xml").unwrap())
        .unwrap()
        .set_blob(conflicting.into_bytes());
    let before = comments_source(&target);
    let ids_before: Vec<_> = target
        .get_part(&target_worksheet)
        .unwrap()
        .rels()
        .iter()
        .map(|relationship| relationship.r_id().to_owned())
        .collect();

    assert!(patch.apply(&mut target).is_err());
    assert_eq!(comments_source(&target), before);
    let ids_after: Vec<_> = target
        .get_part(&target_worksheet)
        .unwrap()
        .rels()
        .iter()
        .map(|relationship| relationship.r_id().to_owned())
        .collect();
    assert_eq!(ids_after, ids_before);
}
