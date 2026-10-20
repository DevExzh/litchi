//! Focused source-read, typed graph edit and package-atomicity coverage.

use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{BlobPart, OpcPackage, PackURI};

use super::{Comment, load_from_worksheet, validate_graph};
use crate::Package;

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
    crate::record::adopt_worksheet(&mut package, &worksheet, false);
    (package, worksheet)
}

fn adopted() -> (Package, PackURI) {
    let (package, worksheet) = fixture();
    (Package::from_opc(package).unwrap(), worksheet)
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
fn source_read_retains_exact_source_and_inert_shape_identity() {
    let (package, worksheet) = fixture();
    let part = load_from_worksheet(&package, &worksheet).unwrap().unwrap();
    assert_eq!(comments_source(&package), COMMENTS_XML);
    let comment = part.comments.get("A1").unwrap();
    assert_eq!(comment.author, "Alice");
    assert_eq!(comment.text, "legacy note");
    assert_eq!(comment.shape_id, Some(1026));
    assert_eq!(vml_source(&package), VML_XML);
}

#[test]
fn semantic_no_op_keeps_exact_bytes_and_relationships() {
    let (package, worksheet) = adopted();
    let source = comments_source(package.opc());
    let relationship_ids: Vec<_> = package
        .opc()
        .get_part(&worksheet)
        .unwrap()
        .rels()
        .iter()
        .map(|relationship| relationship.r_id().to_owned())
        .collect();

    let mut edit = package.edit_comments().unwrap();
    let existing: Comment = edit
        .comments(0)
        .unwrap()
        .unwrap()
        .get("A1")
        .unwrap()
        .clone();
    assert!(!edit.set_comment(0, existing).unwrap());
    let commit = edit.commit().unwrap();
    assert!(!commit.diagnostics().changed());
    assert!(commit.patch().is_noop());
    assert_eq!(comments_source(commit.package().opc()), source);
    let current_ids: Vec<_> = commit
        .package()
        .opc()
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
    let (package, _) = adopted();
    let vml_before = vml_source(package.opc());
    let mut edit = package.edit_comments().unwrap();
    let before = edit.comments(0).unwrap();
    let mut invalid = before.as_ref().unwrap().get("A1").unwrap().clone();
    invalid.cell_ref = "A0".into();
    assert!(edit.set_comment(0, invalid).is_err());
    assert_eq!(edit.comments(0).unwrap(), before);
    assert!(edit.set_author(0, 99, "Nobody").is_err());
    assert_eq!(edit.comments(0).unwrap(), before);

    let mut updated = before.as_ref().unwrap().get("A1").unwrap().clone();
    updated.author = "Bob".into();
    updated.author_id = 1;
    updated.text = "updated".into();
    assert!(edit.set_comment(0, updated).unwrap());
    assert!(edit.rename_author(0, "Bob", "Robert").unwrap());
    let after = edit.commit().unwrap().into_package();
    let part = load_from_worksheet(
        after.opc(),
        &PackURI::new("/xl/worksheets/sheet1.xml").unwrap(),
    )
    .unwrap()
    .unwrap();
    let comment = part.comments.get("A1").unwrap();
    assert_eq!(comment.author, "Robert");
    assert_eq!(comment.text, "updated");
    assert_eq!(comment.author_id, 1);
    assert_eq!(comment.shape_id, Some(1026));
    assert_eq!(vml_source(after.opc()), vml_before);
}

#[test]
fn patch_replays_exact_source_and_inverse_atomically() {
    let (package, _) = adopted();
    let original = comments_source(package.opc());
    let mut edit = package.edit_comments().unwrap();
    let mut changed = edit
        .comments(0)
        .unwrap()
        .unwrap()
        .get("A1")
        .unwrap()
        .clone();
    changed.text = "changed".into();
    assert!(edit.set_comment(0, changed).unwrap());
    let commit = edit.commit().unwrap();
    let patch = commit.patch().clone();
    let changed = comments_source(commit.package().opc());

    let (replay, _) = adopted();
    let replayed = replay.apply_comments(&patch).unwrap().into_package();
    assert_eq!(comments_source(replayed.opc()), changed);
    let restored = replayed
        .apply_comments(&patch.inverse())
        .unwrap()
        .into_package();
    assert_eq!(comments_source(restored.opc()), original);
    assert_eq!(vml_source(restored.opc()), VML_XML);
}

#[test]
fn clear_removes_the_comments_part_and_keeps_the_vml_drawing() {
    let (package, worksheet) = adopted();
    let mut edit = package.edit_comments().unwrap();
    assert!(edit.clear(0).unwrap());
    assert!(!edit.clear(0).unwrap());
    let after = edit.commit().unwrap().into_package();
    assert!(
        load_from_worksheet(after.opc(), &worksheet)
            .unwrap()
            .is_none()
    );
    assert_eq!(vml_source(after.opc()), VML_XML);
}

#[test]
fn patch_conflict_does_not_publish_partial_relationship_changes() {
    let (source, _) = adopted();
    let mut edit = source.edit_comments().unwrap();
    let mut changed = edit
        .comments(0)
        .unwrap()
        .unwrap()
        .get("A1")
        .unwrap()
        .clone();
    changed.text = "changed".into();
    assert!(edit.set_comment(0, changed).unwrap());
    let patch = edit.commit().unwrap().patch().clone();

    let (mut target, _) = fixture();
    let conflicting = String::from_utf8(COMMENTS_XML.to_vec())
        .unwrap()
        .replace("legacy note", "different source");
    target
        .get_part_mut(&PackURI::new("/xl/comments1.xml").unwrap())
        .unwrap()
        .set_blob(conflicting.into_bytes());
    let target = Package::from_opc(target).unwrap();
    let before = comments_source(target.opc());
    assert!(matches!(
        target.apply_comments(&patch),
        Err(crate::Error::PatchConflict { .. })
    ));
    assert_eq!(comments_source(target.opc()), before);
}
