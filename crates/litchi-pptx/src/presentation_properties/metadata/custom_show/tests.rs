#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]

use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::part::BlobPart;
use litchi_opc::{OpcPackage, PackURI};

use super::*;

const PRESENTATION_XML: &[u8] = br#"<?xml version="1.0" encoding="UTF-8"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:x="urn:future">
  <p:sldIdLst>
    <p:sldId id="256" r:id="rIdSlideOne"/>
    <p:sldId id="257" r:id="rIdSlideTwo"/>
    <p:sldId id="258" r:id="rIdSlideThree"/>
  </p:sldIdLst>
  <x:beforeOpaque/>
  <p:custShowLst future="keep">
    <x:listOpaque/>
    <p:custShow name="Opening &amp; More" id="7" future="show">
      <x:showOpaque/>
      <p:sldLst future="list">
        <x:slideOpaque/>
        <p:sld r:id="rIdSlideOne" future="slide"/>
        <p:sld r:id="rIdSlideThree"/>
      </p:sldLst>
    </p:custShow>
    <p:custShow name="Recap" id="8">
      <p:sldLst><p:sld r:id="rIdSlideTwo"/></p:sldLst>
    </p:custShow>
  </p:custShowLst>
  <x:afterOpaque/>
</p:presentation>"#;

fn presentation_name() -> PackURI {
    PackURI::new("/ppt/presentation.xml").unwrap()
}

fn fixture() -> OpcPackage {
    let mut package = OpcPackage::new();
    package.add_part(Box::new(BlobPart::new(
        presentation_name(),
        ct::PML_PRESENTATION_MAIN.to_owned(),
        PRESENTATION_XML.to_vec(),
    )));
    package.relate_to("/ppt/presentation.xml", rt::OFFICE_DOCUMENT);

    let presentation = package.get_part_mut(&presentation_name()).unwrap();
    for (relationship_id, target) in [
        ("rIdSlideOne", "slides/slide1.xml"),
        ("rIdSlideTwo", "slides/slide2.xml"),
        ("rIdSlideThree", "slides/slide3.xml"),
    ] {
        presentation.rels_mut().add_relationship(
            rt::SLIDE.to_owned(),
            target.to_owned(),
            relationship_id.to_owned(),
            false,
        );
    }
    for index in 1..=3 {
        package.add_part(Box::new(BlobPart::new(
            PackURI::new(format!("/ppt/slides/slide{index}.xml")).unwrap(),
            ct::PML_SLIDE.to_owned(),
            b"<p:sld xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\"/>"
                .to_vec(),
        )));
    }
    package
}

#[test]
fn snapshot_crud_preserves_opaque_xml_and_relationships() {
    let mut package = fixture();
    let original = load_snapshot(&package).unwrap();
    assert_eq!(original.list().shows[0].name, "Opening & More");
    assert_eq!(original.list().shows[0].slide_ids, [256, 258]);

    let mut edit = original.edit();
    assert!(edit.set_name(7, "Opening & Updated").unwrap());
    edit.add_slide(7, 257).unwrap();
    let created = edit.create("New Show", vec![258]).unwrap();
    assert_eq!(created, 9);
    let removed = edit.remove(8).unwrap();
    assert_eq!(removed.name, "Recap");
    edit.reorder(&[9, 7]).unwrap();

    let commit = edit.commit().unwrap();
    assert!(commit.is_changed());
    let output = String::from_utf8_lossy(commit.snapshot().source_xml());
    assert!(output.contains("future=\"keep\""));
    assert!(output.contains("<x:listOpaque/>"));
    assert!(output.contains("future=\"show\""));
    assert!(output.contains("<x:showOpaque/>"));
    assert!(output.contains("future=\"list\""));
    assert!(output.contains("<x:slideOpaque/>"));
    assert!(output.contains("<x:afterOpaque/>"));
    apply_commit(&mut package, commit).unwrap();

    let current = load_snapshot(&package).unwrap();
    assert_eq!(
        current
            .list()
            .shows
            .iter()
            .map(|show| show.id)
            .collect::<Vec<_>>(),
        [9, 7]
    );
    assert_eq!(
        current.list().get_by_id(7).unwrap().slide_ids,
        [256, 258, 257]
    );
    assert_eq!(
        package
            .get_part(&presentation_name())
            .unwrap()
            .rels()
            .get("rIdSlideOne")
            .unwrap()
            .target_ref(),
        "slides/slide1.xml"
    );
}

#[test]
fn no_op_commit_is_byte_exact() {
    let mut package = fixture();
    let snapshot = load_snapshot(&package).unwrap();
    let source = snapshot.source_xml().to_vec();
    let commit = snapshot.edit().commit().unwrap();
    assert!(!commit.is_changed());
    assert!(commit.patch().is_empty());
    apply_commit(&mut package, commit).unwrap();
    assert_eq!(
        package.get_part(&presentation_name()).unwrap().blob(),
        source.as_slice()
    );
}

#[test]
fn stale_patch_rejection_is_atomic_and_inverse_restores_bytes() {
    let mut package = fixture();
    let original = load_snapshot(&package).unwrap();
    let mut edit = original.edit();
    edit.set_name(7, "Changed").unwrap();
    let patch = edit.commit().unwrap().into_patch();

    let mut stale = package
        .get_part(&presentation_name())
        .unwrap()
        .blob()
        .to_vec();
    stale.extend_from_slice(b" ");
    package
        .get_part_mut(&presentation_name())
        .unwrap()
        .set_blob(stale.clone());
    assert!(patch.apply(&mut package).is_err());
    assert_eq!(
        package.get_part(&presentation_name()).unwrap().blob(),
        stale.as_slice()
    );

    let mut restored = fixture();
    let source = load_snapshot(&restored).unwrap();
    patch.apply(&mut restored).unwrap();
    patch.inverse().apply(&mut restored).unwrap();
    assert_eq!(
        restored.get_part(&presentation_name()).unwrap().blob(),
        source.source_xml()
    );
}

#[test]
fn invalid_slide_references_leave_staged_list_unchanged() {
    let package = fixture();
    let snapshot = load_snapshot(&package).unwrap();
    let mut edit = snapshot.edit();
    let before = edit.list().clone();
    assert!(edit.set_slides(7, vec![999]).is_err());
    assert_eq!(edit.list(), &before);
}
