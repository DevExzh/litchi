use super::{
    Anchor, Limits, Payload, Relationship, Snapshot, Target, TargetMode, apply_commit, apply_patch,
    load_slide, load_snapshot,
};
use crate::Error;
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{BlobPart, OpcPackage, PackURI, Part};

const PML: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const P14: &str = "http://schemas.microsoft.com/office/powerpoint/2010/main";

#[test]
fn loads_opaque_payload_and_relationship_metadata_losslessly() {
    let payload = b"<?xml version=\"1.0\"?><opaque:root xmlns:opaque=\"urn:opaque\"><opaque:value>raw &amp; bytes</opaque:value></opaque:root>";
    let (package, slide_name) = package_with_internal_payload(payload);
    let slide = package.get_part(&slide_name).expect("slide");
    let mut limits = Limits::default();
    let parts = load_slide(&package, 3, slide, &mut limits).expect("content part");

    assert_eq!(parts.len(), 1);
    let content_part = &parts[0];
    assert_eq!(content_part.slide_index(), 3);
    assert_eq!(content_part.index(), 0);
    assert_eq!(content_part.relationship_id(), "rIdOpaque");
    assert_eq!(
        content_part.anchor().xml(),
        b"<p:contentPart r:id=\"rIdOpaque\"><p14:nvContentPartPr/></p:contentPart>"
    );
    assert_eq!(content_part.relationship().id(), "rIdOpaque");
    assert_eq!(
        content_part.relationship().relationship_type(),
        rt::CUSTOM_XML
    );
    assert_eq!(
        content_part.relationship().target_ref(),
        "../custom/opaque.xml"
    );

    let payload_view = content_part.payload().expect("internal payload");
    assert_eq!(payload_view.part_name().as_str(), "/ppt/custom/opaque.xml");
    assert_eq!(payload_view.content_type(), "application/xml");
    assert_eq!(payload_view.bytes(), payload);
    assert_eq!(payload_view.relationships().len(), 1);
    assert_eq!(payload_view.relationships()[0].id(), "rIdPayloadLink");
    assert_eq!(
        payload_view.relationships()[0].target_ref(),
        "https://example.invalid/opaque"
    );
}

#[test]
fn discovery_is_a_source_preserving_noop() {
    let (package, slide_name) = package_with_internal_payload(b"opaque");
    let before = package
        .get_part(&slide_name)
        .expect("slide")
        .blob()
        .to_vec();
    let slide = package.get_part(&slide_name).expect("slide");
    let _ = load_slide(&package, 0, slide, &mut Limits::default()).expect("content part");
    assert_eq!(package.get_part(&slide_name).unwrap().blob(), before);
}

#[test]
fn retains_external_target_mode_without_following_it() {
    let (package, slide_name) = package_with_external_payload();
    let slide = package.get_part(&slide_name).expect("slide");
    let parts = load_slide(&package, 0, slide, &mut Limits::default()).expect("content part");
    let relationship = parts[0].relationship();

    assert_eq!(relationship.target_mode(), TargetMode::External);
    assert_eq!(relationship.target_ref(), "https://example.invalid/content");
    assert!(relationship.payload().is_none());
    assert_eq!(
        relationship.target().external_ref(),
        Some("https://example.invalid/content")
    );
}

#[test]
fn rejects_missing_anchor_relationship() {
    let package = package_with_anchor(b"<p:contentPart r:id=\"rIdMissing\"/>", None, None);
    let slide_name = PackURI::new("/ppt/slides/slide1.xml").unwrap();
    let slide = package.get_part(&slide_name).expect("slide");
    let error = load_slide(&package, 0, slide, &mut Limits::default()).unwrap_err();
    assert!(matches!(error, Error::Relationship(message) if message.contains("rIdMissing")));
}

#[test]
fn rejects_anchor_without_relationship_id() {
    let package = package_with_anchor(b"<p:contentPart/>", None, None);
    let slide_name = PackURI::new("/ppt/slides/slide1.xml").unwrap();
    let slide = package.get_part(&slide_name).expect("slide");
    let error = load_slide(&package, 0, slide, &mut Limits::default()).unwrap_err();
    assert!(matches!(error, Error::Invalid(message) if message.contains("missing r:id")));
}

#[test]
fn enforces_content_part_count_limit() {
    let package = package_with_anchor(
        b"<p:contentPart r:id=\"rId1\"/><p:contentPart r:id=\"rId2\"/>",
        Some((rt::CUSTOM_XML, "../custom/opaque.xml", "rId1", false)),
        Some((rt::CUSTOM_XML, "../custom/opaque.xml", "rId2", false)),
    );
    let slide_name = PackURI::new("/ppt/slides/slide1.xml").unwrap();
    let slide = package.get_part(&slide_name).expect("slide");
    let mut limits = Limits::new(1, 1024, 4096, 8).unwrap();
    let error = load_slide(&package, 0, slide, &mut limits).unwrap_err();
    assert!(matches!(
        error,
        Error::Limit {
            resource: "content-part count",
            ..
        }
    ));
}

#[test]
fn exact_noop_snapshot_commit_preserves_the_complete_source_graph() {
    let mut package = package_with_internal_payload(b"opaque").0;
    let source = snapshot(&package);
    let slide_before = package
        .get_part(&PackURI::new("/ppt/slides/slide1.xml").unwrap())
        .unwrap()
        .blob()
        .to_vec();
    let payload_before = package
        .get_part(&PackURI::new("/ppt/custom/opaque.xml").unwrap())
        .unwrap()
        .blob()
        .to_vec();
    let commit = source.edit().commit().unwrap();

    assert!(!commit.is_changed());
    assert!(commit.patch().is_empty());
    apply_commit(&mut package, commit).unwrap();
    assert_eq!(
        package
            .get_part(&PackURI::new("/ppt/slides/slide1.xml").unwrap())
            .unwrap()
            .blob(),
        slide_before.as_slice()
    );
    assert_eq!(
        package
            .get_part(&PackURI::new("/ppt/custom/opaque.xml").unwrap())
            .unwrap()
            .blob(),
        payload_before.as_slice()
    );
}

#[test]
fn relationship_and_payload_edits_preserve_unknown_anchor_markup() {
    let mut package = package_with_internal_payload(b"opaque").0;
    let source = snapshot(&package);
    let mut edit = source.edit();
    edit.set_relationship_type(0, "urn:vendor:content-part")
        .unwrap();
    edit.replace_payload(0, "application/vendor+xml", b"edited")
        .unwrap();
    let commit = edit.commit().unwrap();
    let source_xml = std::str::from_utf8(commit.snapshot().source_xml()).unwrap();
    assert!(
        source_xml.contains("<p14:nvContentPartPr/>") || source_xml.contains("p14:nvContentPartPr")
    );

    apply_patch(&mut package, commit.patch()).unwrap();
    let current = snapshot(&package);
    assert_eq!(
        current.parts()[0].relationship().relationship_type(),
        "urn:vendor:content-part"
    );
    assert_eq!(
        current.parts()[0].payload().unwrap().content_type(),
        "application/vendor+xml"
    );
    assert_eq!(current.parts()[0].payload().unwrap().bytes(), b"edited");
    assert!(
        current
            .source_xml()
            .windows(b"p14:nvContentPartPr".len())
            .any(|window| { window == b"p14:nvContentPartPr" })
    );
}

#[test]
fn relationship_id_edit_rewrites_only_the_anchor_attribute() {
    let (mut package, _) = package_with_internal_payload(b"opaque");
    let source = snapshot(&package);
    let mut edit = source.edit();
    edit.set_relationship_id(0, "rIdRenamed").unwrap();
    let commit = edit.commit().unwrap();
    let source_xml = std::str::from_utf8(commit.snapshot().source_xml()).unwrap();
    assert!(source_xml.contains(r#"r:id="rIdRenamed""#));
    assert!(!source_xml.contains(r#"r:id="rIdOpaque""#));
    assert!(source_xml.contains("p14:nvContentPartPr"));

    apply_patch(&mut package, commit.patch()).unwrap();
    let current = snapshot(&package);
    assert_eq!(current.parts()[0].relationship_id(), "rIdRenamed");
    let slide = package
        .get_part(&PackURI::new("/ppt/slides/slide1.xml").unwrap())
        .unwrap();
    assert!(slide.rels().get("rIdOpaque").is_none());
    assert!(slide.rels().get("rIdRenamed").is_some());

    apply_patch(&mut package, &commit.patch().inverse()).unwrap();
    assert_eq!(snapshot(&package).source_xml(), source.source_xml());
}

#[test]
fn add_remove_and_inverse_manage_relationships_and_payload_parts() {
    let (mut package, _) = package_with_internal_payload(b"opaque");
    let source = snapshot(&package);
    let payload_name = PackURI::new("/ppt/custom/added.xml").unwrap();
    let payload = Payload::new(
        payload_name.clone(),
        "application/custom+xml",
        b"added".to_vec(),
    );
    let anchor = Anchor::new(
        "rIdAdded",
        br#"<p:contentPart r:id="rIdAdded"><p14:nvContentPartPr/></p:contentPart>"#,
    );
    let relationship = Relationship::new(
        "rIdAdded",
        litchi_opc::constants::relationship_type::CUSTOM_XML,
        "../custom/added.xml",
        TargetMode::Internal,
        Target::internal(payload),
    );
    let mut edit = source.edit();
    edit.push(anchor, relationship).unwrap();
    let commit = edit.commit().unwrap();
    apply_patch(&mut package, commit.patch()).unwrap();
    assert_eq!(snapshot(&package).parts().len(), 2);
    assert_eq!(package.get_part(&payload_name).unwrap().blob(), b"added");

    let inverse = commit.patch().inverse();
    apply_patch(&mut package, &inverse).unwrap();
    let restored = snapshot(&package);
    assert_eq!(restored.source_xml(), source.source_xml());
    assert_eq!(restored.parts(), source.parts());
    assert!(package.get_part(&payload_name).is_err());
}

#[test]
fn removal_collects_orphan_payload_and_inverse_restores_it() {
    let (mut package, _) = package_with_internal_payload(b"opaque");
    let source = snapshot(&package);
    let mut edit = source.edit();
    edit.remove(0).unwrap();
    let commit = edit.commit().unwrap();
    apply_patch(&mut package, commit.patch()).unwrap();
    assert!(snapshot(&package).parts().is_empty());
    assert!(
        package
            .get_part(&PackURI::new("/ppt/custom/opaque.xml").unwrap())
            .is_err()
    );

    apply_patch(&mut package, &commit.patch().inverse()).unwrap();
    assert_eq!(snapshot(&package).source_xml(), source.source_xml());
}

#[test]
fn stale_source_rejection_is_atomic() {
    let (mut package, _) = package_with_internal_payload(b"opaque");
    let source = snapshot(&package);
    let mut edit = source.edit();
    edit.set_relationship_type(0, "urn:changed").unwrap();
    let patch = edit.commit().unwrap().into_patch();
    let slide_name = PackURI::new("/ppt/slides/slide1.xml").unwrap();
    package
        .get_part_mut(&slide_name)
        .unwrap()
        .set_blob(b"stale".to_vec());
    assert!(patch.apply(&mut package).is_err());
    assert_eq!(package.get_part(&slide_name).unwrap().blob(), b"stale");
}

#[test]
fn invalid_edits_do_not_mutate_the_staged_snapshot() {
    let (package, _) = package_with_internal_payload(b"opaque");
    let source = snapshot(&package);
    let mut edit = source.edit();
    assert!(edit.set_relationship_id(0, "bad\u{0000}").is_err());
    assert_eq!(edit.parts(), source.parts());
}

fn package_with_internal_payload(payload: &[u8]) -> (OpcPackage, PackURI) {
    let relationship = Some((rt::CUSTOM_XML, "../custom/opaque.xml", "rIdOpaque", false));
    let mut package = package_with_anchor(
        b"<p:contentPart r:id=\"rIdOpaque\"><p14:nvContentPartPr/></p:contentPart>",
        relationship,
        None,
    );
    let payload_name = PackURI::new("/ppt/custom/opaque.xml").unwrap();
    package
        .get_part_mut(&payload_name)
        .expect("payload")
        .set_blob(payload.to_vec());
    (package, PackURI::new("/ppt/slides/slide1.xml").unwrap())
}

fn package_with_external_payload() -> (OpcPackage, PackURI) {
    let package = package_with_anchor(
        b"<p:contentPart r:id=\"rIdExternal\"/>",
        Some((
            rt::CUSTOM_XML,
            "https://example.invalid/content",
            "rIdExternal",
            true,
        )),
        None,
    );
    (package, PackURI::new("/ppt/slides/slide1.xml").unwrap())
}

fn package_with_anchor(
    anchors: &[u8],
    relationship: Option<(&str, &str, &str, bool)>,
    relationship_2: Option<(&str, &str, &str, bool)>,
) -> OpcPackage {
    let slide_name = PackURI::new("/ppt/slides/slide1.xml").unwrap();
    let xml = format!(
        "<p:sld xmlns:p=\"{PML}\" xmlns:r=\"{REL}\" xmlns:p14=\"{P14}\"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>{}</p:spTree></p:cSld></p:sld>",
        String::from_utf8(anchors.to_vec()).unwrap()
    );
    let mut slide = BlobPart::new(slide_name.clone(), ct::PML_SLIDE.into(), xml.into_bytes());
    if let Some((relationship_type, target, id, external)) = relationship {
        slide.rels_mut().add_relationship(
            relationship_type.to_owned(),
            target.to_owned(),
            id.to_owned(),
            external,
        );
    }
    if let Some((relationship_type, target, id, external)) = relationship_2 {
        slide.rels_mut().add_relationship(
            relationship_type.to_owned(),
            target.to_owned(),
            id.to_owned(),
            external,
        );
    }
    let mut package = OpcPackage::new();
    package.add_part(Box::new(slide));
    let payload_name = PackURI::new("/ppt/custom/opaque.xml").unwrap();
    let mut payload = BlobPart::new(
        payload_name,
        "application/xml".to_owned(),
        b"opaque payload".to_vec(),
    );
    payload.rels_mut().add_relationship(
        rt::HYPERLINK.to_owned(),
        "https://example.invalid/opaque".to_owned(),
        "rIdPayloadLink".to_owned(),
        true,
    );
    package.add_part(Box::new(payload));
    package
}

fn snapshot(package: &OpcPackage) -> Snapshot {
    let slide_name = PackURI::new("/ppt/slides/slide1.xml").unwrap();
    let slide = package.get_part(&slide_name).expect("slide");
    load_snapshot(package, 0, slide, &mut Limits::default()).expect("snapshot")
}
