#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]

use super::{Snapshot, apply_patch, load};
use crate::presentation::embedded::controls::{
    BINARY_CONTENT_TYPE, BINARY_RELATIONSHIP, CONTROL_RELATIONSHIP, DESCRIPTOR_CONTENT_TYPE, Limits,
};
use litchi_opc::{BlobPart, OpcPackage, PackURI, Part};

const PML: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const AX: &str = "http://schemas.microsoft.com/office/2006/activeX";
const SLIDE: &str = "/ppt/slides/slide1.xml";
const DESCRIPTOR: &str = "/ppt/activeX/activeX1.xml";
const BINARY: &str = "/ppt/activeX/activeX1.bin";

fn package(mce: bool) -> OpcPackage {
    let controls = if mce {
        r#"<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice Requires="v" xmlns:v="urn:schemas-microsoft-com:vml"><p:control name="ChoiceName" r:id="rIdControl" showAsIcon="0" imgW="10" imgH="20"/></mc:Choice><mc:Fallback><p:control name="FallbackName" r:id="rIdControl" showAsIcon="0" imgW="10" imgH="20"/></mc:Fallback></mc:AlternateContent>"#.to_string()
    } else {
        r#"<p:control name="OldName" r:id="rIdControl" showAsIcon="0" imgW="10" imgH="20"><x:opaque xmlns:x="urn:opaque">retain</x:opaque></p:control>"#.into()
    };
    let slide_xml = format!(
        r#"<p:sld xmlns:p="{PML}" xmlns:r="{REL}"><p:cSld><p:controls>{controls}</p:controls></p:cSld></p:sld>"#
    );
    let descriptor_xml = format!(
        r#"<ax:ocx xmlns:ax="{AX}" xmlns:r="{REL}" ax:classid="old-class" ax:license="old-license" ax:persistence="persistStream" r:id="rIdBinary"><x:opaque xmlns:x="urn:opaque"><x:value>retain</x:value></x:opaque></ax:ocx>"#
    );
    let mut slide = BlobPart::new(
        PackURI::new(SLIDE).unwrap(),
        "application/vnd.openxmlformats-officedocument.presentationml.slide+xml".into(),
        slide_xml.into_bytes(),
    );
    slide.rels_mut().add_relationship(
        CONTROL_RELATIONSHIP.into(),
        "../activeX/activeX1.xml".into(),
        "rIdControl".into(),
        false,
    );
    let mut descriptor = BlobPart::new(
        PackURI::new(DESCRIPTOR).unwrap(),
        DESCRIPTOR_CONTENT_TYPE.into(),
        descriptor_xml.into_bytes(),
    );
    descriptor.rels_mut().add_relationship(
        BINARY_RELATIONSHIP.into(),
        "activeX1.bin".into(),
        "rIdBinary".into(),
        false,
    );
    let binary = BlobPart::new(
        PackURI::new(BINARY).unwrap(),
        BINARY_CONTENT_TYPE.into(),
        vec![0, 1, 2, 255],
    );
    let mut package = OpcPackage::new();
    package.add_part(Box::new(slide));
    package.add_part(Box::new(descriptor));
    package.add_part(Box::new(binary));
    package
}

fn snapshot(package: &OpcPackage) -> Snapshot {
    let slide_uri = PackURI::new(SLIDE).unwrap();
    let slide = package.get_part(&slide_uri).unwrap();
    load(package, 0, slide, 0, &mut Limits::default()).unwrap()
}

#[test]
fn exact_noop_preserves_source_and_package_bytes() {
    let mut package = package(false);
    let before = package
        .get_part(&PackURI::new(SLIDE).unwrap())
        .unwrap()
        .blob()
        .to_vec();
    let source = snapshot(&package);
    let commit = source.edit().commit().unwrap();
    assert!(!commit.is_changed());
    assert_eq!(commit.snapshot().source_xml(), before.as_slice());
    apply_patch(&mut package, commit.patch()).unwrap();
    assert_eq!(
        package
            .get_part(&PackURI::new(SLIDE).unwrap())
            .unwrap()
            .blob(),
        before
    );
}

#[test]
fn typed_edits_preserve_opaque_branches_and_binary_payload() {
    let mut package = package(true);
    let source = snapshot(&package);
    let binary_before = source.binary_bytes().unwrap().to_vec();
    let mut edit = source.edit();
    edit.set_name(Some("Edited & Safe".into())).unwrap();
    edit.set_show_as_icon(Some(true));
    edit.set_image_width(Some(42));
    edit.set_license(Some("new-license".into())).unwrap();
    edit.set_persistence(super::super::model::Persistence::Storage)
        .unwrap();
    let commit = edit.commit().unwrap();
    let target_xml = std::str::from_utf8(commit.snapshot().source_xml()).unwrap();
    assert!(target_xml.contains("name=\"Edited &amp; Safe\""));
    assert_eq!(target_xml.matches("name=\"Edited &amp; Safe\"").count(), 2);
    assert!(target_xml.contains("Choice"));
    assert!(target_xml.contains("xmlns:mc"));
    assert!(target_xml.contains("showAsIcon=\"true\""));
    assert!(target_xml.contains("imgW=\"42\""));
    assert_eq!(
        commit.snapshot().binary_bytes(),
        Some(binary_before.as_slice())
    );
    apply_patch(&mut package, commit.patch()).unwrap();
    let applied = snapshot(&package);
    assert_eq!(applied.control().name(), Some("Edited & Safe"));
    assert_eq!(applied.control().show_as_icon(), Some(true));
    assert_eq!(applied.control().image_width(), Some(42));
    assert_eq!(
        applied.control().descriptor().unwrap().license(),
        Some("new-license")
    );
}

#[test]
fn binary_payload_replacement_keeps_the_descriptor_relationship() {
    let mut package = package(false);
    let source = snapshot(&package);
    let mut edit = source.edit();
    edit.replace_binary(vec![9, 8, 7]).unwrap();
    let commit = edit.commit().unwrap();
    apply_patch(&mut package, commit.patch()).unwrap();
    let applied = snapshot(&package);
    assert_eq!(applied.binary_bytes(), Some(&[9, 8, 7][..]));
    assert_eq!(
        applied
            .control()
            .descriptor()
            .unwrap()
            .binary()
            .unwrap()
            .byte_length(),
        3
    );
    assert_eq!(
        package
            .get_part(&PackURI::new(DESCRIPTOR).unwrap())
            .unwrap()
            .rels()
            .get("rIdBinary")
            .unwrap()
            .reltype(),
        BINARY_RELATIONSHIP
    );
}

#[test]
fn optional_descriptor_attributes_can_be_authored_without_rewriting_opaque_children() {
    let mut package = package(false);
    let descriptor_uri = PackURI::new(DESCRIPTOR).unwrap();
    package
        .get_part_mut(&descriptor_uri)
        .unwrap()
        .set_blob(
            format!(
                r#"<ax:ocx xmlns:ax="{AX}" xmlns:r="{REL}" ax:classid="old-class" r:id="rIdBinary"><x:opaque xmlns:x="urn:opaque">keep</x:opaque></ax:ocx>"#
            )
            .into_bytes(),
        );
    let source = snapshot(&package);
    let mut edit = source.edit();
    edit.set_license(Some("added".into())).unwrap();
    edit.set_persistence(super::super::model::Persistence::StreamInit)
        .unwrap();
    let commit = edit.commit().unwrap();
    let descriptor_xml = std::str::from_utf8(commit.snapshot().descriptor_xml().unwrap()).unwrap();
    assert!(descriptor_xml.contains("ax:license=\"added\""));
    assert!(descriptor_xml.contains("ax:persistence=\"persistStreamInit\""));
    assert!(descriptor_xml.contains("<x:opaque"));
    apply_patch(&mut package, commit.patch()).unwrap();
    assert_eq!(
        snapshot(&package)
            .control()
            .descriptor()
            .unwrap()
            .persistence(),
        super::super::model::Persistence::StreamInit
    );
}

#[test]
fn binary_relationship_detach_collects_orphan_and_inverse_restores_it() {
    let mut package = package(false);
    let source = snapshot(&package);
    let mut edit = source.edit();
    edit.remove_binary().unwrap();
    let commit = edit.commit().unwrap();
    apply_patch(&mut package, commit.patch()).unwrap();
    let detached = snapshot(&package);
    assert!(detached.binary_bytes().is_none());
    assert!(package.get_part(&PackURI::new(BINARY).unwrap()).is_err());

    let restored = commit.patch().inverse();
    apply_patch(&mut package, &restored).unwrap();
    let current = snapshot(&package);
    assert_eq!(current.binary_bytes(), Some(&[0, 1, 2, 255][..]));
    assert_eq!(
        current
            .control()
            .descriptor()
            .unwrap()
            .binary()
            .unwrap()
            .byte_length(),
        4
    );
}

#[test]
fn stale_source_rejection_is_atomic() {
    let mut package = package(false);
    let source = snapshot(&package);
    let mut edit = source.edit();
    edit.set_name(Some("new".into())).unwrap();
    let patch = edit.commit().unwrap().into_patch();
    let slide_uri = PackURI::new(SLIDE).unwrap();
    let original = package.get_part(&slide_uri).unwrap().blob().to_vec();
    package
        .get_part_mut(&slide_uri)
        .unwrap()
        .set_blob(b"stale".to_vec());
    assert!(apply_patch(&mut package, &patch).is_err());
    assert_eq!(package.get_part(&slide_uri).unwrap().blob(), b"stale");
    assert_ne!(original, package.get_part(&slide_uri).unwrap().blob());
}

#[test]
fn malformed_edit_fails_before_source_publication() {
    let package = package(false);
    let source = snapshot(&package);
    let mut edit = source.edit();
    assert!(edit.set_name(Some("bad\u{0000}".into())).is_err());
    assert_eq!(snapshot(&package).source_xml(), source.source_xml());
}
