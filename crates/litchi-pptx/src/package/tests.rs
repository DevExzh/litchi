#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]

//! Regression tests for the `PresentationML` package facade.

use litchi_opc::PackURI;
use litchi_opc::constants::relationship_type as rt;

use super::Package;
use crate::custom::Props;
use crate::web::{AddIn, Conformance, Pane, Panes, Reference, Store};

fn with_eocd_comment(mut archive: Vec<u8>, comment: &[u8]) -> Vec<u8> {
    let comment_len = u16::try_from(comment.len()).expect("ZIP comment fits in EOCD");
    let eocd = archive.len().checked_sub(22).expect("archive has an EOCD");
    assert_eq!(&archive[eocd..eocd + 4], b"PK\x05\x06");
    archive[eocd + 20..eocd + 22].copy_from_slice(&comment_len.to_le_bytes());
    archive.extend_from_slice(comment);
    archive
}

#[test]
fn owned_presentation_noop_output_preserves_exact_archive() {
    let source = with_eocd_comment(
        Package::new()
            .expect("new presentation")
            .to_bytes()
            .expect("serialize presentation"),
        b"pptx exact source",
    );
    let package = Package::from_reader(std::io::Cursor::new(source.clone()))
        .expect("open owned presentation");

    assert_eq!(package.to_bytes().expect("serialize no-op package"), source);
}

#[test]
fn replay_index_ingress_does_not_decode_unrelated_slide_payload() {
    let mut builder = crate::builder::Builder::new().expect("new presentation");
    builder
        .presentation()
        .expect("presentation")
        .add_slide()
        .expect("slide");
    let source = builder
        .build()
        .expect("build presentation")
        .to_bytes()
        .expect("serialize presentation");
    let package = Package::from_reader(std::io::Cursor::new(source)).expect("open presentation");
    let slide_name = PackURI::new("/ppt/slides/slide1.xml").expect("slide name");
    let slide = package
        .opc
        .iter_parts()
        .find(|part| part.partname() == &slide_name)
        .expect("slide metadata");
    assert!(
        !slide.payload_is_decoded(),
        "replay ingress must use admitted metadata lengths, not hydrate unrelated slides"
    );
}

fn task_panes() -> Panes {
    let reference = Reference::new("litchi-test-add-in", "1.0.0.0", Store::Omex)
        .expect("valid add-in reference");
    let add_in =
        AddIn::new("{B1C15FE4-84FA-4773-AD36-9EF5444C5A01}", reference).expect("valid add-in");
    let mut panes = Panes::new();
    panes.push(Pane::new(add_in)).expect("valid task pane");
    panes
}

#[test]
fn new_writer_round_trips_the_bounded_slide_graph() {
    let mut builder = crate::builder::Builder::new().expect("new package");
    {
        let presentation = builder.presentation().expect("mutable presentation");
        let slide = presentation.add_slide().expect("slide");
        slide.set_title("Canonical owner");
        slide.add_text_box("Hello & goodbye", 914_400, 914_400, 2_743_200, 914_400);
        presentation.set_widescreen_slide_size();
    }

    let package = builder.build().expect("new package");
    let bytes = package.to_bytes().expect("serialize package");
    let reopened = Package::from_bytes(&bytes).expect("reopen package");
    let presentation = reopened.presentation().expect("presentation");
    assert_eq!(presentation.slide_count().expect("slide count"), 1);
    assert_eq!(
        presentation.slide_size().expect("slide size"),
        (9_144_000, 5_143_500)
    );
    let slide = presentation.slide(0).expect("slide lookup").expect("slide");
    assert_eq!(slide.name().expect("slide name"), "Slide 256");
    assert!(
        slide
            .text()
            .expect("slide text")
            .contains("Hello & goodbye")
    );
    assert_eq!(slide.shape_count().expect("shape count"), 2);
    assert_eq!(presentation.slide_masters().expect("masters").len(), 1);
    assert_eq!(presentation.slide_layouts().expect("layouts").len(), 11);
}

#[test]
fn fresh_and_opened_packages_have_canonical_edit_roots() {
    let package = Package::new().expect("new package");
    assert!(
        package
            .edit()
            .expect("edit")
            .commit()
            .expect("commit")
            .patch()
            .is_noop()
    );
    let bytes = package.to_bytes().expect("serialize package");
    let opened = Package::from_bytes(&bytes).expect("reopen package");
    assert!(package.same_content(&opened).expect("same content"));
}

#[test]
fn custom_property_no_ops_preserve_signatures_and_changes_require_explicit_unsign() {
    let mut package = Package::new().expect("new package");
    std::sync::Arc::make_mut(&mut package.opc)
        .relate_to("_xmlsignatures/origin.sigs", rt::DIGITAL_SIGNATURE_ORIGIN);
    assert!(package.opc.is_signed());

    package
        .put_custom_props(Props::new())
        .expect("empty custom properties are a no-op");
    assert!(package.opc.is_signed());

    // A change is refused with the crate's typed signed error and leaves the
    // package unchanged until the signature is removed explicitly (ADR 0006;
    // the shared owner used to remove it).
    let owner = || {
        let mut props = Props::new();
        props.insert("Owner", "Alice").expect("property");
        props
    };
    assert!(matches!(
        package.put_custom_props(owner()),
        Err(crate::Error::Opc(
            litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy
        ))
    ));
    assert!(package.opc.is_signed());
    assert!(package.custom_props().expect("read").is_empty());
    package = package.unsign().expect("explicit unsign");
    package
        .put_custom_props(owner())
        .expect("write custom properties");
    assert!(!package.opc.is_signed());

    std::sync::Arc::make_mut(&mut package.opc)
        .relate_to("_xmlsignatures/origin.sigs", rt::DIGITAL_SIGNATURE_ORIGIN);
    let props = package.custom_props().expect("read custom properties");
    package
        .put_custom_props(props)
        .expect("unchanged custom properties are a no-op");
    assert!(package.opc.is_signed());
}

#[test]
fn task_panes_facade_publishes_compact_source_checked_reversible_patches() {
    let package = Package::new().expect("new package");
    let mut edit = package.edit_task_panes().expect("edit task panes");
    assert!(
        edit.set(task_panes(), Conformance::Transitional)
            .expect("stage task panes")
    );
    let commit = edit.commit().expect("publish task panes");
    assert!(commit.diagnostics().changed());
    let inverse = commit.patch().inverse();
    let published = commit.into_package();
    let loaded = published
        .task_panes()
        .expect("load task panes")
        .expect("panes");
    assert_eq!(loaded.len(), 1);
    assert_eq!(
        loaded
            .get("{B1C15FE4-84FA-4773-AD36-9EF5444C5A01}")
            .expect("semantic add-in selector")
            .add_in()
            .reference()
            .store(),
        Store::Omex
    );
    for part in published
        .opc
        .try_iter_parts()
        .map(|part| part.expect("part payload decodes"))
        .filter(|part| part.partname().as_str().starts_with("/ppt/webextensions/"))
    {
        assert!(!part.blob().contains(&b'\n'));
        assert!(!part.blob().contains(&b'\r'));
    }
    assert!(package.task_panes().expect("source unchanged").is_none());

    let reverted = published
        .apply_task_panes(&inverse)
        .expect("revert task panes")
        .into_package();
    assert!(reverted.task_panes().expect("load removed panes").is_none());
}

#[test]
fn empty_task_pane_patch_preserves_signatures() {
    let mut package = Package::new().expect("new package");
    std::sync::Arc::make_mut(&mut package.opc)
        .relate_to("_xmlsignatures/origin.sigs", rt::DIGITAL_SIGNATURE_ORIGIN);
    assert!(package.opc.is_signed());

    let mut edit = package.edit_task_panes().expect("edit task panes");
    assert!(!edit.remove().expect("stage empty task-pane removal"));
    let commit = edit.commit().expect("commit empty removal");
    assert!(commit.patch().is_noop());
    assert!(commit.package().opc.is_signed());
    assert!(package.opc.is_signed());
}
