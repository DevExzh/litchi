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
    let mut package = Package::new().expect("new package");
    {
        let presentation = package.presentation_mut().expect("mutable presentation");
        let slide = presentation.add_slide().expect("slide");
        slide.set_title("Canonical owner");
        slide.add_text_box("Hello & goodbye", 914_400, 914_400, 2_743_200, 914_400);
        presentation.set_widescreen_slide_size();
    }

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
fn opened_package_refuses_unsafe_mutable_hydration() {
    let mut package = Package::new().expect("new package");
    let bytes = package.to_bytes().expect("serialize package");
    let mut opened = Package::from_bytes(&bytes).expect("reopen package");
    assert!(opened.presentation_mut().is_err());
}

#[test]
fn failed_typed_edit_restores_pending_presentation_state() {
    let mut package = Package::new().expect("new package");
    let presentation_part = PackURI::new("/ppt/presentation.xml").expect("part URI");
    let before = package
        .opc
        .get_part(&presentation_part)
        .expect("presentation part")
        .blob()
        .to_vec();

    {
        let presentation = package.presentation_mut().expect("mutable presentation");
        presentation
            .add_slide()
            .expect("pending slide")
            .set_title("must remain pending");
    }
    package.opc.rels_mut().add_relationship(
        rt::CUSTOM_PROPERTIES.to_owned(),
        "https://example.invalid/custom.xml".to_owned(),
        "rIdInvalidCustom".to_owned(),
        true,
    );

    let mut props = Props::new();
    props.insert("Owner", "Alice").expect("property");
    assert!(package.put_custom_props(props).is_err());
    assert_eq!(
        package
            .opc
            .get_part(&presentation_part)
            .expect("presentation part")
            .blob(),
        before.as_slice()
    );
    assert!(package.mutable_pres.is_some());

    let _ = package.opc.rels_mut().remove("rIdInvalidCustom");
    let bytes = package.to_bytes().expect("flush restored presentation");
    let reopened = Package::from_bytes(&bytes).expect("reopen package");
    assert_eq!(
        reopened
            .presentation()
            .expect("presentation")
            .slide_count()
            .expect("slide count"),
        1
    );
}

#[test]
fn custom_property_no_ops_preserve_signatures_and_changes_invalidate_them() {
    let mut package = Package::new().expect("new package");
    package
        .opc
        .relate_to("_xmlsignatures/origin.sigs", rt::DIGITAL_SIGNATURE_ORIGIN);
    assert!(package.opc.is_signed());

    package
        .put_custom_props(Props::new())
        .expect("empty custom properties are a no-op");
    assert!(package.opc.is_signed());

    let mut props = Props::new();
    props.insert("Owner", "Alice").expect("property");
    package
        .put_custom_props(props)
        .expect("write custom properties");
    assert!(!package.opc.is_signed());

    package
        .opc
        .relate_to("_xmlsignatures/origin.sigs", rt::DIGITAL_SIGNATURE_ORIGIN);
    let props = package.custom_props().expect("read custom properties");
    package
        .put_custom_props(props)
        .expect("unchanged custom properties are a no-op");
    assert!(package.opc.is_signed());
}

#[test]
fn task_panes_facade_publishes_compact_source_checked_reversible_patches() {
    let mut package = Package::new().expect("new package");
    let patch = package
        .task_panes_patch(task_panes(), Conformance::Transitional)
        .expect("plan task panes");
    assert!(!patch.is_empty());
    let inverse = patch.inverse();

    assert!(
        package
            .apply_task_panes_patch(&patch)
            .expect("publish task panes")
    );
    let loaded = package
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
    for part in package
        .opc
        .iter_parts()
        .filter(|part| part.partname().as_str().starts_with("/ppt/webextensions/"))
    {
        assert!(!part.blob().contains(&b'\n'));
        assert!(!part.blob().contains(&b'\r'));
    }

    assert!(
        package
            .apply_task_panes_patch(&inverse)
            .expect("revert task panes")
    );
    assert!(package.task_panes().expect("load removed panes").is_none());
}

#[test]
fn empty_task_pane_patch_preserves_signatures() {
    let mut package = Package::new().expect("new package");
    package
        .opc
        .relate_to("_xmlsignatures/origin.sigs", rt::DIGITAL_SIGNATURE_ORIGIN);
    assert!(package.opc.is_signed());

    let patch = package
        .clear_task_panes_patch()
        .expect("plan empty task-pane removal");
    assert!(patch.is_empty());
    assert!(
        !package
            .apply_task_panes_patch(&patch)
            .expect("apply empty patch")
    );
    assert!(package.opc.is_signed());
}
