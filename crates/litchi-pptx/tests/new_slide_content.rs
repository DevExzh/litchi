//! Record 0252: slides inserted from a layout instantiate its placeholders,
//! and notes text creates a notes page when the slide has none, through both
//! the eager and the source-backed primary edits.
#![allow(clippy::unwrap_used, clippy::expect_used, reason = "test assertions")]
use std::sync::Arc;

use litchi_core::{OwnedSource, Position};
use litchi_pptx::master_layout::{PlaceholderKind, PlaceholderSpec, SlideLayoutKind};
use litchi_pptx::{Error, Package};

/// A one-slide deck with an extra "Agenda Layout" layout holding a title,
/// a body (idx 1) and a footer, which new slides do not instantiate.
fn deck() -> Package {
    let mut builder = litchi_pptx::builder::Builder::new().unwrap();
    builder.presentation().unwrap().add_slide().unwrap();
    let package = builder.build().unwrap();
    let master = package.presentation().unwrap().slide_masters().unwrap()[0]
        .part()
        .part()
        .partname()
        .clone();
    let mut edit = package.edit().unwrap();
    edit.add_slide_layout(
        &master,
        SlideLayoutKind::Object,
        "Agenda Layout",
        &[
            PlaceholderSpec::new(PlaceholderKind::Title),
            PlaceholderSpec::new(PlaceholderKind::Body).with_index(1),
            PlaceholderSpec::new(PlaceholderKind::Footer).with_index(11),
        ],
    )
    .unwrap();
    edit.commit().unwrap().into_package()
}

fn placeholders(xml: &[u8]) -> Vec<(Option<String>, u32, Option<String>)> {
    let scene = litchi_pptx::shape::Scene::read(xml, litchi_pptx::shape::Limits::DEFAULT).unwrap();
    scene
        .iter()
        .map(|shape| {
            let placeholder = shape.placeholder().expect("placeholder");
            (
                placeholder.kind().map(str::to_owned),
                placeholder.index(),
                shape.name().map(str::to_owned),
            )
        })
        .collect()
}

fn expected() -> Vec<(Option<String>, u32, Option<String>)> {
    vec![
        (Some("title".into()), 0, Some("Title 1".into())),
        (Some("body".into()), 1, Some("Text Placeholder 2".into())),
    ]
}

fn slide_xml(package: &Package, position: usize) -> Vec<u8> {
    let mut edit = package.edit().unwrap();
    let part = edit.slides().unwrap()[position].part_name().clone();
    package
        .opc()
        .unwrap()
        .get_part(&part)
        .unwrap()
        .blob()
        .to_vec()
}

#[test]
fn eager_insertion_instantiates_layout_placeholders() {
    let source = deck();
    let mut edit = source.edit().unwrap();
    edit.insert_slide(Position::new(1), "Agenda Layout")
        .unwrap();
    let package = edit.commit().unwrap().into_package();
    assert_eq!(placeholders(&slide_xml(&package, 1)), expected());
    // An empty placeholder holds no run, so the run-slot text verb refuses
    // it; the text-body verb fills it.
    let mut edit = package.edit().unwrap();
    assert!(edit.set_shape_text(1usize, 0usize, "Agenda").is_err());
    let title = litchi_pptx::opened::TextAssignment::new(
        litchi_pptx::opened::TextTarget::Shape {
            shape: Position::new(0),
        },
        litchi_pptx::opened::TextProperties::default().with_text("Agenda"),
    );
    assert_eq!(edit.set_text_properties(1usize, &[title]).unwrap(), 1);
    let package = edit.commit().unwrap().into_package();
    let reopened = Package::from_bytes(&package.to_bytes().unwrap()).unwrap();
    let scene_xml = slide_xml(&reopened, 1);
    let scene =
        litchi_pptx::shape::Scene::read(&scene_xml, litchi_pptx::shape::Limits::DEFAULT).unwrap();
    assert_eq!(scene.at(0).unwrap().text(), Some("Agenda"));
}

#[test]
fn source_insertion_matches_the_eager_slide() {
    let bytes = deck().to_bytes().unwrap();
    let source =
        litchi_pptx::source::Package::from_read_at(Arc::new(OwnedSource::new(bytes.clone())))
            .unwrap();
    let mut edit = source.edit().unwrap();
    edit.insert_slide(Position::new(1), "Agenda Layout")
        .unwrap();
    let (after, patch) = edit.commit().unwrap().into_parts();
    let reopened = Package::from_bytes(&after.to_bytes().unwrap()).unwrap();
    assert_eq!(placeholders(&slide_xml(&reopened, 1)), expected());
    assert_eq!(
        after
            .apply(&patch.inverse())
            .unwrap()
            .into_package()
            .to_bytes()
            .unwrap(),
        source.to_bytes().unwrap()
    );
}

#[test]
fn notes_text_creates_and_replaces_a_notes_page() {
    let package = deck();
    let mut edit = package.edit().unwrap();
    // Empty text on a slide without notes changes nothing.
    assert!(!edit.set_notes_text(0usize, "").unwrap());
    assert!(edit.set_notes_text(0usize, "Speaker note").unwrap());
    let created = edit.commit().unwrap().into_package();
    let reopened = Package::from_bytes(&created.to_bytes().unwrap()).unwrap();
    let source = litchi_pptx::source::Package::from_read_at(Arc::new(OwnedSource::new(
        reopened.to_bytes().unwrap(),
    )))
    .unwrap();
    let notes = source
        .presentation()
        .unwrap()
        .slide(Position::new(0))
        .unwrap()
        .notes()
        .unwrap()
        .expect("notes page");
    assert_eq!(notes.text(), Some("Speaker note"));
    let mut edit = reopened.edit().unwrap();
    assert!(edit.set_notes_text(0usize, "Replaced").unwrap());
    assert!(!edit.set_notes_text(0usize, "Replaced").unwrap());
    let replaced = edit.commit().unwrap().into_package();
    assert_eq!(
        replaced
            .opc()
            .unwrap()
            .iter_parts()
            .filter(|part| part.partname().as_str().starts_with("/ppt/notesSlides/"))
            .count(),
        1
    );
}

#[test]
fn notes_refusals_leave_the_edit_usable() {
    let package = deck();
    let mut edit = package.edit().unwrap();
    assert!(edit.set_notes_text(0usize, "ok").is_ok());
    // An out-of-range slide is a typed refusal and leaves the edit usable.
    assert!(matches!(
        edit.set_notes_text(9usize, "x"),
        Err(Error::SlideIndexOutOfBounds { .. })
    ));
    assert!(edit.commit().unwrap().diagnostics().changed());
}
