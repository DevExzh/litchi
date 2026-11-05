//! Scene splices and metadata edits through the primary `Package::edit`
//! transaction (record 0278): every changed byte range is a checked splice,
//! the reopened package reads the change back, and every byte outside the
//! changed ranges keeps its source.

use litchi_odp::Package;
use litchi_odp::scene::{Attribute, Change, ObjectRef, Scene, StoryRef};

fn fixture(name: &str) -> Vec<u8> {
    let path = format!("{}/../../test-data/{name}", env!("CARGO_MANIFEST_DIR"));
    std::fs::read(&path).unwrap_or_else(|error| panic!("{path}: {error}"))
}

fn scene(package: &Package) -> Scene {
    Scene::parse(package.presentation().unwrap().content_xml()).unwrap()
}

fn first_text_object(scene: &Scene) -> (usize, usize) {
    for (page_index, page) in scene.pages().iter().enumerate() {
        for (object_index, object) in page.objects().iter().enumerate() {
            if object
                .story()
                .is_some_and(|story| !story.paragraphs().is_empty())
            {
                return (page_index, object_index);
            }
        }
    }
    panic!("no text object");
}

#[test]
fn corpus_scenes_index_every_slide() {
    for name in [
        "odf/corpus/impress-basic.odp",
        "odf/corpus/impress-master-layouts.odp",
        "odf/corpus/impress-embedded-spreadsheet.odp",
        "odf/odp/cellspan.odp",
        "odf/odp/tdf102223.odp",
        "odf/odp/tdf105502.odp",
        "odf/odp/tdf169979.odp",
        "odf/odp/text-in-image.odp",
    ] {
        let package = Package::from_bytes(&fixture(name)).unwrap();
        let scene = scene(&package);
        assert_eq!(
            scene.pages().len(),
            package.presentation().unwrap().slide_count().unwrap(),
            "{name}"
        );
    }
}

#[test]
fn a_paragraph_set_round_trips_through_the_edit() {
    let source = Package::from_bytes(&fixture("odf/corpus/impress-basic.odp")).unwrap();
    let before = scene(&source);
    let (page, object) = first_text_object(&before);
    let story = StoryRef::object(ObjectRef::new(page, object));
    let mut edit = source.edit().unwrap();
    assert!(
        edit.splice_scene(&[Change::SetParagraphText {
            story: story.clone(),
            paragraph: 0,
            text: "Replaced  title".into(),
        }])
        .unwrap()
    );
    let commit = edit.commit().unwrap();
    assert!(commit.diagnostics().changed());
    let package = commit.into_package();
    let after = scene(&package);
    assert_eq!(
        after.story(&story).unwrap().paragraphs()[0].text(),
        "Replaced  title"
    );
    // The source is unchanged and the reopened bytes parse.
    assert_eq!(scene(&source), before);
    let reopened = Package::from_bytes(&package.to_bytes().unwrap()).unwrap();
    assert_eq!(scene(&reopened), after);
}

#[test]
fn exact_noops_return_the_source() {
    let bytes = fixture("odf/corpus/impress-basic.odp");
    let source = Package::from_bytes(&bytes).unwrap();
    let before = scene(&source);
    let (page, object) = first_text_object(&before);
    let story = StoryRef::object(ObjectRef::new(page, object));
    let text = before.story(&story).unwrap().text();
    let mut edit = source.edit().unwrap();
    assert!(
        !edit
            .splice_scene(&[Change::SetStoryText { story, text }])
            .unwrap()
    );
    let metadata = edit.metadata().unwrap();
    assert!(!edit.set_metadata(&metadata).unwrap());
    let commit = edit.commit().unwrap();
    assert!(!commit.diagnostics().changed());
    assert_eq!(commit.into_package().as_bytes(), bytes.as_slice());
}

#[test]
fn geometry_and_metadata_publish_in_one_commit() {
    let source = Package::from_bytes(&fixture("odf/corpus/impress-basic.odp")).unwrap();
    let mut edit = source.edit().unwrap();
    edit.splice_scene(&[Change::SetAttribute {
        object: ObjectRef::new(0, 0),
        attribute: Attribute::Width,
        value: Some("12.5cm".into()),
    }])
    .unwrap();
    let mut metadata = edit.metadata().unwrap();
    metadata.title = Some("A new title".into());
    assert!(edit.set_metadata(&metadata).unwrap());
    let package = edit.commit().unwrap().into_package();
    assert_eq!(
        scene(&package).pages()[0].objects()[0].width(),
        Some("12.5cm")
    );
    assert_eq!(
        package
            .presentation()
            .unwrap()
            .metadata()
            .unwrap()
            .title
            .as_deref(),
        Some("A new title")
    );
    // An unsupported metadata field is refused.
    let mut edit = package.edit().unwrap();
    let mut metadata = edit.metadata().unwrap();
    metadata.company = Some("Litchi".into());
    assert!(edit.set_metadata(&metadata).is_err());
}

#[test]
fn slides_insert_move_and_remove_by_splice() {
    let source = Package::from_bytes(&fixture("odf/corpus/impress-basic.odp")).unwrap();
    let count = scene(&source).pages().len();
    let mut edit = source.edit().unwrap();
    edit.splice_scene(&[Change::InsertPage {
        index: count,
        page: litchi_odp::scene::NewPage::default(),
    }])
    .unwrap();
    edit.splice_scene(&[Change::MovePage { page: count, to: 0 }])
        .unwrap();
    let package = edit.commit().unwrap().into_package();
    let after = scene(&package);
    assert_eq!(after.pages().len(), count + 1);
    assert!(after.pages()[0].objects().is_empty());
    assert_eq!(
        package.presentation().unwrap().slide_count().unwrap(),
        count + 1
    );
    let mut edit = package.edit().unwrap();
    edit.splice_scene(&[Change::RemovePage { page: 0 }])
        .unwrap();
    let restored = edit.commit().unwrap().into_package();
    assert_eq!(scene(&restored).pages().len(), count);
}
