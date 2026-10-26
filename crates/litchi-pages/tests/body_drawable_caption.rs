//! Body-drawable title and caption transactions
//! (`litchi_pages::drawable::caption`, record 0371).

use litchi_pages::drawable::caption::{Error, Kind, Refusal};
use litchi_pages::{DrawableSelector, Package};

const IMAGES: &[u8] =
    include_bytes!("../../../test-data/iwork/pages/image-adjustments-native.pages");
const AUDIO: &[u8] = include_bytes!("../../../test-data/iwork/pages/audio-playback-native.pages");

fn bytes(package: &Package) -> Vec<u8> {
    let mut output = Vec::new();
    package.write_to(&mut output).unwrap();
    output
}

#[test]
fn image_title_is_created_replaced_cleared_and_inverted() {
    let image = DrawableSelector::image(0);
    let package = Package::from_bytes(IMAGES).unwrap();
    assert_eq!(
        package.body_drawable_caption(image, Kind::Title).unwrap(),
        None
    );

    let mut edit = package
        .edit_body_drawable_caption(image, Kind::Title)
        .unwrap();
    assert_eq!(edit.before(), None);
    edit.set("Quarterly portrait").unwrap();
    let created = edit.commit().unwrap();
    assert!(created.diagnostics().changed());
    assert_eq!(created.diagnostics().created_objects(), 4);
    let titled = created.package();
    assert_eq!(
        titled
            .body_drawable_caption(image, Kind::Title)
            .unwrap()
            .as_deref(),
        Some("Quarterly portrait")
    );
    assert_eq!(
        titled.body_drawable_caption(image, Kind::Caption).unwrap(),
        None
    );
    let restored = titled
        .apply_body_drawable_caption(&created.patch().inverse())
        .unwrap();
    assert_eq!(bytes(restored.package()), IMAGES);

    let mut edit = titled
        .edit_body_drawable_caption(image, Kind::Title)
        .unwrap();
    edit.set("Updated").unwrap();
    let replaced = edit.commit().unwrap();
    assert_eq!(replaced.diagnostics().created_objects(), 0);
    assert_eq!(replaced.patch().before(), Some("Quarterly portrait"));
    assert_eq!(replaced.patch().after(), Some("Updated"));

    let mut edit = replaced
        .package()
        .edit_body_drawable_caption(image, Kind::Title)
        .unwrap();
    edit.clear().unwrap();
    let cleared = edit.commit().unwrap();
    assert_eq!(cleared.diagnostics().created_objects(), 1);
    assert_eq!(
        cleared
            .package()
            .body_drawable_caption(image, Kind::Title)
            .unwrap(),
        None
    );
}

#[test]
fn clearing_an_absent_label_changes_nothing() {
    let package = Package::from_bytes(IMAGES).unwrap();
    let mut edit = package
        .edit_body_drawable_caption(DrawableSelector::image(0), Kind::Caption)
        .unwrap();
    edit.clear().unwrap();
    let commit = edit.commit().unwrap();
    assert!(!commit.diagnostics().changed());
    assert!(commit.patch().is_noop());
    assert_eq!(bytes(commit.package()), IMAGES);
}

#[test]
fn refusals_are_typed() {
    let package = Package::from_bytes(IMAGES).unwrap();
    let mut edit = package
        .edit_body_drawable_caption(DrawableSelector::image(0), Kind::Caption)
        .unwrap();
    assert!(matches!(
        edit.set("a\u{fffc}b"),
        Err(Error::Refused(Refusal::ReservedMarker))
    ));
    edit.set("one").unwrap();
    assert!(matches!(
        edit.clear(),
        Err(Error::Refused(Refusal::OperationAlreadyStaged))
    ));
    assert!(matches!(
        package.body_drawable_caption(DrawableSelector::image(9), Kind::Title),
        Err(Error::Refused(Refusal::NotFound { .. }))
    ));
    let audio = Package::from_bytes(AUDIO).unwrap();
    assert!(matches!(
        audio.body_drawable_caption(DrawableSelector::audio(0), Kind::Title),
        Err(Error::Refused(Refusal::NoLabel { .. }))
    ));
}
