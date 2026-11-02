//! Body-media lifecycle transactions (`litchi_pages::drawable::lifecycle`,
//! record 0371).

use litchi_pages::drawable::lifecycle::{Error, Refusal};
use litchi_pages::{DrawableSelector, Package};

const IMAGES: &[u8] =
    include_bytes!("../../../test-data/iwork/pages/image-adjustments-native.pages");

fn bytes(package: &Package) -> Vec<u8> {
    let mut output = Vec::new();
    package.write_to(&mut output).unwrap();
    output
}

#[test]
fn native_image_removal_deletes_the_graph_and_assets_and_inverts() {
    let package = Package::from_bytes(IMAGES).unwrap();
    let mut edit = package.edit_body_media();
    edit.remove(DrawableSelector::image(0)).unwrap();
    let commit = edit.commit().unwrap();
    assert!(commit.diagnostics().changed());
    assert_eq!(commit.diagnostics().removed_objects(), 4);
    assert_eq!(commit.diagnostics().removed_assets(), 4);
    assert!(
        commit
            .package()
            .body_drawable_properties(DrawableSelector::image(0))
            .is_err()
    );
    let restored = commit
        .package()
        .apply_body_media(&commit.patch().inverse())
        .unwrap();
    assert_eq!(bytes(restored.package()), IMAGES);
}

#[test]
fn refusals_are_typed() {
    let package = Package::from_bytes(IMAGES).unwrap();
    // Shapes and text boxes are removable (record 0413); duplication is
    // still media-only.
    let mut edit = package.edit_body_media();
    assert!(matches!(
        edit.duplicate(DrawableSelector::shape(0), litchi_pages::TextPosition::ZERO),
        Err(Error::Refused(Refusal::NotMedia { .. }))
    ));
    let mut edit = package.edit_body_media();
    edit.remove(DrawableSelector::text_box(0)).unwrap();
    assert!(matches!(
        edit.commit(),
        Err(Error::Refused(Refusal::NotFound { .. }))
    ));
    assert!(matches!(
        package.edit_body_media().commit(),
        Err(Error::Refused(Refusal::NoStagedOperation))
    ));
    let mut edit = package.edit_body_media();
    edit.remove(DrawableSelector::image(0)).unwrap();
    assert!(matches!(
        edit.remove(DrawableSelector::image(0)),
        Err(Error::Refused(Refusal::OperationAlreadyStaged))
    ));
    let mut edit = package.edit_body_media();
    edit.remove(DrawableSelector::movie(3)).unwrap();
    assert!(matches!(
        edit.commit(),
        Err(Error::Refused(Refusal::NotFound { .. }))
    ));
}
