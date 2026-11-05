//! Body text-box character and paragraph properties
//! (`litchi_pages::drawable::text_box_text`, record 0413). The host
//! differential cases ran in the retired host's parity harness
//! (`pages_text_box_cases.rs` at commit `4d8a3e161f`); these pin the leaf
//! refusals.

use litchi_pages::drawable::text_box_text::{Error, PropertyKind, Refusal};
use litchi_pages::{DrawableSelector, Package};

const BASIC: &[u8] = include_bytes!("../../../test-data/iwork/pages/basic.pages");
const IMAGES: &[u8] =
    include_bytes!("../../../test-data/iwork/pages/image-adjustments-native.pages");

#[test]
fn a_missing_text_box_is_not_found() {
    let package = Package::from_bytes(BASIC).unwrap();
    let selector = DrawableSelector::text_box(0);
    assert_eq!(
        package.edit_body_text_box_text(selector).err(),
        Some(Error::Refused(Refusal::DrawableNotFound { selector }))
    );
    assert_eq!(
        package
            .body_text_box_text(selector, PropertyKind::Alignment)
            .err(),
        Some(Error::Refused(Refusal::DrawableNotFound { selector }))
    );
}

#[test]
fn media_owns_no_text() {
    let package = Package::from_bytes(IMAGES).unwrap();
    let error = package
        .edit_body_text_box_text(DrawableSelector::image(0))
        .err()
        .unwrap();
    assert_eq!(error, Error::Refused(Refusal::NoText { kind: "image" }));
    let Error::Refused(refusal) = error else {
        unreachable!()
    };
    assert_eq!(
        refusal.kind(),
        litchi_core::edit::RefusalKind::UnsupportedTarget
    );
}

#[test]
fn lists_and_named_styles_refuse_media() {
    let package = Package::from_bytes(IMAGES).unwrap();
    let image = DrawableSelector::image(0);
    assert_eq!(
        package.edit_body_text_box_lists(image).err(),
        Some(litchi_pages::drawable::list::Error::Refused(
            litchi_pages::drawable::list::Refusal::NoText { kind: "image" }
        ))
    );
    assert_eq!(
        package.body_text_box_named_styles(image).err(),
        Some(litchi_pages::drawable::named_style::Error::Refused(
            litchi_pages::drawable::named_style::Refusal::NoText { kind: "image" }
        ))
    );
}
