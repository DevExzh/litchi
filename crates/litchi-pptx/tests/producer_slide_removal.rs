//! Source-backed slide removal on producer decks (record 0293).
//!
//! PowerPoint writes slides with extension metadata (`p14:creationId`),
//! hyperlinks and slide jumps, and presentation roots with guide extensions;
//! LibreOffice declares the markup-compatibility namespace on every root.
//! Removal deletes the slide Part with its relationship Part, so these
//! constructs are admitted, and the inverse patch restores the exact source.
//! Slides that own dependent Parts (notes, pictures, diagrams) stay typed
//! refusals that name the dependency.

use std::sync::Arc;

use litchi_core::{OwnedSource, Position};
use litchi_pptx::source::Package;
use litchi_pptx::{Error, SlideRemovalRefusal};

fn open(bytes: &'static [u8]) -> Package {
    Package::from_read_at(Arc::new(OwnedSource::new(bytes.to_vec()))).unwrap()
}

/// Removes the slide at `position` and proves the inverse restores the
/// source bytes.
fn remove_and_invert(bytes: &'static [u8], position: usize, remaining: usize) {
    let source = open(bytes);
    let mut edit = source.edit().unwrap();
    edit.remove_slide(Position::new(position)).unwrap();
    let (updated, patch) = edit.commit().unwrap().into_parts();
    assert_eq!(
        updated.presentation().unwrap().slide_count().unwrap(),
        remaining
    );
    let reopened =
        Package::from_read_at(Arc::new(OwnedSource::new(updated.to_bytes().unwrap()))).unwrap();
    assert_eq!(
        reopened.presentation().unwrap().slide_count().unwrap(),
        remaining
    );
    assert_eq!(
        updated.apply(&patch.inverse()).unwrap().to_bytes().unwrap(),
        source.to_bytes().unwrap()
    );
}

const SHAPES: &[u8] = include_bytes!("../../../test-data/ooxml/pptx/shapes.pptx");
const LIBREOFFICE_SHAPES: &[u8] =
    include_bytes!("../../../test-data/office-interop/libreoffice-resaved/shapes-litchi.pptx");

#[test]
fn slides_with_extension_metadata_are_removed_and_restored() {
    // Slide 6 holds only a layout edge and a `p14:creationId` extension.
    remove_and_invert(SHAPES, 5, 5);
}

#[test]
fn slides_with_hyperlinks_and_slide_jumps_are_removed_and_restored() {
    // Slide 5 links to a web page and jumps to another slide.
    remove_and_invert(SHAPES, 4, 5);
}

#[test]
fn a_bare_markup_compatibility_declaration_is_admitted() {
    // LibreOffice declares `xmlns:mc` on the presentation root.
    remove_and_invert(LIBREOFFICE_SHAPES, 5, 5);
}

#[test]
fn slides_owning_dependent_parts_are_typed_refusals() {
    // Slide 1 owns a picture.
    let source = open(SHAPES);
    let mut edit = source.edit().unwrap();
    let error = edit.remove_slide(Position::new(0)).unwrap_err();
    assert!(
        matches!(
            &error,
            Error::SlideRemovalPlan {
                kind: SlideRemovalRefusal::UnsupportedRelationship,
                detail,
            } if detail.contains("`image`")
        ),
        "{error:?}"
    );
}
