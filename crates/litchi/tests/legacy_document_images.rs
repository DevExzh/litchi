//! Picture discovery through `litchi::images::doc` on stored Word documents.
#![cfg(all(feature = "doc", feature = "images"))]

use litchi::images::art::Kind;

fn kinds(path: &str) -> Vec<Kind> {
    let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data")
        .join(path);
    litchi::images::doc(&path)
        .unwrap_or_else(|error| panic!("{}: {error}", path.display()))
        .iter()
        .map(litchi::images::File::kind)
        .collect()
}

fn count(kinds: &[Kind], kind: Kind) -> usize {
    kinds.iter().filter(|candidate| **candidate == kind).count()
}

#[test]
fn picture_frames_without_stored_pictures_are_skipped() {
    // Each document holds an inline picture frame whose shape has no BLIP.
    let vector = kinds("microsoft-office/word/word-vector-diagram.doc");
    assert_eq!(vector, [Kind::Png, Kind::Png]);
    let pictures = kinds("ole/doc/testPictures.doc");
    assert_eq!(pictures.len(), 6);
    assert_eq!(count(&pictures, Kind::Wmf), 1);
}

#[test]
fn floating_pictures_are_not_merged_by_their_shared_picture_location() {
    // Floating anchors keep CPicLocation 0, which is also the Data-stream
    // offset of the first inline picture.
    let trigonometry = kinds("ole/doc/trigonometric_ratios.doc");
    assert_eq!(trigonometry.len(), 675);
    assert_eq!(count(&trigonometry, Kind::Png), 3);
    let floating = kinds("ole/doc/FloatingPictures.doc");
    assert_eq!(floating.len(), 20);
    assert_eq!(count(&floating, Kind::Png), 19);
}
