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
fn floating_pictures_are_merged_by_their_image_store_entry_only() {
    // Floating anchors keep CPicLocation 0, which is also the Data-stream
    // offset of the first inline picture, and two consecutive anchors share
    // one run.
    let trigonometry = kinds("ole/doc/trigonometric_ratios.doc");
    assert_eq!(trigonometry.len(), 677);
    assert_eq!(count(&trigonometry, Kind::Png), 4);
    assert_eq!(count(&trigonometry, Kind::Emf), 1);
    // Sixteen inline pictures, and four floating shapes that show one
    // image-store entry.
    let floating = kinds("ole/doc/FloatingPictures.doc");
    assert_eq!(floating.len(), 17);
    assert_eq!(count(&floating, Kind::Png), 16);
}

#[test]
fn duplicated_style_names_do_not_hide_pictures() {
    // The stylesheet repeats a style name, which strict reading rejects.
    assert!(strict_open_fails("ole/doc/picture.doc"));
    assert_eq!(kinds("ole/doc/picture.doc"), [Kind::Emf]);
}

fn strict_open_fails(path: &str) -> bool {
    let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data")
        .join(path);
    litchi::doc::Package::open(path).is_err()
}
