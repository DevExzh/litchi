//! Source-backed slide removal on producer decks (record 0293).
//!
//! PowerPoint writes slides with extension metadata (`p14:creationId`),
//! hyperlinks and slide jumps, and presentation roots with guide extensions;
//! LibreOffice declares the markup-compatibility namespace on every root.
//! Removal deletes the slide Part with its relationship Part, so these
//! constructs are admitted, and the inverse patch restores the exact source.
//! A slide's exclusively owned Parts (notes, pictures, diagrams, comments,
//! VML drawings, embedded objects) leave with it, and the inverse restores
//! them (record 0303).

use std::sync::Arc;

use litchi_core::{OwnedSource, Position};
use litchi_pptx::source::Package;
use litchi_pptx::{Error, SlideRemovalRefusal};
use soapberry_zip::office::{ArchiveReader, StreamingArchiveWriter};

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
        updated
            .apply(&patch.inverse())
            .unwrap()
            .into_package()
            .to_bytes()
            .unwrap(),
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

/// Removes the slide at `position`, proves the Parts in `owned` leave with
/// it and the Parts in `kept` stay, and proves that both the local inverse
/// and the durable inverse replayed on the reopened output restore the
/// source bytes and Parts (record 0303).
fn remove_owned_and_invert(
    bytes: &'static [u8],
    position: usize,
    remaining: usize,
    owned: &[&str],
    kept: &[&str],
) {
    let source = open(bytes);
    let mut edit = source.edit().unwrap();
    edit.remove_slide(Position::new(position)).unwrap();
    let (updated, patch) = edit.commit().unwrap().into_parts();
    let output = updated.to_bytes().unwrap();
    let names = member_names(&output);
    for part in owned {
        assert!(!names.iter().any(|name| name == part), "{part} stayed");
    }
    for part in kept {
        assert!(names.iter().any(|name| name == part), "{part} was removed");
    }
    let reopened = Package::from_read_at(Arc::new(OwnedSource::new(output))).unwrap();
    assert_eq!(
        reopened.presentation().unwrap().slide_count().unwrap(),
        remaining
    );
    assert_eq!(
        updated
            .apply(&patch.inverse())
            .unwrap()
            .into_package()
            .to_bytes()
            .unwrap(),
        source.to_bytes().unwrap()
    );
    // The durable inverse carries the removed Parts' exact images.
    let wire = patch.inverse().to_bytes().unwrap();
    let durable = litchi_pptx::source::Patch::from_bytes(&wire).unwrap();
    let restored = reopened
        .apply(&durable)
        .unwrap()
        .into_package()
        .to_bytes()
        .unwrap();
    let restored_names = member_names(&restored);
    for part in owned.iter().chain(kept) {
        assert!(
            restored_names.iter().any(|name| name == part),
            "{part} was not restored"
        );
        assert_eq!(member(&restored, part), member(bytes, part), "{part}");
    }
    let restored = Package::from_read_at(Arc::new(OwnedSource::new(restored))).unwrap();
    assert_eq!(
        restored.presentation().unwrap().slide_count().unwrap(),
        remaining + 1
    );
    // The forward patch replays on a fresh open of the source.
    let forward = litchi_pptx::source::Patch::from_bytes(&patch.to_bytes().unwrap()).unwrap();
    let replayed = open(bytes)
        .apply(&forward)
        .unwrap()
        .into_package()
        .to_bytes()
        .unwrap();
    assert_eq!(member_names(&replayed), names);
}

fn member_names(bytes: &[u8]) -> Vec<String> {
    let archive =
        ArchiveReader::new(bytes, soapberry_zip::office::ArchiveLimits::default()).unwrap();
    archive.file_names().map(|name| name.to_string()).collect()
}

fn member(bytes: &[u8], name: &str) -> Vec<u8> {
    ArchiveReader::new(bytes, soapberry_zip::office::ArchiveLimits::default())
        .unwrap()
        .read(name)
        .unwrap()
}

const SAMPLE: &[u8] = include_bytes!("../../../test-data/ooxml/pptx/sample.pptx");
const COMMENTS: &[u8] =
    include_bytes!("../../../test-data/poi/test-data/slideshow/45545_Comment.pptx");
const SMARTART: &[u8] = include_bytes!("../../../test-data/ooxml/pptx/smartart-autoTxRot.pptx");

#[test]
fn a_slide_and_its_picture_leave_together() {
    // Slide 1 owns `image1`; no other slide shows it.
    remove_owned_and_invert(
        SHAPES,
        0,
        5,
        &["ppt/slides/slide1.xml", "ppt/media/image1.jpg"],
        &["ppt/slideLayouts/slideLayout1.xml"],
    );
}

#[test]
fn a_slide_and_its_notes_slide_leave_together() {
    // The notes slide names its slide back; the notes master stays.
    remove_owned_and_invert(
        SAMPLE,
        0,
        1,
        &["ppt/slides/slide1.xml", "ppt/notesSlides/notesSlide1.xml"],
        &[
            "ppt/notesMasters/notesMaster1.xml",
            "ppt/notesSlides/notesSlide2.xml",
        ],
    );
}

#[test]
fn comments_vml_and_ole_objects_leave_with_their_slide() {
    remove_owned_and_invert(
        COMMENTS,
        0,
        10,
        &[
            "ppt/slides/slide1.xml",
            "ppt/notesSlides/notesSlide1.xml",
            "ppt/comments/comment1.xml",
            "ppt/drawings/vmlDrawing1.vml",
            "ppt/embeddings/oleObject1.bin",
        ],
        &["ppt/commentAuthors.xml"],
    );
}

#[test]
fn a_slide_whose_edge_names_a_missing_part_is_a_typed_refusal() {
    // The SmartArt slide names a diagram drawing its producer never wrote.
    // Removal itself would be safe, but a durable inverse could not restore
    // an edge to a missing Part after reopening.
    let source = open(SMARTART);
    let mut edit = source.edit().unwrap();
    let error = edit.remove_slide(Position::new(0)).unwrap_err();
    assert!(
        matches!(
            &error,
            Error::Refused(refusal) if refusal.kind() == SlideRemovalRefusal::AmbiguousTopology.kind()
        ),
        "{error:?}"
    );
}

/// `shapes.pptx` with slide 3 also showing slide 1's picture.
fn shapes_with_a_shared_picture() -> &'static [u8] {
    let archive =
        ArchiveReader::new(SHAPES, soapberry_zip::office::ArchiveLimits::default()).unwrap();
    let mut writer =
        StreamingArchiveWriter::new(soapberry_zip::office::StreamingArchiveLimits::default());
    for name in archive.file_names() {
        let mut bytes = archive.read(name).unwrap();
        if name == "ppt/slides/_rels/slide3.xml.rels" {
            let text = String::from_utf8(bytes).unwrap().replace(
                "</Relationships>",
                "<Relationship Id=\"rId9\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"../media/image1.jpg\"/></Relationships>",
            );
            bytes = text.into_bytes();
        }
        writer.write_stored(name, &bytes).unwrap();
    }
    Vec::leak(writer.finish().unwrap().into_inner())
}

#[test]
fn a_picture_another_slide_still_shows_stays() {
    let shared = shapes_with_a_shared_picture();
    remove_owned_and_invert(
        shared,
        0,
        5,
        &["ppt/slides/slide1.xml"],
        &["ppt/media/image1.jpg"],
    );
}

#[test]
fn a_picture_shared_by_two_removed_slides_leaves_with_them() {
    let shared = shapes_with_a_shared_picture();
    let source = open(shared);
    let mut edit = source.edit().unwrap();
    edit.remove_slide(Position::new(0)).unwrap();
    // Slide 3 is now at position 1, and the picture is its alone.
    edit.remove_slide(Position::new(1)).unwrap();
    let (updated, patch) = edit.commit().unwrap().into_parts();
    let output = updated.to_bytes().unwrap();
    let names = member_names(&output);
    assert!(!names.iter().any(|name| name == "ppt/media/image1.jpg"));
    assert_eq!(
        updated
            .apply(&patch.inverse())
            .unwrap()
            .into_package()
            .to_bytes()
            .unwrap(),
        source.to_bytes().unwrap()
    );
    // The composed patch proves joint ownership on a fresh source and
    // its durable inverse restores the picture after reopening.
    let forward = litchi_pptx::source::Patch::from_bytes(&patch.to_bytes().unwrap()).unwrap();
    let replayed = open(shared).apply(&forward).unwrap().into_package();
    assert_eq!(member_names(&replayed.to_bytes().unwrap()), names);
    let reopened = Package::from_read_at(Arc::new(OwnedSource::new(output))).unwrap();
    let inverse =
        litchi_pptx::source::Patch::from_bytes(&patch.inverse().to_bytes().unwrap()).unwrap();
    let restored = reopened
        .apply(&inverse)
        .unwrap()
        .into_package()
        .to_bytes()
        .unwrap();
    assert_eq!(
        member(&restored, "ppt/media/image1.jpg"),
        member(shared, "ppt/media/image1.jpg")
    );
}

#[test]
fn a_slide_jump_into_the_removed_slide_is_still_refused() {
    // Slide 5 jumps to slide 2 (position 1); removing slide 2 would leave
    // the jump dangling.
    let source = open(SHAPES);
    let mut edit = source.edit().unwrap();
    let error = edit.remove_slide(Position::new(1)).unwrap_err();
    assert!(
        matches!(
            error,
            Error::Refused(refusal) if refusal.kind() == SlideRemovalRefusal::SharedOwner.kind()
        ),
        "{error:?}"
    );
}

#[test]
fn the_eager_editor_removes_owned_parts_too() {
    // The eager owner (created or promoted packages) removes the same
    // closure as the source-backed one.
    let package = litchi_pptx::Package::from_reader(std::io::Cursor::new(COMMENTS)).unwrap();
    let mut edit = package.edit().unwrap();
    edit.remove_slide(0_usize).unwrap();
    let (updated, patch) = edit.commit().unwrap().into_parts();
    let output = updated.to_bytes().unwrap();
    let names = member_names(&output);
    for part in [
        "ppt/notesSlides/notesSlide1.xml",
        "ppt/comments/comment1.xml",
        "ppt/drawings/vmlDrawing1.vml",
        "ppt/embeddings/oleObject1.bin",
    ] {
        assert!(!names.iter().any(|name| name == part), "{part} stayed");
    }
    assert!(names.iter().any(|name| name == "ppt/commentAuthors.xml"));
    let restored = updated.apply(&patch.inverse()).unwrap().into_package();
    let restored_names = member_names(&restored.to_bytes().unwrap());
    assert!(
        restored_names
            .iter()
            .any(|name| name == "ppt/embeddings/oleObject1.bin")
    );
    assert_eq!(restored.presentation().unwrap().slide_count().unwrap(), 11);
}
