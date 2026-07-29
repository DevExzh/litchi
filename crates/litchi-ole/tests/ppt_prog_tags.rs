//! Fixture-level integration tests for typed document/slide programmable tags
//! (MS-PPT 2.4.23 and 2.5.19-2.5.22) reachable through `Presentation`, `Slide`,
//! `SpeakerNotes`, and the main-master accessor.

use litchi_ole::ppt::{
    Package, PowerPointProgBinaryTagVersion, PowerPointProgTagScope, PowerPointProgTags,
};
use std::path::PathBuf;

fn fixture(name: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/ole/ppt")
        .join(name)
}

fn poi_fixture(name: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/poi/test-data/slideshow")
        .join(name)
}

#[test]
fn presentation_exposes_document_prog_tags_with_typed_extensions() {
    let mut package = Package::open(fixture("SampleShow.ppt")).unwrap();
    let presentation = package.presentation().unwrap();

    let tags = presentation
        .programmable_tags()
        .unwrap()
        .expect("SampleShow.ppt carries a DocProgTagsContainer");
    assert_eq!(tags.scope, PowerPointProgTagScope::Document);

    // Every binary tag payload decodes as a strict record sequence, and the
    // container serializes back byte-for-byte.
    let limits = Default::default();
    let reparsed = PowerPointProgTags::parse(&tags.to_record(limits).unwrap(), tags.scope, limits)
        .unwrap();
    assert_eq!(reparsed, tags);

    let extensions = tags.document_extensions().unwrap();
    let pp9 = tags.binary_tag(PowerPointProgBinaryTagVersion::PowerPoint9);
    if let Some(pp9) = pp9 {
        let extension = extensions.powerpoint9.as_ref().unwrap();
        assert!(!extension.text_master_styles.is_empty());
        // Extension-level serialization reproduces the retained blob exactly.
        assert_eq!(
            extension.to_payload().unwrap(),
            pp9.payload,
        );
    }

    // Slide, notes, and main-master scopes parse without corruption.
    for slide in presentation.slides().unwrap() {
        if let Some(slide_tags) = slide.programmable_tags().unwrap() {
            assert_eq!(slide_tags.scope, PowerPointProgTagScope::Slide);
            slide_tags.slide_extensions().unwrap();
        }
        if let Some(notes) = slide.speaker_notes().unwrap() {
            if let Some(notes_tags) = notes.programmable_tags().unwrap() {
                assert_eq!(notes_tags.scope, PowerPointProgTagScope::Slide);
                notes_tags.slide_extensions().unwrap();
            }
        }
    }
    for master_tags in presentation.main_master_programmable_tags().unwrap() {
        assert_eq!(master_tags.scope, PowerPointProgTagScope::Slide);
        master_tags.slide_extensions().unwrap();
    }
}

#[test]
fn presentation_prog_tags_cover_poi_fixtures() {
    // A spread of real-world files: every ProgTags container in these scopes
    // must parse strictly and its versioned extensions must decode.
    for name in ["basic_test_ppt_file.ppt", "WithComments.ppt", "datetime.ppt"] {
        let mut package = Package::open(poi_fixture(name)).unwrap();
        let presentation = package.presentation().unwrap();
        if let Some(tags) = presentation.programmable_tags().unwrap() {
            assert_eq!(tags.scope, PowerPointProgTagScope::Document);
            tags.document_extensions().unwrap();
        }
        for slide in presentation.slides().unwrap() {
            if let Some(tags) = slide.programmable_tags().unwrap() {
                tags.slide_extensions().unwrap();
            }
        }
        presentation.main_master_programmable_tags().unwrap();
    }
}
