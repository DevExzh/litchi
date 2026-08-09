#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]

use litchi_pptx::Package;
use tempfile::NamedTempFile;

#[test]
fn inserted_slides_keep_position_and_ids_after_round_trip() {
    let output = NamedTempFile::with_suffix(".pptx").unwrap();
    let mut package = Package::new().unwrap();
    {
        let presentation = package.presentation_mut().unwrap();
        presentation.add_slide().unwrap().set_title("one");
        presentation.add_slide().unwrap().set_title("two");
        presentation.add_slide().unwrap().set_title("three");

        presentation.insert_slide(1).unwrap().set_title("middle");
        presentation.insert_slide(0).unwrap().set_title("start");
        // Appending through insert is equivalent to add_slide.
        presentation.insert_slide(5).unwrap().set_title("end");

        assert!(presentation.insert_slide(7).is_err());
        assert_eq!(presentation.slide_count(), 6);
        assert_eq!(
            presentation
                .slides()
                .iter()
                .map(litchi_pptx::MutableSlide::slide_id)
                .collect::<Vec<_>>(),
            [260, 256, 259, 257, 258, 261]
        );
    }
    package.save(output.path()).unwrap();

    let reopened = Package::open(output.path()).unwrap();
    let presentation = reopened.presentation().unwrap();
    assert_eq!(
        presentation
            .slide_references()
            .unwrap()
            .iter()
            .map(litchi_pptx::parts::SlideReference::id)
            .collect::<Vec<_>>(),
        [260, 256, 259, 257, 258, 261]
    );

    let slides = presentation.slides().unwrap();
    assert_eq!(slides.len(), 6);
    for (slide, expected) in slides
        .iter()
        .zip(["start", "one", "middle", "two", "three", "end"])
    {
        let text = slide.text().unwrap();
        assert!(text.contains(expected), "expected '{expected}' in '{text}'");
    }
}

#[test]
fn insert_slide_into_empty_presentation_is_append() {
    let mut presentation = litchi_pptx::MutablePresentation::new();
    presentation.insert_slide(0).unwrap().set_title("only");
    assert!(presentation.insert_slide(2).is_err());
    assert_eq!(presentation.slide_count(), 1);
    assert_eq!(presentation.slides()[0].slide_id(), 256);
}
