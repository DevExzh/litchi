use litchi_ooxml::pptx::Package;
use tempfile::NamedTempFile;

#[test]
fn duplicated_slide_takes_requested_position_after_round_trip() {
    let output = NamedTempFile::with_suffix(".pptx").unwrap();
    let mut package = Package::new().unwrap();
    {
        let presentation = package.presentation_mut().unwrap();
        presentation.add_slide().unwrap().set_title("one");
        presentation.add_slide().unwrap().set_title("two");
        presentation.add_slide().unwrap().set_title("three");

        let position = presentation.insert_duplicate_slide(0, 2).unwrap();
        assert_eq!(position, 2);
        assert_eq!(presentation.slide_count(), 4);
        // Failed insertions validate before duplicating anything.
        assert!(presentation.insert_duplicate_slide(9, 0).is_err());
        assert!(presentation.insert_duplicate_slide(0, 9).is_err());
        assert_eq!(presentation.slide_count(), 4);
    }
    package.save(output.path()).unwrap();

    let reopened = Package::open(output.path()).unwrap();
    let presentation = reopened.presentation().unwrap();
    // The copy received the next slide ID and sits at position 2.
    assert_eq!(presentation.slide_ids().unwrap(), [256, 257, 259, 258]);

    let slides = presentation.slides().unwrap();
    for (slide, expected) in slides.iter().zip(["one", "two", "one", "three"]) {
        let text = slide.text().unwrap();
        assert!(text.contains(expected), "expected '{expected}' in '{text}'");
    }
}

#[test]
fn insert_duplicate_at_end_matches_duplicate_slide() {
    let mut presentation = litchi_ooxml::pptx::MutablePresentation::new();
    presentation.add_slide().unwrap().set_title("only");

    let position = presentation.insert_duplicate_slide(0, 1).unwrap();
    assert_eq!(position, 1);
    assert_eq!(presentation.slide_count(), 2);
    assert_eq!(presentation.slides()[1].slide_id(), 257);
    assert_eq!(presentation.slide_mut(1).unwrap().title(), Some("only"));

    // Empty decks reject every index without panicking.
    let mut empty = litchi_ooxml::pptx::MutablePresentation::new();
    assert!(empty.insert_duplicate_slide(0, 0).is_err());
}
