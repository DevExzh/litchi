//! Regression tests for the PresentationML package facade.

use super::Package;

#[test]
fn new_writer_round_trips_the_bounded_slide_graph() {
    let mut package = Package::new().expect("new package");
    {
        let presentation = package.presentation_mut().expect("mutable presentation");
        let slide = presentation.add_slide().expect("slide");
        slide.set_title("Canonical owner");
        slide.add_text_box("Hello & goodbye", 914_400, 914_400, 2_743_200, 914_400);
        presentation.set_widescreen_slide_size();
    }

    let bytes = package.to_bytes().expect("serialize package");
    let reopened = Package::from_bytes(&bytes).expect("reopen package");
    let presentation = reopened.presentation().expect("presentation");
    assert_eq!(presentation.slide_count().expect("slide count"), 1);
    assert_eq!(
        presentation.slide_size().expect("slide size"),
        (9_144_000, 5_143_500)
    );
    let slide = presentation.slide(0).expect("slide lookup").expect("slide");
    assert_eq!(slide.name().expect("slide name"), "Slide 256");
    assert!(
        slide
            .text()
            .expect("slide text")
            .contains("Hello & goodbye")
    );
    assert_eq!(slide.shape_count().expect("shape count"), 2);
    assert_eq!(presentation.slide_masters().expect("masters").len(), 1);
    assert_eq!(presentation.slide_layouts().expect("layouts").len(), 11);
}

#[test]
fn opened_package_refuses_unsafe_mutable_hydration() {
    let mut package = Package::new().expect("new package");
    let bytes = package.to_bytes().expect("serialize package");
    let mut opened = Package::from_bytes(&bytes).expect("reopen package");
    assert!(opened.presentation_mut().is_err());
}
