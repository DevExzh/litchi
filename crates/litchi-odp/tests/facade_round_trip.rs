use litchi_odp::{Presentation, PresentationBuilder};

#[test]
fn builder_and_presentation_facade_round_trip() {
    let mut builder = PresentationBuilder::new();
    builder.add_slide_with_title("Welcome", "Hello from ODP").unwrap();
    let bytes = builder.build().unwrap();
    let presentation = Presentation::from_bytes(bytes.clone()).unwrap();
    assert_eq!(presentation.slide_count().unwrap(), 1);
    assert_eq!(presentation.text().unwrap(), "Welcome\nHello from ODP");
    assert_eq!(presentation.into_bytes(), bytes);
}

#[test]
fn presentation_facade_reports_empty_packages_without_slides() {
    let bytes = PresentationBuilder::new().build().unwrap();
    let presentation = Presentation::from_bytes(bytes).unwrap();
    assert_eq!(presentation.slide_count().unwrap(), 0);
    assert_eq!(presentation.text().unwrap(), "");
}
