use litchi_odi::{Builder, Image, frame::Frame, source::Source};

#[test]
fn focused_modules_are_the_canonical_semantic_api() {
    let frame = Frame::new(Source::Linked("Pictures/photo.png".into())).with_name("Photo");
    assert_eq!(frame.name(), Some("Photo"));
    assert!(matches!(frame.source(), Source::Linked(_)));

    let bytes = Builder::new().build().unwrap();
    let image = Image::from_bytes(bytes).unwrap();
    assert!(image.content_xml().contains("<office:image"));
}
