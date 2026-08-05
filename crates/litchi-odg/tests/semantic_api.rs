use litchi_odg::{Builder, Drawing, layer::Layer, page::Page};

#[test]
fn focused_modules_are_the_canonical_semantic_api() {
    assert_eq!(Page::new("Page 1").name(), "Page 1");
    assert_eq!(Layer::new("Foreground").name(), "Foreground");

    let bytes = Builder::new().build().unwrap();
    let drawing = Drawing::from_bytes(bytes).unwrap();
    assert!(drawing.content_xml().contains("<office:drawing"));
}
