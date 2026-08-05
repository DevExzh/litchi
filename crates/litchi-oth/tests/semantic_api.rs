use litchi_oth::{Builder, Template, link::Link, paragraph::Paragraph};

#[test]
fn focused_modules_are_the_canonical_semantic_api() {
    assert_eq!(Paragraph::new("Welcome").text(), "Welcome");
    let link = Link::new("https://example.test", "Example");
    assert_eq!(link.href(), "https://example.test");
    assert_eq!(link.label(), "Example");

    let bytes = Builder::new().build().unwrap();
    let template = Template::from_bytes(bytes).unwrap();
    assert!(template.content_xml().contains("<office:text"));
}
