use litchi_odm::{Builder, Master, section::Section, subdocument::Subdocument};

#[test]
fn focused_modules_are_the_canonical_semantic_api() {
    let mut section = Section::new("Introduction");
    section.push(Subdocument::new("chapter-1.odt"));
    assert_eq!(section.children()[0].href(), "chapter-1.odt");

    let bytes = Builder::new().build().unwrap();
    let master = Master::from_bytes(bytes).unwrap();
    assert!(master.content_xml().contains("<office:text"));
}
