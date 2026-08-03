use litchi_ooxml::PackURI;
use litchi_ooxml::pptx::Package;
use tempfile::NamedTempFile;

const LOCAL_SHAPE_XML: &[u8] =
    include_bytes!("../../../test-data/ooxml/pptx/shapes/shape_identity.xml");

#[test]
fn slide_shapes_use_names_first_with_checked_numeric_fallback() {
    let package = package_with_slide_xml(LOCAL_SHAPE_XML);
    let presentation = package.presentation().unwrap();
    assert_eq!(
        presentation.get_placeholders(0).unwrap(),
        Some(vec!["title".to_string(), "body".to_string()])
    );
    let slide = presentation.slides().unwrap().remove(0);

    assert_eq!(slide.placeholders().unwrap().count(), 2);

    let title = slide.shape("Title").unwrap().unwrap();
    assert_eq!(title.name(), Some("Title"));
    assert_eq!(title.id(), Some(7));
    let title_placeholder = title.placeholder().unwrap();
    assert_eq!(title_placeholder.kind(), Some("title"));
    assert_eq!(title_placeholder.index(), 0);

    let body = slide.shape(1_usize).unwrap().unwrap();
    assert_eq!(body.name(), Some("Body"));
    assert_eq!(body.id(), Some(11));
    let body_placeholder = body.placeholder().unwrap();
    assert_eq!(body_placeholder.kind(), Some("body"));
    assert_eq!(body_placeholder.index(), 3);
    assert!(slide.shape("Missing").unwrap().is_none());
}

#[test]
fn duplicate_semantic_shape_names_are_typed_errors() {
    let xml = std::str::from_utf8(LOCAL_SHAPE_XML)
        .unwrap()
        .replace("name=\"Body\"", "name=\"Title\"");
    let package = package_with_slide_xml(xml.as_bytes());
    let presentation = package.presentation().unwrap();
    let slide = presentation.slides().unwrap().remove(0);

    assert!(matches!(
        slide.shape("Title"),
        Err(litchi_ooxml::OoxmlError::Pptx(
            litchi_pptx::Error::ShapeLookup(
                litchi_pptx::shape::LookupError::AmbiguousName {
                    name,
                    matches: 2,
                },
            ),
        )) if name == "Title"
    ));
}

fn package_with_slide_xml(xml: &[u8]) -> Package {
    let output = NamedTempFile::with_suffix(".pptx").unwrap();
    let mut package = Package::new().unwrap();
    package.presentation_mut().unwrap().add_slide().unwrap();
    package.save(output.path()).unwrap();

    let mut package = Package::open(output.path()).unwrap();
    let part_name = PackURI::new("/ppt/slides/slide1.xml").unwrap();
    package
        .edit_opc(|opc| {
            opc.get_part_mut(&part_name)?.set_blob(xml.to_vec());
            Ok(())
        })
        .unwrap();
    package
}
