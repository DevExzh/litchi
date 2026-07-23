use litchi_ooxml::OoxmlError;
use litchi_ooxml::PackURI;
use litchi_ooxml::pptx::Package;
use tempfile::NamedTempFile;

const LOCAL_SHAPE_XML: &[u8] =
    include_bytes!("../../../test-data/ooxml/pptx/shapes/shape_identity.xml");

#[test]
fn slide_shapes_are_resolved_by_non_visual_id() {
    let package = package_with_slide_xml(LOCAL_SHAPE_XML);
    let presentation = package.presentation().unwrap();
    let slide = presentation.slides().unwrap().remove(0);

    let title = slide.shape_by_id(7).unwrap().unwrap();
    assert_eq!(title.shape_id().unwrap(), Some(7));
    assert_eq!(title.placeholder_type().unwrap(), "title");
    assert_eq!(title.placeholder_index().unwrap(), Some(0));

    let body = slide.shape_by_id(11).unwrap().unwrap();
    assert_eq!(body.shape_id().unwrap(), Some(11));
    assert_eq!(body.placeholder_type().unwrap(), "body");
    assert_eq!(body.placeholder_index().unwrap(), Some(3));
    assert!(slide.shape_by_id(99).unwrap().is_none());
}

#[test]
fn duplicate_non_visual_shape_ids_are_rejected() {
    let xml = std::str::from_utf8(LOCAL_SHAPE_XML)
        .unwrap()
        .replace("id=\"11\"", "id=\"7\"");
    let package = package_with_slide_xml(xml.as_bytes());
    let presentation = package.presentation().unwrap();
    let slide = presentation.slides().unwrap().remove(0);

    assert!(matches!(
        slide.shape_by_id(7),
        Err(OoxmlError::InvalidFormat(message))
            if message.contains("multiple shapes with non-visual ID 7")
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
        .opc_package_mut()
        .get_part_mut(&part_name)
        .unwrap()
        .set_blob(xml.to_vec());
    package
}
