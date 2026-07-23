use litchi_ooxml::PackURI;
use litchi_ooxml::pptx::Package;
use litchi_ooxml::pptx::{ExtendedGuideColorKind, ExtendedGuideOrientation};

const PRESENTATION_XML: &[u8] =
    include_bytes!("../../../test-data/ooxml/pptx/extended-guides/presentation.xml");
const INVALID_PRESENTATION_XML: &[u8] =
    include_bytes!("../../../test-data/ooxml/pptx/extended-guides/invalid-presentation.xml");

#[test]
fn presentation_extended_guides_are_available_at_high_level() {
    let package = package_with_presentation_xml(PRESENTATION_XML);
    let guides = package.presentation().unwrap().extended_guides().unwrap();

    let slide = guides.slide.as_ref().unwrap();
    assert_eq!(slide.guides.len(), 1);
    assert_eq!(slide.guides[0].id, 1);
    assert_eq!(slide.guides[0].name.as_deref(), Some("Middle"));
    assert_eq!(
        slide.guides[0].orientation,
        Some(ExtendedGuideOrientation::Horizontal)
    );
    assert_eq!(slide.guides[0].position, Some(2160));
    assert_eq!(slide.guides[0].user_drawn, Some(true));
    assert_eq!(slide.guides[0].color.kind, ExtendedGuideColorKind::Srgb);
    assert!(guides.notes.as_ref().unwrap().guides.is_empty());
}

#[test]
fn presentation_extended_guides_preserve_validation_errors() {
    let package = package_with_presentation_xml(INVALID_PRESENTATION_XML);

    assert!(package.presentation().unwrap().extended_guides().is_err());
}

fn package_with_presentation_xml(xml: &[u8]) -> Package {
    let mut package = Package::new().unwrap();
    let presentation_name = PackURI::new("/ppt/presentation.xml").unwrap();
    package
        .opc_package_mut()
        .get_part_mut(&presentation_name)
        .unwrap()
        .set_blob(xml.to_vec());
    package
}
