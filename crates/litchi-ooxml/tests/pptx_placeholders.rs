use litchi_ooxml::pptx::Package;
use litchi_ooxml::{OoxmlError, PackURI};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::part::BlobPart;

const PRESENTATION_XML: &[u8] =
    include_bytes!("../../../test-data/ooxml/pptx/placeholders/presentation.xml");
const VALID_SLIDE_XML: &[u8] =
    include_bytes!("../../../test-data/ooxml/pptx/placeholders/slide.xml");
const INVALID_PLACEHOLDER_SLIDE_XML: &[u8] =
    include_bytes!("../../../test-data/ooxml/pptx/placeholders/invalid-placeholder.xml");

#[test]
fn presentation_placeholders_return_declared_types() {
    let package = package_with_slide(VALID_SLIDE_XML);

    assert_eq!(
        package.presentation().unwrap().get_placeholders(0).unwrap(),
        Some(vec!["ctrTitle".to_string()])
    );
}

#[test]
fn presentation_placeholders_preserve_type_decoding_errors() {
    let package = package_with_slide(INVALID_PLACEHOLDER_SLIDE_XML);

    assert!(matches!(
        package.presentation().unwrap().get_placeholders(0),
        Err(OoxmlError::CommonXml(_))
    ));
}

fn package_with_slide(slide_xml: &[u8]) -> Package {
    let mut package = Package::new().unwrap();
    let presentation_name = PackURI::new("/ppt/presentation.xml").unwrap();
    let slide_name = PackURI::new("/ppt/slides/slide1.xml").unwrap();

    {
        let presentation = package
            .opc_package_mut()
            .get_part_mut(&presentation_name)
            .unwrap();
        presentation.set_blob(PRESENTATION_XML.to_vec());
        presentation.rels_mut().add_relationship(
            rt::SLIDE.to_string(),
            "slides/slide1.xml".to_string(),
            "rIdSlideOne".to_string(),
            false,
        );
    }
    package.opc_package_mut().add_part(Box::new(BlobPart::new(
        slide_name,
        ct::PML_SLIDE.to_string(),
        slide_xml.to_vec(),
    )));
    package
}
