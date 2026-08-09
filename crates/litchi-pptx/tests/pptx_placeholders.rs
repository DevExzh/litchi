#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]

use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{BlobPart, OpcPackage, PackURI};
use litchi_pptx::{Error, Package};

const PRESENTATION_XML: &[u8] =
    include_bytes!("../../../test-data/ooxml/pptx/placeholders/presentation.xml");
const VALID_SLIDE_XML: &[u8] =
    include_bytes!("../../../test-data/ooxml/pptx/placeholders/slide.xml");
const INVALID_PLACEHOLDER_SLIDE_XML: &[u8] =
    include_bytes!("../../../test-data/ooxml/pptx/placeholders/invalid-placeholder.xml");

#[test]
fn slide_shape_placeholders_return_declared_types() {
    let package = package_with_slide(VALID_SLIDE_XML);
    let slide = package.presentation().unwrap().slide(0).unwrap().unwrap();
    let shapes = slide.shapes().unwrap();
    let placeholder = shapes.placeholders().next().unwrap().placeholder().unwrap();

    assert_eq!(placeholder.kind(), Some("ctrTitle"));
    assert_eq!(placeholder.index(), 0);
}

#[test]
fn slide_shape_placeholders_preserve_type_decoding_errors() {
    let package = package_with_slide(INVALID_PLACEHOLDER_SLIDE_XML);
    let slide = package.presentation().unwrap().slide(0).unwrap().unwrap();

    assert!(matches!(
        slide.shapes(),
        Err(Error::Decode(_) | Error::Xml(_) | Error::Invalid(_))
    ));
}

fn package_with_slide(slide_xml: &[u8]) -> Package {
    let mut package = Package::new().unwrap();
    let presentation_name = PackURI::new("/ppt/presentation.xml").unwrap();
    let slide_name = PackURI::new("/ppt/slides/slide1.xml").unwrap();

    let package_bytes = package.to_bytes().unwrap();
    let mut opc = OpcPackage::from_bytes(&package_bytes).unwrap();
    let presentation = opc.get_part_mut(&presentation_name).unwrap();
    presentation.set_blob(PRESENTATION_XML.to_vec());
    presentation.rels_mut().add_relationship(
        rt::SLIDE.to_string(),
        "slides/slide1.xml".to_string(),
        "rIdSlideOne".to_string(),
        false,
    );
    opc.add_part(Box::new(BlobPart::new(
        slide_name,
        ct::PML_SLIDE.to_string(),
        slide_xml.to_vec(),
    )));
    Package::from_opc_package(opc).unwrap()
}
