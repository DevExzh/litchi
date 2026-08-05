use std::collections::BTreeMap;

use litchi_opc::{OpcPackage, PackURI, Part};
use litchi_pptx::Package;
use quick_xml::Reader;
use quick_xml::events::Event;

const DEFINED_XML: &[u8] =
    include_bytes!("../../../test-data/ooxml/pptx/presentation-sizes/defined.xml");
const ABSENT_XML: &[u8] =
    include_bytes!("../../../test-data/ooxml/pptx/presentation-sizes/absent.xml");

#[test]
fn presentation_surface_sizes_are_exposed_by_main_and_wire_owners() {
    let package = package_with_presentation_xml(DEFINED_XML);
    let presentation = package.presentation().unwrap();

    assert_eq!(presentation.slide_size().unwrap(), (12_192_000, 6_858_000));
    let attributes = element_attributes(presentation.part().part().blob(), "sldSz").unwrap();
    assert_eq!(attributes.get("cx").map(String::as_str), Some("12192000"));
    assert_eq!(attributes.get("cy").map(String::as_str), Some("6858000"));
    assert_eq!(
        attributes.get("type").map(String::as_str),
        Some("screen16x9")
    );

    let notes = element_attributes(presentation.part().part().blob(), "notesSz").unwrap();
    assert_eq!(notes.get("cx").map(String::as_str), Some("6858000"));
    assert_eq!(notes.get("cy").map(String::as_str), Some("9144000"));
}

#[test]
fn absent_presentation_surface_sizes_return_typed_absence_or_error() {
    let package = package_with_presentation_xml(ABSENT_XML);
    let presentation = package.presentation().unwrap();

    // slide_size is a required PresentationML element in the focused graph
    // API; absence is therefore a typed invalid-document error. notesSz has
    // no standalone projection and is represented by absent wire metadata.
    assert!(presentation.slide_size().is_err());
    assert_eq!(
        element_attributes(presentation.part().part().blob(), "sldSz"),
        None
    );
    assert_eq!(
        element_attributes(presentation.part().part().blob(), "notesSz"),
        None
    );
}

fn package_with_presentation_xml(xml: &[u8]) -> Package {
    let mut package = Package::new().unwrap();
    package
        .presentation_mut()
        .unwrap()
        .add_slide()
        .unwrap();
    let package_bytes = package.to_bytes().unwrap();
    let mut opc = OpcPackage::from_bytes(&package_bytes).unwrap();
    let part_name = PackURI::new("/ppt/presentation.xml").unwrap();
    opc.get_part_mut(&part_name)
        .unwrap()
        .set_blob(xml.to_vec());
    Package::from_opc_package(opc).unwrap()
}

fn element_attributes(xml: &[u8], local_name: &str) -> Option<BTreeMap<String, String>> {
    let mut reader = Reader::from_reader(xml);
    loop {
        match reader.read_event().unwrap() {
            Event::Start(element) | Event::Empty(element)
                if element.local_name().as_ref() == local_name.as_bytes() =>
            {
                return Some(
                    element
                        .attributes()
                        .map(|attribute| {
                            let attribute = attribute.unwrap();
                            (
                                String::from_utf8(attribute.key.as_ref().to_vec()).unwrap(),
                                attribute
                                    .decoded_and_normalized_value(
                                        quick_xml::XmlVersion::Explicit1_0,
                                        reader.decoder(),
                                    )
                                    .unwrap()
                                    .into_owned(),
                            )
                        })
                        .collect(),
                );
            }
            Event::Eof => return None,
            _ => {}
        }
    }
}
