use std::collections::BTreeMap;

use litchi_opc::{OpcPackage, PackURI};
use litchi_pptx::Package;
use quick_xml::Reader;
use quick_xml::events::Event;

const PRESENTATION_XML: &[u8] =
    include_bytes!("../../../test-data/ooxml/pptx/presentation-metadata/presentation.xml");

#[test]
fn presentation_root_metadata_is_retained_by_the_main_owner() {
    let mut package = Package::new().unwrap();
    let package_bytes = package.to_bytes().unwrap();
    let mut opc = OpcPackage::from_bytes(&package_bytes).unwrap();
    let part_name = PackURI::new("/ppt/presentation.xml").unwrap();
    opc.get_part_mut(&part_name)
        .unwrap()
        .set_blob(PRESENTATION_XML.to_vec());
    let package = Package::from_opc_package(opc).unwrap();
    let presentation = package.presentation().unwrap();
    let attributes = root_attributes(presentation.part().part().blob());

    // These root attributes are intentionally retained as wire metadata while
    // the focused graph facade exposes only typed graph operations.
    assert_eq!(
        attributes.get("serverZoom").map(String::as_str),
        Some("125000")
    );
    assert_eq!(
        attributes.get("firstSlideNum").map(String::as_str),
        Some("7")
    );
    assert_eq!(
        attributes
            .get("showSpecialPlsOnTitleSld")
            .map(String::as_str),
        Some("0")
    );
    assert_eq!(attributes.get("rtl").map(String::as_str), Some("1"));
    assert_eq!(
        attributes
            .get("removePersonalInfoOnSave")
            .map(String::as_str),
        Some("true")
    );
    assert_eq!(attributes.get("compatMode").map(String::as_str), Some("1"));
    assert_eq!(
        attributes
            .get("strictFirstAndLastChars")
            .map(String::as_str),
        Some("false")
    );
    assert_eq!(
        attributes.get("embedTrueTypeFonts").map(String::as_str),
        Some("1")
    );
    assert_eq!(
        attributes.get("saveSubsetFonts").map(String::as_str),
        Some("true")
    );
    assert_eq!(
        attributes.get("autoCompressPictures").map(String::as_str),
        Some("0")
    );
    assert_eq!(
        attributes.get("bookmarkIdSeed").map(String::as_str),
        Some("42")
    );
    assert_eq!(
        attributes.get("conformance").map(String::as_str),
        Some("strict")
    );
}

fn root_attributes(xml: &[u8]) -> BTreeMap<String, String> {
    let mut reader = Reader::from_reader(xml);
    loop {
        match reader.read_event().unwrap() {
            Event::Start(element) | Event::Empty(element) => {
                return element
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
                    .collect();
            },
            Event::Decl(_) | Event::Comment(_) => {},
            Event::Eof => panic!("presentation XML has no root element"),
            _ => {},
        }
    }
}
