#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]

use litchi_opc::{OpcPackage, PackURI};
use litchi_pptx::Package;
use quick_xml::Reader;
use quick_xml::events::Event;

const DEFINED_LAYOUT_XML: &[u8] =
    include_bytes!("../../../test-data/ooxml/pptx/layout-metadata/defined.xml");
const DEFAULT_LAYOUT_XML: &[u8] =
    include_bytes!("../../../test-data/ooxml/pptx/layout-metadata/default.xml");

#[test]
fn layout_owner_exposes_name_and_kind_metadata() {
    let package = package_with_layout_xml(DEFINED_LAYOUT_XML);
    let presentation = package.presentation().unwrap();
    let slide = presentation.slides().unwrap().remove(0);
    let layout = slide.layout().unwrap().unwrap();

    // The standalone layout facade owns the contextual cSld name and type.
    assert_eq!(layout.name().unwrap(), "Picture with Caption");
    assert_eq!(layout.kind().unwrap().as_deref(), Some("picTx"));

    // matchingName, preserve, and userDrawn remain wire metadata until a
    // dedicated typed projection is published; inspect them through the
    // borrowed owner part without copying the layout tree.
    let attributes = root_attributes(layout.part().part().blob());
    assert_eq!(
        attributes.get("matchingName").map(String::as_str),
        Some("Picture Caption")
    );
    assert_eq!(attributes.get("preserve").map(String::as_str), Some("1"));
    assert_eq!(attributes.get("userDrawn").map(String::as_str), Some("0"));
}

#[test]
fn omitted_layout_metadata_keeps_wire_defaults_explicit() {
    let package = package_with_layout_xml(DEFAULT_LAYOUT_XML);
    let presentation = package.presentation().unwrap();
    let slide = presentation.slides().unwrap().remove(0);
    let layout = slide.layout().unwrap().unwrap();

    assert_eq!(layout.name().unwrap(), "Custom layout");
    assert_eq!(layout.kind().unwrap(), None);

    let attributes = root_attributes(layout.part().part().blob());
    assert!(!attributes.contains_key("matchingName"));
    assert!(!attributes.contains_key("type"));
    assert!(!attributes.contains_key("preserve"));
    assert!(!attributes.contains_key("userDrawn"));
}

fn package_with_layout_xml(layout_xml: &[u8]) -> Package {
    let mut package = Package::new().unwrap();
    package.presentation_mut().unwrap().add_slide().unwrap();
    let package_bytes = package.to_bytes().unwrap();
    let mut opc = OpcPackage::from_bytes(&package_bytes).unwrap();
    let part_name = PackURI::new("/ppt/slideLayouts/slideLayout1.xml").unwrap();
    opc.get_part_mut(&part_name)
        .unwrap()
        .set_blob(layout_xml.to_vec());
    Package::from_opc_package(opc).unwrap()
}

fn root_attributes(xml: &[u8]) -> std::collections::BTreeMap<String, String> {
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
            Event::Eof => panic!("layout XML has no root element"),
            _ => {},
        }
    }
}
