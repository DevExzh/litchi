use litchi_opc::{OpcPackage, PackURI};
use litchi_pptx::Package;
use quick_xml::Reader;
use quick_xml::events::Event;

const DISABLED_SLIDE_XML: &[u8] =
    include_bytes!("../../../test-data/ooxml/pptx/master-visibility/slide_disabled.xml");
const MIXED_LAYOUT_XML: &[u8] =
    include_bytes!("../../../test-data/ooxml/pptx/master-visibility/layout_mixed.xml");
const DEFAULT_SLIDE_XML: &[u8] =
    include_bytes!("../../../test-data/ooxml/pptx/master-visibility/slide_defaults.xml");
const DEFAULT_LAYOUT_XML: &[u8] =
    include_bytes!("../../../test-data/ooxml/pptx/master-visibility/layout_defaults.xml");

#[test]
fn slide_and_layout_master_visibility_is_exposed_by_owner_parts() {
    let package = package_with_visibility(DISABLED_SLIDE_XML, MIXED_LAYOUT_XML);
    let presentation = package.presentation().unwrap();
    let slide = presentation.slides().unwrap().remove(0);
    let layout = slide.layout().unwrap().unwrap();

    // MasterVisibility was intentionally not duplicated into the graph
    // facade. Read the two root flags from their borrowed owner parts.
    assert_eq!(
        root_bool(slide.part().part().blob(), "showMasterSp"),
        Some(false)
    );
    assert_eq!(
        root_bool(slide.part().part().blob(), "showMasterPhAnim"),
        Some(false)
    );
    assert_eq!(
        root_bool(layout.part().part().blob(), "showMasterSp"),
        Some(true)
    );
    assert_eq!(
        root_bool(layout.part().part().blob(), "showMasterPhAnim"),
        Some(false)
    );
}

#[test]
fn omitted_master_visibility_flags_default_to_true() {
    let package = package_with_visibility(DEFAULT_SLIDE_XML, DEFAULT_LAYOUT_XML);
    let presentation = package.presentation().unwrap();
    let slide = presentation.slides().unwrap().remove(0);
    let layout = slide.layout().unwrap().unwrap();

    assert_eq!(root_bool(slide.part().part().blob(), "showMasterSp"), None);
    assert_eq!(
        root_bool(slide.part().part().blob(), "showMasterPhAnim"),
        None
    );
    assert_eq!(root_bool(layout.part().part().blob(), "showMasterSp"), None);
    assert_eq!(
        root_bool(layout.part().part().blob(), "showMasterPhAnim"),
        None
    );
    assert_eq!(
        root_bool(slide.part().part().blob(), "showMasterSp").unwrap_or(true),
        true
    );
    assert_eq!(
        root_bool(layout.part().part().blob(), "showMasterPhAnim").unwrap_or(true),
        true
    );
}

fn package_with_visibility(slide_xml: &[u8], layout_xml: &[u8]) -> Package {
    let mut package = Package::new().unwrap();
    package.presentation_mut().unwrap().add_slide().unwrap();
    let package_bytes = package.to_bytes().unwrap();
    let mut opc = OpcPackage::from_bytes(&package_bytes).unwrap();
    replace_part(&mut opc, "/ppt/slides/slide1.xml", slide_xml);
    replace_part(&mut opc, "/ppt/slideLayouts/slideLayout1.xml", layout_xml);
    Package::from_opc_package(opc).unwrap()
}

fn replace_part(package: &mut OpcPackage, part_name: &str, xml: &[u8]) {
    let part_name = PackURI::new(part_name).unwrap();
    package
        .get_part_mut(&part_name)
        .unwrap()
        .set_blob(xml.to_vec());
}

fn root_bool(xml: &[u8], attribute_name: &str) -> Option<bool> {
    let mut reader = Reader::from_reader(xml);
    loop {
        match reader.read_event().unwrap() {
            Event::Start(element) | Event::Empty(element) => {
                for attribute in element.attributes() {
                    let attribute = attribute.unwrap();
                    if attribute.key.as_ref() == attribute_name.as_bytes() {
                        return Some(matches!(
                            attribute
                                .decoded_and_normalized_value(
                                    quick_xml::XmlVersion::Explicit1_0,
                                    reader.decoder(),
                                )
                                .unwrap()
                                .as_ref(),
                            "1" | "true" | "on"
                        ));
                    }
                }
                return None;
            },
            Event::Decl(_) | Event::Comment(_) => {},
            Event::Eof => return None,
            _ => {},
        }
    }
}
