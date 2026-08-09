#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]

use litchi_opc::{OpcPackage, PackURI, Part};
use litchi_pptx::Package;
use litchi_pptx::parts::SlidePart;
use quick_xml::Reader;
use quick_xml::events::Event;

const LAYOUT_XML: &[u8] = include_bytes!("../../../test-data/ooxml/pptx/header-footer/layout.xml");
const MASTER_XML: &[u8] = include_bytes!("../../../test-data/ooxml/pptx/header-footer/master.xml");

#[test]
fn layout_and_master_header_footer_visibility_is_exposed_by_the_part_owner() {
    // Header/footer visibility is not currently published as a standalone
    // slide/layout model. Validate the direct PresentationML part owner while
    // preserving the same inheritance inputs and semantic assertions.
    let package = package_with_header_footer_xml();
    let slide = SlidePart::from_part(
        package
            .get_part(&PackURI::new("/ppt/slides/slide1.xml").unwrap())
            .unwrap(),
    )
    .unwrap();
    let layout = slide.layout(&package).unwrap().unwrap();
    let master = layout.master(&package).unwrap();

    let layout_visibility = parse_visibility(layout.part());
    assert!(!layout_visibility.date_time);
    assert!(layout_visibility.footer);
    assert!(!layout_visibility.header);
    assert!(!layout_visibility.slide_number);

    let master_visibility = parse_visibility(master.part());
    assert!(master_visibility.date_time);
    assert!(!master_visibility.footer);
    assert!(!master_visibility.header);
    assert!(master_visibility.slide_number);
}

#[derive(Default)]
struct Visibility {
    date_time: bool,
    footer: bool,
    header: bool,
    slide_number: bool,
}

fn parse_visibility(part: &dyn Part) -> Visibility {
    let xml = litchi_ooxml_common::mce::process_ooxml(part.blob()).unwrap();
    let mut reader = Reader::from_reader(xml.as_ref());
    reader.config_mut().trim_text(true);
    loop {
        match reader.read_event().unwrap() {
            Event::Start(element) | Event::Empty(element)
                if element.local_name().as_ref() == b"hf" =>
            {
                let mut value = Visibility {
                    date_time: true,
                    footer: true,
                    header: true,
                    slide_number: true,
                };
                for attribute in element.attributes().flatten() {
                    let target = match attribute.key.as_ref() {
                        b"dt" => &mut value.date_time,
                        b"ftr" => &mut value.footer,
                        b"hdr" => &mut value.header,
                        b"sldNum" => &mut value.slide_number,
                        _ => continue,
                    };
                    *target =
                        matches!(attribute.value.as_ref(), b"1" | b"true" | b"TRUE" | b"True");
                }
                return value;
            },
            Event::Eof => return Visibility::default(),
            _ => {},
        }
    }
}

fn package_with_header_footer_xml() -> OpcPackage {
    let mut package = Package::new().unwrap();
    package.presentation_mut().unwrap().add_slide().unwrap();
    let bytes = package.to_bytes().unwrap();
    let mut package = OpcPackage::from_vec(bytes).unwrap();
    replace_part(
        &mut package,
        "/ppt/slideLayouts/slideLayout1.xml",
        LAYOUT_XML,
    );
    replace_part(
        &mut package,
        "/ppt/slideMasters/slideMaster1.xml",
        MASTER_XML,
    );
    package
}

fn replace_part(package: &mut OpcPackage, part_name: &str, xml: &[u8]) {
    package
        .get_part_mut(&PackURI::new(part_name).unwrap())
        .unwrap()
        .set_blob(xml.to_vec());
}
