#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]

use litchi_opc::{OpcPackage, PackURI};
use litchi_pptx::Package;

const HIDDEN_SLIDE_XML: &[u8] =
    include_bytes!("../../../test-data/ooxml/pptx/slide-visibility/hidden.xml");
const DEFAULT_SLIDE_XML: &[u8] =
    include_bytes!("../../../test-data/ooxml/pptx/slide-visibility/default.xml");

#[test]
fn hidden_slide_state_is_exposed() {
    let package = package_with_slide_xml(HIDDEN_SLIDE_XML);
    let slide = package.presentation().unwrap().slides().unwrap().remove(0);

    assert!(slide.is_hidden().unwrap());
}

#[test]
fn omitted_slide_show_flag_is_not_hidden() {
    let package = package_with_slide_xml(DEFAULT_SLIDE_XML);
    let slide = package.presentation().unwrap().slides().unwrap().remove(0);

    assert!(!slide.is_hidden().unwrap());
}

fn package_with_slide_xml(xml: &[u8]) -> Package {
    // Opened packages intentionally expose only read-only OPC access. Build
    // the synthetic graph through the public OPC owner and then adopt it.
    let mut authored = Package::new().unwrap();
    authored.presentation_mut().unwrap().add_slide().unwrap();
    let bytes = authored.to_bytes().unwrap();
    let mut opc = OpcPackage::from_bytes(&bytes).unwrap();
    opc.get_part_mut(&PackURI::new("/ppt/slides/slide1.xml").unwrap())
        .unwrap()
        .set_blob(xml.to_vec());
    Package::from_opc_package(opc).unwrap()
}
