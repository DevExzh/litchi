#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]

use litchi_opc::{OpcPackage, PackURI};
use litchi_pptx::Package;

const RETAINED_MASTER_XML: &[u8] =
    include_bytes!("../../../test-data/ooxml/pptx/master-retention/retained.xml");
const DEFAULT_MASTER_XML: &[u8] =
    include_bytes!("../../../test-data/ooxml/pptx/master-retention/default.xml");

#[test]
fn retained_master_state_is_exposed() {
    let package = package_with_master_xml(RETAINED_MASTER_XML);
    let presentation = package.presentation().unwrap();
    let master = presentation.slide_masters().unwrap().remove(0);

    assert!(master.is_preserved().unwrap());
}

#[test]
fn omitted_master_preserve_flag_defaults_to_false() {
    let package = package_with_master_xml(DEFAULT_MASTER_XML);
    let presentation = package.presentation().unwrap();
    let master = presentation.slide_masters().unwrap().remove(0);

    assert!(!master.is_preserved().unwrap());
}

fn package_with_master_xml(master_xml: &[u8]) -> Package {
    let mut package = Package::new().unwrap();
    let package_bytes = package.to_bytes().unwrap();
    let mut opc = OpcPackage::from_bytes(&package_bytes).unwrap();
    let part_name = PackURI::new("/ppt/slideMasters/slideMaster1.xml").unwrap();
    opc.get_part_mut(&part_name)
        .unwrap()
        .set_blob(master_xml.to_vec());
    Package::from_opc_package(opc).unwrap()
}
