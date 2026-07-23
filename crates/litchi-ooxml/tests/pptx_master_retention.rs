use litchi_ooxml::PackURI;
use litchi_ooxml::pptx::Package;
use tempfile::NamedTempFile;

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
    let output = NamedTempFile::with_suffix(".pptx").unwrap();
    let mut package = Package::new().unwrap();
    package.save(output.path()).unwrap();

    let mut package = Package::open(output.path()).unwrap();
    let part_name = PackURI::new("/ppt/slideMasters/slideMaster1.xml").unwrap();
    package
        .opc_package_mut()
        .get_part_mut(&part_name)
        .unwrap()
        .set_blob(master_xml.to_vec());
    package
}
