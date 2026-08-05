use litchi_opc::{OpcPackage, PackURI, Part};
use litchi_pptx::Package;

const MASTER_XML: &[u8] =
    include_bytes!("../../../test-data/ooxml/pptx/layout-references/master.xml");

#[test]
fn master_layout_relationships_keep_stable_ids() {
    let package = package_with_master_xml();
    let presentation = package.presentation().unwrap();
    let slide = presentation.slides().unwrap().remove(0);
    let master = slide.layout().unwrap().unwrap().master().unwrap();

    let relationship = master
        .part()
        .part()
        .rels()
        .get("rId1")
        .expect("layout relationship");
    assert_eq!(relationship.r_id(), "rId1");
    assert_eq!(
        relationship.target_partname().unwrap().as_str(),
        "/ppt/slideLayouts/slideLayout1.xml"
    );
    assert_eq!(master.layouts().unwrap().len(), 1);
}

fn package_with_master_xml() -> Package {
    let mut package = Package::new().unwrap();
    package
        .presentation_mut()
        .unwrap()
        .add_slide()
        .unwrap();
    let package_bytes = package.to_bytes().unwrap();
    let mut opc = OpcPackage::from_bytes(&package_bytes).unwrap();
    let part_name = PackURI::new("/ppt/slideMasters/slideMaster1.xml").unwrap();
    opc.get_part_mut(&part_name)
        .unwrap()
        .set_blob(MASTER_XML.to_vec());
    Package::from_opc_package(opc).unwrap()
}
