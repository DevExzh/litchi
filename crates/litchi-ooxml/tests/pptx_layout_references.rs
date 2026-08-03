use litchi_ooxml::PackURI;
use litchi_ooxml::pptx::Package;
use tempfile::NamedTempFile;

const MASTER_XML: &[u8] =
    include_bytes!("../../../test-data/ooxml/pptx/layout-references/master.xml");

#[test]
fn master_layout_references_include_stable_ids() {
    let package = package_with_master_xml();
    let presentation = package.presentation().unwrap();
    let slide = presentation.slides().unwrap().remove(0);
    let master = slide.master().unwrap();

    let references = master.slide_layout_references().unwrap();
    assert_eq!(references.len(), 1);
    assert_eq!(references[0].layout_id(), Some(2_147_483_648));
    assert_eq!(references[0].relationship_id(), "rId1");
    assert_eq!(master.slide_layout_rids().unwrap(), ["rId1"]);
    assert_eq!(master.slide_layouts().unwrap().len(), 1);
}

fn package_with_master_xml() -> Package {
    let output = NamedTempFile::with_suffix(".pptx").unwrap();
    let mut package = Package::new().unwrap();
    package.presentation_mut().unwrap().add_slide().unwrap();
    package.save(output.path()).unwrap();

    let mut package = Package::open(output.path()).unwrap();
    let part_name = PackURI::new("/ppt/slideMasters/slideMaster1.xml").unwrap();
    package
        .edit_opc(|opc| {
            opc.get_part_mut(&part_name)?.set_blob(MASTER_XML.to_vec());
            Ok(())
        })
        .unwrap();
    package
}
