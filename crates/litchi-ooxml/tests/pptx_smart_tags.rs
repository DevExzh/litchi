use litchi_ooxml::PackURI;
use litchi_ooxml::pptx::Package;
use tempfile::NamedTempFile;

const PRESENTATION_XML: &[u8] =
    include_bytes!("../../../test-data/ooxml/pptx/smart-tags/presentation.xml");

#[test]
fn presentation_smart_tags_relationship_is_exposed() {
    let package = package_with_presentation_xml();

    assert_eq!(
        package
            .presentation()
            .unwrap()
            .smart_tags_relationship_id()
            .unwrap(),
        Some("rIdSmartTags".to_string())
    );
}

fn package_with_presentation_xml() -> Package {
    let output = NamedTempFile::with_suffix(".pptx").unwrap();
    let mut package = Package::new().unwrap();
    package.save(output.path()).unwrap();

    let mut package = Package::open(output.path()).unwrap();
    let part_name = PackURI::new("/ppt/presentation.xml").unwrap();
    package
        .opc_package_mut()
        .get_part_mut(&part_name)
        .unwrap()
        .set_blob(PRESENTATION_XML.to_vec());
    package
}
