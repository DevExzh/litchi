use litchi_ooxml::PackURI;
use litchi_ooxml::pptx::{Package, SlideLayoutMetadata};
use tempfile::NamedTempFile;

const DEFINED_LAYOUT_XML: &[u8] =
    include_bytes!("../../../test-data/ooxml/pptx/layout-metadata/defined.xml");
const DEFAULT_LAYOUT_XML: &[u8] =
    include_bytes!("../../../test-data/ooxml/pptx/layout-metadata/default.xml");

#[test]
fn layout_metadata_is_exposed() {
    let package = package_with_layout_xml(DEFINED_LAYOUT_XML);
    let presentation = package.presentation().unwrap();
    let slide = presentation.slides().unwrap().remove(0);
    let metadata = slide.layout().unwrap().metadata().unwrap();

    assert_eq!(metadata.matching_name(), "Picture Caption");
    assert_eq!(metadata.layout_type(), "picTx");
    assert!(metadata.is_preserved());
    assert!(!metadata.is_user_drawn());
}

#[test]
fn omitted_layout_metadata_uses_schema_defaults() {
    let package = package_with_layout_xml(DEFAULT_LAYOUT_XML);
    let presentation = package.presentation().unwrap();
    let slide = presentation.slides().unwrap().remove(0);
    let metadata = slide.layout().unwrap().metadata().unwrap();

    assert_eq!(metadata, SlideLayoutMetadata::default());
    assert_eq!(metadata.matching_name(), "");
    assert_eq!(metadata.layout_type(), "cust");
    assert!(!metadata.is_preserved());
    assert!(!metadata.is_user_drawn());
}

fn package_with_layout_xml(layout_xml: &[u8]) -> Package {
    let output = NamedTempFile::with_suffix(".pptx").unwrap();
    let mut package = Package::new().unwrap();
    package.presentation_mut().unwrap().add_slide().unwrap();
    package.save(output.path()).unwrap();

    let mut package = Package::open(output.path()).unwrap();
    let part_name = PackURI::new("/ppt/slideLayouts/slideLayout1.xml").unwrap();
    package
        .opc_package_mut()
        .get_part_mut(&part_name)
        .unwrap()
        .set_blob(layout_xml.to_vec());
    package
}
