use litchi_ooxml::PackURI;
use litchi_ooxml::pptx::Package;
use tempfile::NamedTempFile;

const HIDDEN_SLIDE_XML: &[u8] =
    include_bytes!("../../../test-data/ooxml/pptx/slide-visibility/hidden.xml");
const DEFAULT_SLIDE_XML: &[u8] =
    include_bytes!("../../../test-data/ooxml/pptx/slide-visibility/default.xml");

#[test]
fn hidden_slide_state_is_exposed() {
    let package = package_with_slide_xml(HIDDEN_SLIDE_XML);
    let presentation = package.presentation().unwrap();
    let slide = presentation.slides().unwrap().remove(0);

    assert!(slide.is_hidden().unwrap());
}

#[test]
fn omitted_slide_show_flag_is_not_hidden() {
    let package = package_with_slide_xml(DEFAULT_SLIDE_XML);
    let presentation = package.presentation().unwrap();
    let slide = presentation.slides().unwrap().remove(0);

    assert!(!slide.is_hidden().unwrap());
}

fn package_with_slide_xml(xml: &[u8]) -> Package {
    let output = NamedTempFile::with_suffix(".pptx").unwrap();
    let mut package = Package::new().unwrap();
    package.presentation_mut().unwrap().add_slide().unwrap();
    package.save(output.path()).unwrap();

    let mut package = Package::open(output.path()).unwrap();
    let part_name = PackURI::new("/ppt/slides/slide1.xml").unwrap();
    package
        .opc_package_mut()
        .get_part_mut(&part_name)
        .unwrap()
        .set_blob(xml.to_vec());
    package
}
