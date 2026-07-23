use litchi_ooxml::PackURI;
use litchi_ooxml::pptx::Package;
use tempfile::NamedTempFile;

const MASTER_XML: &[u8] =
    include_bytes!("../../../test-data/ooxml/pptx/master-text-styles/master.xml");

#[test]
fn master_text_style_inventories_are_exposed() {
    let package = package_with_master_xml();
    let presentation = package.presentation().unwrap();
    let slide = presentation.slides().unwrap().remove(0);
    let text_styles = slide.master().unwrap().text_styles().unwrap().unwrap();

    let title = text_styles.title_style().unwrap();
    assert!(!title.has_default_paragraph_properties());
    assert_eq!(title.levels(), [1, 2]);
    assert!(title.has_level(2));

    let body = text_styles.body_style().unwrap();
    assert!(body.has_default_paragraph_properties());
    assert_eq!(body.levels(), [1, 9]);

    let other = text_styles.other_style().unwrap();
    assert!(!other.has_default_paragraph_properties());
    assert!(other.levels().is_empty());
}

fn package_with_master_xml() -> Package {
    let output = NamedTempFile::with_suffix(".pptx").unwrap();
    let mut package = Package::new().unwrap();
    package.presentation_mut().unwrap().add_slide().unwrap();
    package.save(output.path()).unwrap();

    let mut package = Package::open(output.path()).unwrap();
    let part_name = PackURI::new("/ppt/slideMasters/slideMaster1.xml").unwrap();
    package
        .opc_package_mut()
        .get_part_mut(&part_name)
        .unwrap()
        .set_blob(MASTER_XML.to_vec());
    package
}
