use litchi_ooxml::PackURI;
use litchi_ooxml::pptx::Package;
use tempfile::NamedTempFile;

const LAYOUT_XML: &[u8] = include_bytes!("../../../test-data/ooxml/pptx/header-footer/layout.xml");
const MASTER_XML: &[u8] = include_bytes!("../../../test-data/ooxml/pptx/header-footer/master.xml");

#[test]
fn layout_and_master_header_footer_visibility_is_exposed() {
    let package = package_with_header_footer_xml();
    let presentation = package.presentation().unwrap();
    let slide = presentation.slides().unwrap().remove(0);

    let layout_visibility = slide.layout().unwrap().header_footer().unwrap().unwrap();
    assert!(!layout_visibility.shows_date_time());
    assert!(layout_visibility.shows_footer());
    assert!(!layout_visibility.shows_header());
    assert!(!layout_visibility.shows_slide_number());

    let master_visibility = slide.master().unwrap().header_footer().unwrap().unwrap();
    assert!(master_visibility.shows_date_time());
    assert!(!master_visibility.shows_footer());
    assert!(!master_visibility.shows_header());
    assert!(master_visibility.shows_slide_number());
}

fn package_with_header_footer_xml() -> Package {
    let output = NamedTempFile::with_suffix(".pptx").unwrap();
    let mut package = Package::new().unwrap();
    package.presentation_mut().unwrap().add_slide().unwrap();
    package.save(output.path()).unwrap();

    let mut package = Package::open(output.path()).unwrap();
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

fn replace_part(package: &mut Package, part_name: &str, xml: &[u8]) {
    let part_name = PackURI::new(part_name).unwrap();
    package
        .edit_opc(|opc| {
            opc.get_part_mut(&part_name).unwrap().set_blob(xml.to_vec());
            Ok(())
        })
        .unwrap();
}
