use litchi_ooxml::PackURI;
use litchi_ooxml::pptx::{MasterVisibility, Package};
use tempfile::NamedTempFile;

const DISABLED_SLIDE_XML: &[u8] =
    include_bytes!("../../../test-data/ooxml/pptx/master-visibility/slide_disabled.xml");
const MIXED_LAYOUT_XML: &[u8] =
    include_bytes!("../../../test-data/ooxml/pptx/master-visibility/layout_mixed.xml");
const DEFAULT_SLIDE_XML: &[u8] =
    include_bytes!("../../../test-data/ooxml/pptx/master-visibility/slide_defaults.xml");
const DEFAULT_LAYOUT_XML: &[u8] =
    include_bytes!("../../../test-data/ooxml/pptx/master-visibility/layout_defaults.xml");

#[test]
fn slide_and_layout_master_visibility_is_exposed() {
    let package = package_with_visibility(DISABLED_SLIDE_XML, MIXED_LAYOUT_XML);
    let presentation = package.presentation().unwrap();
    let slide = presentation.slides().unwrap().remove(0);

    let slide_visibility = slide.master_visibility().unwrap();
    assert!(!slide_visibility.shows_master_shapes());
    assert!(!slide_visibility.shows_master_placeholder_animations());

    let layout_visibility = slide.layout().unwrap().master_visibility().unwrap();
    assert!(layout_visibility.shows_master_shapes());
    assert!(!layout_visibility.shows_master_placeholder_animations());
}

#[test]
fn omitted_master_visibility_flags_default_to_true() {
    let package = package_with_visibility(DEFAULT_SLIDE_XML, DEFAULT_LAYOUT_XML);
    let presentation = package.presentation().unwrap();
    let slide = presentation.slides().unwrap().remove(0);

    let slide_visibility = slide.master_visibility().unwrap();
    assert_eq!(slide_visibility, MasterVisibility::default());
    assert!(slide_visibility.shows_master_shapes());
    assert!(slide_visibility.shows_master_placeholder_animations());

    let layout_visibility = slide.layout().unwrap().master_visibility().unwrap();
    assert_eq!(layout_visibility, MasterVisibility::default());
    assert!(layout_visibility.shows_master_shapes());
    assert!(layout_visibility.shows_master_placeholder_animations());
}

fn package_with_visibility(slide_xml: &[u8], layout_xml: &[u8]) -> Package {
    let output = NamedTempFile::with_suffix(".pptx").unwrap();
    let mut package = Package::new().unwrap();
    package.presentation_mut().unwrap().add_slide().unwrap();
    package.save(output.path()).unwrap();

    let mut package = Package::open(output.path()).unwrap();
    replace_part(&mut package, "/ppt/slides/slide1.xml", slide_xml);
    replace_part(
        &mut package,
        "/ppt/slideLayouts/slideLayout1.xml",
        layout_xml,
    );
    package
}

fn replace_part(package: &mut Package, part_name: &str, xml: &[u8]) {
    let part_name = PackURI::new(part_name).unwrap();
    package
        .edit_opc(|opc| {
            opc.get_part_mut(&part_name)?.set_blob(xml.to_vec());
            Ok(())
        })
        .unwrap();
}
