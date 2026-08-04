use litchi_ooxml::PackURI;
use litchi_ooxml::pptx::{Effect, Package};
use tempfile::NamedTempFile;

const LAYOUT_XML: &[u8] =
    include_bytes!("../../../test-data/ooxml/pptx/master-animations/layout.xml");
const MASTER_XML: &[u8] =
    include_bytes!("../../../test-data/ooxml/pptx/master-animations/master.xml");

#[test]
fn layout_and_master_timing_metadata_is_exposed() {
    let package = package_with_timing_xml();
    let presentation = package.presentation().unwrap();
    let slide = presentation.slides().unwrap().remove(0);

    let layout_animations = slide.layout().unwrap().animations().unwrap();
    assert_eq!(layout_animations.len(), 1);
    assert_eq!(layout_animations.animations[0].shape_id, 3);
    assert_eq!(layout_animations.animations[0].effect, Effect::Fade);

    let master_animations = slide.master().unwrap().animations().unwrap();
    assert_eq!(master_animations.len(), 1);
    assert_eq!(master_animations.animations[0].shape_id, 4);
    assert_eq!(master_animations.animations[0].effect, Effect::Fade);
}

fn package_with_timing_xml() -> Package {
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
            opc.get_part_mut(&part_name)?.set_blob(xml.to_vec());
            Ok(())
        })
        .unwrap();
}
