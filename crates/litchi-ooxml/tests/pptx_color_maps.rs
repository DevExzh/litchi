use litchi_ooxml::PackURI;
use litchi_ooxml::pptx::{ColorMapOverride, ColorMapSlot, Package, ThemeColorRole};
use tempfile::NamedTempFile;

const MASTER_XML: &[u8] = include_bytes!("../../../test-data/ooxml/pptx/color-maps/master.xml");
const LAYOUT_XML: &[u8] =
    include_bytes!("../../../test-data/ooxml/pptx/color-maps/layout_override.xml");
const SLIDE_MASTER_MAPPING_XML: &[u8] =
    include_bytes!("../../../test-data/ooxml/pptx/color-maps/slide_master_mapping.xml");
const SLIDE_OVERRIDE_XML: &[u8] =
    include_bytes!("../../../test-data/ooxml/pptx/color-maps/slide_override.xml");
const SLIDE_WITHOUT_OVERRIDE_XML: &[u8] =
    include_bytes!("../../../test-data/ooxml/pptx/color-maps/slide_without_override.xml");

#[test]
fn master_mapping_is_used_when_a_slide_requests_it() {
    let package = package_with_color_maps(SLIDE_MASTER_MAPPING_XML);
    let presentation = package.presentation().unwrap();
    let master = presentation.slide_masters().unwrap().remove(0);
    assert_eq!(
        master.color_map().unwrap().color(ColorMapSlot::Background1),
        ThemeColorRole::Dark1
    );
    assert_eq!(
        master
            .theme_color(ColorMapSlot::Background1)
            .unwrap()
            .unwrap()
            .name,
        "dk1"
    );

    let slide = presentation.slides().unwrap().remove(0);
    let layout = slide.layout().unwrap();
    assert!(matches!(
        layout.color_map_override().unwrap(),
        Some(ColorMapOverride::Override(map))
            if map.color(ColorMapSlot::Background1) == ThemeColorRole::Accent1
    ));
    assert_eq!(
        layout
            .effective_color_map()
            .unwrap()
            .color(ColorMapSlot::Background1),
        ThemeColorRole::Accent1
    );
    assert_eq!(
        slide.color_map_override().unwrap(),
        Some(ColorMapOverride::Master)
    );
    assert_eq!(
        slide
            .effective_color_map()
            .unwrap()
            .color(ColorMapSlot::Background1),
        ThemeColorRole::Dark1
    );
    assert_eq!(
        slide
            .effective_theme_color(ColorMapSlot::Background1)
            .unwrap()
            .unwrap()
            .name,
        "dk1"
    );
}

#[test]
fn slide_override_wins_and_absent_slide_mapping_uses_layout() {
    let package = package_with_color_maps(SLIDE_OVERRIDE_XML);
    let presentation = package.presentation().unwrap();
    let slide = presentation.slides().unwrap().remove(0);
    assert!(matches!(
        slide.color_map_override().unwrap(),
        Some(ColorMapOverride::Override(map))
            if map.color(ColorMapSlot::Background1) == ThemeColorRole::Accent2
    ));
    assert_eq!(
        slide
            .effective_color_map()
            .unwrap()
            .color(ColorMapSlot::Background1),
        ThemeColorRole::Accent2
    );
    assert_eq!(
        slide
            .effective_theme_color(ColorMapSlot::Background1)
            .unwrap()
            .unwrap()
            .name,
        "accent2"
    );

    let package = package_with_color_maps(SLIDE_WITHOUT_OVERRIDE_XML);
    let presentation = package.presentation().unwrap();
    let slide = presentation.slides().unwrap().remove(0);
    assert_eq!(slide.color_map_override().unwrap(), None);
    assert_eq!(
        slide
            .effective_color_map()
            .unwrap()
            .color(ColorMapSlot::Background1),
        ThemeColorRole::Accent1
    );
}

fn package_with_color_maps(slide_xml: &[u8]) -> Package {
    let output = NamedTempFile::with_suffix(".pptx").unwrap();
    let mut package = Package::new().unwrap();
    package.presentation_mut().unwrap().add_slide().unwrap();
    package.save(output.path()).unwrap();

    let mut package = Package::open(output.path()).unwrap();
    replace_part(
        &mut package,
        "/ppt/slideMasters/slideMaster1.xml",
        MASTER_XML,
    );
    replace_part(
        &mut package,
        "/ppt/slideLayouts/slideLayout1.xml",
        LAYOUT_XML,
    );
    replace_part(&mut package, "/ppt/slides/slide1.xml", slide_xml);
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
