use litchi_pptx::presentation_properties::metadata::color_map::{
    Map, Override, Role, Slot, parse_master, parse_override,
};

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
    let master = parse_master(MASTER_XML).unwrap();
    assert_eq!(master.color(Slot::Background1), Role::Dark1);
    assert_eq!(Role::Dark1.as_str(), "dk1");

    let layout = parse_override(LAYOUT_XML, b"sldLayout", "slide layout")
        .unwrap()
        .unwrap();
    assert!(matches!(
        layout,
        Override::Override(map) if map.color(Slot::Background1) == Role::Accent1
    ));

    let slide = parse_override(SLIDE_MASTER_MAPPING_XML, b"sld", "slide")
        .unwrap()
        .unwrap();
    assert_eq!(slide, Override::Master);

    assert_eq!(
        effective_map(Some(layout), Some(slide), master).color(Slot::Background1),
        Role::Dark1
    );
}

#[test]
fn slide_override_wins_and_absent_slide_mapping_uses_layout() {
    let master = parse_master(MASTER_XML).unwrap();
    let layout = parse_override(LAYOUT_XML, b"sldLayout", "slide layout")
        .unwrap()
        .unwrap();

    let slide = parse_override(SLIDE_OVERRIDE_XML, b"sld", "slide")
        .unwrap()
        .unwrap();
    assert!(matches!(
        slide,
        Override::Override(map) if map.color(Slot::Background1) == Role::Accent2
    ));
    assert_eq!(
        effective_map(Some(layout), Some(slide), master).color(Slot::Background1),
        Role::Accent2
    );

    let slide_without_override =
        parse_override(SLIDE_WITHOUT_OVERRIDE_XML, b"sld", "slide").unwrap();
    assert_eq!(slide_without_override, None);
    assert_eq!(
        effective_map(Some(layout), slide_without_override, master)
            .color(Slot::Background1),
        Role::Accent1
    );
}

fn effective_map(layout: Option<Override>, slide: Option<Override>, master: Map) -> Map {
    match slide {
        Some(Override::Override(map)) => map,
        Some(Override::Master) => master,
        None => match layout {
            Some(Override::Override(map)) => map,
            Some(Override::Master) | None => master,
        },
    }
}
