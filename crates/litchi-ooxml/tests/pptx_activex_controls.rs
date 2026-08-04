use litchi_ooxml::pptx::{
    ControlBinary, ControlDescriptor, ControlPersistence, Package, SlideControl,
};

const ACTIVEX_CHECKBOX: &str = concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/../../test-data/ooxml/pptx/activex/activex_checkbox.pptx"
);

#[test]
fn package_inventory_reports_activex_controls() {
    let package = Package::open(ACTIVEX_CHECKBOX).unwrap();

    let controls = package.controls().unwrap();
    assert_eq!(controls.len(), 3);

    let first: &SlideControl = &controls[0];
    assert_eq!(first.slide_index(), 0);
    assert_eq!(first.control_index(), 0);
    assert_eq!(first.shape_id(), Some("1031"));
    assert_eq!(first.name(), Some("CheckBox1"));
    assert_eq!(first.show_as_icon(), None);
    assert_eq!(first.image_width(), Some(2_685_960));
    assert_eq!(first.image_height(), Some(923_760));
    assert_eq!(first.relationship_id(), Some("rId2"));

    let descriptor: &ControlDescriptor = first.descriptor().unwrap();
    assert_eq!(descriptor.part_name().as_str(), "/ppt/activeX/activeX1.xml");
    assert_eq!(
        descriptor.class_id(),
        "{8BD21D40-EC42-11CE-9E0D-00AA006002F3}"
    );
    assert_eq!(descriptor.license(), None);
    assert_eq!(descriptor.persistence(), ControlPersistence::Storage);

    let binary: &ControlBinary = descriptor.binary().unwrap();
    assert_eq!(binary.relationship_id(), "rId1");
    assert_eq!(binary.part_name().as_str(), "/ppt/activeX/activeX1.bin");
    assert!(binary.byte_length() > 0);

    let second = &controls[1];
    assert_eq!(second.control_index(), 1);
    assert_eq!(second.shape_id(), Some("1032"));
    assert_eq!(second.name(), Some("CheckBox2"));
    assert_eq!(second.relationship_id(), Some("rId3"));
    assert_eq!(
        second.descriptor().unwrap().part_name().as_str(),
        "/ppt/activeX/activeX2.xml"
    );

    let third = &controls[2];
    assert_eq!(third.control_index(), 2);
    assert_eq!(third.name(), Some("CheckBox3"));
    assert_eq!(third.relationship_id(), Some("rId4"));
    assert_eq!(
        third.descriptor().unwrap().part_name().as_str(),
        "/ppt/activeX/activeX3.xml"
    );
    assert!(third.descriptor().unwrap().binary().is_some());
}

#[test]
fn presentation_controls_match_package_inventory() {
    let package = Package::open(ACTIVEX_CHECKBOX).unwrap();

    let from_package = package.controls().unwrap();
    let from_presentation = package.presentation().unwrap().controls().unwrap();
    assert_eq!(from_package.len(), from_presentation.len());
    for (a, b) in from_package.iter().zip(from_presentation.iter()) {
        assert_eq!(a.slide_index(), b.slide_index());
        assert_eq!(a.control_index(), b.control_index());
        assert_eq!(a.name(), b.name());
        assert_eq!(a.relationship_id(), b.relationship_id());
    }
}

#[test]
fn slides_without_controls_yield_empty_inventory() {
    let package = Package::open(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../test-data/ooxml/pptx/sample.pptx"
    ))
    .unwrap();
    assert!(package.controls().unwrap().is_empty());
}
