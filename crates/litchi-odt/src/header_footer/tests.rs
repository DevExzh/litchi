//! Focused regression tests for the layered header/footer owner.

use super::properties::{Color, Length, Region, StyleProperties};

#[test]
fn property_model_round_trips_through_the_canonical_owner() {
    let properties = StyleProperties {
        height: Some(Length::new("1.25cm").unwrap()),
        background_color: Some(Color::Rgb(0x10, 0x20, 0x30)),
        dynamic_spacing: Some(true),
        ..Default::default()
    };
    let fragment = properties.to_region_fragment(Region::Header).unwrap();
    let xml = format!(
        r#"<office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"><office:automatic-styles><style:page-layout style:name="layout">{fragment}</style:page-layout></office:automatic-styles></office:document-styles>"#
    );
    let entry = super::parse_page_layout_header_footer_properties(&xml)
        .unwrap()
        .pop()
        .unwrap();
    assert_eq!(entry.region, Region::Header);
    assert_eq!(entry.properties, properties);
}
