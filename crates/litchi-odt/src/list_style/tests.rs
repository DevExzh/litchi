//! Focused regression tests for the layered list-style owner.

use super::{BulletStyle, ImageSource, Kind, LevelStyle, Style, parse};
use litchi_odf_common::XmlStreamLimits;

const OFFICE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const STYLE: &str = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
const TEXT: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const XLINK: &str = "http://www.w3.org/1999/xlink";

fn styles(xml: &str) -> String {
    format!(
        r#"<office:styles xmlns:office="{OFFICE}" xmlns:style="{STYLE}" xmlns:text="{TEXT}" xmlns:xlink="{XLINK}">{xml}</office:styles>"#
    )
}

#[test]
fn serializes_and_reparses_a_linked_image_level() {
    let mut style = Style::new("Bullets").unwrap();
    style.levels.push(LevelStyle {
        level: 1,
        style_name: None,
        kind: Kind::Image(ImageSource::Linked("Pictures/bullet.png".to_string())),
    });

    let xml = style
        .to_xml_fragment(&XmlStreamLimits::owner_default_value())
        .unwrap();
    let exact = XmlStreamLimits::new(xml.len() as u64, 256, 100_000, 100_000, 64 * 1024).unwrap();
    assert_eq!(style.to_xml_fragment(&exact).unwrap(), xml);
    let parsed = parse(&styles(&xml), &XmlStreamLimits::owner_default_value()).unwrap();
    assert_eq!(parsed.get("Bullets").unwrap(), &style);
    let shallow =
        XmlStreamLimits::new(xml.len() as u64 + 1, 2, 100_000, 100_000, 64 * 1024).unwrap();
    assert!(parse(&styles(&xml), &shallow).is_err());
}

#[test]
fn rejects_control_bullets_before_encoding() {
    assert!(BulletStyle::new('\n').is_err());
}

#[test]
fn list_writer_honors_tiny_output_budget() {
    let mut style = Style::new("Bullets").unwrap();
    style.levels.push(LevelStyle {
        level: 1,
        style_name: None,
        kind: Kind::Bullet(BulletStyle::new('•').unwrap()),
    });
    let tiny = XmlStreamLimits::new(1, 256, 100_000, 100_000, 64 * 1024).unwrap();
    assert!(style.to_xml_fragment(&tiny).is_err());
}

#[test]
fn list_writer_rejects_a_large_level_collection_before_building_output() {
    let data = "A".repeat(64 * 1024);
    let mut style = Style::new("Images").unwrap();
    for level in 1..=256 {
        style.levels.push(LevelStyle {
            level,
            style_name: None,
            kind: Kind::Image(ImageSource::Embedded(data.clone())),
        });
    }
    let limits = XmlStreamLimits::new(1024 * 1024, 256, 100_000, 100_000, 64 * 1024).unwrap();
    assert!(style.to_xml_fragment(&limits).is_err());
}
