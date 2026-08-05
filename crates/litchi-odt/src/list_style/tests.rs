//! Focused regression tests for the layered list-style owner.

use super::{BulletStyle, ImageSource, Kind, LevelStyle, Style, parse};

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

    let xml = style.to_xml_fragment().unwrap();
    let parsed = parse(&styles(&xml)).unwrap();
    assert_eq!(parsed.get("Bullets").unwrap(), &style);
}

#[test]
fn rejects_control_bullets_before_encoding() {
    assert!(BulletStyle::new('\n').is_err());
}
