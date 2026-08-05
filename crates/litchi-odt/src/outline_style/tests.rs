//! Focused regression tests for the outline-style model and codec.

use super::{
    Attribute, LevelStyle, NumberFormat, PositiveInteger, Style, TextAlign, parse_outline_styles,
};

const OFFICE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const TEXT: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const STYLE: &str = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
const FO: &str = "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0";

fn document(outline: &str) -> String {
    format!(
        r#"<o:document xmlns:o="{OFFICE}" xmlns:t="{TEXT}" xmlns:s="{STYLE}" xmlns:f="{FO}"><o:styles>{outline}</o:styles><o:body><o:text/></o:body></o:document>"#
    )
}

#[test]
fn typed_model_serializes_and_reparses() {
    let style = Style {
        name: "Outline".to_string(),
        levels: vec![LevelStyle {
            level: 1,
            text_style_name: Some("Heading_20_1".to_string()),
            number_format: Some(NumberFormat::new("A").unwrap()),
            number_prefix: Some("(".to_string()),
            number_suffix: Some(")".to_string()),
            letter_sync: Some(true),
            display_levels: Some(PositiveInteger::new("2").unwrap()),
            start_value: Some(PositiveInteger::new("1").unwrap()),
            list_level_properties: None,
            text_properties: None,
            extensions: vec![Attribute::new("urn:producer", "num-list-format", "%1%").unwrap()],
        }],
        extensions: Vec::new(),
    };
    let xml = style.to_xml().unwrap();
    let parsed = parse_outline_styles(&document(&xml)).unwrap();
    assert_eq!(parsed.get("Outline"), Some(&style));
}

#[test]
fn invalid_typed_model_is_rejected_before_writing() {
    let style = Style {
        name: "Outline".to_string(),
        levels: vec![LevelStyle {
            level: 1,
            text_style_name: None,
            number_format: Some(NumberFormat::new("1").unwrap()),
            number_prefix: None,
            number_suffix: None,
            letter_sync: Some(true),
            display_levels: None,
            start_value: None,
            list_level_properties: None,
            text_properties: None,
            extensions: Vec::new(),
        }],
        extensions: Vec::new(),
    };
    assert!(style.to_xml().is_err());
    assert_eq!(TextAlign::Start.as_str(), "start");
}
