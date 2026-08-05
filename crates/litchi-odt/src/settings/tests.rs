use super::codec::{DocumentKind, is_datetime, parse};
use super::model::{ConfigItem, ConfigNode, ConfigValue};

const PREFIX: &str = r#"<office:document-settings
        xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
        xmlns:c="urn:oasis:names:tc:opendocument:xmlns:config:1.0"><office:settings>"#;

#[test]
fn parses_typed_sets_and_maps_with_arbitrary_prefix() {
    let xml = format!(
        r#"{PREFIX}<c:config-item-set c:name="ViewSettings">
                <c:config-item c:name="ShowGrid" c:type="boolean">true</c:config-item>
                <c:config-item c:name="Zoom" c:type="int">125</c:config-item>
                <c:config-item-map-named c:name="Views">
                    <c:config-item-map-entry c:name="main">
                        <c:config-item c:name="Label" c:type="string">A&amp;B</c:config-item>
                    </c:config-item-map-entry>
                </c:config-item-map-named>
                <c:config-item-map-indexed c:name="Windows">
                    <c:config-item-map-entry/>
                </c:config-item-map-indexed>
            </c:config-item-set></office:settings></office:document-settings>"#
    );
    let settings = parse(&xml, DocumentKind::Package).unwrap();
    assert_eq!(settings.sets.len(), 1);
    assert_eq!(settings.sets[0].name, "ViewSettings");
    assert_eq!(
        settings.sets[0].children[0],
        ConfigNode::Item(ConfigItem {
            name: "ShowGrid".to_string(),
            value: ConfigValue::Boolean(true),
        })
    );
    let ConfigNode::NamedMap(map) = &settings.sets[0].children[2] else {
        panic!("expected named map");
    };
    assert_eq!(map.entries[0].name.as_deref(), Some("main"));
}

#[test]
fn accepts_empty_flat_settings_and_missing_flat_settings() {
    let empty = r#"<o:document xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><o:settings/></o:document>"#;
    assert!(parse(empty, DocumentKind::Flat).unwrap().sets.is_empty());
    let missing = r#"<o:document xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><o:body>ordinary flat document text</o:body></o:document>"#;
    assert!(parse(missing, DocumentKind::Flat).unwrap().sets.is_empty());
}

#[test]
fn rejects_invalid_types_placement_and_active_xml() {
    let cases = [
        format!(r#"{PREFIX}<c:config-item c:name="x" c:type="int">1</c:config-item></office:settings></office:document-settings>"#),
        format!(r#"{PREFIX}<c:config-item-set c:name="x"><c:config-item c:name="x" c:type="int">NaN</c:config-item></c:config-item-set></office:settings></office:document-settings>"#),
        format!(r#"{PREFIX}<c:config-item-set c:name="x"><c:config-item-map-named c:name="m"><c:config-item-map-entry/></c:config-item-map-named></c:config-item-set></office:settings></office:document-settings>"#),
        r#"<?xml version="1.0"?><!DOCTYPE x><office:document-settings xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><office:settings/></office:document-settings>"#.to_string(),
    ];
    for xml in cases {
        assert!(parse(&xml, DocumentKind::Package).is_err());
    }
}

#[test]
fn validates_xml_schema_datetime_edges() {
    for valid in [
        "2000-02-29T24:00:00Z",
        "2024-02-29T23:59:59.125+14:00",
        "-0001-01-01T00:00:00-08:30",
    ] {
        assert!(is_datetime(valid), "rejected {valid}");
    }
    for invalid in [
        "0000-01-01T00:00:00",
        "2023-02-29T00:00:00",
        "2024-01-01T24:00:00.1",
        "2024-01-01T00:00:60",
        "2024-01-01T00:00:00+14:01",
    ] {
        assert!(!is_datetime(invalid), "accepted {invalid}");
    }
}
