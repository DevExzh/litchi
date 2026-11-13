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

/// Runs `run` over the XML envelope orderings built around one root element
/// and compares each `label => outcome` line with the frozen expectation.
fn assert_envelope_matrix(
    open: &str,
    close: &str,
    empty: &str,
    expected: &[&str],
    run: impl Fn(&str) -> String,
) {
    const DECL: &str = r#"<?xml version="1.0" encoding="UTF-8"?>"#;
    let doc = format!("{open}{close}");
    let rows = [
        ("empty input", String::new()),
        ("whitespace only", " \n".to_owned()),
        ("declaration only", DECL.to_owned()),
        ("empty-element root", empty.to_owned()),
        ("start and end root", doc.clone()),
        ("declaration then root", format!("{DECL}{doc}")),
        ("declaration, whitespace, root", format!("{DECL}\n{doc}")),
        ("root then whitespace", format!("{doc}\n ")),
        ("root then comment", format!("{doc}<!--c-->")),
        ("root then PI", format!("{doc}<?pi x?>")),
        ("root then text", format!("{doc}text")),
        ("root then second root", format!("{doc}{doc}")),
        (
            "root then foreign start",
            format!("{doc}<foreign></foreign>"),
        ),
        ("root then foreign empty", format!("{doc}<foreign/>")),
        (
            "empty-element root then second root",
            format!("{empty}{doc}"),
        ),
        ("empty-element root then text", format!("{empty}text")),
        ("declaration after root", format!("{doc}{DECL}")),
        ("two declarations", format!("{DECL}{DECL}{doc}")),
        ("comment before declaration", format!("<!--c-->{DECL}{doc}")),
        ("comment before root", format!("<!--c-->{doc}")),
        ("PI before root", format!("<?pi x?>{doc}")),
        ("text before root", format!("text{doc}")),
        ("whitespace before declaration", format!(" {DECL}{doc}")),
        ("unterminated root", open.to_owned()),
        ("stray end tag", format!("{doc}</stray>")),
        ("DOCTYPE", format!("{DECL}<!DOCTYPE r>{doc}")),
        ("DOCTYPE after root", format!("{doc}<!DOCTYPE r>")),
        ("top-level CDATA", format!("<![CDATA[x]]>{doc}")),
        ("CDATA after root", format!("{doc}<![CDATA[x]]>")),
        ("entity reference before root", format!("&amp;{doc}")),
        ("entity reference after root", format!("{doc}&amp;")),
    ];
    let actual: Vec<String> = rows
        .iter()
        .map(|(label, xml)| format!("{label} => {}", run(xml)))
        .collect();
    let listing: Vec<String> = actual.iter().map(|line| format!("{line:?},")).collect();
    assert!(
        actual
            .iter()
            .map(String::as_str)
            .eq(expected.iter().copied()),
        "actual envelope matrix:\n{}",
        listing.join("\n")
    );
}

#[test]
fn settings_document_envelope_matrix() {
    let root = "office:document-settings xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\"";
    assert_envelope_matrix(
        &format!("<{root}><office:settings></office:settings>"),
        "</office:document-settings>",
        &format!("<{root}/>"),
        &[
            "empty input => InvalidFormat(\"incomplete settings XML document\")",
            "whitespace only => InvalidFormat(\"incomplete settings XML document\")",
            "declaration only => InvalidFormat(\"incomplete settings XML document\")",
            "empty-element root => InvalidFormat(\"settings.xml has no office:settings element\")",
            "start and end root => Ok(Settings { sets: [] })",
            "declaration then root => Ok(Settings { sets: [] })",
            "declaration, whitespace, root => Ok(Settings { sets: [] })",
            "root then whitespace => Ok(Settings { sets: [] })",
            "root then comment => Ok(Settings { sets: [] })",
            "root then PI => InvalidFormat(\"DOCTYPE and processing instructions are prohibited in settings XML\")",
            "root then text => InvalidFormat(\"unexpected text outside the settings document element\")",
            "root then second root => InvalidFormat(\"settings XML has more than one document element\")",
            "root then foreign start => InvalidFormat(\"settings XML has more than one document element\")",
            "root then foreign empty => InvalidFormat(\"settings XML has more than one document element\")",
            "empty-element root then second root => InvalidFormat(\"settings XML has more than one document element\")",
            "empty-element root then text => InvalidFormat(\"unexpected text outside the settings document element\")",
            "declaration after root => InvalidFormat(\"XML declaration must precede the document element\")",
            "two declarations => Ok(Settings { sets: [] })",
            "comment before declaration => Ok(Settings { sets: [] })",
            "comment before root => Ok(Settings { sets: [] })",
            "PI before root => InvalidFormat(\"DOCTYPE and processing instructions are prohibited in settings XML\")",
            "text before root => InvalidFormat(\"unexpected text outside the settings document element\")",
            "whitespace before declaration => Ok(Settings { sets: [] })",
            "unterminated root => InvalidFormat(\"incomplete settings XML document\")",
            "stray end tag => InvalidFormat(\"invalid settings XML: ill-formed document: close tag `</stray>` does not match any open tag\")",
            "DOCTYPE => InvalidFormat(\"DOCTYPE and processing instructions are prohibited in settings XML\")",
            "DOCTYPE after root => InvalidFormat(\"DOCTYPE and processing instructions are prohibited in settings XML\")",
            "top-level CDATA => InvalidFormat(\"unexpected text outside the settings document element\")",
            "CDATA after root => InvalidFormat(\"unexpected text outside the settings document element\")",
            "entity reference before root => InvalidFormat(\"unexpected text outside the settings document element\")",
            "entity reference after root => InvalidFormat(\"unexpected text outside the settings document element\")",
        ],
        |xml| match parse(xml, DocumentKind::Package) {
            Ok(value) => format!("Ok({value:?})"),
            Err(error) => format!("{error:?}"),
        },
    );
}
