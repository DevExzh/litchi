use litchi_odt::style::paragraph::drop_cap::{Distance, DropCap, Length, Style, parse};
use litchi_odt::style::paragraph::tab_stop::{Position, Stop, Stops, Style as TabStyle};
use litchi_odt::{Builder, Document};
use std::io::Cursor;
const OFFICE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const STYLE: &str = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
fn wrap(body: &str) -> String {
    format!(r#"<o:styles xmlns:o="{OFFICE}" xmlns:s="{STYLE}">{body}</o:styles>"#)
}

#[test]
fn parses_aliases_all_values_inheritance_and_round_trip() {
    let xml = wrap(concat!(
        r#"<s:default-style s:family="paragraph"><s:paragraph-properties><s:drop-cap s:lines="2"/></s:paragraph-properties></s:default-style>"#,
        r#"<s:style s:name="Parent" s:family="paragraph"><s:paragraph-properties><s:drop-cap s:style-name="Drop Character" s:distance="0.15in" s:lines="3" s:length="word"></s:drop-cap></s:paragraph-properties></s:style>"#,
        r#"<s:style s:name="Child" s:family="paragraph" s:parent-style-name="Parent"/>"#
    ));
    let parsed = parse(&xml).unwrap();
    assert_eq!(parsed.styles.len(), 3);
    let parent = parsed.get("Parent").unwrap();
    let cap = parent.drop_cap.as_ref().unwrap();
    assert_eq!(cap.length, Some(Length::Word));
    assert_eq!(cap.lines, Some(3));
    assert_eq!(cap.distance.as_ref().unwrap().as_str(), "0.15in");
    assert_eq!(parsed.resolved_drop_cap("Child").unwrap(), Some(cap));
    let fragment = parent.to_xml_fragment().unwrap();
    let reparsed = parse(&wrap(&fragment)).unwrap();
    assert_eq!(reparsed.get("Parent"), Some(parent));
}

#[test]
fn parses_real_libreoffice_fixture() {
    let bytes = include_bytes!(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../test-data/libreoffice-core/sw/qa/extras/layout/data/drop_asian_word.fodt"
    ));
    let flat = litchi_odt::generic::FlatDocument::from_reader(Cursor::new(bytes)).unwrap();
    let parsed = flat.paragraph_style_drop_caps().unwrap();
    assert!(parsed.styles.iter().any(|style| {
        style
            .drop_cap
            .as_ref()
            .is_some_and(|cap| cap.lines == Some(3) && cap.length == Some(Length::Word))
    }));
}

#[test]
fn rejects_malformed_namespace_structure_cardinality_and_values() {
    let invalid = [
        wrap(r#"<s:style s:name="x" s:family="paragraph"><s:drop-cap/></s:style>"#),
        wrap(
            r#"<s:style s:name="x" s:family="text"><s:paragraph-properties><s:drop-cap/></s:paragraph-properties></s:style>"#,
        ),
        wrap(
            r#"<s:style s:name="x" s:family="paragraph"><s:paragraph-properties><s:drop-cap/><s:drop-cap/></s:paragraph-properties></s:style>"#,
        ),
        wrap(
            r#"<s:style s:name="x" s:family="paragraph"><s:paragraph-properties/><s:paragraph-properties><s:drop-cap/></s:paragraph-properties></s:style>"#,
        ),
        wrap(
            r#"<s:style s:name="x" s:family="paragraph"><s:paragraph-properties><s:drop-cap> </s:drop-cap></s:paragraph-properties></s:style>"#,
        ),
        wrap(
            r#"<s:style s:name="x" s:family="paragraph"><s:paragraph-properties><s:drop-cap s:lines="0"/></s:paragraph-properties></s:style>"#,
        ),
        wrap(
            r#"<s:style s:name="x" s:family="paragraph"><s:paragraph-properties><s:drop-cap s:length="0"/></s:paragraph-properties></s:style>"#,
        ),
        wrap(
            r#"<s:style s:name="x" s:family="paragraph"><s:paragraph-properties><s:drop-cap s:distance="1em"/></s:paragraph-properties></s:style>"#,
        ),
        wrap(
            r#"<s:style s:name="x" s:family="paragraph"><s:paragraph-properties><s:drop-cap s:unknown="x"/></s:paragraph-properties></s:style>"#,
        ),
        format!(
            r#"<o:styles xmlns:o="{OFFICE}" xmlns:s="{STYLE}" xmlns:x="urn:wrong"><s:style s:name="x" s:family="paragraph"><s:paragraph-properties><x:drop-cap/></s:paragraph-properties></s:style></o:styles>"#
        ),
        format!(
            r#"<!DOCTYPE x><o:styles xmlns:o="{OFFICE}" xmlns:s="{STYLE}"><s:style s:name="x" s:family="paragraph"><s:paragraph-properties><s:drop-cap/></s:paragraph-properties></s:style></o:styles>"#
        ),
    ];
    for xml in invalid {
        assert!(parse(&xml).is_err(), "accepted {xml}");
    }
}

#[test]
fn builder_package_composition_and_mutation_round_trip() {
    let cap = DropCap {
        length: Some(Length::Characters(1)),
        lines: Some(3),
        distance: Some(Distance::new("0.2cm").unwrap()),
        style_name: Some("DropLetter".into()),
    };
    let mut drop_style = Style::named("Opening", Some(cap)).unwrap();
    drop_style.parent_style_name = Some("Standard".into());
    let stops = Stops::try_from_vec(vec![Stop::new(Position::new("2cm").unwrap())]).unwrap();
    let mut tab_style = TabStyle::named("Opening", Some(stops)).unwrap();
    tab_style.parent_style_name = Some("Standard".into());
    let mut builder = Builder::new();
    builder.add_paragraph_tab_style(tab_style).unwrap();
    builder
        .add_paragraph_drop_cap_style(drop_style.clone())
        .unwrap();
    builder.add_paragraph("Opening text").unwrap();
    let bytes = builder.build().unwrap();
    let package = litchi_odt::generic::Package::from_bytes(bytes.clone()).unwrap();
    assert_eq!(
        package.paragraph_style_drop_caps().unwrap().get("Opening"),
        Some(&drop_style)
    );
    assert!(
        package
            .paragraph_style_tab_stops()
            .unwrap()
            .get("Opening")
            .is_some()
    );
    let mut mutable =
        litchi_odt::mutable::MutableDocument::from_document(Document::from_bytes(bytes).unwrap())
            .unwrap();
    let replacement = DropCap {
        length: Some(Length::Word),
        lines: Some(4),
        distance: None,
        style_name: None,
    };
    drop_style.drop_cap = Some(replacement.clone());
    mutable.set_paragraph_style_drop_cap(&drop_style).unwrap();
    let package = litchi_odt::generic::Package::from_bytes(mutable.to_bytes().unwrap()).unwrap();
    assert_eq!(
        package
            .paragraph_style_drop_caps()
            .unwrap()
            .get("Opening")
            .unwrap()
            .drop_cap
            .as_ref(),
        Some(&replacement)
    );
}
