use litchi_odt::Builder;
use litchi_odt::style::paragraph::writing_mode::{Properties, Style, WritingMode, parse, set_xml};
use std::io::Cursor;
const O: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const S: &str = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
const F: &str = "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0";
fn wrap(x: &str) -> String {
    format!(r#"<o:styles xmlns:o="{O}" xmlns:s="{S}" xmlns:f="{F}">{x}</o:styles>"#)
}

fn full_properties() -> &'static str {
    r#"<s:style s:name="P1" s:family="paragraph"><s:paragraph-properties s:writing-mode="tb-rl" s:writing-mode-automatic="true" s:register-true="false" s:join-border="true" f:text-align="justify"/></s:style>"#
}

#[test]
fn parses_all_writing_mode_attributes() {
    let set = parse(&wrap(full_properties())).unwrap();
    let style = set.get("P1").unwrap();
    let p = style.properties.as_ref().unwrap();
    assert_eq!(p.writing_mode, Some(WritingMode::TbRl));
    assert_eq!(p.writing_mode_automatic, Some(true));
    assert_eq!(p.register_true, Some(false));
    assert_eq!(p.join_border, Some(true));
}

#[test]
fn ignores_sibling_owned_attributes_and_foreign_families() {
    let xml = wrap(
        r#"<s:style s:name="P" s:family="paragraph"><s:paragraph-properties f:text-align="start" s:writing-mode="lr"/></s:style><s:style s:name="T" s:family="table"><s:paragraph-properties s:writing-mode="rl"/></s:style>"#,
    );
    let set = parse(&xml).unwrap();
    assert_eq!(set.styles.len(), 1);
    let p = set.get("P").unwrap().properties.as_ref().unwrap();
    assert_eq!(p.writing_mode, Some(WritingMode::Lr));
    assert!(p.writing_mode_automatic.is_none());
    assert!(set.get("T").is_none());
}

#[test]
fn parses_every_token() {
    for (token, expect) in [
        ("lr-tb", WritingMode::LrTb),
        ("rl-tb", WritingMode::RlTb),
        ("tb-rl", WritingMode::TbRl),
        ("tb-lr", WritingMode::TbLr),
        ("lr", WritingMode::Lr),
        ("rl", WritingMode::Rl),
        ("tb", WritingMode::Tb),
        ("page", WritingMode::Page),
    ] {
        let xml = wrap(&format!(
            r#"<s:style s:name="P" s:family="paragraph"><s:paragraph-properties s:writing-mode="{token}"/></s:style>"#
        ));
        let set = parse(&xml).unwrap();
        assert_eq!(
            set.get("P")
                .unwrap()
                .properties
                .as_ref()
                .unwrap()
                .writing_mode,
            Some(expect)
        );
    }
}

#[test]
fn round_trip_serialize_reparse() {
    let set = parse(&wrap(full_properties())).unwrap();
    let style = set.get("P1").unwrap();
    let fragment = style.to_xml_fragment().unwrap();
    let reparsed = parse(&wrap(&fragment)).unwrap();
    assert_eq!(reparsed.get("P1"), Some(style));
    let default = Style::default_style(Some(Properties {
        writing_mode: Some(WritingMode::Page),
        register_true: Some(true),
        ..Default::default()
    }));
    let fragment = default.to_xml_fragment().unwrap();
    let reparsed = parse(&wrap(&fragment)).unwrap();
    assert_eq!(reparsed.default_style(), Some(&default));
}

#[test]
fn rejects_malformed_values_and_duplicates() {
    let bad = [
        // Unknown style:writing-mode token.
        wrap(
            r#"<s:style s:name="P" s:family="paragraph"><s:paragraph-properties s:writing-mode="sideways"/></s:style>"#,
        ),
        // XSL relative tokens are not part of the ODF paragraph token set.
        wrap(
            r#"<s:style s:name="P" s:family="paragraph"><s:paragraph-properties s:writing-mode="bt-lr"/></s:style>"#,
        ),
        // Case matters.
        wrap(
            r#"<s:style s:name="P" s:family="paragraph"><s:paragraph-properties s:writing-mode="LR-TB"/></s:style>"#,
        ),
        // Boolean attributes accept only the RNG true/false tokens.
        wrap(
            r#"<s:style s:name="P" s:family="paragraph"><s:paragraph-properties s:writing-mode-automatic="1"/></s:style>"#,
        ),
        wrap(
            r#"<s:style s:name="P" s:family="paragraph"><s:paragraph-properties s:register-true="yes"/></s:style>"#,
        ),
        wrap(
            r#"<s:style s:name="P" s:family="paragraph"><s:paragraph-properties s:join-border="TRUE"/></s:style>"#,
        ),
        // Duplicate owned attribute (same expanded name).
        wrap(
            r#"<s:style s:name="P" s:family="paragraph"><s:paragraph-properties s:writing-mode="lr" s:writing-mode="rl"/></s:style>"#,
        ),
        // Duplicate properties element.
        wrap(
            r#"<s:style s:name="P" s:family="paragraph"><s:paragraph-properties/><s:paragraph-properties/></s:style>"#,
        ),
        // Duplicate style identity.
        wrap(
            r#"<s:style s:name="P" s:family="paragraph"><s:paragraph-properties/></s:style><s:style s:name="P" s:family="paragraph"><s:paragraph-properties/></s:style>"#,
        ),
        // Broken identity: default style with a name.
        wrap(
            r#"<s:default-style s:name="P" s:family="paragraph"><s:paragraph-properties/></s:default-style>"#,
        ),
    ];
    for xml in bad {
        assert!(parse(&xml).is_err(), "accepted {xml}");
    }
}

#[test]
fn mutation_replaces_inserts_and_strips_owned_attributes() {
    let xml = wrap(full_properties());
    // Replace: owned attributes are rewritten, sibling attributes survive.
    let changed = Style::named(
        "P1",
        Some(Properties {
            writing_mode: Some(WritingMode::LrTb),
            join_border: Some(false),
            ..Default::default()
        }),
    )
    .unwrap();
    let updated = set_xml(&xml, &changed).unwrap();
    assert!(updated.contains(r#"style:writing-mode="lr-tb""#));
    assert!(updated.contains(r#"style:join-border="false""#));
    assert!(!updated.contains("writing-mode-automatic"));
    assert!(!updated.contains("register-true"));
    assert!(updated.contains(r#"f:text-align="justify""#));
    let reparsed = parse(&updated).unwrap();
    assert_eq!(
        reparsed.get("P1").unwrap().properties.as_ref().unwrap(),
        changed.properties.as_ref().unwrap()
    );
    // Strip: None removes only this module's attributes.
    let stripped = Style::named("P1", None).unwrap();
    let updated = set_xml(&xml, &stripped).unwrap();
    assert!(!updated.contains("writing-mode"));
    assert!(!updated.contains("register-true"));
    assert!(!updated.contains("join-border"));
    assert!(updated.contains(r#"f:text-align="justify""#));
    let reparsed = parse(&updated).unwrap();
    let p = reparsed.get("P1").unwrap().properties.as_ref().unwrap();
    assert_eq!(p, &Properties::default());
    // Missing target is an error.
    assert!(set_xml(&xml, &changed_named("Other")).is_err());
    // Insert into a style without a properties element.
    fn changed_named(name: &str) -> Style {
        Style::named(
            name,
            Some(Properties {
                register_true: Some(true),
                ..Default::default()
            }),
        )
        .unwrap()
    }
    let bare = wrap(r#"<s:style s:name="B" s:family="paragraph"/>"#);
    let updated = set_xml(&bare, &changed_named("B")).unwrap();
    let reparsed = parse(&updated).unwrap();
    assert_eq!(
        reparsed
            .get("B")
            .unwrap()
            .properties
            .as_ref()
            .unwrap()
            .register_true,
        Some(true)
    );
    // Insert into a style with other children but no properties element.
    let with_text =
        wrap(r#"<s:style s:name="C" s:family="paragraph"><s:text-properties/></s:style>"#);
    let updated = set_xml(&with_text, &changed_named("C")).unwrap();
    assert!(updated.contains("<s:text-properties/>"));
    let reparsed = parse(&updated).unwrap();
    assert!(reparsed.get("C").unwrap().properties.is_some());
    // Nothing to strip and no properties element: unchanged.
    let unchanged = set_xml(&bare, &Style::named("B", None).unwrap()).unwrap();
    assert_eq!(unchanged, bare);
}

#[test]
fn builder_package_round_trip() {
    let style = Style::named(
        "Body",
        Some(Properties {
            writing_mode: Some(WritingMode::TbRl),
            writing_mode_automatic: Some(false),
            register_true: Some(true),
            join_border: Some(false),
        }),
    )
    .unwrap();
    let mut builder = Builder::new();
    builder
        .add_paragraph_writing_mode_style(style.clone())
        .unwrap();
    assert!(
        builder
            .add_paragraph_writing_mode_style(Style::named("Body", None).unwrap())
            .is_err()
    );
    builder.add_paragraph("x").unwrap();
    let package = litchi_odt::generic::Package::from_bytes(builder.build().unwrap()).unwrap();
    assert_eq!(
        package.paragraph_style_writing_modes().unwrap().get("Body"),
        Some(&style)
    );
}

#[test]
fn parses_real_odf_fixture() {
    let xml = include_str!("../../../test-data/odf/odt/note-tracked-changes.fodt");
    let set = parse(xml).unwrap();
    assert!(!set.styles.is_empty());
    assert!(
        set.styles
            .iter()
            .filter_map(|x| x.properties.as_ref())
            .any(|p| p.writing_mode == Some(WritingMode::LrTb))
    );
    let flat = litchi_odt::generic::FlatDocument::from_reader(Cursor::new(xml)).unwrap();
    assert_eq!(
        flat.paragraph_style_writing_modes().unwrap().styles.len(),
        set.styles.len()
    );
}

#[test]
fn parses_register_true_and_join_border_fixtures() {
    let xml = include_str!(
        "../../../test-data/libreoffice-core/sw/qa/extras/pagelinespacing/data/pageColumns.fodt"
    );
    let set = parse(xml).unwrap();
    assert!(
        set.styles
            .iter()
            .filter_map(|x| x.properties.as_ref())
            .any(|p| p.register_true == Some(true))
    );
    assert!(
        set.styles
            .iter()
            .filter_map(|x| x.properties.as_ref())
            .any(|p| p.writing_mode == Some(WritingMode::LrTb))
    );
    let xml = include_str!(
        "../../../test-data/libreoffice-core/vcl/qa/cppunit/pdfexport/data/tdf159817.fodt"
    );
    let set = parse(xml).unwrap();
    assert!(
        set.styles
            .iter()
            .filter_map(|x| x.properties.as_ref())
            .any(|p| p.join_border == Some(true))
    );
}
