use litchi_odt::Builder;
use litchi_odt::style::paragraph::alignment::{
    Horizontal, Properties, Style, Vertical, parse, set_xml,
};
use std::io::Cursor;
const O: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const S: &str = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
const F: &str = "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0";
fn wrap(x: &str) -> String {
    format!(r#"<o:styles xmlns:o="{O}" xmlns:s="{S}" xmlns:f="{F}">{x}</o:styles>"#)
}

fn full_properties() -> &'static str {
    r#"<s:style s:name="P1" s:family="paragraph"><s:paragraph-properties f:text-align="justify" s:vertical-align="middle" f:line-height="115%" s:justify-single-word="false"/></s:style>"#
}

#[test]
fn parses_all_alignment_attributes() {
    let set = parse(&wrap(full_properties())).unwrap();
    let style = set.get("P1").unwrap();
    let p = style.properties.as_ref().unwrap();
    assert_eq!(p.horizontal, Some(Horizontal::Justify));
    assert_eq!(p.vertical, Some(Vertical::Middle));
}

#[test]
fn ignores_sibling_owned_attributes_and_foreign_families() {
    let xml = wrap(
        r#"<s:style s:name="P" s:family="paragraph"><s:paragraph-properties f:line-height="200%" f:text-align="start"/></s:style><s:style s:name="T" s:family="table"><s:paragraph-properties f:text-align="end"/></s:style>"#,
    );
    let set = parse(&xml).unwrap();
    assert_eq!(set.styles.len(), 1);
    let p = set.get("P").unwrap().properties.as_ref().unwrap();
    assert_eq!(p.horizontal, Some(Horizontal::Start));
    assert!(p.vertical.is_none());
    assert!(set.get("T").is_none());
}

#[test]
fn parses_every_token() {
    for (token, expect) in [
        ("start", Horizontal::Start),
        ("end", Horizontal::End),
        ("left", Horizontal::Left),
        ("right", Horizontal::Right),
        ("center", Horizontal::Center),
        ("justify", Horizontal::Justify),
    ] {
        let xml = wrap(&format!(
            r#"<s:style s:name="P" s:family="paragraph"><s:paragraph-properties f:text-align="{token}"/></s:style>"#
        ));
        let set = parse(&xml).unwrap();
        assert_eq!(
            set.get("P")
                .unwrap()
                .properties
                .as_ref()
                .unwrap()
                .horizontal,
            Some(expect)
        );
    }
    for (token, expect) in [
        ("top", Vertical::Top),
        ("middle", Vertical::Middle),
        ("bottom", Vertical::Bottom),
        ("auto", Vertical::Auto),
        ("baseline", Vertical::Baseline),
    ] {
        let xml = wrap(&format!(
            r#"<s:style s:name="P" s:family="paragraph"><s:paragraph-properties s:vertical-align="{token}"/></s:style>"#
        ));
        let set = parse(&xml).unwrap();
        assert_eq!(
            set.get("P").unwrap().properties.as_ref().unwrap().vertical,
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
        horizontal: Some(Horizontal::Center),
        ..Default::default()
    }));
    let fragment = default.to_xml_fragment().unwrap();
    let reparsed = parse(&wrap(&fragment)).unwrap();
    assert_eq!(reparsed.default_style(), Some(&default));
}

#[test]
fn rejects_malformed_values_and_duplicates() {
    let bad = [
        // Unknown fo:text-align token.
        wrap(
            r#"<s:style s:name="P" s:family="paragraph"><s:paragraph-properties f:text-align="justify-single-word"/></s:style>"#,
        ),
        // XSL inside/outside are not part of the ODF paragraph token set.
        wrap(
            r#"<s:style s:name="P" s:family="paragraph"><s:paragraph-properties f:text-align="inside"/></s:style>"#,
        ),
        // Unknown style:vertical-align token.
        wrap(
            r#"<s:style s:name="P" s:family="paragraph"><s:paragraph-properties s:vertical-align="super"/></s:style>"#,
        ),
        // Case matters.
        wrap(
            r#"<s:style s:name="P" s:family="paragraph"><s:paragraph-properties f:text-align="Center"/></s:style>"#,
        ),
        // Duplicate owned attribute (same expanded name).
        wrap(
            r#"<s:style s:name="P" s:family="paragraph"><s:paragraph-properties f:text-align="left" f:text-align="right"/></s:style>"#,
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
            horizontal: Some(Horizontal::End),
            ..Default::default()
        }),
    )
    .unwrap();
    let updated = set_xml(&xml, &changed).unwrap();
    assert!(updated.contains(r#"fo:text-align="end""#));
    assert!(!updated.contains("vertical-align"));
    assert!(updated.contains(r#"f:line-height="115%""#));
    assert!(updated.contains(r#"s:justify-single-word="false""#));
    let reparsed = parse(&updated).unwrap();
    assert_eq!(
        reparsed.get("P1").unwrap().properties.as_ref().unwrap(),
        changed.properties.as_ref().unwrap()
    );
    // Strip: None removes only this module's attributes.
    let stripped = Style::named("P1", None).unwrap();
    let updated = set_xml(&xml, &stripped).unwrap();
    assert!(!updated.contains("text-align"));
    assert!(updated.contains(r#"f:line-height="115%""#));
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
                vertical: Some(Vertical::Baseline),
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
            .vertical,
        Some(Vertical::Baseline)
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
            horizontal: Some(Horizontal::Justify),
            vertical: Some(Vertical::Auto),
        }),
    )
    .unwrap();
    let mut builder = Builder::new();
    builder
        .add_paragraph_alignment_style(style.clone())
        .unwrap();
    assert!(
        builder
            .add_paragraph_alignment_style(Style::named("Body", None).unwrap())
            .is_err()
    );
    builder.add_paragraph("x").unwrap();
    let package =
        litchi_odt::generic::OpenDocumentPackage::from_bytes(builder.build().unwrap()).unwrap();
    assert_eq!(
        package.paragraph_style_alignments().unwrap().get("Body"),
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
            .any(|p| p.horizontal == Some(Horizontal::End))
    );
    let flat = litchi_odt::generic::FlatOpenDocument::from_reader(Cursor::new(xml)).unwrap();
    assert_eq!(
        flat.paragraph_style_alignments().unwrap().styles.len(),
        set.styles.len()
    );
}
