use litchi_odt::{
    Builder, FlatOpenDocument, OpenDocumentPackage, ParagraphHorizontalMargin, ParagraphMargins,
    ParagraphStyleMargins, ParagraphTextIndent, ParagraphVerticalMargin,
    parse_paragraph_style_margins, set_paragraph_style_margins_xml,
};
use std::io::Cursor;
const O: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const S: &str = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
const F: &str = "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0";
fn wrap(x: &str) -> String {
    format!(r#"<o:styles xmlns:o="{O}" xmlns:s="{S}" xmlns:f="{F}">{x}</o:styles>"#)
}

fn full_properties() -> &'static str {
    r#"<s:style s:name="P1" s:family="paragraph"><s:paragraph-properties f:margin="0.2cm" f:margin-left="1cm" f:margin-right="-0.5cm" f:margin-top="10%" f:margin-bottom="0.212cm" f:text-indent="-0.635cm" s:contextual-spacing="true" f:line-height="115%" f:keep-together="always"/></s:style>"#
}

#[test]
fn parses_all_margin_attributes() {
    let set = parse_paragraph_style_margins(&wrap(full_properties())).unwrap();
    let style = set.get("P1").unwrap();
    let p = style.properties.as_ref().unwrap();
    assert_eq!(
        p.margin,
        Some(ParagraphVerticalMargin::new("0.2cm").unwrap())
    );
    assert_eq!(
        p.margin_left,
        Some(ParagraphHorizontalMargin::new("1cm").unwrap())
    );
    assert_eq!(
        p.margin_right,
        Some(ParagraphHorizontalMargin::new("-0.5cm").unwrap())
    );
    assert_eq!(
        p.margin_top,
        Some(ParagraphVerticalMargin::new("10%").unwrap())
    );
    assert_eq!(
        p.margin_bottom,
        Some(ParagraphVerticalMargin::new("0.212cm").unwrap())
    );
    assert_eq!(
        p.text_indent,
        Some(ParagraphTextIndent::new("-0.635cm").unwrap())
    );
    assert_eq!(p.contextual_spacing, Some(true));
}

#[test]
fn ignores_sibling_owned_attributes_and_foreign_families() {
    let xml = wrap(
        r#"<s:style s:name="P" s:family="paragraph"><s:paragraph-properties f:line-height="200%" f:margin-left="1cm"/></s:style><s:style s:name="T" s:family="table"><s:paragraph-properties f:margin-left="9cm"/></s:style>"#,
    );
    let set = parse_paragraph_style_margins(&xml).unwrap();
    assert_eq!(set.styles.len(), 1);
    let p = set.get("P").unwrap().properties.as_ref().unwrap();
    assert_eq!(
        p.margin_left,
        Some(ParagraphHorizontalMargin::new("1cm").unwrap())
    );
    assert!(p.margin_top.is_none());
    assert!(set.get("T").is_none());
}

#[test]
fn round_trip_serialize_reparse() {
    let set = parse_paragraph_style_margins(&wrap(full_properties())).unwrap();
    let style = set.get("P1").unwrap();
    let fragment = style.to_xml_fragment().unwrap();
    let reparsed = parse_paragraph_style_margins(&wrap(&fragment)).unwrap();
    assert_eq!(reparsed.get("P1"), Some(style));
    let default = ParagraphStyleMargins::default_style(Some(ParagraphMargins {
        margin: Some(ParagraphVerticalMargin::new("0cm").unwrap()),
        contextual_spacing: Some(false),
        ..Default::default()
    }));
    let fragment = default.to_xml_fragment().unwrap();
    let reparsed = parse_paragraph_style_margins(&wrap(&fragment)).unwrap();
    assert_eq!(reparsed.default_style(), Some(&default));
}

#[test]
fn rejects_malformed_values_and_duplicates() {
    let bad = [
        // Vertical margins cannot be negative.
        wrap(
            r#"<s:style s:name="P" s:family="paragraph"><s:paragraph-properties f:margin-top="-1cm"/></s:style>"#,
        ),
        // Plain numbers are not lengths.
        wrap(
            r#"<s:style s:name="P" s:family="paragraph"><s:paragraph-properties f:margin-left="1"/></s:style>"#,
        ),
        // Unknown units are not lengths.
        wrap(
            r#"<s:style s:name="P" s:family="paragraph"><s:paragraph-properties f:margin-left="1em"/></s:style>"#,
        ),
        // Percent needs a number.
        wrap(
            r#"<s:style s:name="P" s:family="paragraph"><s:paragraph-properties f:margin="%"/></s:style>"#,
        ),
        // Text indent must be a length, not a percent.
        wrap(
            r#"<s:style s:name="P" s:family="paragraph"><s:paragraph-properties f:text-indent="10%"/></s:style>"#,
        ),
        // Malformed boolean.
        wrap(
            r#"<s:style s:name="P" s:family="paragraph"><s:paragraph-properties s:contextual-spacing="yes"/></s:style>"#,
        ),
        // Duplicate owned attribute (same expanded name).
        wrap(
            r#"<s:style s:name="P" s:family="paragraph"><s:paragraph-properties f:margin-left="1cm" f:margin-left="2cm"/></s:style>"#,
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
        assert!(
            parse_paragraph_style_margins(&xml).is_err(),
            "accepted {xml}"
        );
    }
    assert!(ParagraphVerticalMargin::new("").is_err());
    assert!(ParagraphHorizontalMargin::new("1.2.3cm").is_err());
    assert!(ParagraphTextIndent::new("abc").is_err());
}

#[test]
fn mutation_replaces_inserts_and_strips_owned_attributes() {
    let xml = wrap(full_properties());
    // Replace: owned attributes are rewritten, sibling attributes survive.
    let changed = ParagraphStyleMargins::named(
        "P1",
        Some(ParagraphMargins {
            margin_left: Some(ParagraphHorizontalMargin::new("2cm").unwrap()),
            ..Default::default()
        }),
    )
    .unwrap();
    let updated = set_paragraph_style_margins_xml(&xml, &changed).unwrap();
    assert!(updated.contains(r#"fo:margin-left="2cm""#));
    assert!(!updated.contains("margin-top"));
    assert!(updated.contains(r#"f:line-height="115%""#));
    assert!(updated.contains(r#"f:keep-together="always""#));
    let reparsed = parse_paragraph_style_margins(&updated).unwrap();
    assert_eq!(
        reparsed.get("P1").unwrap().properties.as_ref().unwrap(),
        changed.properties.as_ref().unwrap()
    );
    // Strip: None removes only this module's attributes.
    let stripped = ParagraphStyleMargins::named("P1", None).unwrap();
    let updated = set_paragraph_style_margins_xml(&xml, &stripped).unwrap();
    assert!(!updated.contains("margin-left"));
    assert!(updated.contains(r#"f:line-height="115%""#));
    let reparsed = parse_paragraph_style_margins(&updated).unwrap();
    let p = reparsed.get("P1").unwrap().properties.as_ref().unwrap();
    assert_eq!(p, &ParagraphMargins::default());
    // Missing target is an error.
    assert!(set_paragraph_style_margins_xml(&xml, &changed_named("Other")).is_err());
    // Insert into a style without a properties element.
    fn changed_named(name: &str) -> ParagraphStyleMargins {
        ParagraphStyleMargins::named(
            name,
            Some(ParagraphMargins {
                margin_bottom: Some(ParagraphVerticalMargin::new("0.5cm").unwrap()),
                ..Default::default()
            }),
        )
        .unwrap()
    }
    let bare = wrap(r#"<s:style s:name="B" s:family="paragraph"/>"#);
    let updated = set_paragraph_style_margins_xml(&bare, &changed_named("B")).unwrap();
    let reparsed = parse_paragraph_style_margins(&updated).unwrap();
    assert_eq!(
        reparsed
            .get("B")
            .unwrap()
            .properties
            .as_ref()
            .unwrap()
            .margin_bottom,
        Some(ParagraphVerticalMargin::new("0.5cm").unwrap())
    );
    // Insert into a style with other children but no properties element.
    let with_text =
        wrap(r#"<s:style s:name="C" s:family="paragraph"><s:text-properties/></s:style>"#);
    let updated = set_paragraph_style_margins_xml(&with_text, &changed_named("C")).unwrap();
    assert!(updated.contains("<s:text-properties/>"));
    let reparsed = parse_paragraph_style_margins(&updated).unwrap();
    assert!(reparsed.get("C").unwrap().properties.is_some());
    // Nothing to strip and no properties element: unchanged.
    let unchanged =
        set_paragraph_style_margins_xml(&bare, &ParagraphStyleMargins::named("B", None).unwrap())
            .unwrap();
    assert_eq!(unchanged, bare);
}

#[test]
fn builder_package_round_trip() {
    let style = ParagraphStyleMargins::named(
        "Body",
        Some(ParagraphMargins {
            margin_top: Some(ParagraphVerticalMargin::new("0.423cm").unwrap()),
            margin_bottom: Some(ParagraphVerticalMargin::new("0.212cm").unwrap()),
            contextual_spacing: Some(false),
            ..Default::default()
        }),
    )
    .unwrap();
    let mut builder = Builder::new();
    builder.add_paragraph_margin_style(style.clone()).unwrap();
    assert!(
        builder
            .add_paragraph_margin_style(ParagraphStyleMargins::named("Body", None).unwrap())
            .is_err()
    );
    builder.add_paragraph("x").unwrap();
    let package = OpenDocumentPackage::from_bytes(builder.build().unwrap()).unwrap();
    assert_eq!(
        package.paragraph_style_margins().unwrap().get("Body"),
        Some(&style)
    );
}

#[test]
fn parses_real_odfdo_fixture() {
    let xml = include_str!("../../../test-data/odfdo/tests/samples/images.fodt");
    let set = parse_paragraph_style_margins(xml).unwrap();
    assert!(!set.styles.is_empty());
    assert!(
        set.styles
            .iter()
            .filter_map(|x| x.properties.as_ref())
            .any(|p| p.contextual_spacing == Some(false) && p.margin_bottom.is_some())
    );
    let flat = FlatOpenDocument::from_reader(Cursor::new(xml)).unwrap();
    assert_eq!(
        flat.paragraph_style_margins().unwrap().styles.len(),
        set.styles.len()
    );
}
