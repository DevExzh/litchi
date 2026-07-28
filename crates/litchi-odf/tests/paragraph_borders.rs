use litchi_odf::{
    DocumentBuilder, FlatOpenDocument, OdfNonNegativeLength, OpenDocumentPackage,
    ParagraphBackgroundTransparency, ParagraphBorder, ParagraphBorderProperties,
    ParagraphBorderWidth, ParagraphBorderWidths, ParagraphStyleBorder, TableRowBackgroundColor,
    TableRowBackgroundImage, TableRowBackgroundRepeat, TableRowBackgroundSource, TableShadow,
    parse_paragraph_style_borders, set_paragraph_style_border_xml,
};
use std::io::Cursor;
const O: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const S: &str = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
const F: &str = "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0";
const D: &str = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
const X: &str = "http://www.w3.org/1999/xlink";
fn wrap(x: &str) -> String {
    format!(
        r#"<o:styles xmlns:o="{O}" xmlns:s="{S}" xmlns:f="{F}" xmlns:d="{D}" xmlns:x="{X}">{x}</o:styles>"#
    )
}

fn full_properties() -> &'static str {
    r##"<s:style s:name="P1" s:family="paragraph"><s:paragraph-properties f:border="0.74pt solid #808080" f:border-top="none" f:border-bottom="0.74pt dotted #000000" f:border-left="1pt solid #FF0000" f:border-right="none" s:border-line-width="0.002cm 0.07cm 0.002cm" s:border-line-width-top="0.01cm 0.07cm 0.01cm" f:padding="0.05cm" f:padding-left="0.199cm" s:shadow="none" f:background-color="#FFCC00" s:background-transparency="25%" f:margin-left="1cm" f:line-height="115%"><s:background-image s:repeat="no-repeat" s:position="right" d:opacity="75%" x:type="simple" x:href="Pictures/a.png" x:show="embed" x:actuate="onLoad"/></s:paragraph-properties></s:style>"##
}

#[test]
fn parses_all_border_attributes() {
    let set = parse_paragraph_style_borders(&wrap(full_properties())).unwrap();
    let style = set.get("P1").unwrap();
    let p = style.properties.as_ref().unwrap();
    assert_eq!(p.border.as_ref().unwrap().as_str(), "0.74pt solid #808080");
    assert_eq!(p.border_top.as_ref().unwrap().as_str(), "none");
    assert_eq!(
        p.border_line_width.as_ref().unwrap().space.as_str(),
        "0.07cm"
    );
    assert_eq!(
        p.border_line_width_top
            .as_ref()
            .unwrap()
            .inner_width
            .as_str(),
        "0.01cm"
    );
    assert_eq!(
        p.padding,
        Some(OdfNonNegativeLength::new("0.05cm").unwrap())
    );
    assert_eq!(
        p.padding_left,
        Some(OdfNonNegativeLength::new("0.199cm").unwrap())
    );
    assert_eq!(p.shadow.as_ref().unwrap().as_str(), "none");
    assert_eq!(p.background_color.as_ref().unwrap().as_str(), "#FFCC00");
    assert_eq!(
        p.background_transparency,
        Some(ParagraphBackgroundTransparency::new("25%").unwrap())
    );
    let image = p.background_image.as_ref().unwrap();
    assert_eq!(image.repeat, Some(TableRowBackgroundRepeat::NoRepeat));
    assert_eq!(
        image.source,
        TableRowBackgroundSource::Link {
            href: "Pictures/a.png".to_string(),
            show_embed: true,
            actuate_on_load: true,
        }
    );
}

#[test]
fn round_trip_serialize_reparse() {
    let set = parse_paragraph_style_borders(&wrap(full_properties())).unwrap();
    let style = set.get("P1").unwrap();
    let fragment = style.to_xml_fragment().unwrap();
    let reparsed = parse_paragraph_style_borders(&wrap(&fragment)).unwrap();
    assert_eq!(reparsed.get("P1"), Some(style));
    // Embedded binary-data round trip.
    let embedded = wrap(
        r#"<s:default-style s:family="paragraph"><s:paragraph-properties f:padding="0cm"><s:background-image><o:binary-data>AQIDBA==</o:binary-data></s:background-image></s:paragraph-properties></s:default-style>"#,
    );
    let set = parse_paragraph_style_borders(&embedded).unwrap();
    assert_eq!(
        set.default_style()
            .unwrap()
            .properties
            .as_ref()
            .unwrap()
            .background_image
            .as_ref()
            .unwrap()
            .source,
        TableRowBackgroundSource::Embedded(vec![1, 2, 3, 4])
    );
    let fragment = set.default_style().unwrap().to_xml_fragment().unwrap();
    assert_eq!(
        parse_paragraph_style_borders(&wrap(&fragment)).unwrap(),
        set
    );
}

#[test]
fn ignores_sibling_owned_attributes_and_children() {
    let xml = wrap(
        r#"<s:style s:name="P" s:family="paragraph"><s:paragraph-properties f:margin-left="1cm" f:line-height="200%"><s:tab-stops><s:tab-stop s:position="2cm"/></s:tab-stops><s:drop-cap s:length="2"/></s:paragraph-properties></s:style>"#,
    );
    let set = parse_paragraph_style_borders(&xml).unwrap();
    let p = set.get("P").unwrap().properties.as_ref().unwrap();
    assert_eq!(p, &ParagraphBorderProperties::default());
}

#[test]
fn rejects_malformed_values_and_duplicates() {
    let bad = [
        // Negative padding.
        wrap(
            r#"<s:style s:name="P" s:family="paragraph"><s:paragraph-properties f:padding="-1cm"/></s:style>"#,
        ),
        // Border line width needs three lengths.
        wrap(
            r#"<s:style s:name="P" s:family="paragraph"><s:paragraph-properties s:border-line-width="1cm 2cm"/></s:style>"#,
        ),
        // Border line widths must be positive.
        wrap(
            r#"<s:style s:name="P" s:family="paragraph"><s:paragraph-properties s:border-line-width="0cm 1cm 1cm"/></s:style>"#,
        ),
        // Transparency above 100 percent.
        wrap(
            r#"<s:style s:name="P" s:family="paragraph"><s:paragraph-properties s:background-transparency="101%"/></s:style>"#,
        ),
        // Transparency must be a percentage.
        wrap(
            r#"<s:style s:name="P" s:family="paragraph"><s:paragraph-properties s:background-transparency="10"/></s:style>"#,
        ),
        // Background color must be transparent or #RRGGBB.
        wrap(
            r#"<s:style s:name="P" s:family="paragraph"><s:paragraph-properties f:background-color="red"/></s:style>"#,
        ),
        // Empty border shorthand.
        wrap(
            r#"<s:style s:name="P" s:family="paragraph"><s:paragraph-properties f:border=""/></s:style>"#,
        ),
        // Duplicate owned attribute.
        wrap(
            r#"<s:style s:name="P" s:family="paragraph"><s:paragraph-properties f:padding="1cm" f:padding="2cm"/></s:style>"#,
        ),
        // Duplicate properties element.
        wrap(
            r#"<s:style s:name="P" s:family="paragraph"><s:paragraph-properties/><s:paragraph-properties/></s:style>"#,
        ),
        // Duplicate background image.
        wrap(
            r#"<s:style s:name="P" s:family="paragraph"><s:paragraph-properties><s:background-image/><s:background-image/></s:paragraph-properties></s:style>"#,
        ),
        // Linked image cannot carry binary data.
        wrap(
            r#"<s:style s:name="P" s:family="paragraph"><s:paragraph-properties><s:background-image x:type="simple" x:href="a.png"><o:binary-data>AQID</o:binary-data></s:background-image></s:paragraph-properties></s:style>"#,
        ),
        // Incomplete xlink group.
        wrap(
            r#"<s:style s:name="P" s:family="paragraph"><s:paragraph-properties><s:background-image x:href="a.png"/></s:paragraph-properties></s:style>"#,
        ),
        // Unknown background-image attribute.
        wrap(
            r#"<s:style s:name="P" s:family="paragraph"><s:paragraph-properties><s:background-image s:foo="bar"/></s:paragraph-properties></s:style>"#,
        ),
        // Invalid base64.
        wrap(
            r#"<s:style s:name="P" s:family="paragraph"><s:paragraph-properties><s:background-image><o:binary-data>!!!</o:binary-data></s:background-image></s:paragraph-properties></s:style>"#,
        ),
        // Duplicate style identity.
        wrap(
            r#"<s:style s:name="P" s:family="paragraph"><s:paragraph-properties/></s:style><s:style s:name="P" s:family="paragraph"><s:paragraph-properties/></s:style>"#,
        ),
    ];
    for xml in bad {
        assert!(
            parse_paragraph_style_borders(&xml).is_err(),
            "accepted {xml}"
        );
    }
    assert!(ParagraphBorderWidth::new("1em").is_err());
    assert!(ParagraphBackgroundTransparency::new("-5%").is_err());
}

#[test]
fn mutation_replaces_inserts_and_removes() {
    let xml = wrap(full_properties());
    // Replace: owned attributes rewritten, sibling attributes and tab-stops survive.
    let changed = ParagraphStyleBorder::named(
        "P1",
        Some(ParagraphBorderProperties {
            border: Some(ParagraphBorder::new("none").unwrap()),
            padding: Some(OdfNonNegativeLength::new("0.1cm").unwrap()),
            ..Default::default()
        }),
    )
    .unwrap();
    let updated = set_paragraph_style_border_xml(&xml, &changed).unwrap();
    assert!(updated.contains(r#"fo:border="none""#));
    assert!(!updated.contains("border-line-width"));
    assert!(!updated.contains("background-image"));
    assert!(updated.contains(r#"f:margin-left="1cm""#));
    assert!(updated.contains(r#"f:line-height="115%""#));
    let reparsed = parse_paragraph_style_borders(&updated).unwrap();
    assert_eq!(
        reparsed.get("P1").unwrap().properties.as_ref().unwrap(),
        changed.properties.as_ref().unwrap()
    );
    // Insert a background image into an empty properties element.
    let empty = wrap(
        r#"<s:style s:name="E" s:family="paragraph"><s:paragraph-properties f:margin-left="1cm"/></s:style>"#,
    );
    let with_image = ParagraphStyleBorder::named(
        "E",
        Some(ParagraphBorderProperties {
            background_image: Some(TableRowBackgroundImage {
                repeat: None,
                position: None,
                filter_name: None,
                opacity: None,
                source: TableRowBackgroundSource::Link {
                    href: "Pictures/b.png".to_string(),
                    show_embed: false,
                    actuate_on_load: false,
                },
            }),
            ..Default::default()
        }),
    )
    .unwrap();
    let updated = set_paragraph_style_border_xml(&empty, &with_image).unwrap();
    assert!(updated.contains(r#"f:margin-left="1cm""#));
    let reparsed = parse_paragraph_style_borders(&updated).unwrap();
    assert_eq!(
        reparsed.get("E").unwrap().properties.as_ref().unwrap(),
        with_image.properties.as_ref().unwrap()
    );
    // Insert a whole properties element into a bare style.
    let bare = wrap(r#"<s:style s:name="B" s:family="paragraph"/>"#);
    let updated = set_paragraph_style_border_xml(&bare, &with_image_named("B")).unwrap();
    let reparsed = parse_paragraph_style_borders(&updated).unwrap();
    assert!(reparsed.get("B").unwrap().properties.is_some());
    // Strip: None removes owned attributes and the background image.
    fn with_image_named(name: &str) -> ParagraphStyleBorder {
        ParagraphStyleBorder::named(
            name,
            Some(ParagraphBorderProperties {
                shadow: Some(TableShadow::new("none").unwrap()),
                background_color: Some(TableRowBackgroundColor::new("transparent").unwrap()),
                ..Default::default()
            }),
        )
        .unwrap()
    }
    let stripped = ParagraphStyleBorder::named("P1", None).unwrap();
    let updated = set_paragraph_style_border_xml(&xml, &stripped).unwrap();
    assert!(!updated.contains("fo:border"));
    assert!(!updated.contains("background-image"));
    assert!(updated.contains(r#"f:margin-left="1cm""#));
    // Missing target is an error.
    assert!(set_paragraph_style_border_xml(&xml, &with_image_named("Missing")).is_err());
    // Nothing to change on a bare style: unchanged.
    let unchanged =
        set_paragraph_style_border_xml(&bare, &ParagraphStyleBorder::named("B", None).unwrap())
            .unwrap();
    assert_eq!(unchanged, bare);
}

#[test]
fn builder_package_round_trip() {
    let style = ParagraphStyleBorder::named(
        "Box",
        Some(ParagraphBorderProperties {
            border: Some(ParagraphBorder::new("0.74pt solid #808080").unwrap()),
            border_line_width: Some(ParagraphBorderWidths {
                inner_width: ParagraphBorderWidth::new("0.002cm").unwrap(),
                space: ParagraphBorderWidth::new("0.07cm").unwrap(),
                outer_width: ParagraphBorderWidth::new("0.002cm").unwrap(),
            }),
            padding: Some(OdfNonNegativeLength::new("0cm").unwrap()),
            background_color: Some(TableRowBackgroundColor::new("transparent").unwrap()),
            ..Default::default()
        }),
    )
    .unwrap();
    let mut builder = DocumentBuilder::new();
    builder.add_paragraph_border_style(style.clone()).unwrap();
    assert!(
        builder
            .add_paragraph_border_style(ParagraphStyleBorder::named("Box", None).unwrap())
            .is_err()
    );
    builder.add_paragraph("x").unwrap();
    let package = OpenDocumentPackage::from_bytes(builder.build().unwrap()).unwrap();
    assert_eq!(
        package.paragraph_style_borders().unwrap().get("Box"),
        Some(&style)
    );
}

#[test]
fn parses_real_libreoffice_fixture() {
    let bytes = include_bytes!(
        "../../../test-data/libreoffice-core/vcl/qa/cppunit/pdfexport/data/tdf159817.fodt"
    );
    let flat = FlatOpenDocument::from_reader(Cursor::new(bytes)).unwrap();
    let set = flat.paragraph_style_borders().unwrap();
    assert!(!set.styles.is_empty());
    assert!(
        set.styles
            .iter()
            .filter_map(|x| x.properties.as_ref())
            .any(|p| p.border_left.is_some() && p.padding_top.is_some())
    );
    // The margins module reads the same element through its own lens.
    let margins = flat.paragraph_style_margins().unwrap();
    assert!(
        margins
            .styles
            .iter()
            .filter_map(|x| x.properties.as_ref())
            .any(|p| p.margin_right.is_some())
    );
}
