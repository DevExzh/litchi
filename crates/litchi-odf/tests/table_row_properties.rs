use litchi_odf::{
    DocumentBuilder, FlatOpenDocument, HorizontalBackgroundPosition, OpenDocumentPackage,
    TableRowBackgroundColor, TableRowBackgroundImage, TableRowBackgroundPosition,
    TableRowBackgroundRepeat, TableRowBackgroundSource, TableRowBreak, TableRowKeepTogether,
    TableRowLength, TableRowOpacity, TableRowProperties, TableRowStyleProperties,
    VerticalBackgroundPosition, parse_table_row_style_properties,
    set_table_row_style_properties_xml,
};
use std::io::Cursor;
const O: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const S: &str = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
const F: &str = "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0";
const D: &str = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
const X: &str = "http://www.w3.org/1999/xlink";
fn wrap(value: &str) -> String {
    format!(
        r#"<o:document-styles xmlns:o="{O}" xmlns:s="{S}" xmlns:f="{F}" xmlns:d="{D}" xmlns:x="{X}"><o:styles>{value}</o:styles></o:document-styles>"#
    )
}

#[test]
fn complete_linked_round_trip_and_mutation() {
    let xml = wrap(
        r##"<s:style s:name="Row" s:family="table-row"><s:table-row-properties s:row-height="1.2cm" s:min-row-height="0cm" s:use-optimal-row-height="1" f:background-color="#a0B1c2" f:break-before="column" f:break-after="page" f:keep-together="always"><s:background-image s:repeat="stretch" s:position="top left" s:filter-name="png" d:opacity="42.5%" x:type="simple" x:href="Pictures/a.png" x:show="embed" x:actuate="onLoad"/></s:table-row-properties></s:style>"##,
    );
    let set = parse_table_row_style_properties(&xml).unwrap();
    let style = set.get("Row").unwrap();
    let props = style.properties.as_ref().unwrap();
    assert_eq!(props.use_optimal_row_height, Some(true));
    assert_eq!(props.break_before, Some(TableRowBreak::Column));
    assert_eq!(
        props.background_image.as_ref().unwrap().position,
        Some(TableRowBackgroundPosition::Pair(
            HorizontalBackgroundPosition::Left,
            VerticalBackgroundPosition::Top
        ))
    );
    let fragment = style.to_xml_fragment().unwrap();
    assert_eq!(
        parse_table_row_style_properties(&wrap(&fragment))
            .unwrap()
            .get("Row"),
        Some(style)
    );
    let changed = TableRowStyleProperties::named(
        "Row",
        Some(TableRowProperties {
            keep_together: Some(TableRowKeepTogether::Auto),
            ..Default::default()
        }),
    )
    .unwrap();
    let updated = set_table_row_style_properties_xml(&xml, &changed).unwrap();
    assert!(!updated.contains("f:background-color"));
    assert!(updated.contains("keep-together=\"auto\""));
}

#[test]
fn embedded_binary_and_empty_image_round_trip() {
    let xml = wrap(
        r#"<s:default-style s:family="table-row"><s:table-row-properties><s:background-image><o:binary-data>AQIDBA==</o:binary-data></s:background-image></s:table-row-properties></s:default-style>"#,
    );
    let set = parse_table_row_style_properties(&xml).unwrap();
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
    let out = set.default_style().unwrap().to_xml_fragment().unwrap();
    assert_eq!(parse_table_row_style_properties(&wrap(&out)).unwrap(), set);
}

#[test]
fn parses_real_odfdo_and_libreoffice_fixtures() {
    let odfdo = include_str!("../../../test-data/odfdo/tests/samples/test_flat_lo.fods");
    let rows = parse_table_row_style_properties(odfdo).unwrap();
    assert!(rows.get("ro1").is_some());
    let lo = include_bytes!(
        "../../../test-data/libreoffice-core/xmloff/qa/unit/data/tdf162686_3D_metal_type_ODF.fods"
    );
    let flat = FlatOpenDocument::from_reader(Cursor::new(lo)).unwrap();
    assert!(!flat.table_row_style_properties().unwrap().styles.is_empty());
}

#[test]
fn builder_package_round_trip() {
    let properties = TableRowProperties {
        row_height: Some(TableRowLength::positive("5mm").unwrap()),
        min_row_height: Some(TableRowLength::non_negative("0cm").unwrap()),
        background_color: Some(TableRowBackgroundColor::new("transparent").unwrap()),
        break_after: Some(TableRowBreak::Auto),
        background_image: Some(TableRowBackgroundImage {
            repeat: Some(TableRowBackgroundRepeat::NoRepeat),
            position: Some(TableRowBackgroundPosition::Right),
            filter_name: None,
            opacity: Some(TableRowOpacity::new("100%").unwrap()),
            source: TableRowBackgroundSource::Empty,
        }),
        ..Default::default()
    };
    let style = TableRowStyleProperties::named("Row", Some(properties)).unwrap();
    let mut builder = DocumentBuilder::new();
    builder.add_table_row_property_style(style.clone()).unwrap();
    builder.add_paragraph("x").unwrap();
    let package = OpenDocumentPackage::from_bytes(builder.build().unwrap()).unwrap();
    assert_eq!(
        package.table_row_style_properties().unwrap().get("Row"),
        Some(&style)
    );
}

#[test]
fn rejects_malformed_and_unsafe_forms() {
    let bad = [
        wrap(
            r#"<s:style s:name="R" s:family="table-row"><s:table-row-properties s:row-height="0cm"/></s:style>"#,
        ),
        wrap(
            r#"<s:style s:name="R" s:family="table-row"><s:table-row-properties f:break-before="odd-page"/></s:style>"#,
        ),
        wrap(
            r#"<s:style s:name="R" s:family="table-row"><s:table-row-properties/><s:table-row-properties/></s:style>"#,
        ),
        wrap(
            r#"<s:style s:name="R" s:family="table-row"><s:table-row-properties><s:background-image x:href="a"/></s:table-row-properties></s:style>"#,
        ),
        wrap(
            r#"<s:style s:name="R" s:family="table-row"><s:table-row-properties><s:background-image d:opacity="101%"/></s:table-row-properties></s:style>"#,
        ),
        wrap(
            r#"<s:style s:name="R" s:family="table-row"><s:table-row-properties><s:background-image><o:binary-data>***=</o:binary-data></s:background-image></s:table-row-properties></s:style>"#,
        ),
        format!(
            "<!DOCTYPE x>{}",
            wrap(r#"<s:style s:name="R" s:family="table-row"/>"#)
        ),
    ];
    for xml in bad {
        assert!(
            parse_table_row_style_properties(&xml).is_err(),
            "accepted {xml}"
        );
    }
}
