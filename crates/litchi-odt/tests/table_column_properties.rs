use litchi_odt::{
    Builder, FlatOpenDocument, OpenDocumentPackage, TableColumnLength, TableColumnProperties,
    TableColumnRelWidth, TableColumnStyleProperties, TableRowBreak,
    parse_table_column_style_properties, set_table_column_style_properties_xml,
};
use std::io::Cursor;
const O: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const S: &str = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
const F: &str = "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0";
fn wrap(value: &str) -> String {
    format!(
        r#"<o:document-styles xmlns:o="{O}" xmlns:s="{S}" xmlns:f="{F}"><o:styles>{value}</o:styles></o:document-styles>"#
    )
}

#[test]
fn complete_round_trip_and_mutation() {
    let xml = wrap(
        r#"<s:style s:name="Col" s:family="table-column"><s:table-column-properties s:column-width="2.5cm" s:rel-column-width="12*" s:use-optimal-column-width="true" f:break-before="page" f:break-after="auto"/></s:style>"#,
    );
    let set = parse_table_column_style_properties(&xml).unwrap();
    let style = set.get("Col").unwrap();
    let props = style.properties.as_ref().unwrap();
    assert_eq!(props.column_width.as_ref().unwrap().as_str(), "2.5cm");
    assert_eq!(props.rel_column_width.as_ref().unwrap().as_str(), "12*");
    assert_eq!(props.use_optimal_column_width, Some(true));
    assert_eq!(props.break_before, Some(TableRowBreak::Page));
    assert_eq!(props.break_after, Some(TableRowBreak::Auto));
    let fragment = style.to_xml_fragment().unwrap();
    assert_eq!(
        parse_table_column_style_properties(&wrap(&fragment))
            .unwrap()
            .get("Col"),
        Some(style)
    );
    let changed = TableColumnStyleProperties::named(
        "Col",
        Some(TableColumnProperties {
            rel_column_width: Some(TableColumnRelWidth::new("3*").unwrap()),
            ..Default::default()
        }),
    )
    .unwrap();
    let updated = set_table_column_style_properties_xml(&xml, &changed).unwrap();
    assert!(!updated.contains("style:column-width"));
    assert!(updated.contains("rel-column-width=\"3*\""));
    assert_eq!(
        parse_table_column_style_properties(&updated)
            .unwrap()
            .get("Col"),
        Some(&changed)
    );
}

#[test]
fn default_style_and_insertion_into_empty_style() {
    let xml = wrap(r#"<s:default-style s:family="table-column"/>"#);
    let set = parse_table_column_style_properties(&xml).unwrap();
    assert!(set.default_style().unwrap().properties.is_none());
    let changed = TableColumnStyleProperties::default_style(Some(TableColumnProperties {
        column_width: Some(TableColumnLength::new("0.889in").unwrap()),
        break_before: Some(TableRowBreak::Auto),
        ..Default::default()
    }));
    let updated = set_table_column_style_properties_xml(&xml, &changed).unwrap();
    assert_eq!(
        parse_table_column_style_properties(&updated)
            .unwrap()
            .default_style(),
        Some(&changed)
    );
    let removed = set_table_column_style_properties_xml(
        &updated,
        &TableColumnStyleProperties::default_style(None),
    )
    .unwrap();
    assert!(removed.contains("default-style"));
    assert!(!removed.contains("table-column-properties"));
}

#[test]
fn parses_real_odfdo_and_libreoffice_fixtures() {
    let odfdo = include_str!("../../../test-data/odfdo/tests/samples/test_flat_lo.fods");
    let columns = parse_table_column_style_properties(odfdo).unwrap();
    let co1 = columns.get("co1").unwrap();
    assert_eq!(
        co1.properties
            .as_ref()
            .unwrap()
            .column_width
            .as_ref()
            .unwrap()
            .as_str(),
        "0.889in"
    );
    let lo = include_bytes!(
        "../../../test-data/libreoffice-core/sc/qa/unit/data/functions/mathematical/fods/aggregate.fods"
    );
    let flat = FlatOpenDocument::from_reader(Cursor::new(lo)).unwrap();
    assert!(
        !flat
            .table_column_style_properties()
            .unwrap()
            .styles
            .is_empty()
    );
}

#[test]
fn builder_package_round_trip() {
    let properties = TableColumnProperties {
        column_width: Some(TableColumnLength::new("4cm").unwrap()),
        rel_column_width: Some(TableColumnRelWidth::new("65535*").unwrap()),
        use_optimal_column_width: Some(false),
        break_after: Some(TableRowBreak::Column),
        ..Default::default()
    };
    let style = TableColumnStyleProperties::named("Col", Some(properties)).unwrap();
    let mut builder = Builder::new();
    builder
        .add_table_column_property_style(style.clone())
        .unwrap();
    builder.add_paragraph("x").unwrap();
    let package = OpenDocumentPackage::from_bytes(builder.build().unwrap()).unwrap();
    assert_eq!(
        package.table_column_style_properties().unwrap().get("Col"),
        Some(&style)
    );
}

#[test]
fn rejects_malformed_and_unsafe_forms() {
    let bad = [
        wrap(
            r#"<s:style s:name="C" s:family="table-column"><s:table-column-properties s:column-width="0cm"/></s:style>"#,
        ),
        wrap(
            r#"<s:style s:name="C" s:family="table-column"><s:table-column-properties s:column-width="-1cm"/></s:style>"#,
        ),
        wrap(
            r#"<s:style s:name="C" s:family="table-column"><s:table-column-properties s:rel-column-width="1.5*"/></s:style>"#,
        ),
        wrap(
            r#"<s:style s:name="C" s:family="table-column"><s:table-column-properties s:use-optimal-column-width="1"/></s:style>"#,
        ),
        wrap(
            r#"<s:style s:name="C" s:family="table-column"><s:table-column-properties f:break-before="odd-page"/></s:style>"#,
        ),
        wrap(
            r#"<s:style s:name="C" s:family="table-column"><s:table-column-properties s:display="true"/></s:style>"#,
        ),
        wrap(
            r#"<s:style s:name="C" s:family="table-column"><s:table-column-properties/><s:table-column-properties/></s:style>"#,
        ),
        wrap(
            r#"<s:style s:name="C" s:family="table-column"><s:table-column-properties><s:background-image/></s:table-column-properties></s:style>"#,
        ),
        wrap(
            r#"<s:style s:name="C" s:family="table-column"><s:table-column-properties>text</s:table-column-properties></s:style>"#,
        ),
        format!(
            "<!DOCTYPE x>{}",
            wrap(r#"<s:style s:name="C" s:family="table-column"/>"#)
        ),
    ];
    for xml in bad {
        assert!(
            parse_table_column_style_properties(&xml).is_err(),
            "accepted {xml}"
        );
    }
    assert!(
        set_table_column_style_properties_xml(
            &wrap(""),
            &TableColumnStyleProperties::named("C", None).unwrap()
        )
        .is_err()
    );
}
