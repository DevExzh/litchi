use litchi_odt::Builder;
use litchi_odt::style::table::{column, row};
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
    let set = column::parse(&xml).unwrap();
    let style = set.get("Col").unwrap();
    let props = style.properties.as_ref().unwrap();
    assert_eq!(props.column_width.as_ref().unwrap().as_str(), "2.5cm");
    assert_eq!(props.rel_column_width.as_ref().unwrap().as_str(), "12*");
    assert_eq!(props.use_optimal_column_width, Some(true));
    assert_eq!(props.break_before, Some(row::Break::Page));
    assert_eq!(props.break_after, Some(row::Break::Auto));
    let fragment = style.to_xml_fragment().unwrap();
    assert_eq!(
        column::parse(&wrap(&fragment)).unwrap().get("Col"),
        Some(style)
    );
    let changed = column::Style::named(
        "Col",
        Some(column::Properties {
            rel_column_width: Some(column::RelativeWidth::new("3*").unwrap()),
            ..Default::default()
        }),
    )
    .unwrap();
    let updated = column::set_xml(&xml, &changed).unwrap();
    assert!(!updated.contains("style:column-width"));
    assert!(updated.contains("rel-column-width=\"3*\""));
    assert_eq!(column::parse(&updated).unwrap().get("Col"), Some(&changed));
}

#[test]
fn default_style_and_insertion_into_empty_style() {
    let xml = wrap(r#"<s:default-style s:family="table-column"/>"#);
    let set = column::parse(&xml).unwrap();
    assert!(set.default_style().unwrap().properties.is_none());
    let changed = column::Style::default_style(Some(column::Properties {
        column_width: Some(column::Length::new("0.889in").unwrap()),
        break_before: Some(row::Break::Auto),
        ..Default::default()
    }));
    let updated = column::set_xml(&xml, &changed).unwrap();
    assert_eq!(
        column::parse(&updated).unwrap().default_style(),
        Some(&changed)
    );
    let removed = column::set_xml(&updated, &column::Style::default_style(None)).unwrap();
    assert!(removed.contains("default-style"));
    assert!(!removed.contains("table-column-properties"));
}

#[test]
fn parses_real_odfdo_and_libreoffice_fixtures() {
    let odfdo = include_str!("../../../test-data/odfdo/tests/samples/test_flat_lo.fods");
    let columns = column::parse(odfdo).unwrap();
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
    let flat = litchi_odt::generic::FlatDocument::from_reader(Cursor::new(lo)).unwrap();
    assert!(!flat.column_style_properties().unwrap().styles.is_empty());
}

#[test]
fn builder_package_round_trip() {
    let properties = column::Properties {
        column_width: Some(column::Length::new("4cm").unwrap()),
        rel_column_width: Some(column::RelativeWidth::new("65535*").unwrap()),
        use_optimal_column_width: Some(false),
        break_after: Some(row::Break::Column),
        ..Default::default()
    };
    let style = column::Style::named("Col", Some(properties)).unwrap();
    let mut builder = Builder::new();
    builder
        .add_table_column_property_style(style.clone())
        .unwrap();
    builder.add_paragraph("x").unwrap();
    let package = litchi_odt::generic::Package::from_bytes(builder.build().unwrap()).unwrap();
    assert_eq!(
        package.column_style_properties().unwrap().get("Col"),
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
        assert!(column::parse(&xml).is_err(), "accepted {xml}");
    }
    assert!(column::set_xml(&wrap(""), &column::Style::named("C", None).unwrap()).is_err());
}
