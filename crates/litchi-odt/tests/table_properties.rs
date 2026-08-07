use litchi_odt::Builder;
use litchi_odt::style::table::{row, table};
use std::io::Cursor;
const O: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const S: &str = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
const F: &str = "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0";
const T: &str = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
const D: &str = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
const X: &str = "http://www.w3.org/1999/xlink";
fn wrap(x: &str) -> String {
    format!(
        r#"<o:document-styles xmlns:o="{O}" xmlns:s="{S}" xmlns:f="{F}" xmlns:t="{T}" xmlns:d="{D}" xmlns:x="{X}"><o:styles>{x}</o:styles></o:document-styles>"#
    )
}
#[test]
fn complete_round_trip_and_lossless_mutation() {
    let xml = wrap(
        r##"<s:style s:name="Table" s:family="table"><s:table-properties s:width="6.5cm" s:rel-width="80%" t:align="margins" f:margin-left="-1cm" f:margin-right="2%" f:margin-top="0cm" f:margin-bottom="3%" f:margin="1cm" s:page-number="4" f:break-before="column" f:break-after="page" f:background-color="#A0b1C2" s:shadow="none" f:keep-with-next="always" s:may-break-between-rows="false" t:border-model="separating" s:writing-mode="rl-tb" t:display="true"><s:background-image s:repeat="no-repeat" s:position="right" x:type="simple" x:href="Pictures/a.png"/></s:table-properties></s:style>"##,
    );
    let set = table::parse(&xml).unwrap();
    let style = set.get("Table").unwrap();
    let p = style.properties.as_ref().unwrap();
    assert_eq!(p.align, Some(table::Alignment::Margins));
    assert_eq!(p.page_number, Some(table::PageNumber::Number(4)));
    assert_eq!(p.writing_mode, Some(table::WritingMode::RlTb));
    let out = style.to_xml_fragment().unwrap();
    assert_eq!(table::parse(&wrap(&out)).unwrap().get("Table"), Some(style));
    let changed = table::Style::named(
        "Table",
        Some(table::Properties {
            display: Some(false),
            ..Default::default()
        }),
    )
    .unwrap();
    let updated = table::set_xml(&xml, &changed).unwrap();
    assert!(!updated.contains("t:align"));
    assert!(updated.contains("table:display=\"false\""));
}
#[test]
fn embedded_image_round_trip() {
    let xml = wrap(
        r#"<s:default-style s:family="table"><s:table-properties><s:background-image><o:binary-data>AQID</o:binary-data></s:background-image></s:table-properties></s:default-style>"#,
    );
    let set = table::parse(&xml).unwrap();
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
        row::BackgroundSource::Embedded(vec![1, 2, 3])
    );
    assert_eq!(
        table::parse(&wrap(
            &set.default_style().unwrap().to_xml_fragment().unwrap()
        ))
        .unwrap(),
        set
    );
}
#[test]
fn parses_real_odfdo_and_libreoffice() {
    let odfdo = include_str!("../../../test-data/odfdo/tests/samples/example.xml");
    assert!(!table::parse(odfdo).unwrap().styles.is_empty());
    let lo = include_bytes!(
        "../../../test-data/libreoffice-core/xmloff/qa/unit/data/floattable-wrap-all-pages2.fodt"
    );
    let flat = litchi_odt::generic::FlatDocument::from_reader(Cursor::new(lo)).unwrap();
    let styles = flat.table_style_properties().unwrap();
    assert!(styles.styles.iter().any(|x| {
        x.properties
            .as_ref()
            .is_some_and(|p| p.width.as_ref().is_some_and(|w| w.as_str() == "6.743cm"))
    }));
}
#[test]
fn builder_package_round_trip() {
    let p = table::Properties {
        width: Some(table::Width::new("8cm").unwrap()),
        relative_width: Some(table::Percent::new("100%").unwrap()),
        align: Some(table::Alignment::Center),
        margin_left: Some(table::Measure::new("-1cm").unwrap()),
        background_color: Some(row::BackgroundColor::new("transparent").unwrap()),
        background_image: Some(row::BackgroundImage {
            repeat: Some(row::Repeat::Repeat),
            position: Some(row::BackgroundPosition::Center),
            filter_name: None,
            opacity: None,
            source: row::BackgroundSource::Empty,
        }),
        break_before: Some(row::Break::Auto),
        keep_with_next: Some(row::KeepTogether::Auto),
        shadow: Some(table::Shadow::new("none").unwrap()),
        border_model: Some(table::BorderModel::Collapsing),
        ..Default::default()
    };
    let style = table::Style::named("Table", Some(p)).unwrap();
    let mut b = Builder::new();
    b.add_table_property_style(style.clone()).unwrap();
    b.add_paragraph("x").unwrap();
    let package = litchi_odt::generic::Package::from_bytes(b.build().unwrap()).unwrap();
    assert_eq!(
        package.table_style_properties().unwrap().get("Table"),
        Some(&style)
    );
}
#[test]
fn accepts_rng_trailing_dot_and_rejects_identity_namespace_lookalikes() {
    let xml = wrap(
        r#"<s:style s:name="Edge" s:family="table"><s:table-properties s:width="1.cm" s:rel-width="80.%"/></s:style>"#,
    );
    let p = table::parse(&xml).unwrap();
    assert_eq!(
        p.get("Edge")
            .unwrap()
            .properties
            .as_ref()
            .unwrap()
            .width
            .as_ref()
            .unwrap()
            .as_str(),
        "1.cm"
    );
    assert!(
        table::parse(&wrap(
            r#"<s:style s:name="X" family="table"><s:table-properties/></s:style>"#
        ))
        .is_err()
    );
}
#[test]
fn rejects_malformed() {
    let bad = [
        wrap(
            r#"<s:style s:name="X" s:family="table"><s:table-properties s:width="0cm"/></s:style>"#,
        ),
        wrap(
            r#"<s:style s:name="X" s:family="table"><s:table-properties f:margin-top="-1cm"/></s:style>"#,
        ),
        wrap(
            r#"<s:style s:name="X" s:family="table"><s:table-properties s:page-number="0"/></s:style>"#,
        ),
        wrap(
            r#"<s:style s:name="X" s:family="table"><s:table-properties t:display="1"/></s:style>"#,
        ),
        wrap(
            r#"<s:style s:name="X" s:family="table"><s:table-properties/><s:table-properties/></s:style>"#,
        ),
        wrap(
            r#"<s:style s:name="X" s:family="table"><s:table-properties f:border="1pt solid"><s:background-image x:href="a"/></s:table-properties></s:style>"#,
        ),
        wrap(
            r#"<s:style s:name="X" s:family="table"><s:table-properties><s:background-image><o:binary-data>***=</o:binary-data></s:background-image></s:table-properties></s:style>"#,
        ),
        format!(
            "<!DOCTYPE x>{}",
            wrap(r#"<s:style s:name="X" s:family="table"/>"#)
        ),
    ];
    for x in bad {
        assert!(table::parse(&x).is_err(), "accepted {x}");
    }
}
