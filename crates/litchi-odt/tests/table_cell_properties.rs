use litchi_odt::{
    CellBorder, CellBorderWidths, CellDirection, CellLength, CellProtect, CellRotationAlign,
    CellRotationAngle, CellTextAlignSource, CellVerticalAlign, CellWrapOption, DocumentBuilder,
    FlatOpenDocument, OpenDocumentPackage, TableCellProperties, TableCellStyleProperties,
    TableRowBackgroundColor, TableRowBackgroundImage, TableRowBackgroundRepeat,
    TableRowBackgroundSource, TableShadow, TableWritingMode, parse_table_cell_style_properties,
    set_table_cell_style_properties_xml,
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
fn complete_round_trip_and_mutation() {
    let xml = wrap(
        r##"<s:style s:name="Cell" s:family="table-cell"><s:table-cell-properties s:vertical-align="middle" s:text-align-source="value-type" s:direction="ttb" s:glyph-orientation-vertical="0deg" s:writing-mode="tb-rl" s:shadow="none" f:background-color="#FFCC00" f:border="0.74pt solid #808080" f:border-top="none" s:diagonal-tl-br="0.74pt solid #808080" s:diagonal-tl-br-widths="0.01cm 0.07cm 0.01cm" s:diagonal-bl-tr="none" s:border-line-width="0.002cm 0.07cm 0.002cm" f:padding="0.05cm" f:padding-left="0.1cm" f:wrap-option="wrap" s:rotation-angle="270" s:rotation-align="center" s:cell-protect="protected formula-hidden" s:print-content="false" s:decimal-places="3" s:repeat-content="true" s:shrink-to-fit="true"/></s:style>"##,
    );
    let set = parse_table_cell_style_properties(&xml).unwrap();
    let style = set.get("Cell").unwrap();
    let props = style.properties.as_ref().unwrap();
    assert_eq!(props.vertical_align, Some(CellVerticalAlign::Middle));
    assert_eq!(
        props.text_align_source,
        Some(CellTextAlignSource::ValueType)
    );
    assert_eq!(props.direction, Some(CellDirection::Ttb));
    assert_eq!(
        props.glyph_orientation_vertical.as_ref().unwrap().as_str(),
        "0deg"
    );
    assert_eq!(props.writing_mode, Some(TableWritingMode::TbRl));
    assert_eq!(props.shadow.as_ref().unwrap().as_str(), "none");
    assert_eq!(props.background_color.as_ref().unwrap().as_str(), "#FFCC00");
    assert_eq!(
        props.border.as_ref().unwrap().as_str(),
        "0.74pt solid #808080"
    );
    assert_eq!(
        props.diagonal_tl_br_widths.as_ref().unwrap().space.as_str(),
        "0.07cm"
    );
    assert_eq!(props.padding_left.as_ref().unwrap().as_str(), "0.1cm");
    assert_eq!(props.wrap_option, Some(CellWrapOption::Wrap));
    assert_eq!(props.rotation_angle.unwrap().degrees(), 270);
    assert_eq!(props.rotation_align, Some(CellRotationAlign::Center));
    assert_eq!(
        props.cell_protect,
        Some(CellProtect::ProtectedFormulaHidden)
    );
    assert_eq!(props.print_content, Some(false));
    assert_eq!(props.decimal_places, Some(3));
    let fragment = style.to_xml_fragment().unwrap();
    assert_eq!(
        parse_table_cell_style_properties(&wrap(&fragment))
            .unwrap()
            .get("Cell"),
        Some(style)
    );
    let changed = TableCellStyleProperties::named(
        "Cell",
        Some(TableCellProperties {
            vertical_align: Some(CellVerticalAlign::Bottom),
            cell_protect: Some(CellProtect::HiddenAndProtected),
            ..Default::default()
        }),
    )
    .unwrap();
    let updated = set_table_cell_style_properties_xml(&xml, &changed).unwrap();
    assert!(!updated.contains("fo:border"));
    assert!(updated.contains("style:cell-protect=\"hidden-and-protected\""));
    assert_eq!(
        parse_table_cell_style_properties(&updated)
            .unwrap()
            .get("Cell"),
        Some(&changed)
    );
}

#[test]
fn background_image_embedded_and_linked_round_trip() {
    let xml = wrap(
        r#"<s:style s:name="Bg" s:family="table-cell"><s:table-cell-properties><s:background-image s:repeat="no-repeat" s:position="right" d:opacity="75%" x:type="simple" x:href="Pictures/a.png" x:show="embed" x:actuate="onLoad"/></s:table-cell-properties></s:style>"#,
    );
    let set = parse_table_cell_style_properties(&xml).unwrap();
    let image = set
        .get("Bg")
        .unwrap()
        .properties
        .as_ref()
        .unwrap()
        .background_image
        .as_ref()
        .unwrap();
    assert_eq!(image.repeat, Some(TableRowBackgroundRepeat::NoRepeat));
    let fragment = set.get("Bg").unwrap().to_xml_fragment().unwrap();
    assert_eq!(
        parse_table_cell_style_properties(&wrap(&fragment)).unwrap(),
        set
    );
    let embedded = wrap(
        r#"<s:default-style s:family="table-cell"><s:table-cell-properties><s:background-image><o:binary-data>AQIDBA==</o:binary-data></s:background-image></s:table-cell-properties></s:default-style>"#,
    );
    let set = parse_table_cell_style_properties(&embedded).unwrap();
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
    assert_eq!(parse_table_cell_style_properties(&wrap(&out)).unwrap(), set);
}

#[test]
fn cell_protect_token_forms() {
    let xml = wrap(
        r#"<s:style s:name="P1" s:family="table-cell"><s:table-cell-properties s:cell-protect="formula-hidden protected"/></s:style><s:style s:name="P2" s:family="table-cell"><s:table-cell-properties s:cell-protect="protected"/></s:style><s:style s:name="P3" s:family="table-cell"><s:table-cell-properties s:cell-protect="none"/></s:style>"#,
    );
    let set = parse_table_cell_style_properties(&xml).unwrap();
    assert_eq!(
        set.get("P1")
            .unwrap()
            .properties
            .as_ref()
            .unwrap()
            .cell_protect,
        Some(CellProtect::ProtectedFormulaHidden)
    );
    assert_eq!(
        set.get("P2")
            .unwrap()
            .properties
            .as_ref()
            .unwrap()
            .cell_protect,
        Some(CellProtect::Protected)
    );
    assert_eq!(
        set.get("P3")
            .unwrap()
            .properties
            .as_ref()
            .unwrap()
            .cell_protect,
        Some(CellProtect::None)
    );
}

#[test]
fn parses_real_odfdo_and_libreoffice_fixtures() {
    let odfdo = include_str!("../../../test-data/odfdo/tests/samples/test_flat_lo.fods");
    let cells = parse_table_cell_style_properties(odfdo).unwrap();
    assert!(!cells.styles.is_empty());
    assert!(
        cells
            .styles
            .iter()
            .filter_map(|x| x.properties.as_ref())
            .any(|p| p.wrap_option == Some(CellWrapOption::NoWrap))
    );
    let lo = include_bytes!(
        "../../../test-data/libreoffice-core/sc/qa/unit/data/functions/mathematical/fods/aggregate.fods"
    );
    let flat = FlatOpenDocument::from_reader(Cursor::new(lo)).unwrap();
    let cells = flat.table_cell_style_properties().unwrap();
    assert!(!cells.styles.is_empty());
    assert!(
        cells
            .styles
            .iter()
            .filter_map(|x| x.properties.as_ref())
            .any(|p| p.rotation_angle == CellRotationAngle::new(90).ok())
    );
}

#[test]
fn builder_package_round_trip() {
    let properties = TableCellProperties {
        vertical_align: Some(CellVerticalAlign::Automatic),
        background_color: Some(TableRowBackgroundColor::new("transparent").unwrap()),
        border: Some(CellBorder::new("0.74pt solid #808080").unwrap()),
        border_line_width: Some(CellBorderWidths {
            inner_width: CellLength::positive("0.002cm").unwrap(),
            space: CellLength::positive("0.07cm").unwrap(),
            outer_width: CellLength::positive("0.002cm").unwrap(),
        }),
        padding: Some(CellLength::non_negative("0cm").unwrap()),
        decimal_places: Some(2),
        shrink_to_fit: Some(true),
        background_image: Some(TableRowBackgroundImage {
            repeat: None,
            position: None,
            filter_name: None,
            opacity: None,
            source: TableRowBackgroundSource::Empty,
        }),
        ..Default::default()
    };
    let style = TableCellStyleProperties::named("Cell", Some(properties)).unwrap();
    let mut builder = DocumentBuilder::new();
    builder
        .add_table_cell_property_style(style.clone())
        .unwrap();
    builder.add_paragraph("x").unwrap();
    let package = OpenDocumentPackage::from_bytes(builder.build().unwrap()).unwrap();
    assert_eq!(
        package.table_cell_style_properties().unwrap().get("Cell"),
        Some(&style)
    );
}

#[test]
fn rejects_malformed_and_unsafe_forms() {
    let bad = [
        wrap(
            r#"<s:style s:name="C" s:family="table-cell"><s:table-cell-properties s:vertical-align="baseline"/></s:style>"#,
        ),
        wrap(
            r#"<s:style s:name="C" s:family="table-cell"><s:table-cell-properties s:direction="rtl"/></s:style>"#,
        ),
        wrap(
            r#"<s:style s:name="C" s:family="table-cell"><s:table-cell-properties s:glyph-orientation-vertical="90"/></s:style>"#,
        ),
        wrap(
            r#"<s:style s:name="C" s:family="table-cell"><s:table-cell-properties f:padding="-1cm"/></s:style>"#,
        ),
        wrap(
            r#"<s:style s:name="C" s:family="table-cell"><s:table-cell-properties s:border-line-width="1cm 2cm"/></s:style>"#,
        ),
        wrap(
            r#"<s:style s:name="C" s:family="table-cell"><s:table-cell-properties s:border-line-width="0cm 1cm 1cm"/></s:style>"#,
        ),
        wrap(
            r#"<s:style s:name="C" s:family="table-cell"><s:table-cell-properties s:rotation-angle="360"/></s:style>"#,
        ),
        wrap(
            r#"<s:style s:name="C" s:family="table-cell"><s:table-cell-properties s:rotation-angle="-90"/></s:style>"#,
        ),
        wrap(
            r#"<s:style s:name="C" s:family="table-cell"><s:table-cell-properties s:cell-protect="protected protected"/></s:style>"#,
        ),
        wrap(
            r#"<s:style s:name="C" s:family="table-cell"><s:table-cell-properties s:cell-protect="hidden"/></s:style>"#,
        ),
        wrap(
            r#"<s:style s:name="C" s:family="table-cell"><s:table-cell-properties s:decimal-places="-1"/></s:style>"#,
        ),
        wrap(
            r#"<s:style s:name="C" s:family="table-cell"><s:table-cell-properties s:print-content="1"/></s:style>"#,
        ),
        wrap(
            r#"<s:style s:name="C" s:family="table-cell"><s:table-cell-properties s:foo="bar"/></s:style>"#,
        ),
        wrap(
            r#"<s:style s:name="C" s:family="table-cell"><s:table-cell-properties/><s:table-cell-properties/></s:style>"#,
        ),
        wrap(
            r#"<s:style s:name="C" s:family="table-cell"><s:table-cell-properties><s:background-image x:href="a"/></s:table-cell-properties></s:style>"#,
        ),
        wrap(
            r#"<s:style s:name="C" s:family="table-cell"><s:table-cell-properties>text</s:table-cell-properties></s:style>"#,
        ),
        format!(
            "<!DOCTYPE x>{}",
            wrap(r#"<s:style s:name="C" s:family="table-cell"/>"#)
        ),
    ];
    for xml in bad {
        assert!(
            parse_table_cell_style_properties(&xml).is_err(),
            "accepted {xml}"
        );
    }
    let props = TableCellProperties {
        decimal_places: Some(101),
        ..Default::default()
    };
    assert!(props.validate().is_err());
    let shadow = TableShadow::new("none").unwrap();
    assert_eq!(shadow.as_str(), "none");
    assert!(
        set_table_cell_style_properties_xml(
            &wrap(""),
            &TableCellStyleProperties::named("C", None).unwrap()
        )
        .is_err()
    );
}
