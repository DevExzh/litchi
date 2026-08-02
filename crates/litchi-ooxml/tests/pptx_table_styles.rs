use litchi_ooxml::pptx::shape::Shape;
use litchi_ooxml::pptx::table::Table;
use litchi_ooxml::pptx::{Package, TableStylePartKind};

const SHAPES: &str = concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/../../test-data/ooxml/pptx/shapes.pptx"
);

#[test]
fn package_loads_table_styles_part() {
    let package = Package::open(SHAPES).unwrap();

    let styles = package.table_styles().unwrap().unwrap();
    assert_eq!(
        styles.default_style_id.as_deref(),
        Some("{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}")
    );
    assert_eq!(styles.styles.len(), 2);

    let medium = styles
        .find("{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}")
        .unwrap();
    assert_eq!(
        medium.style_name.as_deref(),
        Some("Medium Style 2 - Accent 1")
    );
    for part in [
        TableStylePartKind::WholeTable,
        TableStylePartKind::OddRowBand,
        TableStylePartKind::EvenRowBand,
        TableStylePartKind::OddColumnBand,
        TableStylePartKind::EvenColumnBand,
        TableStylePartKind::FirstColumn,
        TableStylePartKind::FirstRow,
        TableStylePartKind::LastColumn,
        TableStylePartKind::LastRow,
    ] {
        assert!(medium.has(part), "missing part style {}", part.xml_name());
    }
    assert!(!medium.has(TableStylePartKind::NorthWestCell));

    let plain = styles
        .find("{5940675A-B579-460E-94D1-54222C63F5DA}")
        .unwrap();
    assert!(plain.has(TableStylePartKind::WholeTable));
    assert!(!plain.has(TableStylePartKind::FirstRow));
}

#[test]
fn slide_tables_report_style_switches_and_references() {
    let package = Package::open(SHAPES).unwrap();
    let presentation = package.presentation().unwrap();

    let mut found = Vec::new();
    for slide in presentation.slides().unwrap() {
        for shape in slide.shapes().unwrap().iter() {
            if let Shape::Table(shape) = shape {
                let table = Table::from_graphic_frame_xml(shape.common().xml().unwrap()).unwrap();
                let properties = table.properties().unwrap().unwrap();
                found.push(properties);
            }
        }
    }

    assert!(!found.is_empty());
    assert!(
        found
            .iter()
            .all(|properties| properties.first_row && properties.band_row)
    );
    let medium = found.iter().find(|properties| {
        properties.style_id.as_deref() == Some("{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}")
    });
    let plain = found.iter().find(|properties| {
        properties.style_id.as_deref() == Some("{5940675A-B579-460E-94D1-54222C63F5DA}")
    });
    assert!(medium.is_some());
    assert!(plain.is_some());

    // Every referenced style resolves in the package's table styles part.
    let styles = package.table_styles().unwrap().unwrap();
    for properties in &found {
        let style_id = properties.style_id.as_deref().unwrap();
        assert!(
            styles.find(style_id).is_some(),
            "unresolved style {style_id}"
        );
    }
}

#[test]
fn presentation_table_styles_match_package_level() {
    let package = Package::open(SHAPES).unwrap();
    let from_package = package.table_styles().unwrap().unwrap();
    let from_presentation = package
        .presentation()
        .unwrap()
        .table_styles()
        .unwrap()
        .unwrap();
    assert_eq!(from_package, from_presentation);
}
