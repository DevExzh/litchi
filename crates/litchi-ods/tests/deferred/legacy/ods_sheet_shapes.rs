//! General drawing-shape support in ODS `table:shapes` containers.
//!
//! Covers authoring through `Builder` and `MutableSpreadsheet`,
//! semantic reparsing, ODS-specific end-cell anchoring, and lossless
//! round-trips of packaged and flat spreadsheets.

use litchi_ods::{
    DrawingAttribute, DrawingAttributeNamespace, DrawingShapeKind, EnhancedGeometry,
    EnhancedGeometryChild, EnhancedGeometryChildKind, FlatSpreadsheet, MutableSpreadsheet, Shape,
    SheetShape, SheetShapeAnchor, Spreadsheet, Builder,
};

const SHEET_SHAPES_FODS: &str = include_str!("../../../test-data/odf/ods/sheet-shapes.fods");

fn draw_attribute(local_name: &str, value: &str) -> DrawingAttribute {
    DrawingAttribute::new(DrawingAttributeNamespace::Drawing, local_name, value).unwrap()
}

fn svg_attribute(local_name: &str, value: &str) -> DrawingAttribute {
    DrawingAttribute::new(DrawingAttributeNamespace::Svg, local_name, value).unwrap()
}

fn rectangle_shape() -> Shape {
    Shape {
        shape_type: litchi_core::ShapeType::AutoShape,
        drawing_kind: Some(DrawingShapeKind::Rectangle),
        name: Some("Box".to_string()),
        style_name: Some("gr1".to_string()),
        x: Some("1cm".to_string()),
        y: Some("2cm".to_string()),
        width: Some("4cm".to_string()),
        height: Some("2cm".to_string()),
        text: "Quarterly totals".to_string(),
        ..Shape::new()
    }
}

fn anchored(shape: Shape, end_cell_address: &str) -> SheetShape {
    let mut anchor = SheetShapeAnchor::new();
    anchor
        .set_end_cell_address(Some(end_cell_address.to_string()))
        .unwrap();
    anchor.set_end_x(Some("0.4cm".to_string())).unwrap();
    anchor.set_end_y(Some("0.2cm".to_string())).unwrap();
    anchor.set_table_background(Some(false));
    SheetShape::with_anchor(shape, anchor).unwrap()
}

fn authored_shapes(sheet_name: &str) -> Vec<SheetShape> {
    let mut star = Shape {
        shape_type: litchi_core::ShapeType::AutoShape,
        drawing_kind: Some(DrawingShapeKind::CustomShape),
        name: Some("Star".to_string()),
        width: Some("2cm".to_string()),
        height: Some("2cm".to_string()),
        text: "Go".to_string(),
        ..Shape::new()
    };
    let mut geometry = EnhancedGeometry::new();
    geometry.attributes_mut().extend([
        draw_attribute("type", "star5"),
        svg_attribute("viewBox", "0 0 21600 21600"),
    ]);
    let mut equation = EnhancedGeometryChild::new(EnhancedGeometryChildKind::Equation);
    equation.attributes_mut().extend([
        draw_attribute("name", "f0"),
        draw_attribute("formula", "$0 *2"),
    ]);
    geometry.children_mut().push(equation);
    star.enhanced_geometry = Some(geometry);

    let group = Shape {
        shape_type: litchi_core::ShapeType::Group,
        drawing_kind: Some(DrawingShapeKind::Group),
        name: Some("Cluster".to_string()),
        children: vec![
            Shape {
                shape_type: litchi_core::ShapeType::AutoShape,
                drawing_kind: Some(DrawingShapeKind::Polyline),
                name: Some("Trend".to_string()),
                width: Some("3cm".to_string()),
                height: Some("2cm".to_string()),
                drawing_attributes: vec![
                    svg_attribute("viewBox", "0 0 3000 2000"),
                    draw_attribute("points", "0,2000 1000,500 3000,0"),
                ],
                ..Shape::new()
            },
            Shape {
                shape_type: litchi_core::ShapeType::AutoShape,
                drawing_kind: Some(DrawingShapeKind::Polygon),
                name: Some("Wedge".to_string()),
                width: Some("2cm".to_string()),
                height: Some("2cm".to_string()),
                drawing_attributes: vec![
                    svg_attribute("viewBox", "0 0 2000 2000"),
                    draw_attribute("points", "0,2000 1000,0 2000,2000"),
                ],
                ..Shape::new()
            },
        ],
        ..Shape::new()
    };

    let line = Shape {
        shape_type: litchi_core::ShapeType::Line,
        drawing_kind: Some(DrawingShapeKind::Line),
        name: Some("Divider".to_string()),
        x: Some("0.5cm".to_string()),
        y: Some("4cm".to_string()),
        width: Some("8cm".to_string()),
        height: Some("4cm".to_string()),
        ..Shape::new()
    };

    let connector = Shape {
        shape_type: litchi_core::ShapeType::Connector,
        drawing_kind: Some(DrawingShapeKind::Connector),
        name: Some("Link".to_string()),
        x: Some("1cm".to_string()),
        y: Some("6cm".to_string()),
        width: Some("6cm".to_string()),
        height: Some("8cm".to_string()),
        drawing_attributes: vec![draw_attribute("type", "line")],
        ..Shape::new()
    };

    let text_box = Shape {
        shape_type: litchi_core::ShapeType::TextBox,
        drawing_kind: Some(DrawingShapeKind::Frame),
        name: Some("Note".to_string()),
        x: Some("2cm".to_string()),
        y: Some("9cm".to_string()),
        width: Some("5cm".to_string()),
        height: Some("1.5cm".to_string()),
        text: "Review before publishing".to_string(),
        ..Shape::new()
    };
    let mut note_anchor = SheetShapeAnchor::new();
    note_anchor
        .set_end_cell_address(Some(format!("{sheet_name}.B12")))
        .unwrap();

    vec![
        anchored(rectangle_shape(), &format!("{sheet_name}.C4")),
        SheetShape::new(line).unwrap(),
        SheetShape::new(star).unwrap(),
        SheetShape::new(group).unwrap(),
        SheetShape::new(connector).unwrap(),
        SheetShape::with_anchor(text_box, note_anchor).unwrap(),
    ]
}

fn assert_shape_matches(actual: &Shape, expected: &Shape, label: &str) {
    assert_eq!(actual.drawing_kind(), expected.drawing_kind(), "{label}");
    assert_eq!(actual.shape_type(), expected.shape_type(), "{label}");
    assert_eq!(actual.name(), expected.name(), "{label}");
    assert_eq!(actual.position(), expected.position(), "{label}");
    assert_eq!(actual.dimensions(), expected.dimensions(), "{label}");
    assert_eq!(actual.text, expected.text, "{label}");
    assert_eq!(actual.z_index(), expected.z_index(), "{label}");
    assert_eq!(
        actual.drawing_attributes(),
        expected.drawing_attributes(),
        "{label}"
    );
    assert_eq!(
        actual.enhanced_geometry(),
        expected.enhanced_geometry(),
        "{label}"
    );
    assert_eq!(
        actual.children().len(),
        expected.children().len(),
        "{label}"
    );
    for (actual_child, expected_child) in actual.children().iter().zip(expected.children()) {
        assert_shape_matches(actual_child, expected_child, label);
    }
}

#[test]
fn builder_authors_general_shapes_that_reparse() {
    let expected = authored_shapes("Data");
    let mut builder = Builder::new();
    builder.add_sheet("Data").unwrap();
    builder.add_row_with_values(&["Region"]).unwrap();
    for shape in &expected {
        builder.add_sheet_shape(shape.clone()).unwrap();
    }

    let bytes = builder.build().unwrap();
    let mut spreadsheet = Spreadsheet::from_bytes(bytes).unwrap();
    let sheets = spreadsheet.sheets().unwrap();
    assert_eq!(sheets.len(), 1);
    let shapes = sheets[0].shapes();
    assert_eq!(shapes.len(), expected.len());
    for (actual, wanted) in shapes.iter().zip(&expected) {
        assert_eq!(actual.anchor(), wanted.anchor());
        assert_shape_matches(actual.shape(), wanted.shape(), "builder round trip");
    }
    assert_eq!(
        shapes[0].anchor().end_cell_address(),
        Some("Data.C4"),
        "end-cell anchoring must survive the packaged round trip"
    );
    assert_eq!(shapes[0].anchor().end_x(), Some("0.4cm"));
    assert_eq!(shapes[0].anchor().end_y(), Some("0.2cm"));
    assert_eq!(shapes[0].anchor().table_background(), Some(false));
}

#[test]
fn mutable_sheet_shape_crud_round_trips() {
    let mut mutable = MutableSpreadsheet::new();
    mutable.add_sheet("Board").unwrap();
    mutable
        .add_sheet_shape(0, anchored(rectangle_shape(), "Board.C4"))
        .unwrap();

    let ellipse = Shape {
        shape_type: litchi_core::ShapeType::AutoShape,
        drawing_kind: Some(DrawingShapeKind::Ellipse),
        name: Some("Bubble".to_string()),
        width: Some("2cm".to_string()),
        height: Some("2cm".to_string()),
        ..Shape::new()
    };
    mutable
        .insert_sheet_shape(0, 0, SheetShape::new(ellipse).unwrap())
        .unwrap();
    assert_eq!(mutable.sheet_shapes(0).unwrap().len(), 2);
    assert_eq!(
        mutable.sheet_shapes(0).unwrap()[0].shape().name(),
        Some("Bubble")
    );

    let measure = Shape {
        shape_type: litchi_core::ShapeType::Line,
        drawing_kind: Some(DrawingShapeKind::Measure),
        name: Some("Ruler".to_string()),
        x: Some("1cm".to_string()),
        y: Some("1cm".to_string()),
        width: Some("4cm".to_string()),
        height: Some("1cm".to_string()),
        ..Shape::new()
    };
    let replaced = mutable
        .set_sheet_shape(0, 0, SheetShape::new(measure).unwrap())
        .unwrap();
    assert_eq!(replaced.shape().name(), Some("Bubble"));

    let bytes = mutable.to_bytes().unwrap();
    let reparsed = Spreadsheet::from_bytes(bytes).unwrap();
    let mut mutated = MutableSpreadsheet::from_spreadsheet(reparsed).unwrap();
    {
        let shapes = mutated.sheet_shapes(0).unwrap();
        assert_eq!(shapes.len(), 2);
        assert_eq!(shapes[0].shape().name(), Some("Ruler"));
        assert_eq!(
            shapes[0].shape().drawing_kind(),
            Some(DrawingShapeKind::Measure)
        );
        assert_eq!(shapes[1].anchor().end_cell_address(), Some("Board.C4"));
    }

    let removed = mutated.remove_sheet_shape(0, 0).unwrap();
    assert_eq!(removed.shape().name(), Some("Ruler"));
    let bytes = mutated.to_bytes().unwrap();
    let mut final_parse = Spreadsheet::from_bytes(bytes).unwrap();
    let sheets = final_parse.sheets().unwrap();
    assert_eq!(sheets[0].shapes().len(), 1);
    assert_eq!(sheets[0].shapes()[0].shape().name(), Some("Box"));
}

#[test]
fn fixture_shapes_reparse_losslessly_through_flat_rewrite() {
    let mut document = FlatSpreadsheet::from_bytes(SHEET_SHAPES_FODS.as_bytes().to_vec()).unwrap();
    let original: Vec<SheetShape> = {
        let sheets = document.spreadsheet_mut().sheets().unwrap();
        assert_eq!(sheets.len(), 1);
        assert_eq!(
            sheets[0].images().len(),
            1,
            "picture frames belong to the sheet-image model"
        );
        assert_eq!(
            sheets[0].images()[0]
                .frame
                .as_ref()
                .unwrap()
                .end_cell_address
                .as_deref(),
            Some("Shapes.F2")
        );
        sheets[0].shapes().to_vec()
    };
    assert_eq!(original.len(), 6);
    assert_eq!(
        original[0].shape().drawing_kind(),
        Some(DrawingShapeKind::Rectangle)
    );
    assert_eq!(original[0].anchor().end_cell_address(), Some("Shapes.C4"));
    assert_eq!(original[0].anchor().end_x(), Some("0.4cm"));
    assert_eq!(original[0].anchor().end_y(), Some("0.2cm"));
    assert_eq!(original[0].anchor().table_background(), Some(false));
    assert_eq!(original[0].shape().text, "Quarterly totals");
    assert_eq!(
        original[0].shape().drawing_attributes(),
        &[draw_attribute("corner-radius", "0.15cm")]
    );
    assert_eq!(
        original[1].shape().drawing_kind(),
        Some(DrawingShapeKind::Line)
    );
    assert_eq!(original[1].anchor().end_cell_address(), Some("Shapes.E6"));
    let star = original[2].shape();
    assert_eq!(star.drawing_kind(), Some(DrawingShapeKind::CustomShape));
    let geometry = star.enhanced_geometry().unwrap();
    assert_eq!(geometry.children().len(), 2);
    assert_eq!(
        geometry.children()[0].kind(),
        EnhancedGeometryChildKind::Equation
    );
    assert_eq!(
        geometry.children()[1].kind(),
        EnhancedGeometryChildKind::Handle
    );
    let group = original[3].shape();
    assert_eq!(group.drawing_kind(), Some(DrawingShapeKind::Group));
    assert_eq!(group.children().len(), 2);
    assert_eq!(
        group.children()[0].drawing_kind(),
        Some(DrawingShapeKind::Polyline)
    );
    assert_eq!(
        group.children()[1].drawing_kind(),
        Some(DrawingShapeKind::Polygon)
    );
    assert_eq!(
        original[4].shape().drawing_kind(),
        Some(DrawingShapeKind::Connector)
    );
    assert_eq!(
        original[5].shape().drawing_kind(),
        Some(DrawingShapeKind::Frame)
    );
    assert_eq!(original[5].shape().text, "Review before publishing");
    assert_eq!(original[5].anchor().end_cell_address(), Some("Shapes.B12"));
    assert_eq!(original[5].anchor().end_x(), Some("1cm"));
    assert_eq!(original[5].anchor().end_y(), Some("0.5cm"));

    // Rewrite through the flat mutable wrapper and reparse both writers.
    let mutable = document.into_mutable().unwrap();
    let flat_bytes = mutable.to_bytes().unwrap();
    let packaged_bytes = mutable.spreadsheet().to_bytes().unwrap();

    let mut flat_reparse = FlatSpreadsheet::from_bytes(flat_bytes).unwrap();
    let mut packaged_reparse = Spreadsheet::from_bytes(packaged_bytes).unwrap();
    for (label, sheets) in [
        ("flat", flat_reparse.spreadsheet_mut().sheets().unwrap()),
        ("packaged", packaged_reparse.sheets().unwrap()),
    ] {
        let reparsed = sheets[0].shapes();
        assert_eq!(reparsed.len(), original.len(), "{label} writer");
        for (actual, wanted) in reparsed.iter().zip(&original) {
            assert_eq!(actual.anchor(), wanted.anchor(), "{label} writer");
            assert_shape_matches(actual.shape(), wanted.shape(), label);
        }
        assert_eq!(sheets[0].images().len(), 1, "{label} writer keeps images");
    }
}

/// Reads `content.xml` out of a packaged spreadsheet.
fn packaged_content(bytes: Vec<u8>) -> String {
    use std::io::Read;
    let mut archive = zip::ZipArchive::new(std::io::Cursor::new(bytes)).unwrap();
    let mut content = String::new();
    archive
        .by_name("content.xml")
        .unwrap()
        .read_to_string(&mut content)
        .unwrap();
    content
}

/// The ODF 1.3 `table:table` content model places `table:shapes` after the
/// table preamble but ahead of the column and row groups.
fn assert_shapes_precede_structure(content: &str, label: &str) {
    let shapes = content
        .find("<table:shapes>")
        .unwrap_or_else(|| panic!("{label} writer emitted no table:shapes container"));
    let column = content
        .find("<table:table-column")
        .unwrap_or_else(|| panic!("{label} writer emitted no columns"));
    let row = content
        .find("<table:table-row")
        .unwrap_or_else(|| panic!("{label} writer emitted no rows"));
    assert!(
        shapes < column,
        "{label} writer must place table:shapes before the column group"
    );
    assert!(
        shapes < row,
        "{label} writer must place table:shapes before the row group"
    );
}

#[test]
fn table_shapes_precede_column_and_row_groups() {
    let mut builder = Builder::new();
    builder.add_sheet("Data").unwrap();
    builder.add_row_with_values(&["Region"]).unwrap();
    builder
        .add_sheet_shape(anchored(rectangle_shape(), "Data.C4"))
        .unwrap();
    assert_shapes_precede_structure(&packaged_content(builder.build().unwrap()), "builder");

    let mut mutable = MutableSpreadsheet::new();
    mutable.add_sheet("Data").unwrap();
    mutable
        .add_row(
            0,
            vec![litchi_ods::SCell::new(
                litchi_ods::CellValue::Text("Region".to_string()),
                "Region",
                0,
                0,
            )],
        )
        .unwrap();
    mutable
        .add_sheet_shape(0, anchored(rectangle_shape(), "Data.C4"))
        .unwrap();
    assert_shapes_precede_structure(&packaged_content(mutable.to_bytes().unwrap()), "mutable");
}

#[test]
fn rejects_untyped_anchors_and_reserved_shape_kinds() {
    let mut raw = rectangle_shape();
    raw.drawing_attributes.push(
        DrawingAttribute::new(
            DrawingAttributeNamespace::Table,
            "end-cell-address",
            "Data.A1",
        )
        .unwrap(),
    );
    assert!(SheetShape::new(raw).is_err());

    let mut kindless = rectangle_shape();
    kindless.drawing_kind = None;
    assert!(SheetShape::new(kindless).is_err());

    let picture = rectangle_shape().with_image_href("Pictures/logo.png");
    assert!(SheetShape::new(picture).is_err());

    let mut anchor = SheetShapeAnchor::new();
    assert!(anchor.set_end_x(Some("not-a-length".to_string())).is_err());
    assert!(anchor.set_end_cell_address(Some(String::new())).is_err());
}
