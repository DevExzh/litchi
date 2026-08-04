//! Integration tests for typed DrawingML custom geometry (`a:custGeom`) in
//! XLSX worksheet drawings: parsing, round-tripping through the authoring
//! pipeline, authoring into a saved workbook, and validation failures.

use litchi_ooxml::xlsx::shape_geometry::{
    AdjustHandle, AdjustValue, ConnectionSite, CustomGeometry, Formula, Guide, Path, PathCommand,
    PathFillMode, Point, PolarAdjustHandle, Rectangle, XyAdjustHandle,
};
use litchi_ooxml::xlsx::writer::ShapeSpec;
use litchi_ooxml::xlsx::{
    CellMarker, DrawingObject, EditAs, Emu, Geometry, ShapeAnchor, Workbook, parse_drawing_shapes,
};
use litchi_opc::{OpcPackage, PackURI, constants::relationship_type as rt};

const XDR_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
const A_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";

fn drawing(body: &str) -> String {
    format!(r#"<xdr:wsDr xmlns:xdr="{XDR_NS}" xmlns:a="{A_NS}">{body}</xdr:wsDr>"#)
}

fn anchored_shape(sp_pr_children: &str) -> String {
    drawing(&format!(
        "<xdr:twoCellAnchor><xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff>\
         <xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>\
         <xdr:to><xdr:col>5</xdr:col><xdr:colOff>0</xdr:colOff>\
         <xdr:row>9</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>\
         <xdr:sp><xdr:nvSpPr><xdr:cNvPr id=\"2\" name=\"Custom\"/><xdr:cNvSpPr/></xdr:nvSpPr>\
         <xdr:spPr>{sp_pr_children}</xdr:spPr></xdr:sp>\
         <xdr:clientData/></xdr:twoCellAnchor>"
    ))
}

fn value(value: i64) -> AdjustValue {
    AdjustValue::Value(value)
}

fn marker(column: u32, row: u32) -> CellMarker {
    CellMarker {
        column,
        column_offset: Emu(0),
        row,
        row_offset: Emu(0),
    }
}

fn two_cell() -> ShapeAnchor {
    ShapeAnchor::TwoCell {
        from: marker(1, 1),
        to: marker(6, 12),
        edit_as: EditAs::TwoCell,
    }
}

/// A geometry exercising every custGeom construct: adjust values, derived
/// guides, XY and polar handles, connection sites, a text rectangle, and a
/// path with every command kind plus non-default path attributes.
fn full_geometry() -> CustomGeometry {
    CustomGeometry {
        adjust_values: vec![Guide::new("adj1", Formula::literal(50))],
        guides: vec![
            Guide::new("x1", "*/ adj1 21600 100".parse().unwrap()),
            Guide::new("y1", "pin 0 x1 21600".parse().unwrap()),
        ],
        adjust_handles: vec![
            AdjustHandle::Xy(XyAdjustHandle {
                horizontal_guide: Some("adj1".to_string()),
                minimum_x: Some(value(0)),
                maximum_x: Some(value(100)),
                position: Point::new(AdjustValue::guide("x1"), value(0)),
                ..XyAdjustHandle::default()
            }),
            AdjustHandle::Polar(PolarAdjustHandle {
                radius_guide: Some("adj1".to_string()),
                minimum_radius: Some(value(0)),
                maximum_radius: Some(value(10_800)),
                angle_guide: Some("adj1".to_string()),
                minimum_angle: Some(value(0)),
                maximum_angle: Some(value(21_600_000)),
                position: Point::new(10_800, 10_800),
            }),
        ],
        connection_sites: vec![
            ConnectionSite {
                angle: value(0),
                position: Point::new(21_600, 10_800),
            },
            ConnectionSite {
                angle: AdjustValue::guide("adj1"),
                position: Point::new(AdjustValue::guide("x1"), value(0)),
            },
        ],
        text_rectangle: Some(Rectangle {
            left: value(3_600),
            top: AdjustValue::guide("y1"),
            right: value(18_000),
            bottom: value(21_600),
        }),
        paths: vec![
            Path::new(21_600, 21_600)
                .with_command(PathCommand::MoveTo(Point::new(0, 10_800)))
                .with_command(PathCommand::LineTo(Point::new(10_800, 0)))
                .with_command(PathCommand::ArcTo {
                    width_radius: value(10_800),
                    height_radius: value(10_800),
                    start_angle: value(16_200_000),
                    swing_angle: value(5_400_000),
                })
                .with_command(PathCommand::QuadraticBezierTo {
                    control: Point::new(21_600, 21_600),
                    end: Point::new(10_800, 21_600),
                })
                .with_command(PathCommand::CubicBezierTo {
                    control1: Point::new(7_200, 21_600),
                    control2: Point::new(0, 18_000),
                    end: Point::new(AdjustValue::guide("x1"), AdjustValue::guide("y1")),
                })
                .with_command(PathCommand::Close),
            Path {
                width: 100,
                height: 100,
                fill_mode: PathFillMode::Darken,
                stroked: false,
                extrusion_allowed: false,
                commands: vec![
                    PathCommand::MoveTo(Point::new(0, 0)),
                    PathCommand::LineTo(Point::new(100, 100)),
                ],
            },
        ],
    }
}

fn parse_single_shape(xml: &str) -> litchi_ooxml::xlsx::Shape {
    let objects = parse_drawing_shapes(xml).unwrap().unwrap();
    assert_eq!(objects.len(), 1);
    let DrawingObject::Shape(shape) = objects.into_iter().next().unwrap().object else {
        panic!("expected a shape");
    };
    shape
}

#[test]
fn parses_custom_geometry_from_drawing_xml() {
    let xml = anchored_shape(
        "<a:custGeom>\
         <a:avLst><a:gd name=\"adj1\" fmla=\"val 50\"/></a:avLst>\
         <a:gdLst><a:gd name=\"x1\" fmla=\"*/ adj1 21600 100\"/></a:gdLst>\
         <a:ahLst><a:ahXY gdRefX=\"adj1\" minX=\"0\" maxX=\"100\"><a:pos x=\"x1\" y=\"0\"/></a:ahXY>\
         <a:ahPolar gdRefAng=\"adj1\" minAng=\"0\" maxAng=\"21600000\"><a:pos x=\"10800\" y=\"10800\"/></a:ahPolar></a:ahLst>\
         <a:cxnLst><a:cxn ang=\"5400000\"><a:pos x=\"10800\" y=\"21600\"/></a:cxn></a:cxnLst>\
         <a:rect l=\"0\" t=\"0\" r=\"21600\" b=\"21600\"/>\
         <a:pathLst><a:path w=\"21600\" h=\"21600\" fill=\"lightenLess\" stroke=\"0\" extrusionOk=\"0\">\
         <a:moveTo><a:pt x=\"0\" y=\"0\"/></a:moveTo>\
         <a:lnTo><a:pt x=\"21600\" y=\"0\"/></a:lnTo>\
         <a:arcTo wR=\"10800\" hR=\"10800\" stAng=\"0\" swAng=\"10800000\"/>\
         <a:quadBezTo><a:pt x=\"10800\" y=\"21600\"/><a:pt x=\"0\" y=\"21600\"/></a:quadBezTo>\
         <a:cubicBezTo><a:pt x=\"0\" y=\"14400\"/><a:pt x=\"0\" y=\"7200\"/><a:pt x=\"0\" y=\"0\"/></a:cubicBezTo>\
         <a:close/></a:path></a:pathLst>\
         </a:custGeom>",
    );
    let shape = parse_single_shape(&xml);
    assert_eq!(shape.preset(), None);
    let geometry = shape
        .geometry
        .and_then(Geometry::into_custom)
        .expect("custom geometry parsed");

    assert_eq!(
        geometry.adjust_values,
        vec![Guide::new("adj1", Formula::literal(50))]
    );
    assert_eq!(
        geometry.guides,
        vec![Guide::new(
            "x1",
            Formula::MultiplyDivide {
                x: AdjustValue::guide("adj1"),
                y: value(21_600),
                z: value(100),
            }
        )]
    );
    assert_eq!(geometry.adjust_handles.len(), 2);
    let AdjustHandle::Xy(xy) = &geometry.adjust_handles[0] else {
        panic!("expected an XY handle");
    };
    assert_eq!(xy.horizontal_guide.as_deref(), Some("adj1"));
    assert_eq!(xy.minimum_x, Some(value(0)));
    assert_eq!(xy.maximum_x, Some(value(100)));
    assert_eq!(xy.vertical_guide, None);
    assert_eq!(xy.position, Point::new(AdjustValue::guide("x1"), value(0)));
    let AdjustHandle::Polar(polar) = &geometry.adjust_handles[1] else {
        panic!("expected a polar handle");
    };
    assert_eq!(polar.angle_guide.as_deref(), Some("adj1"));
    assert_eq!(polar.maximum_angle, Some(value(21_600_000)));
    assert_eq!(polar.position, Point::new(10_800, 10_800));

    assert_eq!(
        geometry.connection_sites,
        vec![ConnectionSite {
            angle: value(5_400_000),
            position: Point::new(10_800, 21_600),
        }]
    );
    assert_eq!(
        geometry.text_rectangle,
        Some(Rectangle {
            left: value(0),
            top: value(0),
            right: value(21_600),
            bottom: value(21_600),
        })
    );

    assert_eq!(geometry.paths.len(), 1);
    let path = &geometry.paths[0];
    assert_eq!((path.width, path.height), (21_600, 21_600));
    assert_eq!(path.fill_mode, PathFillMode::LightenLess);
    assert!(!path.stroked);
    assert!(!path.extrusion_allowed);
    assert_eq!(
        path.commands,
        vec![
            PathCommand::MoveTo(Point::new(0, 0)),
            PathCommand::LineTo(Point::new(21_600, 0)),
            PathCommand::ArcTo {
                width_radius: value(10_800),
                height_radius: value(10_800),
                start_angle: value(0),
                swing_angle: value(10_800_000),
            },
            PathCommand::QuadraticBezierTo {
                control: Point::new(10_800, 21_600),
                end: Point::new(0, 21_600),
            },
            PathCommand::CubicBezierTo {
                control1: Point::new(0, 14_400),
                control2: Point::new(0, 7_200),
                end: Point::new(0, 0),
            },
            PathCommand::Close,
        ]
    );
}

#[test]
fn parses_path_attribute_defaults() {
    let xml = anchored_shape(
        "<a:custGeom><a:pathLst><a:path>\
         <a:moveTo><a:pt x=\"0\" y=\"0\"/></a:moveTo></a:path></a:pathLst></a:custGeom>",
    );
    let geometry = parse_single_shape(&xml)
        .geometry
        .and_then(Geometry::into_custom)
        .unwrap();
    let path = &geometry.paths[0];
    assert_eq!((path.width, path.height), (0, 0));
    assert_eq!(path.fill_mode, PathFillMode::Normal);
    assert!(path.stroked);
    assert!(path.extrusion_allowed);
}

#[test]
fn schema_valid_geometry_guide_tokens_are_canonicalized_without_rejection() {
    let xml = anchored_shape(
        "<a:custGeom><a:avLst>\
         <a:gd name=\"  my&#x9; guide  \" fmla=\"val 1\"/>\
         <a:gd name=\"123\" fmla=\"val 2\"/>\
         <a:gd name=\"\" fmla=\"val 3\"/>\
         </a:avLst><a:ahLst>\
         <a:ahXY gdRefX=\" my&#xA; guide \" ><a:pos x=\"0\" y=\"0\"/></a:ahXY>\
         </a:ahLst><a:pathLst/></a:custGeom>",
    );
    let geometry = parse_single_shape(&xml)
        .geometry
        .and_then(Geometry::into_custom)
        .unwrap();

    assert_eq!(geometry.adjust_values[0].name, "my guide");
    assert_eq!(geometry.adjust_values[1].name, "123");
    assert_eq!(geometry.adjust_values[2].name, "");
    let AdjustHandle::Xy(handle) = &geometry.adjust_handles[0] else {
        panic!("expected XY handle");
    };
    assert_eq!(handle.horizontal_guide.as_deref(), Some("my guide"));
}

#[test]
fn authored_geometry_round_trips_through_the_parser() {
    let geometry = full_geometry();
    let mut spec = ShapeSpec::custom("Wave", two_cell(), geometry.clone(), "wave text");
    spec.description = Some("custom wave".to_string());

    let mut workbook = Workbook::create().unwrap();
    let worksheet = workbook.worksheet_mut(0).unwrap();
    worksheet.add_shape(spec).unwrap();
    let xml = worksheet.generate_drawing_xml().unwrap().unwrap();
    assert!(xml.contains("<a:custGeom>"));
    assert!(!xml.contains("prstGeom"));

    let shape = parse_single_shape(&xml);
    assert_eq!(shape.non_visual.name.as_deref(), Some("Wave"));
    assert_eq!(shape.preset(), None);
    assert_eq!(shape.geometry, Some(geometry.into()));
    assert_eq!(shape.text_body.unwrap().text(), "wave text");
}

#[test]
fn reparsed_geometry_serializes_identically() {
    let geometry = full_geometry();
    let mut workbook = Workbook::create().unwrap();
    let worksheet = workbook.worksheet_mut(0).unwrap();
    worksheet
        .add_shape(ShapeSpec::custom("G", two_cell(), geometry, ""))
        .unwrap();
    let first = worksheet.generate_drawing_xml().unwrap().unwrap();

    // parse -> author again -> serialize must reproduce the same markup.
    let reparsed = parse_single_shape(&first)
        .geometry
        .and_then(Geometry::into_custom)
        .unwrap();
    let mut workbook = Workbook::create().unwrap();
    let worksheet = workbook.worksheet_mut(0).unwrap();
    worksheet
        .add_shape(ShapeSpec::custom("G", two_cell(), reparsed, ""))
        .unwrap();
    let second = worksheet.generate_drawing_xml().unwrap().unwrap();
    assert_eq!(first, second);
}

#[test]
fn authors_workbook_with_custom_geometry_shape() {
    let directory = tempfile::tempdir().unwrap();
    let path = directory.path().join("custom-geometry.xlsx");
    let geometry = full_geometry();
    let sheet_name;
    {
        let mut workbook = Workbook::create().unwrap();
        let worksheet = workbook.worksheet_mut(0).unwrap();
        sheet_name = worksheet.name().to_string();
        worksheet
            .add_shape(ShapeSpec::custom(
                "Wave",
                two_cell(),
                geometry.clone(),
                "hello",
            ))
            .unwrap();
        workbook.save(&path).unwrap();
    }

    let package = OpcPackage::open(&path).unwrap();
    let worksheet_part = package
        .get_part(&PackURI::new("/xl/worksheets/sheet1.xml").unwrap())
        .unwrap();
    let drawing_relationship = worksheet_part
        .rels()
        .iter()
        .find(|relationship| relationship.reltype() == rt::DRAWING)
        .unwrap();
    let worksheet_xml = std::str::from_utf8(worksheet_part.blob()).unwrap();
    assert!(worksheet_xml.contains(&format!(
        r#"<drawing r:id="{}"/>"#,
        drawing_relationship.r_id()
    )));
    assert!(
        package
            .get_part(&drawing_relationship.target_partname().unwrap())
            .is_ok()
    );

    let workbook = Workbook::open(&path).unwrap();
    let inventory = workbook.shapes_on_sheet(&sheet_name).unwrap();
    assert_eq!(inventory.objects.len(), 1);
    let DrawingObject::Shape(shape) = &inventory.objects[0].object else {
        panic!("expected a shape");
    };
    assert_eq!(shape.non_visual.name.as_deref(), Some("Wave"));
    assert_eq!(shape.preset(), None);
    assert_eq!(shape.custom_geometry(), Some(&geometry));
    assert_eq!(shape.text_body.as_ref().unwrap().text(), "hello");
}

#[test]
fn preset_shapes_still_author_preset_geometry() {
    let mut workbook = Workbook::create().unwrap();
    let worksheet = workbook.worksheet_mut(0).unwrap();
    worksheet
        .add_shape(ShapeSpec::shape(
            "Plain",
            two_cell(),
            litchi_ooxml::xlsx::Preset::Ellipse,
            "",
        ))
        .unwrap();
    let xml = worksheet.generate_drawing_xml().unwrap().unwrap();
    assert!(xml.contains(r#"<a:prstGeom prst="ellipse">"#));
    assert!(!xml.contains("custGeom"));
}

#[test]
fn validation_rejects_invalid_or_ambiguous_authored_geometry() {
    let mut workbook = Workbook::create().unwrap();
    let worksheet = workbook.worksheet_mut(0).unwrap();

    // The schema always requires a path list element; authored geometry
    // must also draw at least one path.
    let empty = ShapeSpec::custom("Empty", two_cell(), CustomGeometry::new(), "");
    assert!(worksheet.add_shape(empty).is_err());

    let numeric_guide = CustomGeometry::new()
        .with_adjust_value(Guide::new("123", Formula::literal(1)))
        .with_path(Path::new(0, 0).with_command(PathCommand::MoveTo(Point::new(0, 0))));
    let spec = ShapeSpec::custom("Numeric guide", two_cell(), numeric_guide, "");
    assert!(worksheet.add_shape(spec).is_err());

    let mut negative_path = full_geometry();
    negative_path.paths[0].width = -1;
    let spec = ShapeSpec::custom("Negative", two_cell(), negative_path, "");
    assert!(worksheet.add_shape(spec).is_err());

    assert!(worksheet.shapes().is_empty());
}

#[test]
fn parsing_rejects_structurally_invalid_geometry() {
    // A custGeom without its required path list.
    let missing_path_list = anchored_shape("<a:custGeom><a:avLst/></a:custGeom>");
    assert!(parse_drawing_shapes(&missing_path_list).is_err());

    // A quadratic Bezier with only one point.
    let short_bezier = anchored_shape(
        "<a:custGeom><a:pathLst><a:path>\
         <a:quadBezTo><a:pt x=\"0\" y=\"0\"/></a:quadBezTo></a:path></a:pathLst></a:custGeom>",
    );
    assert!(parse_drawing_shapes(&short_bezier).is_err());

    // A guide with an unknown formula operation.
    let bad_formula = anchored_shape(
        "<a:custGeom><a:avLst><a:gd name=\"adj1\" fmla=\"frob 1 2\"/></a:avLst>\
         <a:pathLst><a:path/></a:pathLst></a:custGeom>",
    );
    assert!(parse_drawing_shapes(&bad_formula).is_err());

    // An arc missing a required attribute.
    let bad_arc = anchored_shape(
        "<a:custGeom><a:pathLst><a:path>\
         <a:arcTo wR=\"1\" hR=\"1\" stAng=\"0\"/></a:path></a:pathLst></a:custGeom>",
    );
    assert!(parse_drawing_shapes(&bad_arc).is_err());

    // A connection site without its required position.
    let missing_position = anchored_shape(
        "<a:custGeom><a:cxnLst><a:cxn ang=\"0\"/></a:cxnLst>\
         <a:pathLst><a:path/></a:pathLst></a:custGeom>",
    );
    assert!(parse_drawing_shapes(&missing_position).is_err());

    for attributes in [
        "fill=\"sparkly\"",
        "stroke=\"on\"",
        "extrusionOk=\"sometimes\"",
        "w=\"-1\"",
        "h=\"27273042316901\"",
    ] {
        let malformed_domain = anchored_shape(&format!(
            "<a:custGeom><a:pathLst><a:path {attributes}/></a:pathLst></a:custGeom>"
        ));
        assert!(
            parse_drawing_shapes(&malformed_domain).is_err(),
            "accepted fixed-domain attributes {attributes}"
        );
    }
}

#[test]
fn empty_path_list_parses_but_cannot_be_authored() {
    // CT_Path2DList allows zero paths; parsing keeps the empty geometry.
    let xml = anchored_shape("<a:custGeom><a:pathLst/></a:custGeom>");
    let geometry = parse_single_shape(&xml)
        .geometry
        .and_then(Geometry::into_custom)
        .unwrap();
    assert!(geometry.paths.is_empty());

    // Re-authoring it is rejected: authored geometry must draw something.
    let mut workbook = Workbook::create().unwrap();
    let worksheet = workbook.worksheet_mut(0).unwrap();
    let spec = ShapeSpec::custom("Empty", two_cell(), geometry, "");
    assert!(worksheet.add_shape(spec).is_err());
}

#[test]
fn custom_geometry_on_unknown_containers_is_skipped_inertly() {
    // custGeom inside a connection shape is not modeled; the connector still
    // parses with its other properties intact.
    let xml = drawing(
        "<xdr:twoCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff>\
         <xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>\
         <xdr:to><xdr:col>2</xdr:col><xdr:colOff>0</xdr:colOff>\
         <xdr:row>2</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>\
         <xdr:cxnSp><xdr:nvCxnSpPr><xdr:cNvPr id=\"3\" name=\"Line\"/><xdr:cNvCxnSpPr/>\
         </xdr:nvCxnSpPr><xdr:spPr><a:custGeom><a:pathLst><a:path>\
         <a:moveTo><a:pt x=\"0\" y=\"0\"/></a:moveTo></a:path></a:pathLst></a:custGeom>\
         </xdr:spPr></xdr:cxnSp><xdr:clientData/></xdr:twoCellAnchor>",
    );
    let objects = parse_drawing_shapes(&xml).unwrap().unwrap();
    assert_eq!(objects.len(), 1);
    let DrawingObject::ConnectionShape(connection) = &objects[0].object else {
        panic!("expected a connection shape");
    };
    assert_eq!(connection.non_visual.name.as_deref(), Some("Line"));
}
