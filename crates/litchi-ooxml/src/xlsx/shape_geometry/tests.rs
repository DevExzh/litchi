//! Unit tests for the typed custom geometry model, formulas, validation,
//! and serialization.

use super::*;

fn value(value: i64) -> XlsxAdjustValue {
    XlsxAdjustValue::Value(value)
}

fn triangle_path() -> XlsxGeometryPath {
    XlsxGeometryPath::new(21_600, 21_600)
        .with_command(XlsxPathCommand::MoveTo(XlsxGeometryPoint::new(10_800, 0)))
        .with_command(XlsxPathCommand::LineTo(XlsxGeometryPoint::new(
            21_600, 21_600,
        )))
        .with_command(XlsxPathCommand::LineTo(XlsxGeometryPoint::new(0, 21_600)))
        .with_command(XlsxPathCommand::Close)
}

fn triangle_geometry() -> XlsxCustomGeometry {
    XlsxCustomGeometry::new().with_path(triangle_path())
}

#[test]
fn formulas_round_trip_through_their_token_form() {
    let cases = [
        ("*/ w 1 2", "*/"),
        ("+- w h ss", "+-"),
        ("+/ 10 hd2 2", "+/"),
        ("?: adj1 10 20", "?:"),
        ("abs adj1", "abs"),
        ("at2 w h", "at2"),
        ("cat2 hd2 w h", "cat2"),
        ("cos hd2 adj1", "cos"),
        ("max w h", "max"),
        ("min w h", "min"),
        ("mod w h 0", "mod"),
        ("pin 0 adj1 21600", "pin"),
        ("sat2 hd2 w h", "sat2"),
        ("sin hd2 adj1", "sin"),
        ("sqrt w", "sqrt"),
        ("tan hd2 adj1", "tan"),
        ("val 10800", "val"),
    ];
    for (text, operation) in cases {
        let formula: XlsxGeometryFormula = text.parse().unwrap();
        assert_eq!(formula.operation(), operation, "operation of '{text}'");
        assert_eq!(formula.to_string(), text, "serialization of '{text}'");
        let reparsed: XlsxGeometryFormula = formula.to_string().parse().unwrap();
        assert_eq!(reparsed, formula, "round trip of '{text}'");
    }
}

#[test]
fn formula_operands_distinguish_literals_from_guide_references() {
    let formula: XlsxGeometryFormula = "pin 0 adj1 21600".parse().unwrap();
    let operands: Vec<_> = formula.operands().cloned().collect();
    assert_eq!(
        operands,
        vec![value(0), XlsxAdjustValue::guide("adj1"), value(21_600)]
    );
}

#[test]
fn formula_parsing_rejects_malformed_tokens() {
    assert!("".parse::<XlsxGeometryFormula>().is_err());
    assert!("frobnicate 1 2".parse::<XlsxGeometryFormula>().is_err());
    assert!("val".parse::<XlsxGeometryFormula>().is_err());
    assert!("val 1 2".parse::<XlsxGeometryFormula>().is_err());
    assert!("*/ w 1".parse::<XlsxGeometryFormula>().is_err());
    assert!("pin 0 adj1 21600 9".parse::<XlsxGeometryFormula>().is_err());
}

#[test]
fn adjust_values_parse_numbers_and_guide_names() {
    assert_eq!("42".parse::<XlsxAdjustValue>().unwrap(), value(42));
    assert_eq!("-7".parse::<XlsxAdjustValue>().unwrap(), value(-7));
    assert_eq!(
        "adj1".parse::<XlsxAdjustValue>().unwrap(),
        XlsxAdjustValue::guide("adj1")
    );
    assert!("".parse::<XlsxAdjustValue>().is_err());
    assert!("   ".parse::<XlsxAdjustValue>().is_err());
}

#[test]
fn path_fill_mode_tokens_round_trip() {
    for mode in [
        XlsxPathFillMode::None,
        XlsxPathFillMode::Normal,
        XlsxPathFillMode::Lighten,
        XlsxPathFillMode::LightenLess,
        XlsxPathFillMode::Darken,
        XlsxPathFillMode::DarkenLess,
    ] {
        assert_eq!(XlsxPathFillMode::from_token(mode.as_str()), Some(mode));
    }
    assert_eq!(XlsxPathFillMode::from_token("sparkly"), None);
}

#[test]
fn validation_accepts_a_complete_geometry() {
    let geometry = XlsxCustomGeometry {
        adjust_values: vec![XlsxGeometryGuide::new(
            "adj1",
            XlsxGeometryFormula::literal(50),
        )],
        guides: vec![XlsxGeometryGuide::new(
            "x1",
            "*/ adj1 21600 100".parse().unwrap(),
        )],
        adjust_handles: vec![
            XlsxAdjustHandle::Xy(XlsxXyAdjustHandle {
                horizontal_guide: Some("adj1".to_string()),
                minimum_x: Some(value(0)),
                maximum_x: Some(value(21_600)),
                position: XlsxGeometryPoint::new(XlsxAdjustValue::guide("x1"), value(0)),
                ..XlsxXyAdjustHandle::default()
            }),
            XlsxAdjustHandle::Polar(XlsxPolarAdjustHandle {
                radius_guide: Some("adj1".to_string()),
                minimum_angle: Some(value(0)),
                maximum_angle: Some(value(21_600_000)),
                position: XlsxGeometryPoint::new(10_800, 10_800),
                ..XlsxPolarAdjustHandle::default()
            }),
        ],
        connection_sites: vec![XlsxConnectionSite {
            angle: value(5_400_000),
            position: XlsxGeometryPoint::new(10_800, 21_600),
        }],
        text_rectangle: Some(XlsxGeometryRectangle {
            left: value(5_400),
            top: value(5_400),
            right: value(16_200),
            bottom: value(21_600),
        }),
        paths: vec![triangle_path()],
    };
    assert!(validate_custom_geometry(&geometry).is_ok());
}

#[test]
fn validation_rejects_an_empty_path_list() {
    let geometry = XlsxCustomGeometry::new();
    let error = validate_custom_geometry(&geometry).unwrap_err();
    assert!(
        error.contains("at least one path"),
        "unexpected error: {error}"
    );
}

#[test]
fn validation_rejects_invalid_guide_names() {
    for name in ["", "has space", "123", "-45"] {
        let geometry = triangle_geometry().with_adjust_value(XlsxGeometryGuide::new(
            name,
            XlsxGeometryFormula::literal(1),
        ));
        assert!(
            validate_custom_geometry(&geometry).is_err(),
            "guide name '{name}' should be rejected"
        );
    }
}

#[test]
fn validation_rejects_out_of_range_values() {
    let mut oversized_path = triangle_geometry();
    oversized_path.paths[0].width = MAX_POSITIVE_COORDINATE + 1;
    assert!(validate_custom_geometry(&oversized_path).is_err());

    let mut negative_path = triangle_geometry();
    negative_path.paths[0].height = -1;
    assert!(validate_custom_geometry(&negative_path).is_err());

    let out_of_range_point =
        triangle_geometry().with_path(XlsxGeometryPath::new(0, 0).with_command(
            XlsxPathCommand::MoveTo(XlsxGeometryPoint::new(MAX_COORDINATE + 1, 0)),
        ));
    assert!(validate_custom_geometry(&out_of_range_point).is_err());

    let out_of_range_angle = triangle_geometry().with_path(
        XlsxGeometryPath::new(0, 0).with_command(XlsxPathCommand::ArcTo {
            width_radius: value(100),
            height_radius: value(100),
            start_angle: value(MAX_ANGLE + 1),
            swing_angle: value(5_400_000),
        }),
    );
    assert!(validate_custom_geometry(&out_of_range_angle).is_err());
}

#[test]
fn serialization_omits_defaults_and_empty_lists() {
    let mut xml = String::new();
    write::write_custom_geometry(&mut xml, &triangle_geometry());
    assert!(xml.starts_with("<a:custGeom><a:pathLst><a:path w=\"21600\" h=\"21600\">"));
    assert!(!xml.contains("avLst"));
    assert!(!xml.contains("gdLst"));
    assert!(!xml.contains("ahLst"));
    assert!(!xml.contains("cxnLst"));
    assert!(!xml.contains("a:rect"));
    assert!(!xml.contains("fill="));
    assert!(!xml.contains("stroke="));
    assert!(!xml.contains("extrusionOk="));

    let mut plain = String::new();
    write::write_custom_geometry(
        &mut plain,
        &XlsxCustomGeometry::new().with_path(XlsxGeometryPath::default()),
    );
    assert_eq!(
        plain,
        "<a:custGeom><a:pathLst><a:path></a:path></a:pathLst></a:custGeom>"
    );
}

#[test]
fn serialization_writes_non_default_path_attributes() {
    let path = XlsxGeometryPath {
        width: 100,
        height: 200,
        fill_mode: XlsxPathFillMode::Lighten,
        stroked: false,
        extrusion_allowed: false,
        commands: vec![XlsxPathCommand::Close],
    };
    let mut xml = String::new();
    write::write_custom_geometry(&mut xml, &XlsxCustomGeometry::new().with_path(path));
    assert!(xml.contains(
        r#"<a:path w="100" h="200" fill="lighten" stroke="0" extrusionOk="0"><a:close/></a:path>"#
    ));
}
