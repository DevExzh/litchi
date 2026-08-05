//! Unit tests for the typed custom geometry model, formulas, validation,
//! and serialization.

use super::*;

fn value(value: i64) -> AdjustValue {
    AdjustValue::Value(value)
}

fn triangle_path() -> Path {
    Path::new(21_600, 21_600)
        .with_command(PathCommand::MoveTo(Point::new(10_800, 0)))
        .with_command(PathCommand::LineTo(Point::new(21_600, 21_600)))
        .with_command(PathCommand::LineTo(Point::new(0, 21_600)))
        .with_command(PathCommand::Close)
}

fn triangle_geometry() -> CustomGeometry {
    CustomGeometry::new().with_path(triangle_path())
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
        let formula: Formula = text.parse().unwrap();
        assert_eq!(formula.operation(), operation, "operation of '{text}'");
        assert_eq!(formula.to_string(), text, "serialization of '{text}'");
        let reparsed: Formula = formula.to_string().parse().unwrap();
        assert_eq!(reparsed, formula, "round trip of '{text}'");
    }
}

#[test]
fn formula_operands_distinguish_literals_from_guide_references() {
    let formula: Formula = "pin 0 adj1 21600".parse().unwrap();
    let operands: Vec<_> = formula.operands().cloned().collect();
    assert_eq!(
        operands,
        vec![value(0), AdjustValue::guide("adj1"), value(21_600)]
    );
}

#[test]
fn formula_parsing_rejects_malformed_tokens() {
    assert!("".parse::<Formula>().is_err());
    assert!("frobnicate 1 2".parse::<Formula>().is_err());
    assert!("val".parse::<Formula>().is_err());
    assert!("val 1 2".parse::<Formula>().is_err());
    assert!("*/ w 1".parse::<Formula>().is_err());
    assert!("pin 0 adj1 21600 9".parse::<Formula>().is_err());
}

#[test]
fn adjust_values_parse_numbers_and_guide_names() {
    assert_eq!("42".parse::<AdjustValue>().unwrap(), value(42));
    assert_eq!("-7".parse::<AdjustValue>().unwrap(), value(-7));
    assert_eq!(
        "adj1".parse::<AdjustValue>().unwrap(),
        AdjustValue::guide("adj1")
    );
    assert!("".parse::<AdjustValue>().is_err());
    assert!("   ".parse::<AdjustValue>().is_err());
}

#[test]
fn path_fill_mode_tokens_round_trip() {
    for mode in [
        PathFillMode::None,
        PathFillMode::Normal,
        PathFillMode::Lighten,
        PathFillMode::LightenLess,
        PathFillMode::Darken,
        PathFillMode::DarkenLess,
    ] {
        assert_eq!(mode.as_str().parse::<PathFillMode>().unwrap(), mode);
    }
    assert!("sparkly".parse::<PathFillMode>().is_err());
}

#[test]
fn validation_accepts_a_complete_geometry() {
    let geometry = CustomGeometry {
        adjust_values: vec![Guide::new("adj1", Formula::literal(50))],
        guides: vec![Guide::new("x1", "*/ adj1 21600 100".parse().unwrap())],
        adjust_handles: vec![
            AdjustHandle::Xy(XyAdjustHandle {
                horizontal_guide: Some("adj1".to_string()),
                minimum_x: Some(value(0)),
                maximum_x: Some(value(21_600)),
                position: Point::new(AdjustValue::guide("x1"), value(0)),
                ..XyAdjustHandle::default()
            }),
            AdjustHandle::Polar(PolarAdjustHandle {
                radius_guide: Some("adj1".to_string()),
                minimum_angle: Some(value(0)),
                maximum_angle: Some(value(21_600_000)),
                position: Point::new(10_800, 10_800),
                ..PolarAdjustHandle::default()
            }),
        ],
        connection_sites: vec![ConnectionSite {
            angle: value(5_400_000),
            position: Point::new(10_800, 21_600),
        }],
        text_rectangle: Some(Rectangle {
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
    let geometry = CustomGeometry::new();
    let error = validate_custom_geometry(&geometry).unwrap_err();
    assert!(
        error.contains("at least one path"),
        "unexpected error: {error}"
    );
}

#[test]
fn validation_rejects_invalid_guide_names() {
    for name in ["", "has space", "123", "-45"] {
        let geometry = triangle_geometry().with_adjust_value(Guide::new(name, Formula::literal(1)));
        assert!(
            validate_custom_geometry(&geometry).is_err(),
            "guide name '{name}' should be rejected"
        );
    }
}

#[test]
fn parsed_validation_preserves_schema_valid_open_guide_names() {
    for name in ["", "has space", "123", "-45"] {
        let geometry =
            CustomGeometry::new().with_adjust_value(Guide::new(name, Formula::literal(1)));
        assert!(
            validate_parsed_custom_geometry(&geometry).is_ok(),
            "schema-valid parsed guide name '{name}' should be preserved"
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

    let out_of_range_point = triangle_geometry().with_path(
        Path::new(0, 0).with_command(PathCommand::MoveTo(Point::new(MAX_COORDINATE + 1, 0))),
    );
    assert!(validate_custom_geometry(&out_of_range_point).is_err());

    let out_of_range_angle =
        triangle_geometry().with_path(Path::new(0, 0).with_command(PathCommand::ArcTo {
            width_radius: value(100),
            height_radius: value(100),
            start_angle: value(MAX_ANGLE + 1),
            swing_angle: value(5_400_000),
        }));
    assert!(validate_custom_geometry(&out_of_range_angle).is_err());
}

#[test]
fn serialization_omits_defaults_and_empty_lists() {
    let mut xml = String::new();
    writer::write_custom_geometry(&mut xml, &triangle_geometry());
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
    writer::write_custom_geometry(
        &mut plain,
        &CustomGeometry::new().with_path(Path::default()),
    );
    assert_eq!(
        plain,
        "<a:custGeom><a:pathLst><a:path></a:path></a:pathLst></a:custGeom>"
    );
}

#[test]
fn serialization_writes_non_default_path_attributes() {
    let path = Path {
        width: 100,
        height: 200,
        fill_mode: PathFillMode::Lighten,
        stroked: false,
        extrusion_allowed: false,
        commands: vec![PathCommand::Close],
    };
    let mut xml = String::new();
    writer::write_custom_geometry(&mut xml, &CustomGeometry::new().with_path(path));
    assert!(xml.contains(
        r#"<a:path w="100" h="200" fill="lighten" stroke="0" extrusionOk="0"><a:close/></a:path>"#
    ));
}
