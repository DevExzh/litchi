use litchi_drawingml::geometry::{
    AdjustValue, CustomGeometry, Formula, Guide, Path, PathCommand, Point,
};

#[test]
fn contextual_geometry_values_and_guides_are_typed() {
    let guide = Guide::new("  width\t", Formula::literal(21_600));
    assert_eq!(guide.name, "width");
    assert_eq!(guide.formula, Formula::literal(21_600));
    assert_eq!(
        AdjustValue::guide(" width "),
        AdjustValue::Guide("width".into())
    );
}

#[test]
fn geometry_paths_retain_typed_commands() {
    let geometry = CustomGeometry {
        adjust_values: Vec::new(),
        guides: vec![Guide::new("x", Formula::literal(10))],
        adjust_handles: Vec::new(),
        connection_sites: Vec::new(),
        text_rectangle: None,
        paths: vec![
            Path::new(100, 100)
                .with_command(PathCommand::MoveTo(Point::new(0, 0)))
                .with_command(PathCommand::LineTo(Point::new(
                    100,
                    AdjustValue::guide("x"),
                ))),
        ],
    };
    assert_eq!(geometry.paths.len(), 1);
    assert_eq!(geometry.paths[0].commands.len(), 2);
}
