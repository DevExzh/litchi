//! Public facade checks for the shared DrawingML transform owner.

use litchi_drawingml::transform::{Angle, Limits, Package, Point, Size, Transform};

#[test]
fn transform_facade_is_typed_and_package_based() {
    let value = Transform::new()
        .with_offset(Point::emu(-10, 20).unwrap())
        .with_extent(Size::emu(300, 400).unwrap())
        .with_rotation(Angle::new(90_000));
    let package =
        Package::new(value, &Limits::DEFAULT).expect("typed transform must create a package");
    assert_eq!(package.value().offset().unwrap().x().as_emu(), Some(-10));
    assert_eq!(package.value().extent().unwrap().height().as_emu(), 400);
    assert_eq!(package.value().rotation(), Angle::new(90_000));
}
