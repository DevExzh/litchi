//! Public facade checks for the shared DrawingML transform owner.

use litchi_drawingml::transform::{Angle, Point, Size, Snapshot, Transform};

#[test]
fn transform_facade_is_typed_and_snapshot_based() {
    let value = Transform::new()
        .with_offset(Point::emu(-10, 20).unwrap())
        .with_extent(Size::emu(300, 400).unwrap())
        .with_rotation(Angle::new(90_000));
    let snapshot = Snapshot::new(value).expect("typed transform must create a snapshot");
    assert_eq!(snapshot.value().offset().unwrap().x().as_emu(), Some(-10));
    assert_eq!(snapshot.value().extent().unwrap().height().as_emu(), 400);
    assert_eq!(snapshot.value().rotation(), Angle::new(90_000));
}
