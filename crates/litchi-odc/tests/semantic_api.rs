use litchi_odc::{Builder, Chart, axis::Dimension, legend::Position, series::Series};

#[test]
fn focused_modules_are_the_canonical_semantic_api() {
    assert_eq!(Dimension::Z, Dimension::Z);
    assert_eq!(Position::BottomEnd, Position::BottomEnd);
    assert_eq!(Series::new("Sheet1.A1:A3").range(), "Sheet1.A1:A3");

    let bytes = Builder::new().build().unwrap();
    let chart = Chart::from_bytes(bytes).unwrap();
    assert!(chart.content_xml().contains("<office:chart"));
}
