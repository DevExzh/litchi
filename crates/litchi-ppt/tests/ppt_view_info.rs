use litchi_ppt::{Package, PowerPointGuideOrientation};
use std::path::Path;

#[test]
fn apache_poi_slide_view_zoom_and_guides_are_exposed() {
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/poi/test-data/slideshow/54880_chinese.ppt");
    let mut package = Package::open(path).unwrap();
    let presentation = package.presentation().unwrap();
    let information = presentation.slide_view_information().unwrap();
    let view = information.slide().unwrap();

    assert!(view.preferences().snap_to_grid());
    assert!(!view.preferences().snap_to_shape());
    let zoom = view.zoom().unwrap();
    assert_eq!(
        (zoom.x_scale().numerator(), zoom.x_scale().denominator()),
        (86, 100)
    );
    assert_eq!(
        (zoom.y_scale().numerator(), zoom.y_scale().denominator()),
        (86, 100)
    );
    assert_eq!((zoom.origin().x(), zoom.origin().y()), (-1542, -96));
    assert!(zoom.uses_variable_scale());
    assert!(!zoom.is_draft_mode());
    assert_eq!(view.guides().len(), 2);
    assert_eq!(
        view.guides()[0].orientation(),
        PowerPointGuideOrientation::Horizontal
    );
    assert_eq!(view.guides()[0].position(), 2160);
    assert_eq!(
        view.guides()[1].orientation(),
        PowerPointGuideOrientation::Vertical
    );
    assert_eq!(view.guides()[1].position(), 2880);
}
