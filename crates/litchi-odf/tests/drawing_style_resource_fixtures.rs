use litchi_odf::{
    drawing_gradient::{OdfDrawingGradient, parse_drawing_gradients},
    drawing_hatch::parse_drawing_hatches,
    drawing_stroke_dash::parse_drawing_stroke_dashes,
};

const MULTICOLOR_GRADIENT: &str =
    include_str!("../../../test-data/odf/drawing/multicolor-gradient.fodp");
const HATCH_ANGLES: &str = include_str!("../../../test-data/odf/drawing/hatch-angles.fodg");
const DASHED_LINE: &str = include_str!("../../../test-data/odf/drawing/dashed-line.fodg");

#[test]
fn local_gradient_fixture_preserves_multicolor_roundtrip_coverage() {
    let gradients = parse_drawing_gradients(MULTICOLOR_GRADIENT).unwrap();
    assert!(gradients.gradients.len() >= 6);
    let OdfDrawingGradient::Legacy(first) = &gradients.gradients[0] else {
        panic!("local fixture should begin with a legacy gradient");
    };
    assert_eq!(first.extension_stops.len(), 2);
    assert!(
        !parse_drawing_gradients(&gradients.to_xml().unwrap())
            .unwrap()
            .gradients
            .is_empty()
    );
}

#[test]
fn local_hatch_fixture_preserves_angle_units() {
    let hatches = parse_drawing_hatches(HATCH_ANGLES).unwrap();
    assert_eq!(hatches.hatches.len(), 4);
    assert_eq!(
        hatches.hatches[0].rotation.as_ref().unwrap().as_str(),
        "58.5deg"
    );
    assert_eq!(
        hatches.hatches[1].rotation.as_ref().unwrap().as_str(),
        "65grad"
    );
    assert_eq!(
        hatches.hatches[2].rotation.as_ref().unwrap().as_str(),
        "1.02101761241558rad"
    );
    assert_eq!(
        hatches.hatches[3].rotation.as_ref().unwrap().as_str(),
        "585"
    );
    assert_eq!(
        parse_drawing_hatches(&hatches.to_xml().unwrap())
            .unwrap()
            .hatches
            .len(),
        4
    );
}

#[test]
fn local_dash_fixture_preserves_segments() {
    let parsed = parse_drawing_stroke_dashes(DASHED_LINE).unwrap();
    let dash = parsed.get("DoubleDashDotDot").unwrap();
    assert_eq!(dash.dots1, Some(1));
    assert_eq!(dash.dots2, Some(2));
    assert_eq!(dash.dots1_length.unwrap().value(), 800.0);
    assert_eq!(dash.distance.unwrap().value(), 300.0);
}
