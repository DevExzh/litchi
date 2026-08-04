use litchi_odt::{
    FlatOpenDocument,
    drawing_fill_image::FillImageLinkKind,
    drawing_gradient::{DrawingGradient, parse_drawing_gradients},
    drawing_hatch::parse_drawing_hatches,
    drawing_marker::MarkerViewBox,
    drawing_opacity::OpacityStyle,
    drawing_stroke_dash::parse_drawing_stroke_dashes,
};

const MULTICOLOR_GRADIENT: &str =
    include_str!("../../../test-data/odf/drawing/multicolor-gradient.fodp");
const HATCH_ANGLES: &str = include_str!("../../../test-data/odf/drawing/hatch-angles.fodg");
const DASHED_LINE: &str = include_str!("../../../test-data/odf/drawing/dashed-line.fodg");
const FILL_IMAGE_LINKED: &str =
    include_str!("../../../test-data/odf/drawing/fill-image-linked.fodp");
const FILL_IMAGE_INLINE: &str =
    include_str!("../../../test-data/odf/drawing/fill-image-inline.fodg");
const MARKER_FLAT: &str = include_str!("../../../test-data/odf/drawing/marker-flat.fods");
const OPACITY_ANGLES: &str = include_str!("../../../test-data/odf/drawing/opacity-angles.fodg");
const OPACITY_EXTENSION_STOPS: &str =
    include_str!("../../../test-data/odf/drawing/opacity-extension-stops.fodt");

#[test]
fn local_gradient_fixture_preserves_multicolor_roundtrip_coverage() {
    let gradients = parse_drawing_gradients(MULTICOLOR_GRADIENT).unwrap();
    assert!(gradients.gradients.len() >= 6);
    let DrawingGradient::Legacy(first) = &gradients.gradients[0] else {
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

#[test]
fn local_fill_image_fixtures_preserve_link_and_inline_bytes() {
    let linked = FlatOpenDocument::from_bytes(FILL_IMAGE_LINKED.as_bytes().to_vec()).unwrap();
    let images = linked.drawing_fill_images().unwrap();
    assert_eq!(
        images
            .get("remote_bg")
            .unwrap()
            .source
            .link()
            .unwrap()
            .kind(),
        FillImageLinkKind::InertExternal
    );

    let inline = FlatOpenDocument::from_bytes(FILL_IMAGE_INLINE.as_bytes().to_vec()).unwrap();
    let images = inline.drawing_fill_images().unwrap();
    let bytes = images
        .get("libreoffice_5f_0")
        .unwrap()
        .source
        .inline_bytes()
        .unwrap();
    assert_eq!(&bytes[..8], b"\x89PNG\r\n\x1a\n");
}

#[test]
fn local_marker_fixture_preserves_view_box_and_path() {
    let document = FlatOpenDocument::from_bytes(MARKER_FLAT.as_bytes().to_vec()).unwrap();
    let markers = document.drawing_markers().unwrap();
    let marker = markers.get("Arrowheads_20_1").unwrap();
    assert_eq!(marker.view_box, MarkerViewBox::new(0, 0, 20, 30));
    assert_eq!(marker.path_data.as_str(), "M10 0l-10 30h20z");
}

#[test]
fn local_opacity_fixtures_preserve_angles_and_extension_stops() {
    let angles = FlatOpenDocument::from_bytes(OPACITY_ANGLES.as_bytes().to_vec()).unwrap();
    let values = angles.drawing_opacities().unwrap();
    assert_eq!(values.opacities.len(), 6);
    assert_eq!(
        values.opacities[0].angle.as_ref().unwrap().as_str(),
        "90deg"
    );
    assert_eq!(
        values.opacities[2].angle.as_ref().unwrap().as_str(),
        "1.0rad"
    );
    assert_eq!(
        values.opacities[3].angle.as_ref().unwrap().as_str(),
        "1000grad"
    );

    let stops = FlatOpenDocument::from_bytes(OPACITY_EXTENSION_STOPS.as_bytes().to_vec()).unwrap();
    let values = stops.drawing_opacities().unwrap();
    let value = values.get("Transparency_20_1").unwrap();
    assert_eq!(value.style, OpacityStyle::Ellipsoid);
    assert_eq!(value.extension_stops.len(), 2);
}
