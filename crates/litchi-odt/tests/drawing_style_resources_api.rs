use litchi_odt::Document;
use litchi_odt::drawing::resources::{fill_image, gradient, hatch, marker, opacity, stroke_dash};
mod support;

const CONTENT: &str =
    include_str!("../../../test-data/odf/odt/drawing-style-resources-content.xml");
const STYLES: &str = include_str!("../../../test-data/odf/odt/drawing-style-resources-styles.xml");
const FLAT: &str = include_str!("../../../test-data/odf/odt/drawing-style-resources.fodt");
const MIMETYPE: &str = "application/vnd.oasis.opendocument.text";

fn document() -> Document {
    Document::from_bytes(support::package(
        MIMETYPE,
        &[
            ("content.xml", CONTENT.as_bytes()),
            ("styles.xml", STYLES.as_bytes()),
        ],
    ))
    .unwrap()
}

fn assert_resources(
    fill_images: &fill_image::Collection,
    gradients: &gradient::Collection,
    hatches: &hatch::Collection,
    markers: &marker::Collection,
    opacities: &opacity::Collection,
    dashes: &stroke_dash::Collection,
) {
    assert_eq!(fill_images.images.len(), 1);
    assert_eq!(
        fill_images
            .get("Fill")
            .unwrap()
            .source
            .link()
            .unwrap()
            .href(),
        "Pictures/fill.png"
    );
    assert_eq!(gradients.gradients.len(), 2);
    assert!(gradients.get("Legacy").is_some());
    assert!(matches!(
        gradients.get("Linear"),
        Some(gradient::Definition::Linear(_))
    ));
    assert_eq!(hatches.hatches.len(), 1);
    assert!(hatches.get("Hatch").is_some());
    assert_eq!(markers.markers.len(), 1);
    assert_eq!(
        markers.get("Arrow").unwrap().path_data.as_str(),
        "M 0 0 L 10 10"
    );
    assert_eq!(opacities.opacities.len(), 1);
    assert_eq!(opacities.get("Fade").unwrap().name.as_deref(), Some("Fade"));
    assert_eq!(dashes.dashes.len(), 1);
    assert_eq!(
        dashes.get("Dash").unwrap().effective_style(),
        stroke_dash::Style::Round
    );
}

#[test]
fn document_generic_package_and_mutable_document_expose_named_style_resources() {
    let source = document();
    let fill_images = source.drawing_fill_images().unwrap();
    let gradients = source.drawing_gradients().unwrap();
    let hatches = source.drawing_hatches().unwrap();
    let markers = source.drawing_markers().unwrap();
    let opacities = source.drawing_opacities().unwrap();
    let dashes = source.drawing_stroke_dashes().unwrap();
    assert_resources(
        &fill_images,
        &gradients,
        &hatches,
        &markers,
        &opacities,
        &dashes,
    );

    let package = litchi_odt::generic::Package::from_bytes(source.to_bytes().unwrap()).unwrap();
    assert_eq!(package.drawing_fill_images().unwrap(), fill_images);
    assert_eq!(package.drawing_gradients().unwrap(), gradients);
    assert_eq!(package.drawing_hatches().unwrap(), hatches);
    assert_eq!(package.drawing_markers().unwrap(), markers);
    assert_eq!(package.drawing_opacities().unwrap(), opacities);
    assert_eq!(package.drawing_stroke_dashes().unwrap(), dashes);

    let mutable = litchi_odt::mutable::MutableDocument::from_document(source).unwrap();
    assert_eq!(mutable.drawing_fill_images().unwrap(), fill_images);
    assert_eq!(mutable.drawing_gradients().unwrap(), gradients);
    assert_eq!(mutable.drawing_hatches().unwrap(), hatches);
    assert_eq!(mutable.drawing_markers().unwrap(), markers);
    assert_eq!(mutable.drawing_opacities().unwrap(), opacities);
    assert_eq!(mutable.drawing_stroke_dashes().unwrap(), dashes);
}

#[test]
fn flat_document_exposes_named_style_resources() {
    let document = litchi_odt::generic::FlatDocument::from_bytes(FLAT.as_bytes().to_vec()).unwrap();
    let fill_images = document.drawing_fill_images().unwrap();
    let gradients = document.drawing_gradients().unwrap();
    let hatches = document.drawing_hatches().unwrap();
    let markers = document.drawing_markers().unwrap();
    let opacities = document.drawing_opacities().unwrap();
    let dashes = document.drawing_stroke_dashes().unwrap();

    assert_resources(
        &fill_images,
        &gradients,
        &hatches,
        &markers,
        &opacities,
        &dashes,
    );
}
