#![allow(
    clippy::unwrap_used,
    reason = "integration-test assertions panic on failure by design"
)]

use litchi_odp::{
    Presentation, constants,
    drawing::resources::{fill_image, gradient, hatch, marker, opacity, stroke_dash},
};
use soapberry_zip::office::StreamingArchiveWriter;

const CONTENT: &str =
    include_str!("../../../test-data/odf/odp/drawing-style-resources-content.xml");
const STYLES: &str = include_str!("../../../test-data/odf/odp/drawing-style-resources-styles.xml");

fn presentation() -> Presentation {
    const MANIFEST: &[u8] = br#"<?xml version="1.0"?><manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.presentation"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/><manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"/></manifest:manifest>"#;
    let mut archive = StreamingArchiveWriter::new();
    archive
        .write_stored("mimetype", constants::ODF_PRESENTATION.as_bytes())
        .unwrap();
    archive
        .write_deflated("content.xml", CONTENT.as_bytes())
        .unwrap();
    archive
        .write_deflated("styles.xml", STYLES.as_bytes())
        .unwrap();
    archive
        .write_deflated("META-INF/manifest.xml", MANIFEST)
        .unwrap();
    Presentation::from_bytes(archive.finish_to_bytes().unwrap()).unwrap()
}

#[test]
fn presentation_exposes_named_drawing_style_resources() {
    let presentation = presentation();
    let fill_images = presentation.drawing_fill_images().unwrap();
    let gradients = presentation.drawing_gradients().unwrap();
    let hatches = presentation.drawing_hatches().unwrap();
    let markers = presentation.drawing_markers().unwrap();
    let opacities = presentation.drawing_opacities().unwrap();
    let dashes = presentation.drawing_stroke_dashes().unwrap();

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

    let _: &fill_image::Collection = &fill_images;
    let _: &hatch::Collection = &hatches;
    let _: &marker::Collection = &markers;
    let _: &opacity::Collection = &opacities;
}
