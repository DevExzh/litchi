use litchi_odf::{
    MutablePresentation, Presentation, drawing_gradient::OdfDrawingGradient,
    drawing_stroke_dash::OdfStrokeDashStyle,
};
use std::io::{Cursor, Write};

const CONTENT: &str =
    include_str!("../../../test-data/odf/odp/drawing-style-resources-content.xml");
const STYLES: &str = include_str!("../../../test-data/odf/odp/drawing-style-resources-styles.xml");
const MIMETYPE: &str = "application/vnd.oasis.opendocument.presentation";

fn presentation() -> Presentation {
    let mut output = Cursor::new(Vec::new());
    let mut zip = zip::ZipWriter::new(&mut output);
    let stored =
        zip::write::SimpleFileOptions::default().compression_method(zip::CompressionMethod::Stored);
    let deflated = zip::write::SimpleFileOptions::default()
        .compression_method(zip::CompressionMethod::Deflated);
    zip.start_file("mimetype", stored).unwrap();
    zip.write_all(MIMETYPE.as_bytes()).unwrap();
    zip.start_file("content.xml", deflated).unwrap();
    zip.write_all(CONTENT.as_bytes()).unwrap();
    zip.start_file("styles.xml", deflated).unwrap();
    zip.write_all(STYLES.as_bytes()).unwrap();
    zip.start_file("META-INF/manifest.xml", deflated).unwrap();
    write!(
        zip,
        r#"<m:manifest xmlns:m="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" m:version="1.3"><m:file-entry m:full-path="/" m:media-type="{MIMETYPE}"/><m:file-entry m:full-path="content.xml" m:media-type="text/xml"/><m:file-entry m:full-path="styles.xml" m:media-type="text/xml"/></m:manifest>"#
    )
    .unwrap();
    zip.finish().unwrap();
    Presentation::from_bytes(output.into_inner()).unwrap()
}

fn assert_resources(
    fill_images: &litchi_odf::drawing_fill_image::OdfDrawingFillImages,
    gradients: &litchi_odf::drawing_gradient::OdfDrawingGradients,
    hatches: &litchi_odf::drawing_hatch::OdfDrawingHatches,
    markers: &litchi_odf::drawing_marker::OdfDrawingMarkers,
    opacities: &litchi_odf::drawing_opacity::OdfDrawingOpacities,
    dashes: &litchi_odf::drawing_stroke_dash::OdfDrawingStrokeDashes,
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
    assert!(matches!(
        gradients.get("Linear"),
        Some(OdfDrawingGradient::Linear(_))
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
    assert_eq!(
        dashes.get("Dash").unwrap().effective_style(),
        OdfStrokeDashStyle::Round
    );
}

#[test]
fn presentation_and_mutable_presentation_expose_named_style_resources() {
    let source = presentation();
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

    let mutable = MutablePresentation::from_presentation(source).unwrap();
    assert_eq!(mutable.drawing_fill_images().unwrap(), fill_images);
    assert_eq!(mutable.drawing_gradients().unwrap(), gradients);
    assert_eq!(mutable.drawing_hatches().unwrap(), hatches);
    assert_eq!(mutable.drawing_markers().unwrap(), markers);
    assert_eq!(mutable.drawing_opacities().unwrap(), opacities);
    assert_eq!(mutable.drawing_stroke_dashes().unwrap(), dashes);
}
