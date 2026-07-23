use litchi_odf::{
    Document, FlatOpenDocument, MutableDocument, OpenDocumentPackage,
    drawing_gradient::OdfDrawingGradient, drawing_stroke_dash::OdfStrokeDashStyle,
};
use std::io::{Cursor, Write};

const CONTENT: &str =
    include_str!("../../../test-data/odf/odt/drawing-style-resources-content.xml");
const STYLES: &str = include_str!("../../../test-data/odf/odt/drawing-style-resources-styles.xml");
const FLAT: &str = include_str!("../../../test-data/odf/odt/drawing-style-resources.fodt");
const MIMETYPE: &str = "application/vnd.oasis.opendocument.text";

fn document() -> Document {
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
    Document::from_bytes(output.into_inner()).unwrap()
}

fn assert_resources(
    gradients: &litchi_odf::drawing_gradient::OdfDrawingGradients,
    hatches: &litchi_odf::drawing_hatch::OdfDrawingHatches,
    dashes: &litchi_odf::drawing_stroke_dash::OdfDrawingStrokeDashes,
) {
    assert_eq!(gradients.gradients.len(), 2);
    assert!(gradients.get("Legacy").is_some());
    assert!(matches!(
        gradients.get("Linear"),
        Some(OdfDrawingGradient::Linear(_))
    ));
    assert_eq!(hatches.hatches.len(), 1);
    assert!(hatches.get("Hatch").is_some());
    assert_eq!(dashes.dashes.len(), 1);
    assert_eq!(
        dashes.get("Dash").unwrap().effective_style(),
        OdfStrokeDashStyle::Round
    );
}

#[test]
fn document_generic_package_and_mutable_document_expose_named_style_resources() {
    let source = document();
    let gradients = source.drawing_gradients().unwrap();
    let hatches = source.drawing_hatches().unwrap();
    let dashes = source.drawing_stroke_dashes().unwrap();
    assert_resources(&gradients, &hatches, &dashes);

    let package = OpenDocumentPackage::from_bytes(source.to_bytes().unwrap()).unwrap();
    assert_eq!(package.drawing_gradients().unwrap(), gradients);
    assert_eq!(package.drawing_hatches().unwrap(), hatches);
    assert_eq!(package.drawing_stroke_dashes().unwrap(), dashes);

    let mutable = MutableDocument::from_document(source).unwrap();
    assert_eq!(mutable.drawing_gradients().unwrap(), gradients);
    assert_eq!(mutable.drawing_hatches().unwrap(), hatches);
    assert_eq!(mutable.drawing_stroke_dashes().unwrap(), dashes);
}

#[test]
fn flat_document_exposes_named_style_resources() {
    let document = FlatOpenDocument::from_bytes(FLAT.as_bytes().to_vec()).unwrap();
    let gradients = document.drawing_gradients().unwrap();
    let hatches = document.drawing_hatches().unwrap();
    let dashes = document.drawing_stroke_dashes().unwrap();

    assert_resources(&gradients, &hatches, &dashes);
}
