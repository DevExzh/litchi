#![allow(
    clippy::unwrap_used,
    reason = "tests are expected to panic on unexpected fixture failures"
)]

use litchi_odg::Drawing;
use soapberry_zip::office::StreamingArchiveWriter;

const CONTENT: &str =
    include_str!("../../../test-data/odf/odg/drawing-style-resources-content.xml");
const STYLES: &str = include_str!("../../../test-data/odf/odg/drawing-style-resources-styles.xml");
const LIBREOFFICE_ODG: &[u8] = include_bytes!(
    "../../../3rdparty/libreoffice-core/xmlsecurity/doc/OpenDocumentSignatures-Workflow.odg"
);
const REAL_DRAW_CORPUS: &[(&str, &[u8])] = &[
    (
        "blank",
        include_bytes!("../../../3rdparty/libreoffice-core/desktop/qa/data/BlankDrawDocument.odg"),
    ),
    (
        "three-page",
        include_bytes!("../../../3rdparty/libreoffice-core/desktop/qa/data/3page.odg"),
    ),
    (
        "transparent-fill",
        include_bytes!(
            "../../../3rdparty/libreoffice-core/filter/qa/unit/data/semi-transparent-fill.odg"
        ),
    ),
    (
        "fit-frame-text",
        include_bytes!("../../../3rdparty/libreoffice-core/sd/qa/unit/data/odg/FitToFrameText.odg"),
    ),
    (
        "fontwork",
        include_bytes!("../../../3rdparty/libreoffice-core/svx/qa/unit/data/FontWork.odg"),
    ),
    (
        "complex-groups",
        include_bytes!("../../../3rdparty/libreoffice-core/sd/qa/unit/data/odg/rhbz1870501.odg"),
    ),
];

#[test]
fn real_drawing_resource_xml_remains_exact_and_opaque() {
    const MIMETYPE: &[u8] = b"application/vnd.oasis.opendocument.graphics";
    const MANIFEST: &[u8] = br#"<?xml version="1.0"?><manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.graphics"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/><manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"/></manifest:manifest>"#;
    let mut writer = StreamingArchiveWriter::new();
    writer.write_stored("mimetype", MIMETYPE).unwrap();
    writer
        .write_deflated("content.xml", CONTENT.as_bytes())
        .unwrap();
    writer
        .write_deflated("styles.xml", STYLES.as_bytes())
        .unwrap();
    writer
        .write_deflated("META-INF/manifest.xml", MANIFEST)
        .unwrap();
    let bytes = writer.finish_to_bytes().unwrap();

    let drawing = Drawing::from_bytes(bytes.clone()).unwrap();
    assert_eq!(drawing.as_bytes(), bytes.as_slice());
    assert_eq!(drawing.content_xml(), CONTENT);
    assert_eq!(drawing.styles_xml(), Some(STYLES));
    assert!(CONTENT.contains('\n') || STYLES.contains('\n'));

    let raw = format!("{}{}", drawing.content_xml(), drawing.styles_xml().unwrap());
    for element in [
        "draw:fill-image",
        "draw:gradient",
        "draw:hatch",
        "draw:marker",
        "draw:opacity",
        "draw:stroke-dash",
    ] {
        assert!(raw.contains(element), "fixture lost {element}");
    }
}

#[test]
fn real_libreoffice_odg_opens_with_exact_bytes_and_declared_layers() {
    let drawing = Drawing::from_bytes(LIBREOFFICE_ODG.to_vec()).unwrap();
    assert_eq!(drawing.as_bytes(), LIBREOFFICE_ODG);
    assert!(!drawing.pages().is_empty());
    assert!(drawing.pages().iter().any(|page| !page.shapes().is_empty()));
    for expected in [
        "layout",
        "background",
        "backgroundobjects",
        "controls",
        "measurelines",
    ] {
        assert!(
            drawing
                .layers()
                .iter()
                .any(|layer| layer.name() == expected)
        );
    }
}

#[test]
fn multiple_real_libreoffice_drawings_open_exactly_with_typed_pages() {
    for (name, bytes) in REAL_DRAW_CORPUS {
        let drawing = Drawing::from_bytes(bytes.to_vec())
            .unwrap_or_else(|error| panic!("real Draw fixture {name} failed: {error}"));
        assert_eq!(drawing.as_bytes(), *bytes, "fixture {name} lost provenance");
        assert!(!drawing.pages().is_empty(), "fixture {name} lost pages");
    }
}

#[test]
fn complex_real_fontwork_change_reopens_and_inverts_exactly() {
    let bytes = REAL_DRAW_CORPUS
        .iter()
        .find_map(|(name, bytes)| (*name == "fontwork").then_some(*bytes))
        .unwrap();
    let drawing = Drawing::from_bytes(bytes.to_vec()).unwrap();
    let shape = drawing.pages()[0]
        .shapes()
        .iter()
        .position(|shape| shape.name().is_some())
        .unwrap();
    let mut edit = drawing.edit();
    edit.set_shape_name(0, shape, "Fontwork changed by Litchi")
        .unwrap();
    let commit = edit.commit().unwrap();
    assert_eq!(
        commit.snapshot().pages()[0].shapes()[shape].name(),
        Some("Fontwork changed by Litchi")
    );
    let fully_reopened = Drawing::from_bytes(commit.snapshot().as_bytes().to_vec()).unwrap();
    assert_eq!(fully_reopened.as_bytes(), commit.snapshot().as_bytes());
    assert_eq!(
        commit
            .patch()
            .durable()
            .unwrap()
            .inverse()
            .apply(commit.snapshot())
            .unwrap()
            .as_bytes(),
        drawing.as_bytes()
    );
}

#[test]
fn complex_real_nested_group_geometry_change_reopens_and_inverts_exactly() {
    let bytes = REAL_DRAW_CORPUS
        .iter()
        .find_map(|(name, bytes)| (*name == "complex-groups").then_some(*bytes))
        .unwrap();
    let drawing = Drawing::from_bytes(bytes.to_vec()).unwrap();
    let (group_shape, descendant, y, width, height) = drawing.pages()[0]
        .shapes()
        .iter()
        .enumerate()
        .filter(|(_, shape)| shape.kind() == litchi_odg::shape::ShapeKind::Group)
        .find_map(|(group_shape, _)| {
            let group = drawing.group(0, group_shape).ok()?;
            group.descendants().iter().find_map(|descendant| {
                let shape = &drawing.pages()[0].shapes()[*descendant];
                shape.x()?;
                Some((
                    group_shape,
                    *descendant,
                    shape.y()?.to_owned(),
                    shape.width()?.to_owned(),
                    shape.height()?.to_owned(),
                ))
            })
        })
        .unwrap();
    let mut edit = drawing.edit();
    edit.set_group_descendant_geometry(0, group_shape, descendant, "9cm", y, width, height)
        .unwrap();
    let commit = edit.commit().unwrap();
    assert_eq!(
        commit.snapshot().pages()[0].shapes()[descendant].x(),
        Some("9cm")
    );
    let reopened = Drawing::from_bytes(commit.snapshot().as_bytes().to_vec()).unwrap();
    assert_eq!(reopened.as_bytes(), commit.snapshot().as_bytes());
    assert_eq!(
        commit
            .patch()
            .durable()
            .unwrap()
            .inverse()
            .apply(commit.snapshot())
            .unwrap()
            .as_bytes(),
        drawing.as_bytes()
    );
}
