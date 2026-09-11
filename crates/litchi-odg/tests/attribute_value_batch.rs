#![allow(clippy::unwrap_used, reason = "test assertions use unwrap for clarity")]

use litchi_odf_common::core::PackageWriter;
use litchi_odg::{Drawing, shape::ShapeKind};
use soapberry_zip::office::StreamingArchiveWriter;

const NAMESPACE_DECLARATIONS: &str = concat!(
    r#"xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" "#,
    r#"xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" "#,
    r#"xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" "#,
    r#"xmlns:dr3d="urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0" "#,
    r#"xmlns:drawAlias="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" "#,
    r#"xmlns:svgAlias="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" "#,
    r#"xmlns:foreign="urn:example:foreign""#,
);

fn document(shape: &str) -> String {
    format!(
        r#"<office:document-content {NAMESPACE_DECLARATIONS} office:version="1.4"><office:body><office:drawing><draw:page draw:name="Page">{shape}</draw:page></office:drawing></office:body></office:document-content>"#
    )
}

fn package(content: &str) -> Vec<u8> {
    let mut writer = PackageWriter::new();
    writer
        .set_mimetype("application/vnd.oasis.opendocument.graphics")
        .unwrap();
    writer.add_file("content.xml", content.as_bytes()).unwrap();
    writer.finish_to_bytes().unwrap()
}

fn raw_package(content: &str) -> Vec<u8> {
    const MIMETYPE: &[u8] = b"application/vnd.oasis.opendocument.graphics";
    const MANIFEST: &[u8] = br#"<?xml version="1.0"?><manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.graphics"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/></manifest:manifest>"#;
    let mut archive = StreamingArchiveWriter::new();
    archive.write_stored("mimetype", MIMETYPE).unwrap();
    archive
        .write_deflated("content.xml", content.as_bytes())
        .unwrap();
    archive
        .write_deflated("META-INF/manifest.xml", MANIFEST)
        .unwrap();
    archive.finish_to_bytes().unwrap()
}

fn error_text(content: &str) -> String {
    match Drawing::from_bytes(raw_package(content)) {
        Ok(_) => panic!("fixture unexpectedly opened successfully"),
        Err(error) => error.to_string(),
    }
}

const VALUE_CONTENT: &str = concat!(
    r#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" office:version="1.4">"#,
    r#"<office:body><office:drawing><draw:page draw:name="Page"><draw:layer-set><draw:layer draw:name="front"/></draw:layer-set>"#,
    r#"<draw:polygon draw:name="Poly" draw:z-index="7" draw:control="ctl&#xA;ref" draw:layer="front" svg:d="M&#xA;0&#xA;0 L 1 1" draw:transform="translate&#xA;(1cm 2cm)" draw:points="0,0 10,0 0,10" svg:viewBox="0 0 10 10" draw:style-name="gr1&#x9;" draw:text-style-name="P1&#xA;" svg:x="1cm" svg:y="2cm" svg:width="3cm" svg:height="4cm"/>"#,
    r#"<draw:line draw:name="Line" draw:z-index="8" draw:control="line&#xA;ref" draw:layer="front" svg:d="M 0 0 L 1 1" draw:transform="rotate&#xA;(5)" draw:style-name="gr2" draw:text-style-name="P2" svg:x="1cm" svg:y="2cm" svg:width="3cm" svg:height="4cm" svg:x1="1cm" svg:y1="2cm" svg:x2="8cm" svg:y2="9cm"/>"#,
    r#"</draw:page></office:drawing></office:body></office:document-content>"#,
);

#[test]
fn shape_values_are_normalized_and_noop_edit_and_inverse_preserve_source() {
    let source = Drawing::from_bytes(package(VALUE_CONTENT)).unwrap();
    assert_eq!(source.pages()[0].shapes().len(), 2);

    let polygon = &source.pages()[0].shapes()[0];
    assert_eq!(polygon.kind(), ShapeKind::Polygon);
    assert_eq!(polygon.z_index(), Some(7));
    assert_eq!(polygon.control_reference(), Some("ctl\nref"));
    assert_eq!(polygon.layer(), Some("front"));
    assert_eq!(polygon.path_data(), Some("M\n0\n0 L 1 1"));
    assert_eq!(polygon.transform(), Some("translate\n(1cm 2cm)"));
    assert_eq!(polygon.points(), Some("0,0 10,0 0,10"));
    assert_eq!(polygon.view_box(), Some("0 0 10 10"));
    assert_eq!(polygon.style_name(), Some("gr1\t"));
    assert_eq!(polygon.text_style_name(), Some("P1\n"));
    assert_eq!(
        [polygon.x(), polygon.y(), polygon.width(), polygon.height()],
        [Some("1cm"), Some("2cm"), Some("3cm"), Some("4cm")]
    );

    let line = &source.pages()[0].shapes()[1];
    assert_eq!(line.kind(), ShapeKind::Line);
    assert_eq!(line.z_index(), Some(8));
    assert_eq!(line.control_reference(), Some("line\nref"));
    assert_eq!(line.transform(), Some("rotate\n(5)"));
    assert_eq!(
        line.line_geometry(),
        [Some("1cm"), Some("2cm"), Some("8cm"), Some("9cm")]
    );

    let mut noop = source.edit();
    noop.set_shape_geometry(0, 0, "1cm", "2cm", "3cm", "4cm")
        .unwrap();
    noop.set_shape_points(0, 0, "0 0 10 10", "0,0 10,0 0,10")
        .unwrap();
    noop.set_shape_style_name(0, 0, "gr1\t").unwrap();
    let noop_commit = noop.commit().unwrap();
    assert!(!noop_commit.changed());
    assert_eq!(noop_commit.snapshot().as_bytes(), source.as_bytes());

    let mut edit = source.edit();
    edit.set_shape_style_name(0, 0, "gr3").unwrap();
    let commit = edit.commit().unwrap();
    assert!(commit.changed());
    assert_eq!(
        commit.snapshot().pages()[0].shapes()[0].style_name(),
        Some("gr3")
    );
    let expected =
        VALUE_CONTENT.replace(r#"draw:style-name="gr1&#x9;""#, r#"draw:style-name="gr3""#);
    assert_eq!(commit.snapshot().content_xml(), expected);
    assert_eq!(
        commit
            .patch()
            .inverse()
            .apply(commit.snapshot())
            .unwrap()
            .as_bytes(),
        source.as_bytes()
    );
    let reopened = Drawing::from_bytes(commit.snapshot().as_bytes().to_vec()).unwrap();
    assert_eq!(reopened.content_xml(), expected);
}

#[test]
fn literal_attribute_whitespace_is_normalized_to_spaces() {
    let content = document(
        r#"<draw:polygon draw:name="Literal" draw:control="ctl
ref" draw:layer="front	" svg:d="M
0	0" draw:transform="translate
(1cm 2cm)" draw:points="0,0 10,0 0,10" svg:viewBox="0 0 10 10" draw:style-name="gr1
" draw:text-style-name="P1	" svg:x="1cm" svg:y="2cm" svg:width="3cm" svg:height="4cm"/>"#,
    );
    let drawing = Drawing::from_bytes(raw_package(&content)).unwrap();
    let shape = &drawing.pages()[0].shapes()[0];
    assert_eq!(shape.control_reference(), Some("ctl ref"));
    assert_eq!(shape.layer(), Some("front "));
    assert_eq!(shape.path_data(), Some("M 0 0"));
    assert_eq!(shape.transform(), Some("translate (1cm 2cm)"));
    assert_eq!(shape.style_name(), Some("gr1 "));
    assert_eq!(shape.text_style_name(), Some("P1 "));
}

#[test]
fn z_index_error_precedes_bad_geometry_value_decode() {
    let content = document(
        r#"<draw:rect draw:name="Box" svg:x="&bogus;" draw:z-index="invalid" svg:y="2cm" svg:width="3cm" svg:height="4cm"/>"#,
    );
    assert_eq!(
        error_text(&content),
        "Invalid format: ODG integer attribute is invalid"
    );

    let geometry_only = content.replace(r#"draw:z-index="invalid""#, r#"draw:z-index="7""#);
    assert!(error_text(&geometry_only).starts_with("Invalid format: invalid ODG attribute value:"));

    let valid = content
        .replace(r#"svg:x="&bogus;""#, r#"svg:x="1cm""#)
        .replace(r#"draw:z-index="invalid""#, r#"draw:z-index="7""#);
    assert!(Drawing::from_bytes(raw_package(&valid)).is_ok());
}

#[test]
fn frame_attribute_error_precedes_z_index_error() {
    let content = document(r#"<draw:frame draw:z-index="invalid" svg:x="1cm" svgAlias:x="2cm"/>"#);
    assert_eq!(
        error_text(&content),
        "Invalid format: ODG element has a duplicate namespaced attribute"
    );

    let valid = content
        .replace(r#" svgAlias:x="2cm""#, "")
        .replace(r#"draw:z-index="invalid""#, r#"draw:z-index="7""#);
    let drawing = Drawing::from_bytes(raw_package(&valid)).unwrap();
    assert_eq!(drawing.pages()[0].shapes()[0].kind(), ShapeKind::Frame);
    assert_eq!(drawing.pages()[0].shapes()[0].z_index(), Some(7));
}

#[test]
fn three_dimensional_validation_precedes_z_index_error() {
    let content = document(r#"<dr3d:scene draw:z-index="invalid" dr3d:projection="neither"/>"#);
    assert_eq!(
        error_text(&content),
        "Invalid format: ODG dr3d projection is invalid"
    );

    let valid = content
        .replace(r#"draw:z-index="invalid""#, r#"draw:z-index="7""#)
        .replace(
            r#"dr3d:projection="neither""#,
            r#"dr3d:projection="parallel""#,
        );
    let drawing = Drawing::from_bytes(raw_package(&valid)).unwrap();
    assert_eq!(
        drawing.pages()[0].shapes()[0].kind(),
        ShapeKind::ThreeDimensionalScene
    );
    assert_eq!(drawing.pages()[0].shapes()[0].z_index(), Some(7));
}

#[test]
fn geometry_lexical_validation_precedes_bad_second_group_value_decode() {
    let content = document(
        r#"<draw:rect draw:name="Box" draw:control="&bogus;" svg:x="not-a-length" svg:y="2cm" svg:width="3cm" svg:height="4cm"/>"#,
    );
    assert_eq!(
        error_text(&content),
        "Invalid format: ODG shape coordinate or size is not an ODF length"
    );

    let valid_geometry = content.replace(r#"svg:x="not-a-length""#, r#"svg:x="1cm""#);
    assert!(
        error_text(&valid_geometry).starts_with("Invalid format: invalid ODG attribute value:")
    );

    let valid = valid_geometry.replace(r#"draw:control="&bogus;""#, r#"draw:control="ctl""#);
    assert!(Drawing::from_bytes(raw_package(&valid)).is_ok());
}

#[test]
fn matching_alias_duplicate_and_malformed_unrelated_attribute_are_rejected() {
    let content = document(
        r#"<draw:rect draw:name="Box" draw:layer="front" drawAlias:layer="back" svg:x="1cm" svg:y="2cm" svg:width="3cm" svg:height="4cm" foreign:opaque/>"#,
    );

    // The public scanner's earlier draw:name lookup still checks every raw
    // attribute, so the malformed tail is rejected before either value batch.
    // The duplicate-only control below verifies the matching expanded-name
    // error independently.
    assert!(error_text(&content).starts_with("Invalid format: invalid ODG attribute:"));

    let duplicate_only = content.replace(" foreign:opaque/>", "/>");
    assert_eq!(
        error_text(&duplicate_only),
        "Invalid format: ODG element has a duplicate namespaced attribute"
    );

    let trailing_only = content.replace(r#" drawAlias:layer="back""#, "");
    assert!(error_text(&trailing_only).starts_with("Invalid format: invalid ODG attribute:"));

    let valid = trailing_only.replace(" foreign:opaque/>", "/>");
    let drawing = Drawing::from_bytes(raw_package(&valid)).unwrap();
    assert_eq!(drawing.pages()[0].shapes()[0].layer(), Some("front"));
}
