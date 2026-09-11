#![allow(clippy::unwrap_used, reason = "test assertions use unwrap for clarity")]

use litchi_odf_common::core::PackageWriter;
use litchi_odg::Drawing;
use soapberry_zip::office::StreamingArchiveWriter;

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

const SPAN_CONTENT: &str = concat!(
    r#"<office:document-content xmlns="urn:example:default" xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:foreign="urn:example:foreign" office:version='1.4'>"#,
    r#"<office:body><office:drawing><draw:page draw:name='Page 1'><draw:layer-set><draw:layer draw:name='front'/><draw:layer draw:name="back"/></draw:layer-set>"#,
    r#"<draw:rect foreign:opaque='&quot;keep&amp;me&quot;' name='default-name' draw:style-name='gr1' svg:height='4cm' draw:layer='front' svg:width='3cm' draw:name='Box' svg:y='2cm' svg:x='1cm' draw:transform='translate (1cm 2cm)' foreign:width='foreign-width'/>"#,
    r#"<draw:line svg:x2='8cm' foreign:x1="foreign-x1" draw:name='Axis' svg:y1="2cm" draw:layer="front" svg:x1='1cm' draw:style-name='gr1' svg:y2='9cm' foreign:opaque="line &amp; raw" draw:transform='rotate (5)'/>"#,
    r#"<draw:polygon draw:points='0,0 10,0 10,10' foreign:points="foreign-points" draw:name='Poly' svg:viewBox="0 0 100 100" draw:transform='translate (1cm)' draw:layer='front' draw:style-name='gr1' foreign:viewBox='foreign-view-box'/>"#,
    r#"<draw:path svg:d='M 0 0 L 10 10' draw:style-name='gr1' foreign:d='foreign-path' draw:name='Path' draw:transform='skewX (1)' draw:layer='front'/>"#,
    r#"<draw:control foreign:control='foreign-control' draw:control='control&amp;old' svg:height='1cm' draw:name='Control' draw:layer='front' draw:style-name='gr1' svg:width='2cm' svg:x='1cm' svg:y='2cm'/>"#,
    r#"</draw:page></office:drawing></office:body></office:document-content>"#,
);

const ALIAS_CONTENT: &str = concat!(
    r#"<office:document-content xmlns="urn:example:default" xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:d="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:drawAlias="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:s="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:svgAlias="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:foreign="urn:example:foreign" office:version='1.4'>"#,
    r#"<office:body><office:drawing><drawAlias:page drawAlias:name='Alias page'><drawAlias:layer-set><drawAlias:layer drawAlias:name='front'/><drawAlias:layer drawAlias:name="back"/></drawAlias:layer-set>"#,
    r#"<drawAlias:rect foreign:opaque='&quot;keep&amp;exact&quot;' name='default-name' foreign:name='foreign-name' drawAlias:style-name='gr1' foreign:style-name='foreign-style' svgAlias:height='4cm' foreign:height='foreign-height' drawAlias:layer='front' svgAlias:width='3cm' foreign:width='foreign-width' drawAlias:name='old&amp;name' svgAlias:y='2cm' foreign:y='foreign-y' svgAlias:x='1cm' foreign:x='foreign-x' drawAlias:transform='translate (1cm 2cm)' foreign:transform='foreign-transform'/>"#,
    r#"</drawAlias:page></office:drawing></office:body></office:document-content>"#,
);

#[test]
fn all_shape_attribute_spans_edit_exactly_and_inverse_reopens() {
    let source = Drawing::from_bytes(package(SPAN_CONTENT)).unwrap();
    assert_eq!(source.pages()[0].shapes().len(), 5);
    assert_eq!(source.pages()[0].shapes()[0].name(), Some("Box"));
    assert_eq!(
        source.pages()[0].shapes()[1].line_geometry(),
        [Some("1cm"), Some("2cm"), Some("8cm"), Some("9cm"),]
    );
    assert_eq!(
        source.pages()[0].shapes()[2].view_box(),
        Some("0 0 100 100")
    );
    assert_eq!(
        source.pages()[0].shapes()[3].path_data(),
        Some("M 0 0 L 10 10")
    );
    assert_eq!(
        source.pages()[0].shapes()[4].control_reference(),
        Some("control&old")
    );

    let mut edit = source.edit();
    edit.set_shape_name(0, 0, "Renamed & <Box>").unwrap();
    edit.set_shape_layer(0, 0, "back").unwrap();
    edit.set_shape_geometry(0, 0, "10cm", "11cm", "12cm", "13cm")
        .unwrap();
    edit.set_shape_transform(0, 0, "rotate (20) translate (3cm 4cm)")
        .unwrap();
    edit.set_shape_style_name(0, 0, "gr2 & styled").unwrap();
    edit.set_shape_line_geometry(0, 1, "11cm", "12cm", "18cm", "19cm")
        .unwrap();
    edit.set_shape_points(0, 2, "0 0 200 200", "0,0 100,200 200,0")
        .unwrap();
    edit.set_shape_path_data(0, 3, "M 1 2 C 3 4 5 6 7 8")
        .unwrap();
    edit.set_shape_control_reference(0, 4, "control & updated")
        .unwrap();
    let commit = edit.commit().unwrap();

    let expected = SPAN_CONTENT
        .replace("draw:name='Box'", "draw:name='Renamed &amp; &lt;Box&gt;'")
        .replacen("draw:layer='front'", "draw:layer='back'", 1)
        .replacen("svg:x='1cm'", "svg:x='10cm'", 1)
        .replacen("svg:y='2cm'", "svg:y='11cm'", 1)
        .replace("svg:width='3cm'", "svg:width='12cm'")
        .replace("svg:height='4cm'", "svg:height='13cm'")
        .replace(
            "draw:transform='translate (1cm 2cm)'",
            "draw:transform='rotate (20) translate (3cm 4cm)'",
        )
        .replacen(
            "draw:style-name='gr1'",
            "draw:style-name='gr2 &amp; styled'",
            1,
        )
        .replace("svg:x1='1cm'", "svg:x1='11cm'")
        .replace("svg:y1=\"2cm\"", "svg:y1=\"12cm\"")
        .replace("svg:x2='8cm'", "svg:x2='18cm'")
        .replace("svg:y2='9cm'", "svg:y2='19cm'")
        .replace("svg:viewBox=\"0 0 100 100\"", "svg:viewBox=\"0 0 200 200\"")
        .replace(
            "draw:points='0,0 10,0 10,10'",
            "draw:points='0,0 100,200 200,0'",
        )
        .replace("svg:d='M 0 0 L 10 10'", "svg:d='M 1 2 C 3 4 5 6 7 8'")
        .replace(
            "draw:control='control&amp;old'",
            "draw:control='control &amp; updated'",
        );
    assert_eq!(commit.snapshot().content_xml(), expected);

    for preserved in [
        "foreign:opaque='&quot;keep&amp;me&quot;'",
        "name='default-name'",
        "foreign:width='foreign-width'",
        "foreign:x1=\"foreign-x1\"",
        "foreign:opaque=\"line &amp; raw\"",
        "foreign:points=\"foreign-points\"",
        "foreign:viewBox='foreign-view-box'",
        "foreign:d='foreign-path'",
        "foreign:control='foreign-control'",
    ] {
        assert!(
            commit.snapshot().content_xml().contains(preserved),
            "unrelated lexical attribute was changed: {preserved}"
        );
    }
    assert_eq!(
        commit.snapshot().pages()[0].shapes()[0].name(),
        Some("Renamed & <Box>")
    );
    assert_eq!(
        commit.snapshot().pages()[0].shapes()[0].layer(),
        Some("back")
    );
    assert_eq!(
        commit.snapshot().pages()[0].shapes()[0].line_geometry(),
        [None, None, None, None]
    );
    assert_eq!(
        commit.snapshot().pages()[0].shapes()[1].line_geometry(),
        [Some("11cm"), Some("12cm"), Some("18cm"), Some("19cm")]
    );
    assert_eq!(
        commit.snapshot().pages()[0].shapes()[2].points(),
        Some("0,0 100,200 200,0")
    );
    assert_eq!(
        commit.snapshot().pages()[0].shapes()[3].path_data(),
        Some("M 1 2 C 3 4 5 6 7 8")
    );
    assert_eq!(
        commit.snapshot().pages()[0].shapes()[4].control_reference(),
        Some("control & updated")
    );

    let durable = commit.patch().durable().unwrap();
    assert_eq!(
        durable.apply(source.snapshot()).unwrap().as_bytes(),
        commit.snapshot().as_bytes()
    );
    assert_eq!(
        durable
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
fn aliased_shape_attributes_ignore_same_local_foreign_fields() {
    let source = Drawing::from_bytes(package(ALIAS_CONTENT)).unwrap();
    let shape = &source.pages()[0].shapes()[0];
    assert_eq!(source.pages()[0].name(), Some("Alias page"));
    assert_eq!(shape.name(), Some("old&name"));
    assert_eq!(shape.layer(), Some("front"));
    assert_eq!(shape.style_name(), Some("gr1"));
    assert_eq!(shape.x(), Some("1cm"));
    assert_eq!(shape.y(), Some("2cm"));
    assert_eq!(shape.width(), Some("3cm"));
    assert_eq!(shape.height(), Some("4cm"));
    assert_eq!(shape.transform(), Some("translate (1cm 2cm)"));

    let mut edit = source.edit();
    edit.set_shape_name(0, 0, "new & <name>").unwrap();
    edit.set_shape_layer(0, 0, "back").unwrap();
    edit.set_shape_geometry(0, 0, "10cm", "11cm", "12cm", "13cm")
        .unwrap();
    edit.set_shape_transform(0, 0, "rotate (25)").unwrap();
    edit.set_shape_style_name(0, 0, "gr2 & alias").unwrap();
    let commit = edit.commit().unwrap();

    let expected = ALIAS_CONTENT
        .replace(
            "drawAlias:name='old&amp;name'",
            "drawAlias:name='new &amp; &lt;name&gt;'",
        )
        .replace("drawAlias:layer='front'", "drawAlias:layer='back'")
        .replace("svgAlias:x='1cm'", "svgAlias:x='10cm'")
        .replace("svgAlias:y='2cm'", "svgAlias:y='11cm'")
        .replace("svgAlias:width='3cm'", "svgAlias:width='12cm'")
        .replace("svgAlias:height='4cm'", "svgAlias:height='13cm'")
        .replace(
            "drawAlias:transform='translate (1cm 2cm)'",
            "drawAlias:transform='rotate (25)'",
        )
        .replace(
            "drawAlias:style-name='gr1'",
            "drawAlias:style-name='gr2 &amp; alias'",
        );
    assert_eq!(commit.snapshot().content_xml(), expected);
    for preserved in [
        "name='default-name'",
        "foreign:name='foreign-name'",
        "foreign:style-name='foreign-style'",
        "foreign:height='foreign-height'",
        "foreign:width='foreign-width'",
        "foreign:y='foreign-y'",
        "foreign:x='foreign-x'",
        "foreign:transform='foreign-transform'",
        "foreign:opaque='&quot;keep&amp;exact&quot;'",
    ] {
        assert!(
            commit.snapshot().content_xml().contains(preserved),
            "same-local foreign attribute was not preserved: {preserved}"
        );
    }
    let reopened = Drawing::from_bytes(commit.snapshot().as_bytes().to_vec()).unwrap();
    assert_eq!(reopened.pages()[0].shapes()[0].name(), Some("new & <name>"));
    assert_eq!(reopened.pages()[0].shapes()[0].layer(), Some("back"));
    assert_eq!(reopened.pages()[0].shapes()[0].x(), Some("10cm"));
    assert_eq!(reopened.pages()[0].shapes()[0].height(), Some("13cm"));
}

#[test]
fn malformed_trailing_shape_attribute_is_still_rejected() {
    const CONTENT: &str = r#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" office:version="1.4"><office:body><office:drawing><draw:page draw:name="Page"><draw:rect draw:name="Box" svg:x="1cm" svg:y="2cm" svg:width="3cm" svg:height="4cm" malformed/></draw:page></office:drawing></office:body></office:document-content>"#;

    let valid = CONTENT.replace(" malformed/>", "/>");
    assert!(Drawing::from_bytes(raw_package(&valid)).is_ok());
    assert!(Drawing::from_bytes(raw_package(CONTENT)).is_err());
}

#[test]
fn duplicate_trailing_shape_attribute_is_still_rejected() {
    const CONTENT: &str = r#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:foreign="urn:example:foreign" office:version="1.4"><office:body><office:drawing><draw:page draw:name="Page"><draw:rect draw:name="Box" svg:x="1cm" svg:y="2cm" svg:width="3cm" svg:height="4cm" foreign:opaque='first' foreign:opaque="second"/></draw:page></office:drawing></office:body></office:document-content>"#;

    let valid = CONTENT.replace(" foreign:opaque=\"second\"", "");
    assert!(Drawing::from_bytes(raw_package(&valid)).is_ok());
    assert!(Drawing::from_bytes(raw_package(CONTENT)).is_err());
}
