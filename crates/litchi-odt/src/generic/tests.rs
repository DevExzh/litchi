//! Regression tests for the generic ODF snapshot owners.

use super::{Family, FlatDocument, Package};
use crate::constants;
use crate::core::PackageWriter;
use soapberry_zip::office::{ArchiveReader, StreamingArchiveWriter};

fn replace_zip_member_raw(package: &[u8], path: &str, replacement: &[u8]) -> Vec<u8> {
    let archive = ArchiveReader::new(package).unwrap();
    let mut writer = StreamingArchiveWriter::new();
    let mut replaced = false;
    for name in archive.file_names() {
        let data = if name == path {
            replaced = true;
            replacement.to_vec()
        } else {
            archive.read(name).unwrap()
        };
        writer.write_stored(name, &data).unwrap();
    }
    assert!(replaced, "test ZIP member {path} must exist");
    writer.finish_to_bytes().unwrap()
}

fn package(mimetype: &str) -> Vec<u8> {
    let mut writer = PackageWriter::new();
    writer.set_mimetype(mimetype).unwrap();
    writer
        .add_file(
            constants::ODF_CONTENT,
            br#"<?xml version="1.0"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><office:body/></office:document-content>"#,
        )
        .unwrap();
    writer
        .add_file_with_media_type("Pictures/pixel.png", b"PNG", "image/png")
        .unwrap();
    writer.finish_to_bytes().unwrap()
}

#[test]
fn opens_every_document_family_and_template_losslessly() {
    for (mimetype, family, template) in [
        (constants::ODF_TEXT, Family::Text, false),
        (constants::ODF_TEXT_TEMPLATE, Family::Text, true),
        (constants::ODF_SPREADSHEET, Family::Spreadsheet, false),
        (
            constants::ODF_SPREADSHEET_TEMPLATE,
            Family::Spreadsheet,
            true,
        ),
        (constants::ODF_PRESENTATION, Family::Presentation, false),
        (
            constants::ODF_PRESENTATION_TEMPLATE,
            Family::Presentation,
            true,
        ),
        (constants::ODF_DRAWING, Family::Drawing, false),
        (constants::ODF_DRAWING_TEMPLATE, Family::Drawing, true),
        (constants::ODF_CHART, Family::Chart, false),
        (constants::ODF_CHART_TEMPLATE, Family::Chart, true),
        (constants::ODF_FORMULA, Family::Formula, false),
        (constants::ODF_FORMULA_TEMPLATE, Family::Formula, true),
        (constants::ODF_IMAGE, Family::Image, false),
        (constants::ODF_IMAGE_TEMPLATE, Family::Image, true),
        (constants::ODF_MASTER, Family::Master, false),
        (constants::ODF_MASTER_TEMPLATE, Family::Master, true),
        (constants::ODF_WEB, Family::Web, true),
        (constants::ODF_DATABASE, Family::Database, false),
    ] {
        let bytes = package(mimetype);
        let document = Package::from_bytes(bytes.clone()).unwrap();
        assert_eq!(document.family(), family);
        assert_eq!(document.is_template(), template);
        assert_eq!(document.mimetype(), mimetype);
        assert!(document.content_xml().unwrap().contains("office:body"));
        assert!(document.odf_metadata().unwrap().is_none());
        assert_eq!(document.media_files().unwrap(), ["Pictures/pixel.png"]);
        assert_eq!(document.to_bytes(), bytes);
        assert_eq!(document.into_bytes(), bytes);
    }
}

#[test]
fn rejects_non_odf_missing_content_and_invalid_xml_bytes() {
    let mut writer = PackageWriter::new();
    writer.set_mimetype("application/zip").unwrap();
    writer.add_file(constants::ODF_CONTENT, b"<x/>").unwrap();
    assert!(Package::from_bytes(writer.finish_to_bytes().unwrap()).is_err());

    let mut writer = PackageWriter::new();
    writer.set_mimetype(constants::ODF_DRAWING).unwrap();
    assert!(Package::from_bytes(writer.finish_to_bytes().unwrap()).is_err());

    let mut writer = PackageWriter::new();
    writer.set_mimetype(constants::ODF_CHART).unwrap();
    writer.add_file(constants::ODF_CONTENT, b"<x/>").unwrap();
    let package = writer.finish_to_bytes().unwrap();
    let invalid_xml = replace_zip_member_raw(&package, constants::ODF_CONTENT, &[0xff]);
    assert!(Package::from_bytes(invalid_xml).is_err());
}

#[test]
fn opens_standard_and_odfdo_compatible_flat_documents_losslessly() {
    for (mimetype, body, family, extension) in [
        (constants::ODF_TEXT, "text", Family::Text, "fodt"),
        (
            constants::ODF_SPREADSHEET,
            "spreadsheet",
            Family::Spreadsheet,
            "fods",
        ),
        (
            constants::ODF_PRESENTATION,
            "presentation",
            Family::Presentation,
            "fodp",
        ),
        (constants::ODF_DRAWING, "drawing", Family::Drawing, "fodg"),
        (constants::ODF_CHART, "chart", Family::Chart, "fodc"),
        (constants::ODF_FORMULA, "formula", Family::Formula, "fodf"),
        (constants::ODF_IMAGE, "image", Family::Image, "fodi"),
    ] {
        let xml = format!(
            r#"<?xml version="1.0"?><!-- keep --><o:document xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" o:mimetype="{mimetype}" o:version="1.3"><o:body><o:{body}/></o:body></o:document>"#
        );
        let document = FlatDocument::from_bytes(xml.clone().into_bytes()).unwrap();
        assert_eq!(document.family(), family);
        assert_eq!(document.mimetype(), mimetype);
        assert_eq!(document.extension(), extension);
        assert_eq!(document.xml(), xml);
        assert_eq!(document.to_bytes(), xml.as_bytes());
        assert_eq!(document.into_bytes(), xml.into_bytes());
    }
}

#[test]
fn rejects_flat_mimetype_body_mismatch_and_incomplete_xml() {
    for xml in [
        r#"<o:document xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" o:mimetype="application/vnd.oasis.opendocument.text"><o:body><o:spreadsheet/></o:body></o:document>"#,
        r#"<o:document xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" o:mimetype="application/vnd.oasis.opendocument.text"><o:body><o:text/></o:body>"#,
        r#"<o:document xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" o:mimetype="application/vnd.oasis.opendocument.text-template"><o:body><o:text/></o:body></o:document>"#,
        r#"<o:document xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" o:mimetype="application/vnd.oasis.opendocument.text"><o:body><o:text/></o:body></o:document><o:document/>"#,
    ] {
        assert!(
            FlatDocument::from_bytes(xml.as_bytes().to_vec()).is_err(),
            "accepted invalid flat document {xml}"
        );
    }
}

#[test]
fn flat_document_exposes_namespace_aware_metadata() {
    let xml = br#"<o:document xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
        xmlns:d="http://purl.org/dc/elements/1.1/"
        o:mimetype="application/vnd.oasis.opendocument.text">
        <o:meta><d:title>A &amp; B</d:title></o:meta>
        <o:body><o:text/></o:body>
    </o:document>"#;
    let document = FlatDocument::from_bytes(xml.to_vec()).unwrap();
    assert_eq!(
        document.odf_metadata().unwrap().title.as_deref(),
        Some("A & B")
    );
    assert_eq!(document.metadata().unwrap().title.as_deref(), Some("A & B"));
}

#[test]
fn flat_variable_declarations_expand_replace_and_remove_atomically() {
    let xml = format!(
        r#"<o:document xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" o:mimetype="{}"><o:body><o:text/></o:body></o:document>"#,
        constants::ODF_TEXT,
    );
    let mut document = FlatDocument::from_bytes(xml.into_bytes()).unwrap();
    let scope = crate::variable_declaration::Scope::Body(crate::variable_declaration::Body::Text);
    let first = crate::variable_declaration::Group {
        kind: crate::variable_declaration::Kind::Simple,
        part: crate::variable_declaration::Part::Flat,
        scope: scope.clone(),
        declarations: vec![crate::variable_declaration::Declaration::Simple {
            name: "counter".to_string(),
            value_type: crate::variable_declaration::ValueType::Float,
        }],
    };
    assert!(
        document
            .set_variable_declaration_group(&first)
            .unwrap()
            .is_none()
    );
    assert!(document.xml().contains("<o:text><text:variable-decls"));
    assert!(
        document
            .variable_declarations()
            .unwrap()
            .find(crate::variable_declaration::Kind::Simple, "counter")
            .is_some()
    );

    let second = crate::variable_declaration::Group {
        declarations: vec![crate::variable_declaration::Declaration::Simple {
            name: "replacement".to_string(),
            value_type: crate::variable_declaration::ValueType::String,
        }],
        ..first.clone()
    };
    assert_eq!(
        document.set_variable_declaration_group(&second).unwrap(),
        Some(first.clone()),
    );
    assert!(
        document
            .variable_declarations()
            .unwrap()
            .find(crate::variable_declaration::Kind::Simple, "replacement")
            .is_some()
    );
    assert_eq!(
        document
            .remove_variable_declaration_group(&scope, crate::variable_declaration::Kind::Simple)
            .unwrap(),
        Some(second),
    );
    assert!(document.variable_declarations().unwrap().groups.is_empty());
}
