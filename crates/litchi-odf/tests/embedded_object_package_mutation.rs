use litchi_odf::{
    Document, ImageSource, OdfEmbeddedObjectSource, OdfEmbeddedResource, OdfEmbeddedResourceFile,
    OdfEmbeddedResourceKind, OdfEmbeddedResourceSource, OwnedPackage, PackageWriter, Presentation,
    Spreadsheet, constants,
};
const OFFICE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const DRAW: &str = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
const TEXT: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const TABLE: &str = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
const XLINK: &str = "http://www.w3.org/1999/xlink";
fn package(mimetype: &str, family: &str, inner: &str, files: &[(&str, &[u8], &str)]) -> Vec<u8> {
    let body = match family {
        "text" => format!(
            r#"<office:text><text:p litchi:unknown="preserved">sentinel</text:p>{inner}</office:text>"#
        ),
        "spreadsheet" => format!(
            r#"<office:spreadsheet><table:table table:name="Sheet1"><table:table-row><table:table-cell/></table:table-row>{inner}</table:table></office:spreadsheet>"#
        ),
        "presentation" => format!(
            r#"<office:presentation><draw:page draw:name="Slide1">{inner}</draw:page></office:presentation>"#
        ),
        _ => unreachable!(),
    };
    let xml = format!(
        r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="{OFFICE}" xmlns:draw="{DRAW}" xmlns:text="{TEXT}" xmlns:table="{TABLE}" xmlns:xlink="{XLINK}" xmlns:litchi="urn:litchi:unknown" office:version="1.3"><office:body>{body}</office:body></office:document-content>"#
    );
    let mut writer = PackageWriter::new();
    writer.set_mimetype(mimetype).unwrap();
    writer
        .add_file(constants::ODF_CONTENT, xml.as_bytes())
        .unwrap();
    for (path, bytes, media_type) in files {
        writer
            .add_file_with_media_type(path, bytes, media_type)
            .unwrap();
    }
    writer.finish_to_bytes().unwrap()
}
fn linked(kind: OdfEmbeddedResourceKind, href: &str) -> OdfEmbeddedResource {
    OdfEmbeddedResource {
        kind,
        source: OdfEmbeddedResourceSource::Linked {
            href: href.to_string(),
        },
        frame_name: None,
        xml_id: None,
        class_id: None,
    }
}
fn packaged(kind: OdfEmbeddedResourceKind, bytes: &[u8], media_type: &str) -> OdfEmbeddedResource {
    OdfEmbeddedResource {
        kind,
        source: OdfEmbeddedResourceSource::PackageFile {
            bytes: bytes.to_vec(),
            media_type: media_type.to_string(),
            preferred_path: None,
        },
        frame_name: Some("authored".to_string()),
        xml_id: None,
        class_id: (kind == OdfEmbeddedResourceKind::ObjectOle).then(|| "urn:uuid:test".to_string()),
    }
}
#[test]
fn odt_subdocument_ole_link_reorder_and_atomic_validation() {
    let mut document = Document::from_bytes(package(constants::ODF_TEXT, "text", "", &[])).unwrap();
    let subdocument = OdfEmbeddedResource {
        kind: OdfEmbeddedResourceKind::Object,
        source: OdfEmbeddedResourceSource::PackageSubdocument {
            files: vec![
                OdfEmbeddedResourceFile {
                    path: "content.xml".to_string(),
                    bytes: format!(r#"<office:document-content xmlns:office="{OFFICE}"><office:body><office:text/></office:body></office:document-content>"#).into_bytes(),
                    media_type: "text/xml".to_string(),
                },
                OdfEmbeddedResourceFile {
                    path: "settings.xml".to_string(),
                    bytes: b"<settings/>".to_vec(),
                    media_type: "text/xml".to_string(),
                },
            ],
            media_type: constants::ODF_TEXT.to_string(),
            preferred_root: Some("Object_1".to_string()),
        },
        frame_name: Some("subdocument".to_string()),
        xml_id: Some("embedded-one".to_string()),
        class_id: None,
    };
    assert_eq!(document.add_embedded_resource(&subdocument).unwrap(), 0);
    assert_eq!(
        document
            .add_embedded_resource(&packaged(
                OdfEmbeddedResourceKind::ObjectOle,
                b"opaque-ole",
                "application/vnd.ms-ole"
            ))
            .unwrap(),
        1
    );
    assert_eq!(
        document
            .add_embedded_resource(&linked(
                OdfEmbeddedResourceKind::Object,
                "https://example.invalid/inert"
            ))
            .unwrap(),
        2
    );
    document.move_embedded_object(2, 0).unwrap();
    let objects = document.embedded_objects().unwrap();
    assert!(
        matches!(&objects[0].source, OdfEmbeddedObjectSource::Linked { href } if href == "https://example.invalid/inert")
    );
    assert!(
        matches!(&objects[1].source, OdfEmbeddedObjectSource::PackageSubdocument { root_path, .. } if root_path == "Object_1/")
    );
    assert!(
        matches!(&objects[2].source, OdfEmbeddedObjectSource::PackageFile { manifest_media_type: Some(media), .. } if media == "application/vnd.ms-ole")
    );
    let archive = OwnedPackage::from_bytes(document.to_bytes().unwrap()).unwrap();
    assert_eq!(
        archive
            .package()
            .unwrap()
            .manifest()
            .get_media_type("Object_1/"),
        Some(constants::ODF_TEXT)
    );
    assert_eq!(
        archive
            .package()
            .unwrap()
            .manifest()
            .get_media_type("Object_1/content.xml"),
        Some("text/xml")
    );
    assert!(
        String::from_utf8(archive.get_file(constants::ODF_CONTENT).unwrap())
            .unwrap()
            .contains("litchi:unknown=\"preserved\"")
    );
    let before = document.to_bytes().unwrap();
    let mut traversal = packaged(
        OdfEmbeddedResourceKind::ObjectOle,
        b"x",
        "application/vnd.ms-ole",
    );
    if let OdfEmbeddedResourceSource::PackageFile { preferred_path, .. } = &mut traversal.source {
        *preferred_path = Some("../escape.bin".to_string());
    }
    assert!(document.replace_embedded_object(2, &traversal).is_err());
    assert_eq!(document.to_bytes().unwrap(), before);
    let executable = packaged(
        OdfEmbeddedResourceKind::ObjectOle,
        b"MZ",
        "application/x-msdownload",
    );
    assert!(document.replace_embedded_object(2, &executable).is_err());
    assert_eq!(document.to_bytes().unwrap(), before);
    document.remove_embedded_object(1).unwrap();
    let archive = OwnedPackage::from_bytes(document.to_bytes().unwrap()).unwrap();
    assert!(
        !archive
            .files()
            .unwrap()
            .iter()
            .any(|path| path.starts_with("Object_1/"))
    );
}
#[test]
fn shared_package_payload_is_removed_only_after_last_reference() {
    let object = r#"<draw:frame><draw:object xlink:href="Shared.bin"/></draw:frame>"#;
    let bytes = package(
        constants::ODF_TEXT,
        "text",
        &format!("{object}{object}"),
        &[("Shared.bin", b"shared", "application/octet-stream")],
    );
    let mut document = Document::from_bytes(bytes).unwrap();
    document.remove_embedded_object(0).unwrap();
    assert!(
        OwnedPackage::from_bytes(document.to_bytes().unwrap())
            .unwrap()
            .has_file("Shared.bin")
            .unwrap()
    );
    document.remove_embedded_object(0).unwrap();
    assert!(
        !OwnedPackage::from_bytes(document.to_bytes().unwrap())
            .unwrap()
            .has_file("Shared.bin")
            .unwrap()
    );
}
#[test]
fn ods_and_odp_image_package_and_inline_mutation() {
    let mut sheet =
        Spreadsheet::from_bytes(package(constants::ODF_SPREADSHEET, "spreadsheet", "", &[]))
            .unwrap();
    let png = packaged(OdfEmbeddedResourceKind::Image, b"\x89PNG\r\n", "image/png");
    assert_eq!(sheet.add_embedded_resource("Sheet1", &png).unwrap(), 0);
    assert!(
        matches!(&sheet.images().unwrap()[0].source, ImageSource::PackagePart { path, manifest_media_type: Some(media), .. } if path == "Pictures/Image_1.png" && media == "image/png")
    );
    let inline = OdfEmbeddedResource {
        kind: OdfEmbeddedResourceKind::Image,
        source: OdfEmbeddedResourceSource::InlineBinary {
            bytes: b"inline".to_vec(),
            media_type: Some("image/png".to_string()),
        },
        frame_name: None,
        xml_id: None,
        class_id: None,
    };
    sheet.replace_embedded_image(0, &inline).unwrap();
    assert!(
        matches!(&sheet.images().unwrap()[0].source, ImageSource::Inline { bytes, .. } if bytes == b"inline")
    );
    assert!(
        !OwnedPackage::from_bytes(sheet.to_bytes().unwrap())
            .unwrap()
            .has_file("Pictures/Image_1.png")
            .unwrap()
    );
    let mut slides = Presentation::from_bytes(package(
        constants::ODF_PRESENTATION,
        "presentation",
        "",
        &[],
    ))
    .unwrap();
    slides.add_embedded_resource("Slide1", &png).unwrap();
    slides.add_embedded_resource("Slide1", &inline).unwrap();
    slides.move_embedded_image(1, 0).unwrap();
    assert!(matches!(
        &slides.images().unwrap()[0].source,
        ImageSource::Inline { .. }
    ));
    slides.remove_embedded_image(0).unwrap();
    slides.remove_embedded_image(0).unwrap();
    assert!(slides.images().unwrap().is_empty());
}
#[test]
fn mutations_drop_stale_signatures() {
    let bytes = package(
        constants::ODF_TEXT,
        "text",
        "",
        &[(
            "META-INF/documentsignatures.xml",
            b"<signatures/>",
            "text/xml",
        )],
    );
    let mut document = Document::from_bytes(bytes).unwrap();
    document
        .add_embedded_resource(&linked(
            OdfEmbeddedResourceKind::Object,
            "urn:example:inert",
        ))
        .unwrap();
    let archive = OwnedPackage::from_bytes(document.to_bytes().unwrap()).unwrap();
    assert!(!archive.has_file("META-INF/documentsignatures.xml").unwrap());
}
#[test]
fn libreoffice_ole_fixtures_remain_mutable_without_resolution() {
    let root = concat!(env!("CARGO_MANIFEST_DIR"), "/../../");
    let mut document = Document::open(format!(
        "{root}test-data/libreoffice-core/sw/qa/uibase/shells/data/ole-save-preview-update.odt"
    ))
    .unwrap();
    assert!(!document.embedded_objects().unwrap().is_empty());
    document
        .replace_embedded_object(
            0,
            &linked(
                OdfEmbeddedResourceKind::ObjectOle,
                "https://example.invalid/not-fetched",
            ),
        )
        .unwrap();
    assert!(matches!(
        &document.embedded_objects().unwrap()[0].source,
        OdfEmbeddedObjectSource::Linked { .. }
    ));
    let mut slides = Presentation::open(format!(
        "{root}test-data/libreoffice-core/sd/qa/unit/data/odp/ole_icon.odp"
    ))
    .unwrap();
    if !slides.embedded_objects().unwrap().is_empty() {
        slides.remove_embedded_object(0).unwrap();
        let _ = Presentation::from_bytes(slides.to_bytes().unwrap()).unwrap();
    }
}
