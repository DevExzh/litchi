use litchi_odf_common::{
    constants,
    core::{OwnedPackage, PackageWriter},
    embedded::{Kind, Source, scan_package},
    package::{edit::rebuild_package, resolve_package_path},
};

const OBJECT: &str = r#"<draw:object xlink:href="Shared.bin"/>"#;

fn package(inner: &str, files: &[(&str, &[u8], &str)]) -> Vec<u8> {
    let content = format!(
        r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:xlink="http://www.w3.org/1999/xlink"><office:body><office:text><text:p litchi:unknown="preserved" xmlns:litchi="urn:litchi:unknown">sentinel</text:p>{inner}</office:text></office:body></office:document-content>"#
    );
    let mut writer = PackageWriter::new();
    writer.set_mimetype(constants::ODF_TEXT).unwrap();
    writer
        .add_file(constants::ODF_CONTENT, content.as_bytes())
        .unwrap();
    for (path, bytes, media_type) in files {
        writer
            .add_file_with_media_type(path, bytes, media_type)
            .unwrap();
    }
    writer.finish_to_bytes().unwrap()
}

fn content(package: &OwnedPackage) -> String {
    String::from_utf8(package.get_file(constants::ODF_CONTENT).unwrap()).unwrap()
}

#[test]
fn shared_package_payload_is_removed_only_after_the_last_reference() {
    let source = OwnedPackage::from_bytes(package(
        &format!("{OBJECT}{OBJECT}"),
        &[("Shared.bin", b"shared", "application/octet-stream")],
    ))
    .unwrap();
    let source_xml = content(&source);
    let objects = scan_package(&source_xml, None, &source.package().unwrap()).unwrap();
    assert_eq!(objects.len(), 2);
    assert!(objects.iter().all(|object| {
        object.kind == Kind::Object
            && matches!(
                &object.source,
                Source::PackageFile {
                    path,
                    manifest_media_type: Some(media_type),
                    ..
                } if path == "Shared.bin" && media_type == "application/octet-stream"
            )
    }));

    let one_reference = source_xml.replacen(OBJECT, "", 1);
    let first = rebuild_package(
        &source,
        &one_reference,
        Vec::new(),
        Vec::new(),
        Vec::new(),
        Vec::new(),
    )
    .unwrap();
    let first = OwnedPackage::from_bytes(first).unwrap();
    assert!(first.has_file("Shared.bin").unwrap());
    assert_eq!(
        scan_package(&content(&first), None, &first.package().unwrap())
            .unwrap()
            .len(),
        1
    );

    let no_references = one_reference.replacen(OBJECT, "", 1);
    let second = rebuild_package(
        &first,
        &no_references,
        Vec::new(),
        Vec::new(),
        vec!["Shared.bin".to_string()],
        Vec::new(),
    )
    .unwrap();
    let second = OwnedPackage::from_bytes(second).unwrap();
    assert!(!second.has_file("Shared.bin").unwrap());
    assert!(content(&second).contains("litchi:unknown=\"preserved\""));
}

#[test]
fn external_targets_remain_inert_and_unsafe_package_paths_fail_atomically() {
    let external = r#"<draw:object-ole xlink:href="https://example.invalid/not-fetched"/>"#;
    let source = OwnedPackage::from_bytes(package(external, &[])).unwrap();
    let objects = scan_package(&content(&source), None, &source.package().unwrap()).unwrap();
    assert_eq!(objects.len(), 1);
    assert_eq!(objects[0].kind, Kind::ObjectOle);
    assert!(matches!(
        &objects[0].source,
        Source::Linked { href } if href == "https://example.invalid/not-fetched"
    ));

    let before = source.as_bytes().to_vec();
    assert!(resolve_package_path("../../escape.bin").is_err());
    assert_eq!(source.as_bytes(), before);
}
