#![cfg_attr(
    test,
    allow(
        clippy::unwrap_used,
        reason = "Fixed in-memory package fixtures use direct assertion setup."
    )
)]

use litchi_odf_common::{
    constants,
    core::{OwnedPackage, PackageWriter},
    package::{
        edit::{Addition, rebuild_package, splice},
        is_linked_href, resolve_package_path,
    },
};

const MODULE_PATH: &str = "Basic/Standard/Module1.xml";
const LINK: &str = "javascript:alert(1)";
const CONTENT: &str = r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:script="urn:oasis:names:tc:opendocument:xmlns:script:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:litchi="urn:litchi:test"><office:scripts><office:script script:language="ooo:Basic"><litchi:payload>one</litchi:payload></office:script></office:scripts><office:event-listeners><script:event-listener script:event-name="dom:load" script:language="ooo:Basic" xlink:href="javascript:alert(1)"/></office:event-listeners><office:body><office:text><text:p>preserved body</text:p></office:text></office:body></office:document-content>"#;

fn package() -> Vec<u8> {
    let mut writer = PackageWriter::new();
    writer.set_mimetype(constants::ODF_TEXT).unwrap();
    writer
        .add_file(constants::ODF_CONTENT, CONTENT.as_bytes())
        .unwrap();
    writer
        .add_file_with_media_type(
            MODULE_PATH,
            b"<module xmlns=\"urn:test\">one</module>",
            "text/xml",
        )
        .unwrap();
    writer
        .add_file_with_media_type(
            "Scripts/payload.bin",
            &[0, 255, 1, 2],
            "application/octet-stream",
        )
        .unwrap();
    writer.finish_to_bytes().unwrap()
}

#[test]
fn script_markup_and_resources_are_updated_as_inert_package_bytes() {
    let source = OwnedPackage::from_bytes(package()).unwrap();
    let content = String::from_utf8(source.get_file(constants::ODF_CONTENT).unwrap()).unwrap();
    let value = content.find(">one<").unwrap() + 1;
    let updated_content = splice(&content, value, value + "one".len(), "two").unwrap();
    let updated_bytes = rebuild_package(
        &source,
        &updated_content,
        vec![Addition {
            path: MODULE_PATH.to_string(),
            bytes: b"<module xmlns=\"urn:test\">updated</module>".to_vec(),
            media_type: "text/xml".to_string(),
        }],
        Vec::new(),
        vec![MODULE_PATH.to_string()],
        Vec::new(),
    )
    .unwrap();
    let updated = OwnedPackage::from_bytes(updated_bytes).unwrap();
    let updated_xml = String::from_utf8(updated.get_file(constants::ODF_CONTENT).unwrap()).unwrap();
    assert!(updated_xml.contains("<litchi:payload>two</litchi:payload>"));
    assert!(updated_xml.contains(LINK));
    assert_eq!(
        updated.get_file(MODULE_PATH).unwrap(),
        b"<module xmlns=\"urn:test\">updated</module>"
    );
    assert_eq!(
        updated.get_file("Scripts/payload.bin").unwrap(),
        [0, 255, 1, 2]
    );
}

#[test]
fn executable_uris_stay_linked_and_unsafe_resource_paths_are_rejected() {
    let source = OwnedPackage::from_bytes(package()).unwrap();
    let before = source.as_bytes().to_vec();
    assert!(is_linked_href(LINK));
    assert!(resolve_package_path("../../content.xml").is_err());
    assert_eq!(source.as_bytes(), before);
    assert_eq!(
        source.get_file(MODULE_PATH).unwrap(),
        b"<module xmlns=\"urn:test\">one</module>"
    );
}
