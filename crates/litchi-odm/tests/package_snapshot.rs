#![allow(
    clippy::unwrap_used,
    reason = "tests are expected to panic on unexpected fixture failures"
)]

use litchi_odf_common::core::PackageWriter;
use litchi_odm::Master;

const CONTENT: &str =
    include_str!("../../litchi-odt/tests/fixtures/libreoffice-master-document-content.xml");

#[test]
fn real_master_document_xml_and_auxiliary_entries_remain_exact_and_opaque() {
    let mut writer = PackageWriter::new();
    writer
        .set_mimetype("application/vnd.oasis.opendocument.text-master")
        .unwrap();
    writer.add_file("content.xml", CONTENT.as_bytes()).unwrap();
    writer
        .add_file_with_media_type("custom/cache.bin", b"cached", "application/octet-stream")
        .unwrap();
    let bytes = writer.finish_to_bytes().unwrap();

    let master = Master::from_bytes(bytes.clone()).unwrap();
    assert_eq!(master.as_bytes(), bytes.as_slice());
    assert_eq!(master.content_xml(), CONTENT);
    assert!(master.content_xml().contains("Chapters/a.odt"));
    assert!(master.content_xml().contains("loext:opaque"));
    assert!(
        master
            .files()
            .unwrap()
            .iter()
            .any(|path| path == "custom/cache.bin")
    );
}
