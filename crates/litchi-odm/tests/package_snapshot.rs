#![allow(
    clippy::unwrap_used,
    reason = "tests are expected to panic on unexpected fixture failures"
)]

use litchi_odf_common::core::PackageWriter;
use litchi_odm::Package;

const CONTENT: &str = include_str!("fixtures/libreoffice-master-document-content.xml");

#[test]
fn real_master_document_xml_and_auxiliary_entries_remain_exact_and_opaque() {
    let mut writer = PackageWriter::new(litchi_odf_common::core::StreamingArchiveLimits::default());
    writer
        .set_mimetype("application/vnd.oasis.opendocument.text-master")
        .unwrap();
    writer.add_file("content.xml", CONTENT.as_bytes()).unwrap();
    writer
        .add_file_with_media_type("custom/cache.bin", b"cached", "application/octet-stream")
        .unwrap();
    let bytes = writer.finish_to_bytes().unwrap();

    let master = Package::from_bytes(&bytes).unwrap();
    assert_eq!(master.as_bytes(), bytes.as_slice());
    assert_eq!(master.document().content_xml(), CONTENT);
    assert!(master.document().content_xml().contains("Chapters/a.odt"));
    assert!(master.document().content_xml().contains("loext:opaque"));
    assert!(
        master
            .document()
            .files()
            .unwrap()
            .iter()
            .any(|path| path == "custom/cache.bin")
    );
}
