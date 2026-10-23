//! Record 0290 F6: the metadata edit of a package without `meta.xml` (OTH and
//! ODM leaf builders write none) creates the part instead of refusing.

use litchi_odt::Package;
use litchi_odt::core::{OwnedPackage, PackageWriter};
use litchi_odt::metadata::Field;

const MIMETYPE: &str = "application/vnd.oasis.opendocument.text";

fn package_without_meta() -> Package {
    let mut writer = PackageWriter::new(litchi_odf_common::core::StreamingArchiveLimits::default());
    writer.set_mimetype(MIMETYPE).unwrap();
    writer
        .add_file(
            "content.xml",
            concat!(
                "<office:document-content xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" ",
                "xmlns:text=\"urn:oasis:names:tc:opendocument:xmlns:text:1.0\" office:version=\"1.2\">",
                "<office:body><office:text><text:p>Body</text:p></office:text></office:body>",
                "</office:document-content>"
            )
            .as_bytes(),
        )
        .unwrap();
    let bytes = writer.finish_to_bytes().unwrap();
    assert!(
        !OwnedPackage::from_bytes(bytes.clone())
            .unwrap()
            .has_file("meta.xml")
            .unwrap()
    );
    Package::from_bytes(&bytes).unwrap()
}

#[test]
fn metadata_set_creates_meta_xml() {
    let source = package_without_meta();
    let mut edit = source.edit_metadata();
    edit.set(Field::Title, Some("Created"))
        .unwrap()
        .set(Field::Language, Some("de-DE"))
        .unwrap();
    let commit = edit.commit().unwrap();
    assert!(commit.diagnostics().changed());
    let next = commit.into_package();
    let metadata = next.document().unwrap().metadata().unwrap();
    assert_eq!(metadata.title.as_deref(), Some("Created"));
    let owned = OwnedPackage::from_bytes(next.to_bytes().unwrap()).unwrap();
    let meta = String::from_utf8(owned.get_file("meta.xml").unwrap()).unwrap();
    assert!(meta.contains("office:version=\"1.2\""), "{meta}");
    assert!(meta.contains("<dc:language>de-DE</dc:language>"), "{meta}");
    // The body is untouched.
    assert_eq!(
        owned.get_file("content.xml").unwrap(),
        OwnedPackage::from_bytes(source.to_bytes().unwrap())
            .unwrap()
            .get_file("content.xml")
            .unwrap()
    );

    // A second edit splices the created part.
    let mut second = next.edit_metadata();
    second.set(Field::Title, Some("Again")).unwrap();
    let again = second.commit().unwrap().into_package();
    assert_eq!(
        again
            .document()
            .unwrap()
            .metadata()
            .unwrap()
            .title
            .as_deref(),
        Some("Again")
    );
}

#[test]
fn metadata_removal_without_meta_xml_is_a_no_op() {
    let source = package_without_meta();
    let mut edit = source.edit_metadata();
    edit.set(Field::Title, None).unwrap();
    let commit = edit.commit().unwrap();
    assert!(!commit.diagnostics().changed());
    assert_eq!(commit.package().as_bytes(), source.as_bytes());
}
