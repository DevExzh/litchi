//! Admission and source checks for shared XML publication.
#![allow(clippy::unwrap_used, clippy::expect_used, reason = "test assertions")]

use std::sync::Arc;

use litchi_opc::{OpcPackage, PackURI, ReadLimits};

fn source_package() -> OpcPackage {
    let mut writer = soapberry_zip::office::StreamingArchiveWriter::new();
    writer.write_stored("[Content_Types].xml", br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>"#).unwrap();
    writer.write_stored("item.xml", b"<root/>").unwrap();
    let limits = ReadLimits::builder()
        .max_part_bytes(1024)
        .unwrap()
        .build()
        .unwrap();
    OpcPackage::from_vec_with_limits(writer.finish_to_bytes().unwrap(), limits).unwrap()
}

fn xml_of_size(size: usize) -> Arc<Vec<u8>> {
    let mut xml = b"<root><!--".to_vec();
    xml.resize(size - b"--></root>".len(), b'x');
    xml.extend_from_slice(b"--></root>");
    Arc::new(xml)
}

#[test]
fn shared_xml_add_checks_exact_caller_cap_and_retains_allocation() {
    let mut package = source_package();
    let name = PackURI::new("/added.xml").unwrap();
    assert!(
        package
            .try_add_owned_xml_part_bytes(name.clone(), "application/xml".into(), xml_of_size(1025))
            .is_err()
    );
    assert!(package.get_part(&name).is_err());
    let exact = xml_of_size(1024);
    package
        .try_add_owned_xml_part_bytes(name.clone(), "application/xml".into(), Arc::clone(&exact))
        .unwrap();
    assert!(Arc::ptr_eq(
        &package.get_part(&name).unwrap().blob_arc(),
        &exact
    ));
}

#[test]
fn shared_xml_replace_checks_source_and_limits_before_mutation() {
    let mut package = source_package();
    let name = PackURI::new("/item.xml").unwrap();
    let original = package.get_part(&name).unwrap().blob_arc();
    assert!(
        package
            .try_replace_owned_xml_part_bytes(&name, b"<stale/>", xml_of_size(1024))
            .is_err()
    );
    assert!(
        package
            .try_replace_owned_xml_part_bytes(&name, &original, xml_of_size(1025))
            .is_err()
    );
    assert!(
        package
            .try_replace_owned_xml_part_bytes(
                &name,
                &original,
                Arc::new(b"<root>&unknown;</root>".to_vec())
            )
            .is_err()
    );
    assert!(Arc::ptr_eq(
        &package.get_part(&name).unwrap().blob_arc(),
        &original
    ));
    let exact = xml_of_size(1024);
    package
        .try_replace_owned_xml_part_bytes(&name, &original, Arc::clone(&exact))
        .unwrap();
    assert!(Arc::ptr_eq(
        &package.get_part(&name).unwrap().blob_arc(),
        &exact
    ));
}
