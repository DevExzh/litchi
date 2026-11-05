//! Admission and source checks for shared XML publication.
#![allow(clippy::unwrap_used, clippy::expect_used, reason = "test assertions")]

use std::sync::Arc;

use litchi_opc::constants::relationship_type;
use litchi_opc::{OpcError, OpcPackage, PackURI, SourceLimitKind, SourceLimits};

fn source_package() -> OpcPackage {
    let mut writer = soapberry_zip::office::StreamingArchiveWriter::new(
        soapberry_zip::office::StreamingArchiveLimits::default(),
    );
    writer.write_stored("[Content_Types].xml", br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>"#).unwrap();
    writer.write_stored("item.xml", b"<root/>").unwrap();
    let limits = SourceLimits::builder()
        .max_part_bytes(1024)
        .unwrap()
        .build()
        .unwrap();
    OpcPackage::from_vec(writer.finish_to_bytes().unwrap(), limits).unwrap()
}

fn case_alias_source_package() -> OpcPackage {
    let mut writer = soapberry_zip::office::StreamingArchiveWriter::new(
        soapberry_zip::office::StreamingArchiveLimits::default(),
    );
    writer
        .write_stored(
            "[Content_Types].xml",
            br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>"#,
        )
        .unwrap();
    writer
        .write_stored("CUSTOMUI.XML", b"<root>\n  <old/>\n</root>")
        .unwrap();
    OpcPackage::from_vec(writer.finish_to_bytes().unwrap(), SourceLimits::default()).unwrap()
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

#[test]
fn authored_xml_add_checks_attribute_part_and_aggregate_limits_atomically() {
    let name = PackURI::new("/added.xml").unwrap();
    let authored = Arc::new(b"<root a=\"1234\"/>".to_vec());

    let mut package = source_package();
    let item = PackURI::new("/item.xml").unwrap();
    let before = package.get_part(&item).unwrap().blob_arc();
    // `SourceLimitsBuilder::build` keeps relationship targets within the XML
    // attribute ceiling, so lowering the attribute ceiling lowers both.
    let limits = SourceLimits::builder()
        .max_xml_attribute_bytes(3)
        .unwrap()
        .max_relationship_target_bytes(3)
        .unwrap()
        .build()
        .unwrap();
    assert!(matches!(
        package.try_add_authored_xml_part_bytes(
            name.clone(),
            "application/xml".into(),
            Arc::clone(&authored),
            limits,
        ),
        Err(OpcError::Limit {
            kind: SourceLimitKind::XmlAttributeBytes,
            ..
        })
    ));
    assert!(package.get_part(&name).is_err());
    assert!(Arc::ptr_eq(
        &package.get_part(&item).unwrap().blob_arc(),
        &before
    ));

    let mut package = source_package();
    let limits = SourceLimits::builder()
        .max_parts(1)
        .unwrap()
        .build()
        .unwrap();
    assert!(matches!(
        package.try_add_authored_xml_part_bytes(
            name.clone(),
            "application/xml".into(),
            Arc::new(b"<root/>".to_vec()),
            limits,
        ),
        Err(OpcError::Limit {
            kind: SourceLimitKind::Parts,
            ..
        })
    ));
    assert_eq!(package.part_count(), 1);

    let mut package = source_package();
    let limits = SourceLimits::builder()
        .max_total_part_bytes(7)
        .unwrap()
        .build()
        .unwrap();
    assert!(matches!(
        package.try_add_authored_xml_part_bytes(
            name,
            "application/xml".into(),
            Arc::new(b"<root/>".to_vec()),
            limits,
        ),
        Err(OpcError::Limit {
            kind: SourceLimitKind::TotalPartBytes,
            ..
        })
    ));
    assert_eq!(package.part_count(), 1);
}

#[test]
fn authored_xml_replacement_is_stale_safe_and_signed_noop_preserving() {
    let mut package = source_package();
    let name = PackURI::new("/item.xml").unwrap();
    let original = package.get_part(&name).unwrap().blob_arc();

    assert!(matches!(
        package.try_replace_authored_xml_part_bytes(
            &name,
            b"<stale/>",
            Arc::new(b"<root changed=\"1\"/>".to_vec()),
            SourceLimits::default(),
        ),
        Err(OpcError::XmlError(message)) if message.contains("stale")
    ));
    assert!(Arc::ptr_eq(
        &package.get_part(&name).unwrap().blob_arc(),
        &original
    ));

    assert!(
        package
            .try_replace_authored_xml_part_bytes(
                &name,
                &original,
                Arc::new(b"<root>&unknown;</root>".to_vec()),
                SourceLimits::default(),
            )
            .is_err()
    );
    assert!(Arc::ptr_eq(
        &package.get_part(&name).unwrap().blob_arc(),
        &original
    ));

    package.relate_to(
        "_xmlsignatures/origin.sigs",
        relationship_type::DIGITAL_SIGNATURE_ORIGIN,
    );
    assert!(package.is_signed());
    package
        .try_replace_authored_xml_part_bytes(
            &name,
            &original,
            Arc::clone(&original),
            SourceLimits::default(),
        )
        .unwrap();
    assert!(package.is_signed());
    assert!(Arc::ptr_eq(
        &package.get_part(&name).unwrap().blob_arc(),
        &original
    ));
}

#[test]
fn xml_replacement_provenance_uses_canonical_name_for_case_aliases() {
    let alias = PackURI::new("/customui.xml").unwrap();
    let canonical = PackURI::new("/CUSTOMUI.XML").unwrap();
    let replacement = Arc::new(b"<root>\n  <new/>\n</root>".to_vec());

    let mut authored = case_alias_source_package();
    let original = authored.get_part(&alias).unwrap().blob_arc();
    authored
        .try_replace_authored_xml_part_bytes(
            &alias,
            &original,
            Arc::clone(&replacement),
            SourceLimits::default(),
        )
        .unwrap();
    assert_eq!(
        authored.source_xml_part(&canonical).unwrap().bytes(),
        replacement.as_slice()
    );

    let mut strict = case_alias_source_package();
    let original = strict.get_part(&alias).unwrap().blob_arc();
    strict
        .try_replace_owned_xml_part_bytes(&alias, &original, Arc::clone(&replacement))
        .unwrap();
    assert_eq!(
        strict.source_xml_part(&canonical).unwrap().bytes(),
        replacement.as_slice()
    );
}
