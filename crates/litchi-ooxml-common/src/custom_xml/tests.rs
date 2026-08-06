//! Focused regression coverage for the Custom XML Data Storage model, codecs, and package graph.

use super::*;
use crate::Error;
use litchi_opc::part::BlobPart;
use litchi_opc::{OpcPackage, PackURI, Part};

const POI_XLSX: &[u8] =
    include_bytes!("../../../../test-data/poi/test-data/spreadsheet/customIndexedColors.xlsx");
const LO_DOCX: &[u8] = include_bytes!(
    "../../../../test-data/libreoffice-core/sw/qa/core/objectpositioning/data/do-not-capture-draw-objs-on-page-draw-wrap-none.docx"
);

#[test]
fn loads_poi_and_libreoffice_reference_fixtures() {
    let poi = OpcPackage::from_bytes(POI_XLSX).unwrap();
    let items = discover(&poi).unwrap();
    assert_eq!(items.len(), 1);
    assert_eq!(items[0].root().local_name, "easyPacket");
    assert!(items[0].props().unwrap().schemas.is_empty());

    let libreoffice = OpcPackage::from_bytes(LO_DOCX).unwrap();
    let items = discover(&libreoffice).unwrap();
    assert_eq!(items.len(), 1);
    assert_eq!(items[0].root().local_name, "Sources");
    assert_eq!(
        items[0].props().unwrap().schemas,
        ["http://schemas.openxmlformats.org/officeDocument/2006/bibliography"]
    );
}

#[test]
fn strict_writer_is_deterministic_and_round_trips() {
    let props = sample_props();
    let first = write_props(&props, Conformance::Strict).unwrap();
    let second = write_props(&props, Conformance::Strict).unwrap();
    assert_eq!(first, second);
    assert!(
        std::str::from_utf8(&first)
            .unwrap()
            .contains(STRICT_NAMESPACE)
    );
    assert_eq!(read_props(&first).unwrap(), props);
}

#[test]
fn attribute_whitespace_round_trips_without_normalization_loss() {
    let mut props = sample_props();
    props.schemas = vec!["urn:line\nnext\tlast\rreturn".into()];
    let xml = write_props(&props, Conformance::Transitional).unwrap();
    assert_eq!(read_props(&xml).unwrap(), props);
}

#[test]
fn mce_selects_fallback_schema_reference() {
    let xml = format!(
        r#"<ds:datastoreItem xmlns:ds="{TRANSITIONAL_NAMESPACE}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:unsupported" ds:itemID="{{11111111-1111-1111-1111-111111111111}}"><ds:schemaRefs><mc:AlternateContent><mc:Choice Requires="x"><ds:schemaRef ds:uri="urn:wrong"/></mc:Choice><mc:Fallback><ds:schemaRef ds:uri="urn:right"/></mc:Fallback></mc:AlternateContent></ds:schemaRefs></ds:datastoreItem>"#
    );
    assert_eq!(read_props(xml.as_bytes()).unwrap().schemas, ["urn:right"]);
}

#[test]
fn package_writer_round_trips_without_interpreting_payload() {
    let mut package = package_with_source();
    package.relate_to(
        "_xmlsignatures/origin.sigs",
        litchi_opc::constants::relationship_type::DIGITAL_SIGNATURE_ORIGIN,
    );
    assert!(package.is_signed());
    add(
        &mut package,
        NewItem {
            source: PackURI::new("/word/document.xml").unwrap(),
            rel_id: "rIdData".into(),
            part: PackURI::new("/customXml/item1.xml").unwrap(),
            content_type: "application/xml".into(),
            xml: b"<customer xmlns=\"urn:customer\" id=\"7\"/>".to_vec(),
            props: Some(NewProps {
                part: PackURI::new("/customXml/itemProps1.xml").unwrap(),
                rel_id: "rIdProps".into(),
                value: sample_props(),
            }),
            conformance: Conformance::Transitional,
        },
    )
    .unwrap();
    assert!(!package.is_signed());
    let items = discover(&package).unwrap();
    assert_eq!(items.len(), 1);
    assert_eq!(
        items[0].xml(),
        b"<customer xmlns=\"urn:customer\" id=\"7\"/>"
    );
    assert_eq!(items[0].props().unwrap(), &sample_props());
}

#[test]
fn failed_add_is_transactional() {
    let mut package = package_with_source();
    let source = PackURI::new("/word/document.xml").unwrap();
    let before_parts = package.part_count();
    let before_rels = package.get_part(&source).unwrap().rels().len();
    let same_part = PackURI::new("/customXml/item1.xml").unwrap();
    let error = add(
        &mut package,
        NewItem {
            source: source.clone(),
            rel_id: "rIdData".into(),
            part: same_part.clone(),
            content_type: "application/xml".into(),
            xml: b"<root/>".to_vec(),
            props: Some(NewProps {
                part: same_part,
                rel_id: "rIdProps".into(),
                value: sample_props(),
            }),
            conformance: Conformance::Transitional,
        },
    )
    .unwrap_err();
    assert!(matches!(error, Error::Opc(_)));
    assert_eq!(package.part_count(), before_parts);
    assert_eq!(package.get_part(&source).unwrap().rels().len(), before_rels);

    let error = add(
        &mut package,
        NewItem {
            source: source.clone(),
            rel_id: "1 invalid".into(),
            part: PackURI::new("/customXml/item2.xml").unwrap(),
            content_type: "application/xml".into(),
            xml: b"<root/>".to_vec(),
            props: None,
            conformance: Conformance::Transitional,
        },
    )
    .unwrap_err();
    assert!(matches!(error, Error::Relationship(_)));
    assert_eq!(package.part_count(), before_parts);
    assert_eq!(package.get_part(&source).unwrap().rels().len(), before_rels);
}

#[test]
fn rejects_malformed_properties_payloads_and_package_graphs() {
    assert!(read_props(br#"<!DOCTYPE x><x/>"#).is_err());
    let missing_id = format!(r#"<ds:datastoreItem xmlns:ds="{TRANSITIONAL_NAMESPACE}"/>"#);
    assert!(read_props(missing_id.as_bytes()).is_err());
    let duplicate_refs = format!(
        r#"<ds:datastoreItem xmlns:ds="{TRANSITIONAL_NAMESPACE}" ds:itemID="{{11111111-1111-1111-1111-111111111111}}"><ds:schemaRefs/><ds:schemaRefs/></ds:datastoreItem>"#
    );
    assert!(read_props(duplicate_refs.as_bytes()).is_err());
    assert!(validate_payload(br#"<!DOCTYPE x><x/>"#).is_err());
    assert!(validate_payload(b"<a><b></a>").is_err());
    assert!(validate_payload(b"&#32;<root/>").is_err());
    assert!(validate_payload(b"<![CDATA[ ]]><root/>").is_err());
    assert!(validate_payload(b"<root>&unknown;</root>").is_err());
    assert!(validate_payload(b"<root>&#x110000;</root>").is_err());
    assert!(validate_payload(b"<root>\0</root>").is_err());
    assert!(validate_payload(b"<1root/>").is_err());

    let mut package = OpcPackage::new();
    let mut source = BlobPart::new(
        PackURI::new("/word/document.xml").unwrap(),
        "application/xml".into(),
        b"<document/>".to_vec(),
    );
    source.rels_mut().add_relationship(
        TRANSITIONAL_RELATIONSHIP.into(),
        "https://example.invalid/data.xml".into(),
        "rId1".into(),
        true,
    );
    package.add_part(Box::new(source));
    assert!(discover(&package).is_err());
}

#[test]
fn enforces_guid_depth_size_and_content_type_caps() {
    let mut invalid_guid = sample_props();
    invalid_guid.id = "not-a-guid".into();
    assert!(write_props(&invalid_guid, Conformance::Transitional).is_err());

    let too_deep = format!(
        "{}<leaf/>{}",
        "<x>".repeat(MAX_DEPTH),
        "</x>".repeat(MAX_DEPTH)
    );
    let error = validate_payload(too_deep.as_bytes()).unwrap_err();
    assert!(matches!(
        error,
        Error::Limit {
            resource: "custom XML depth",
            ..
        }
    ));
    assert!(validate_payload(&vec![b' '; MAX_PART_BYTES + 1]).is_err());
    assert!(validate_content_type("not-a-media-type+xml").is_err());
    assert!(validate_content_type("application/vnd.example+xml").is_ok());
}

#[test]
fn rejects_invalid_declarations_attributes_and_xml_characters() {
    assert!(validate_payload(br#" <!--before--><?xml version="1.0"?><root/>"#).is_err());
    assert!(validate_payload(br#"<?xml version="1.0" bad="x"?><root/>"#).is_err());
    assert!(
        validate_payload(br#"<?xml version="1.0" standalone="yes" encoding="UTF-8"?><root/>"#)
            .is_err()
    );
    assert!(validate_payload(br#"<?xml version="1.0" standalone="maybe"?><root/>"#).is_err());
    assert!(validate_payload(br#"<p:root/>"#).is_err());
    assert!(validate_payload(br#"<root 1id="value"/>"#).is_err());
    assert!(
        validate_payload(br#"<root xmlns:a="urn:x" xmlns:b="urn:x" a:id="1" b:id="2"/>"#).is_err()
    );
    let mut props = sample_props();
    props.schemas = vec!["urn:\0bad".into()];
    assert!(write_props(&props, Conformance::Transitional).is_err());
}

#[test]
fn transaction_preserves_opaque_payload_and_relationship_topology() {
    let mut package = package_with_source();
    add(
        &mut package,
        NewItem {
            source: PackURI::new("/word/document.xml").unwrap(),
            rel_id: "rIdData".into(),
            part: PackURI::new("/customXml/item1.xml").unwrap(),
            content_type: "application/xml".into(),
            xml: br#"<customer xmlns="urn:customer"><future marker="keep"/></customer>"#.to_vec(),
            props: Some(NewProps {
                part: PackURI::new("/customXml/itemProps1.xml").unwrap(),
                rel_id: "rIdProps".into(),
                value: sample_props(),
            }),
            conformance: Conformance::Transitional,
        },
    )
    .unwrap();
    let data = PackURI::new("/customXml/item1.xml").unwrap();
    let props = PackURI::new("/customXml/itemProps1.xml").unwrap();
    let props_before = package.get_part(&props).unwrap().blob().to_vec();
    package
        .get_part_mut(&data)
        .unwrap()
        .rels_mut()
        .add_relationship(
            "urn:future:custom-xml".into(),
            "https://example.invalid/opaque".into(),
            "rIdFuture".into(),
            true,
        );

    let before = Snapshot::load(&package).unwrap();
    assert_eq!(before.items()[0].relationships().len(), 2);
    let mut transaction = Transaction::new(&mut package).unwrap();
    assert!(
        transaction
            .set_item_xml(
                0,
                br#"<customer xmlns="urn:customer"><future marker="changed"/></customer>"#.to_vec(),
            )
            .unwrap()
    );
    let commit = transaction.commit().unwrap();
    assert!(commit.changed());
    assert_eq!(package.get_part(&props).unwrap().blob(), props_before);
    assert_eq!(
        package
            .get_part(&data)
            .unwrap()
            .rels()
            .get("rIdFuture")
            .unwrap()
            .target_ref(),
        "https://example.invalid/opaque"
    );

    let patch = commit.patch().clone();
    assert!(patch.inverse().apply(&mut package).unwrap());
    assert_eq!(Snapshot::load(&package).unwrap(), before);
}

#[test]
fn properties_edits_retain_future_markup() {
    let mut package = package_with_source();
    add(
        &mut package,
        NewItem {
            source: PackURI::new("/word/document.xml").unwrap(),
            rel_id: "rIdData".into(),
            part: PackURI::new("/customXml/item1.xml").unwrap(),
            content_type: "application/xml".into(),
            xml: b"<root/>".to_vec(),
            props: Some(NewProps {
                part: PackURI::new("/customXml/itemProps1.xml").unwrap(),
                rel_id: "rIdProps".into(),
                value: sample_props(),
            }),
            conformance: Conformance::Transitional,
        },
    )
    .unwrap();
    let props_part = PackURI::new("/customXml/itemProps1.xml").unwrap();
    let source = format!(
        r#"<ds:datastoreItem xmlns:ds="{TRANSITIONAL_NAMESPACE}" xmlns:x="urn:future" x:marker="keep" ds:itemID="{{11111111-1111-1111-1111-111111111111}}"><x:future>opaque</x:future><ds:schemaRefs><ds:schemaRef ds:uri="urn:customer"/></ds:schemaRefs></ds:datastoreItem>"#
    );
    package
        .get_part_mut(&props_part)
        .unwrap()
        .set_blob(source.into_bytes());
    let mut transaction = Transaction::new(&mut package).unwrap();
    let mut updated = sample_props();
    updated.id = "{22222222-2222-2222-2222-222222222222}".into();
    assert!(transaction.set_properties(0, updated).unwrap());
    transaction.commit().unwrap();
    let xml = std::str::from_utf8(package.get_part(&props_part).unwrap().blob()).unwrap();
    assert!(xml.contains("x:marker=\"keep\""));
    assert!(xml.contains("<x:future>opaque</x:future>"));
    assert!(xml.contains("22222222-2222-2222-2222-222222222222"));
}

#[test]
fn transaction_crud_is_source_checked_and_failure_atomic() {
    let mut package = package_with_source();
    let source = PackURI::new("/word/document.xml").unwrap();
    let initial_parts = package.part_count();
    let initial_relationships = package.get_part(&source).unwrap().rels().len();
    let mut transaction = Transaction::new(&mut package).unwrap();
    let index = transaction
        .insert(NewItem {
            source: source.clone(),
            rel_id: "rIdInserted".into(),
            part: PackURI::new("/customXml/item2.xml").unwrap(),
            content_type: "application/xml".into(),
            xml: b"<opaque xmlns=\"urn:opaque\"><x:future xmlns:x=\"urn:x\"/></opaque>".to_vec(),
            props: Some(NewProps {
                part: PackURI::new("/customXml/itemProps2.xml").unwrap(),
                rel_id: "rIdProps".into(),
                value: sample_props(),
            }),
            conformance: Conformance::Strict,
        })
        .unwrap();
    assert_eq!(index, 0);
    let commit = transaction.commit().unwrap();
    assert_eq!(discover(&package).unwrap().len(), 1);
    let patch = commit.patch().clone();
    assert!(patch.inverse().apply(&mut package).unwrap());
    assert!(discover(&package).unwrap().is_empty());
    assert_eq!(package.part_count(), initial_parts);
    assert_eq!(
        package.get_part(&source).unwrap().rels().len(),
        initial_relationships
    );

    let mut stale = package_with_source();
    let mut source_edit = Transaction::new(&mut stale).unwrap();
    source_edit
        .insert(NewItem {
            source: source.clone(),
            rel_id: "rIdStale".into(),
            part: PackURI::new("/customXml/stale.xml").unwrap(),
            content_type: "application/xml".into(),
            xml: b"<stale/>".to_vec(),
            props: None,
            conformance: Conformance::Transitional,
        })
        .unwrap();
    let stale_patch = source_edit.commit().unwrap().patch().clone();
    package
        .get_part_mut(&source)
        .unwrap()
        .set_blob(b"<changed/>".to_vec());
    let before_parts = package.part_count();
    assert!(stale_patch.apply(&mut package).is_err());
    assert_eq!(package.part_count(), before_parts);
    assert_eq!(package.get_part(&source).unwrap().blob(), b"<changed/>");
}

#[test]
fn exact_noop_does_not_unsign_or_rewrite() {
    let mut package = package_with_source();
    add(
        &mut package,
        NewItem {
            source: PackURI::new("/word/document.xml").unwrap(),
            rel_id: "rIdData".into(),
            part: PackURI::new("/customXml/item1.xml").unwrap(),
            content_type: "application/xml".into(),
            xml: b"<root/>".to_vec(),
            props: None,
            conformance: Conformance::Transitional,
        },
    )
    .unwrap();
    package.relate_to(
        "_xmlsignatures/origin.sigs",
        litchi_opc::constants::relationship_type::DIGITAL_SIGNATURE_ORIGIN,
    );
    let before = Snapshot::load(&package).unwrap();
    let mut transaction = Transaction::new(&mut package).unwrap();
    assert!(!transaction.set_xml(0, b"<root/>".to_vec()).unwrap());
    let commit = transaction.commit().unwrap();
    assert!(!commit.changed());
    assert!(commit.patch().is_empty());
    assert!(package.is_signed());
    assert_eq!(Snapshot::load(&package).unwrap(), before);
}

fn package_with_source() -> OpcPackage {
    let mut package = OpcPackage::new();
    package.add_part(Box::new(BlobPart::new(
        PackURI::new("/word/document.xml").unwrap(),
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml".into(),
        b"<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"/>"
            .to_vec(),
    )));
    package
}

fn sample_props() -> Props {
    Props {
        id: "{11111111-1111-1111-1111-111111111111}".into(),
        schemas: vec!["urn:customer".into()],
    }
}
