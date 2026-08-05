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
