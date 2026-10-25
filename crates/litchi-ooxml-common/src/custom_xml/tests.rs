//! Focused regression coverage for the Custom XML Data Storage model, codecs, and package graph.

use super::*;
use crate::Error;
use litchi_opc::part::BlobPart;
use litchi_opc::{OpcPackage, PackURI, Part};

fn source_limits() -> litchi_opc::SourceLimits {
    litchi_opc::SourceLimits::default()
}

const POI_XLSX: &[u8] =
    include_bytes!("../../../../test-data/poi/test-data/spreadsheet/customIndexedColors.xlsx");
const LO_DOCX: &[u8] = include_bytes!(
    "../../../../test-data/libreoffice-core/sw/qa/core/objectpositioning/data/do-not-capture-draw-objs-on-page-draw-wrap-none.docx"
);

#[test]
fn loads_poi_and_libreoffice_reference_fixtures() {
    let poi = OpcPackage::from_bytes(POI_XLSX, litchi_opc::SourceLimits::default()).unwrap();
    let items = discover(&poi).unwrap();
    assert_eq!(items.len(), 1);
    assert_eq!(items[0].root().local_name, "easyPacket");
    assert!(items[0].props().unwrap().schemas.is_empty());

    let libreoffice = OpcPackage::from_bytes(LO_DOCX, litchi_opc::SourceLimits::default()).unwrap();
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
    let first = write_props(&props, Conformance::Strict, &source_limits()).unwrap();
    let second = write_props(&props, Conformance::Strict, &source_limits()).unwrap();
    assert_eq!(first, second);
    assert!(
        std::str::from_utf8(&first)
            .unwrap()
            .contains(STRICT_NAMESPACE)
    );
    assert_eq!(read_props(&first, &source_limits()).unwrap(), props);
}

#[test]
fn properties_codec_refuses_when_the_retained_part_budget_is_tight() {
    let limits = litchi_opc::SourceLimits::builder()
        .max_part_bytes(64)
        .unwrap()
        .build()
        .unwrap();
    let props = sample_props();
    let error = write_props(&props, Conformance::Strict, &limits).unwrap_err();
    assert!(matches!(
        error,
        Error::Limit {
            kind: crate::LimitKind::CustomXml(LimitKind::SerializedPropertiesBytes),
            observed,
            maximum: 64,
        } if observed > 64
    ));

    let xml = write_props(&props, Conformance::Strict, &source_limits()).unwrap();
    let error = read_props(&xml, &limits).unwrap_err();
    assert!(matches!(
        error,
        Error::Limit {
            kind: crate::LimitKind::CustomXml(LimitKind::PropertiesBytes),
            observed,
            maximum: 64,
        } if observed > 64
    ));

    let depth_limited = litchi_opc::SourceLimits::builder()
        .max_depth(1)
        .unwrap()
        .build()
        .unwrap();
    let error = read_props(&xml, &depth_limited).unwrap_err();
    assert!(matches!(
        error,
        Error::Limit {
            kind: crate::LimitKind::CustomXml(LimitKind::Depth),
            maximum: 1,
            ..
        }
    ));
}

#[test]
fn attribute_whitespace_round_trips_without_normalization_loss() {
    let mut props = sample_props();
    props.schemas = vec!["urn:line\nnext\tlast\rreturn".into()];
    let xml = write_props(&props, Conformance::Transitional, &source_limits()).unwrap();
    assert_eq!(read_props(&xml, &source_limits()).unwrap(), props);
}

#[test]
fn mce_selects_fallback_schema_reference() {
    let xml = format!(
        r#"<ds:datastoreItem xmlns:ds="{TRANSITIONAL_NAMESPACE}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:unsupported" ds:itemID="{{11111111-1111-1111-1111-111111111111}}"><ds:schemaRefs><mc:AlternateContent><mc:Choice Requires="x"><ds:schemaRef ds:uri="urn:wrong"/></mc:Choice><mc:Fallback><ds:schemaRef ds:uri="urn:right"/></mc:Fallback></mc:AlternateContent></ds:schemaRefs></ds:datastoreItem>"#
    );
    assert_eq!(
        read_props(xml.as_bytes(), &source_limits())
            .unwrap()
            .schemas,
        ["urn:right"]
    );
}

#[test]
fn package_writer_round_trips_without_interpreting_payload() {
    let mut package = package_with_source();
    package.relate_to(
        "_xmlsignatures/origin.sigs",
        litchi_opc::constants::relationship_type::DIGITAL_SIGNATURE_ORIGIN,
    );
    assert!(package.is_signed());
    let item = || NewItem {
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
    };
    // Adding an item changes the signed package, so it is refused before any
    // part is inserted until the signature is removed explicitly (ADR 0006).
    let parts = package.part_count();
    assert!(matches!(
        add(&mut package, item()),
        Err(Error::Opc(
            litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy
        ))
    ));
    assert!(package.is_signed());
    assert_eq!(package.part_count(), parts);
    package.unsign();
    add(&mut package, item()).unwrap();
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
    assert!(read_props(br#"<!DOCTYPE x><x/>"#, &source_limits()).is_err());
    let missing_id = format!(r#"<ds:datastoreItem xmlns:ds="{TRANSITIONAL_NAMESPACE}"/>"#);
    assert!(read_props(missing_id.as_bytes(), &source_limits()).is_err());
    let duplicate_refs = format!(
        r#"<ds:datastoreItem xmlns:ds="{TRANSITIONAL_NAMESPACE}" ds:itemID="{{11111111-1111-1111-1111-111111111111}}"><ds:schemaRefs/><ds:schemaRefs/></ds:datastoreItem>"#
    );
    assert!(read_props(duplicate_refs.as_bytes(), &source_limits()).is_err());
    assert!(validate_payload(br#"<!DOCTYPE x><x/>"#, &source_limits()).is_err());
    assert!(validate_payload(b"<a><b></a>", &source_limits()).is_err());
    assert!(validate_payload(b"&#32;<root/>", &source_limits()).is_err());
    assert!(validate_payload(b"<![CDATA[ ]]><root/>", &source_limits()).is_err());
    assert!(validate_payload(b"<root>&unknown;</root>", &source_limits()).is_err());
    assert!(validate_payload(b"<root>&#x110000;</root>", &source_limits()).is_err());
    assert!(validate_payload(b"<root>\0</root>", &source_limits()).is_err());
    assert!(validate_payload(b"<1root/>", &source_limits()).is_err());

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
    assert!(write_props(&invalid_guid, Conformance::Transitional, &source_limits()).is_err());

    let too_deep = format!(
        "{}<leaf/>{}",
        "<x>".repeat(MAX_DEPTH),
        "</x>".repeat(MAX_DEPTH)
    );
    let error = validate_payload(too_deep.as_bytes(), &source_limits()).unwrap_err();
    assert!(matches!(
        error,
        Error::Limit {
            kind: crate::LimitKind::CustomXml(LimitKind::Depth),
            ..
        }
    ));
    assert!(validate_payload(&vec![b' '; MAX_PART_BYTES + 1], &source_limits()).is_err());
    assert!(validate_content_type("not-a-media-type+xml").is_err());
    assert!(validate_content_type("application/vnd.example+xml").is_ok());
}

#[test]
fn payload_caps_follow_the_retained_source_profile() {
    let payload = b"<root><child/></root>";
    let tight_bytes = litchi_opc::SourceLimits::builder()
        .max_part_bytes((payload.len() - 1) as u64)
        .unwrap()
        .build()
        .unwrap();
    assert!(matches!(
        validate_payload(payload, &tight_bytes),
        Err(Error::Limit {
            kind: crate::LimitKind::CustomXml(LimitKind::PayloadBytes),
            ..
        })
    ));

    let tight_depth = litchi_opc::SourceLimits::builder()
        .max_depth(1)
        .unwrap()
        .build()
        .unwrap();
    assert!(matches!(
        validate_payload(payload, &tight_depth),
        Err(Error::Limit {
            kind: crate::LimitKind::CustomXml(LimitKind::Depth),
            ..
        })
    ));

    let tight_events = litchi_opc::SourceLimits::builder()
        .max_events(1)
        .unwrap()
        .build()
        .unwrap();
    assert!(matches!(
        validate_payload(payload, &tight_events),
        Err(Error::Limit {
            kind: crate::LimitKind::CustomXml(LimitKind::Elements),
            ..
        })
    ));
}

#[test]
fn properties_writer_checks_before_allocating_and_noop_retain_budget() {
    let props = sample_props();
    let output = write_props(&props, Conformance::Strict, &source_limits()).unwrap();
    let tight = litchi_opc::SourceLimits::builder()
        .max_part_bytes((output.len() - 1) as u64)
        .unwrap()
        .build()
        .unwrap();
    assert!(matches!(
        write_props(&props, Conformance::Strict, &tight),
        Err(Error::Limit {
            kind: crate::LimitKind::CustomXml(LimitKind::SerializedPropertiesBytes),
            ..
        })
    ));
    assert!(matches!(
        codec::rewrite_props(&output, &props, &props, Conformance::Strict, &tight,),
        Err(Error::Limit {
            kind: crate::LimitKind::CustomXml(LimitKind::SerializedPropertiesBytes),
            ..
        })
    ));
}

#[test]
fn rejects_invalid_declarations_attributes_and_xml_characters() {
    assert!(
        validate_payload(
            br#" <!--before--><?xml version="1.0"?><root/>"#,
            &source_limits()
        )
        .is_err()
    );
    assert!(
        validate_payload(br#"<?xml version="1.0" bad="x"?><root/>"#, &source_limits()).is_err()
    );
    assert!(
        validate_payload(
            br#"<?xml version="1.0" standalone="yes" encoding="UTF-8"?><root/>"#,
            &source_limits()
        )
        .is_err()
    );
    assert!(
        validate_payload(
            br#"<?xml version="1.0" standalone="maybe"?><root/>"#,
            &source_limits()
        )
        .is_err()
    );
    assert!(validate_payload(br#"<p:root/>"#, &source_limits()).is_err());
    assert!(validate_payload(br#"<root 1id="value"/>"#, &source_limits()).is_err());
    assert!(
        validate_payload(
            br#"<root xmlns:a="urn:x" xmlns:b="urn:x" a:id="1" b:id="2"/>"#,
            &source_limits()
        )
        .is_err()
    );
    let mut props = sample_props();
    props.schemas = vec!["urn:\0bad".into()];
    assert!(write_props(&props, Conformance::Transitional, &source_limits()).is_err());
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
    let mut transaction = Edit::new(&package).unwrap();
    assert!(
        transaction
            .set_xml(
                0,
                br#"<customer xmlns="urn:customer"><future marker="changed"/></customer>"#.to_vec(),
            )
            .unwrap()
    );
    let commit = transaction.commit().unwrap();
    assert!(commit.diagnostics().changed());
    let package = commit.package().clone();
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
    let restored = patch.inverse().apply(&package).unwrap();
    assert!(restored.diagnostics().changed());
    assert_eq!(Snapshot::load(restored.package()).unwrap(), before);
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
    let mut transaction = Edit::new(&package).unwrap();
    let mut updated = sample_props();
    updated.id = "{22222222-2222-2222-2222-222222222222}".into();
    assert!(transaction.set_properties(0, updated).unwrap());
    let package = transaction.commit().unwrap().into_package();
    let xml = std::str::from_utf8(package.get_part(&props_part).unwrap().blob()).unwrap();
    assert!(xml.contains("x:marker=\"keep\""));
    assert!(xml.contains("<x:future>opaque</x:future>"));
    assert!(xml.contains("22222222-2222-2222-2222-222222222222"));
}

#[test]
fn transaction_crud_is_source_checked_and_failure_atomic() {
    let original = package_with_source();
    let source = PackURI::new("/word/document.xml").unwrap();
    let initial_parts = original.part_count();
    let initial_relationships = original.get_part(&source).unwrap().rels().len();
    let mut transaction = Edit::new(&original).unwrap();
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
    assert!(discover(&original).unwrap().is_empty());
    assert_eq!(discover(commit.package()).unwrap().len(), 1);
    let patch = commit.patch().clone();
    let restored = patch.inverse().apply(commit.package()).unwrap();
    assert!(restored.diagnostics().changed());
    let mut package = restored.into_package();
    assert!(discover(&package).unwrap().is_empty());
    assert_eq!(package.part_count(), initial_parts);
    assert_eq!(
        package.get_part(&source).unwrap().rels().len(),
        initial_relationships
    );

    let stale = package_with_source();
    let mut source_edit = Edit::new(&stale).unwrap();
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
    assert!(stale_patch.apply(&package).is_err());
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
    let mut transaction = Edit::new(&package).unwrap();
    assert!(!transaction.set_xml(0, b"<root/>".to_vec()).unwrap());
    let commit = transaction.commit().unwrap();
    assert!(!commit.diagnostics().changed());
    assert!(commit.patch().is_noop());
    assert!(commit.package().is_signed());
    assert_eq!(Snapshot::load(commit.package()).unwrap(), before);
    assert!(package.is_signed());
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

#[test]
fn changed_transactions_and_patches_refuse_a_signed_package() {
    let source = PackURI::new("/word/document.xml").unwrap();
    let item = || NewItem {
        source: source.clone(),
        rel_id: "rIdData".into(),
        part: PackURI::new("/customXml/item1.xml").unwrap(),
        content_type: "application/xml".into(),
        xml: b"<root/>".to_vec(),
        props: None,
        conformance: Conformance::Transitional,
    };
    let mut package = package_with_source();
    package.relate_to(
        "_xmlsignatures/origin.sigs",
        litchi_opc::constants::relationship_type::DIGITAL_SIGNATURE_ORIGIN,
    );
    assert!(package.is_signed());
    let patch = {
        let mut unsigned = package.clone();
        unsigned.unsign();
        let mut edit = Edit::new(&unsigned).unwrap();
        edit.insert(item()).unwrap();
        edit.commit().unwrap().patch().clone()
    };
    let parts = package.part_count();

    // An unchanged transaction publishes nothing and keeps the signature.
    let unchanged = Edit::new(&package).unwrap().commit().unwrap();
    assert!(!unchanged.diagnostics().changed());
    assert!(unchanged.package().is_signed());
    assert!(
        unchanged
            .patch()
            .inverse()
            .apply(&package)
            .is_ok_and(|commit| !commit.diagnostics().changed())
    );

    // A changed transaction and a changed patch are refused (ADR 0006).
    let mut edit = Edit::new(&package).unwrap();
    edit.insert(item()).unwrap();
    assert!(matches!(
        edit.commit(),
        Err(Error::Opc(
            litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy
        ))
    ));
    assert!(matches!(
        patch.apply(&package),
        Err(Error::Opc(
            litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy
        ))
    ));
    assert!(package.is_signed());
    assert_eq!(package.part_count(), parts);

    package.unsign();
    let commit = patch.apply(&package).unwrap();
    assert!(commit.diagnostics().changed());
    assert_eq!(discover(commit.package()).unwrap().len(), 1);
    assert!(discover(&package).unwrap().is_empty());
}

#[test]
fn reset_discards_staged_changes_and_the_source_stays_unchanged() {
    let package = package_with_source();
    let mut edit = Edit::new(&package).unwrap();
    edit.insert(NewItem {
        source: PackURI::new("/word/document.xml").unwrap(),
        rel_id: "rIdData".into(),
        part: PackURI::new("/customXml/item1.xml").unwrap(),
        content_type: "application/xml".into(),
        xml: b"<root/>".to_vec(),
        props: None,
        conformance: Conformance::Transitional,
    })
    .unwrap();
    assert_eq!(edit.items().len(), 1);
    edit.reset();
    assert!(edit.items().is_empty());
    let commit = edit.commit().unwrap();
    assert!(!commit.diagnostics().changed());
    assert!(commit.patch().is_noop());
    assert!(discover(&package).unwrap().is_empty());
}
