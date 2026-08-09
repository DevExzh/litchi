use litchi_docx::Package;
use litchi_docx::content_control::{BindingFlavor, Checksum, Limits, Snapshot};
use litchi_docx::custom_xml::NewStore;
use litchi_ooxml_common::custom_xml::{Conformance, TRANSITIONAL_RELATIONSHIP};
use litchi_opc::constants::content_type as ct;
use litchi_opc::constants::relationship_type as rt;
use litchi_opc::packuri::PackURI;
use litchi_opc::part::{BlobPart, Part};

const ITEM_A: &str = "{11111111-1111-4111-8111-111111111111}";
const ITEM_B: &str = "{22222222-2222-4222-8222-222222222222}";
const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const HASH: &str = "http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash";
const MCE: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";

fn store(id: &str, value: &str) -> NewStore {
    NewStore {
        xml: format!(r#"<root xmlns="urn:test"><value>{value}</value></root>"#).into_bytes(),
        content_type: "application/xml".to_string(),
        id: id.to_string(),
        schemas: vec!["urn:test:schema".to_string()],
        conformance: Conformance::Transitional,
    }
}

fn bound_story(root: &str, body: &str, id: &str, checksum: &str) -> Vec<u8> {
    let body_end = if body.is_empty() { "" } else { "</w:body>" };
    format!(
        r#"<{root} xmlns:w="{W}" xmlns:mc="{MCE}" xmlns:h="{HASH}" mc:Ignorable="h">{body}<w:sdt><w:sdtPr><w:dataBinding w:xpath="/root/value" w:storeItemID="{id}" h:storeItemChecksum="{checksum}"/></w:sdtPr><w:sdtContent/></w:sdt>{body_end}</{root}>"#
    )
    .into_bytes()
}

fn mark_signed(package: &mut Package) {
    package
        .edit_opc(|opc| {
            opc.try_add_part(Box::new(BlobPart::new(
                PackURI::new("/_xmlsignatures/origin.sigs").unwrap(),
                ct::OPC_DIGITAL_SIGNATURE_ORIGIN.to_owned(),
                Vec::new(),
            )))?;
            opc.rels_mut().add_relationship(
                rt::DIGITAL_SIGNATURE_ORIGIN.to_owned(),
                "_xmlsignatures/origin.sigs".to_owned(),
                "rSignature".to_owned(),
                false,
            );
            Ok(())
        })
        .unwrap();
}

fn checksum_at(package: &Package, part: &str) -> Checksum {
    let part = package
        .opc_package()
        .get_part(&PackURI::new(part).unwrap())
        .unwrap();
    let snapshot = Snapshot::from_xml(part.blob().to_vec()).unwrap();
    snapshot.inventory().occurrences()[0]
        .control()
        .data_binding()
        .unwrap()
        .checksum()
        .unwrap()
        .clone()
}

fn main_relationships(package: &Package) -> Vec<(String, String, String, bool)> {
    let document = package
        .opc_package()
        .main_document_part()
        .unwrap()
        .partname()
        .clone();
    let mut relationships = package
        .opc_package()
        .get_part(&document)
        .unwrap()
        .rels()
        .iter()
        .map(|relationship| {
            (
                relationship.r_id().to_owned(),
                relationship.reltype().to_owned(),
                relationship.target_ref().to_owned(),
                relationship.is_external(),
            )
        })
        .collect::<Vec<_>>();
    relationships.sort_unstable();
    relationships
}

#[test]
fn generated_add_find_update_replace_reorder_remove_round_trip() {
    let directory = tempfile::tempdir().unwrap();
    let path = directory.path().join("custom-xml.docx");
    let mut package = Package::new().unwrap();
    package.add_custom_xml(store(ITEM_A, "a")).unwrap();
    package.add_custom_xml(store(ITEM_B, "b")).unwrap();
    assert_eq!(package.custom_xml().unwrap().len(), 2);

    package
        .set_custom_xml(ITEM_A, b"<updated/>".to_vec())
        .unwrap();
    let mut replacement = store(ITEM_B, "replacement");
    replacement.content_type = "application/vnd.example.data+xml".to_string();
    replacement.schemas.push("urn:second".to_string());
    package.replace_custom_xml(ITEM_B, replacement).unwrap();
    package
        .order_custom_xml(&[ITEM_B.to_string(), ITEM_A.to_string()])
        .unwrap();
    let items = package.custom_xml().unwrap();
    assert_eq!(items[0].props().unwrap().id, ITEM_B);
    assert_eq!(items[0].content_type(), "application/vnd.example.data+xml");
    assert_eq!(items[1].xml(), b"<updated/>");
    package.save(&path).unwrap();

    let mut reopened = Package::open(&path).unwrap();
    assert_eq!(reopened.custom_xml().unwrap().len(), 2);
    assert!(reopened.remove_custom_xml(ITEM_A).unwrap());
    assert!(reopened.custom_xml_by_id(ITEM_A).unwrap().is_none());
    assert!(!reopened.remove_custom_xml(ITEM_A).unwrap());
}

#[test]
fn binding_integrity_scans_word_containers_without_executing_xpath() {
    let mut package = Package::new().unwrap();
    package.add_custom_xml(store(ITEM_A, "a")).unwrap();
    let header_xml = format!(
        r#"<w:hdr xmlns:w="{W}"><w:sdt><w:sdtPr><w:id w:val="17"/><w:dataBinding w:prefixMappings="xmlns:x='urn:test'" w:xpath="/x:root/x:value" w:storeItemID="{ITEM_A}"/></w:sdtPr><w:sdtContent/></w:sdt></w:hdr>"#
    );
    package
        .edit_opc(|opc| {
            let document = opc.main_document_part()?.partname().clone();
            let header = PackURI::new("/word/header42.xml").unwrap();
            opc.add_part(Box::new(BlobPart::new(
                header.clone(),
                ct::WML_HEADER.to_string(),
                header_xml.into_bytes(),
            )));
            let target = header.relative_ref(document.base_uri());
            opc.get_part_mut(&document)?.rels_mut().add_relationship(
                rt::HEADER.to_owned(),
                target,
                "rIdHeader42".to_owned(),
                false,
            );
            Ok(())
        })
        .unwrap();
    let bindings = package.custom_xml_bindings().unwrap();
    assert_eq!(bindings.len(), 1);
    assert_eq!(bindings[0].control_id, Some(17));
    package.validate_custom_xml_bindings().unwrap();
    assert!(package.remove_custom_xml(ITEM_A).is_err());
    assert!(package.custom_xml_by_id(ITEM_A).unwrap().is_some());
}

#[test]
fn malformed_binding_and_replacement_fail_without_mutation() {
    let mut package = Package::new().unwrap();
    package.add_custom_xml(store(ITEM_A, "original")).unwrap();
    let before = package.custom_xml_by_id(ITEM_A).unwrap().unwrap();
    let bad_header = format!(
        r#"<w:hdr xmlns:w="{W}"><w:sdtPr><w:id w:val="1"/><w:dataBinding w:prefixMappings="xmlns:x=urn:test" w:xpath="/x" w:storeItemID="{ITEM_A}"/></w:sdtPr></w:hdr>"#
    );
    package
        .edit_opc(|opc| {
            opc.add_part(Box::new(BlobPart::new(
                PackURI::new("/word/headerBad.xml").unwrap(),
                ct::WML_HEADER.to_string(),
                bad_header.into_bytes(),
            )));
            Ok(())
        })
        .unwrap();
    assert!(package.custom_xml_bindings().is_err());
    assert!(package.remove_custom_xml(ITEM_A).is_err());

    let mut invalid = store(ITEM_A, "bad");
    invalid.xml = b"<!DOCTYPE root><root/>".to_vec();
    assert!(package.replace_custom_xml(ITEM_A, invalid).is_err());
    let after = package.custom_xml_by_id(ITEM_A).unwrap().unwrap();
    assert_eq!(after.xml(), before.xml());
    assert_eq!(after.props(), before.props());
}

#[test]
fn removal_preserves_a_data_part_with_an_unrelated_shared_reference() {
    let mut package = Package::new().unwrap();
    let item = package.add_custom_xml(store(ITEM_A, "shared")).unwrap();
    let footer_uri = PackURI::new("/word/footerShared.xml").unwrap();
    let mut footer = BlobPart::new(
        footer_uri.clone(),
        ct::WML_FOOTER.to_string(),
        format!(r#"<w:ftr xmlns:w="{W}"/>"#).into_bytes(),
    );
    footer.rels_mut().add_relationship(
        "urn:test:shared".to_string(),
        item.part().relative_ref(footer_uri.base_uri()),
        "rIdShared".to_string(),
        false,
    );
    package
        .edit_opc(|opc| {
            let document = opc.main_document_part()?.partname().clone();
            let target = footer_uri.relative_ref(document.base_uri());
            opc.add_part(Box::new(footer));
            opc.get_part_mut(&document)?.rels_mut().add_relationship(
                rt::FOOTER.to_owned(),
                target,
                "rIdFooterShared".to_owned(),
                false,
            );
            Ok(())
        })
        .unwrap();
    assert!(package.remove_custom_xml(ITEM_A).unwrap());
    assert!(package.opc_package().get_part(item.part()).is_ok());
    assert!(
        package
            .opc_package()
            .get_part(item.props_part().unwrap())
            .is_ok()
    );
}

#[test]
fn malformed_external_data_relationship_is_rejected_before_crud() {
    let mut package = Package::new().unwrap();
    let item = package.add_custom_xml(store(ITEM_A, "a")).unwrap();
    let source_name = item.source().clone();
    let relationship_id = item.rel_id().to_string();
    package
        .edit_opc(|opc| {
            let source = opc.get_part_mut(&source_name)?;
            source.rels_mut().remove(&relationship_id);
            source.rels_mut().add_relationship(
                TRANSITIONAL_RELATIONSHIP.to_string(),
                "https://example.invalid/data.xml".to_string(),
                relationship_id,
                true,
            );
            Ok(())
        })
        .unwrap();
    let part_count = package.opc_package().part_count();
    assert!(package.custom_xml().is_err());
    assert!(package.remove_custom_xml(ITEM_A).is_err());
    assert_eq!(package.opc_package().part_count(), part_count);
}

#[test]
fn identical_payload_is_an_exact_signed_noop_even_with_malformed_checksum() {
    let mut package = Package::new().unwrap();
    package.add_custom_xml(store(ITEM_A, "same")).unwrap();
    let exact = package
        .custom_xml_by_id(ITEM_A)
        .unwrap()
        .unwrap()
        .xml()
        .to_vec();
    package
        .edit_opc(|opc| {
            let document = opc.main_document_part()?.partname().clone();
            opc.get_part_mut(&document)?.set_blob(bound_story(
                "w:document",
                "<w:body>",
                ITEM_A,
                "malformed",
            ));
            Ok(())
        })
        .unwrap();
    mark_signed(&mut package);

    package.set_custom_xml(ITEM_A, exact.clone()).unwrap();
    package
        .replace_custom_xml(ITEM_A, store(ITEM_A, "same"))
        .unwrap();

    assert!(package.is_signed());
    assert_eq!(
        package.custom_xml_by_id(ITEM_A).unwrap().unwrap().xml(),
        exact
    );
    assert!(
        package
            .opc_package()
            .contains_part(&PackURI::new("/_xmlsignatures/origin.sigs").unwrap())
    );
}

#[test]
fn changed_payload_refreshes_main_header_and_reachable_comment_checksums() {
    let mut package = Package::new().unwrap();
    package.add_custom_xml(store(ITEM_A, "old")).unwrap();
    package
        .edit_opc(|opc| {
            let document = opc.main_document_part()?.partname().clone();
            opc.get_part_mut(&document)?.set_blob(bound_story(
                "w:document",
                "<w:body>",
                ITEM_A,
                "AAAAAA==",
            ));
            let header = PackURI::new("/word/headerChecksum.xml").unwrap();
            let comments = PackURI::new("/word/commentsChecksum.xml").unwrap();
            opc.add_part(Box::new(BlobPart::new(
                header.clone(),
                ct::WML_HEADER.to_owned(),
                bound_story("w:hdr", "", ITEM_A, "AAAAAA=="),
            )));
            opc.add_part(Box::new(BlobPart::new(
                comments.clone(),
                ct::WML_COMMENTS.to_owned(),
                bound_story("w:comments", "", ITEM_A, "AAAAAA=="),
            )));
            let header_target = header.relative_ref(document.base_uri());
            let comments_target = comments.relative_ref(document.base_uri());
            let relationships = opc.get_part_mut(&document)?.rels_mut();
            relationships.add_relationship(
                rt::HEADER.to_owned(),
                header_target,
                "rIdHeaderChecksum".to_owned(),
                false,
            );
            relationships.add_relationship(
                rt::COMMENTS.to_owned(),
                comments_target,
                "rIdCommentsChecksum".to_owned(),
                false,
            );
            Ok(())
        })
        .unwrap();
    mark_signed(&mut package);
    let updated = b"<root xmlns=\"urn:test\"><value>new</value></root>".to_vec();
    let expected = Checksum::compute(&updated, &Limits::default()).unwrap();

    assert!(package.set_custom_xml(ITEM_A, updated.clone()).is_err());
    assert!(package.is_signed());
    assert_ne!(
        package.custom_xml_by_id(ITEM_A).unwrap().unwrap().xml(),
        updated
    );

    package.unsign();
    package.set_custom_xml(ITEM_A, updated).unwrap();

    assert!(!package.is_signed());
    for part in [
        "/word/document.xml",
        "/word/headerChecksum.xml",
        "/word/commentsChecksum.xml",
    ] {
        assert_eq!(checksum_at(&package, part).as_bytes(), expected.as_bytes());
    }
    let bindings = package.custom_xml_bindings().unwrap();
    assert_eq!(bindings.len(), 3);
    assert!(
        bindings
            .iter()
            .any(|binding| binding.source.as_str() == "/word/commentsChecksum.xml")
    );
}

#[test]
fn malformed_matching_checksum_rolls_back_payload_stories_and_signature() {
    let mut package = Package::new().unwrap();
    package.add_custom_xml(store(ITEM_A, "old")).unwrap();
    package
        .edit_opc(|opc| {
            let document = opc.main_document_part()?.partname().clone();
            opc.get_part_mut(&document)?.set_blob(bound_story(
                "w:document",
                "<w:body>",
                ITEM_A,
                "malformed",
            ));
            Ok(())
        })
        .unwrap();
    mark_signed(&mut package);
    let payload_before = package
        .custom_xml_by_id(ITEM_A)
        .unwrap()
        .unwrap()
        .xml()
        .to_vec();
    let document_before = package
        .opc_package()
        .main_document_part()
        .unwrap()
        .blob()
        .to_vec();

    assert!(
        package
            .set_custom_xml(ITEM_A, b"<changed/>".to_vec())
            .is_err()
    );

    assert!(package.is_signed());
    assert_eq!(
        package.custom_xml_by_id(ITEM_A).unwrap().unwrap().xml(),
        payload_before
    );
    assert_eq!(
        package.opc_package().main_document_part().unwrap().blob(),
        document_before
    );

    package.unsign();
    assert!(
        package
            .set_custom_xml(ITEM_A, b"<changed/>".to_vec())
            .is_err()
    );
    assert_eq!(
        package.custom_xml_by_id(ITEM_A).unwrap().unwrap().xml(),
        payload_before
    );
    assert_eq!(
        package.opc_package().main_document_part().unwrap().blob(),
        document_before
    );
}

#[test]
fn duplicate_item_guid_and_orphan_story_are_refused_without_publication() {
    let mut duplicate = Package::new().unwrap();
    duplicate.add_custom_xml(store(ITEM_A, "first")).unwrap();
    duplicate.add_custom_xml(store(ITEM_B, "second")).unwrap();
    duplicate
        .edit_opc(|opc| {
            let props = PackURI::new("/customXml/itemProps2.xml").unwrap();
            let xml = String::from_utf8(opc.get_part(&props)?.blob().to_vec())
                .unwrap()
                .replace(ITEM_B, ITEM_A)
                .into_bytes();
            opc.get_part_mut(&props)?.set_blob(xml);
            Ok(())
        })
        .unwrap();
    let duplicate_parts = [
        PackURI::new("/customXml/item1.xml").unwrap(),
        PackURI::new("/customXml/item2.xml").unwrap(),
    ];
    let parts_before = duplicate_parts
        .iter()
        .map(|part| {
            duplicate
                .opc_package()
                .get_part(part)
                .unwrap()
                .blob()
                .to_vec()
        })
        .collect::<Vec<_>>();
    assert!(
        duplicate
            .set_custom_xml(ITEM_A, b"<changed/>".to_vec())
            .is_err()
    );
    let parts_after = duplicate_parts
        .iter()
        .map(|part| {
            duplicate
                .opc_package()
                .get_part(part)
                .unwrap()
                .blob()
                .to_vec()
        })
        .collect::<Vec<_>>();
    assert_eq!(parts_after, parts_before);

    let mut orphan = Package::new().unwrap();
    orphan.add_custom_xml(store(ITEM_A, "old")).unwrap();
    orphan
        .edit_opc(|opc| {
            opc.add_part(Box::new(BlobPart::new(
                PackURI::new("/word/orphanComments.xml").unwrap(),
                ct::WML_COMMENTS.to_owned(),
                bound_story("w:comments", "", ITEM_A, "AAAAAA=="),
            )));
            Ok(())
        })
        .unwrap();
    mark_signed(&mut orphan);
    let before = orphan
        .custom_xml_by_id(ITEM_A)
        .unwrap()
        .unwrap()
        .xml()
        .to_vec();
    assert!(
        orphan
            .set_custom_xml(ITEM_A, b"<changed/>".to_vec())
            .is_err()
    );
    assert!(orphan.is_signed());
    assert_eq!(
        orphan.custom_xml_by_id(ITEM_A).unwrap().unwrap().xml(),
        before
    );
}

#[test]
fn add_resolves_declared_checksum_and_duplicate_add_rolls_back() {
    let mut package = Package::new().unwrap();
    package
        .edit_opc(|opc| {
            let document = opc.main_document_part()?.partname().clone();
            opc.get_part_mut(&document)?.set_blob(bound_story(
                "w:document",
                "<w:body>",
                ITEM_A,
                "AAAAAA==",
            ));
            Ok(())
        })
        .unwrap();
    mark_signed(&mut package);
    let new_store = store(ITEM_A, "resolved");
    let expected = Checksum::compute(&new_store.xml, &Limits::default()).unwrap();

    assert!(package.add_custom_xml(new_store).is_err());
    assert!(package.is_signed());
    assert!(package.custom_xml_by_id(ITEM_A).unwrap().is_none());

    package.unsign();
    let added = package.add_custom_xml(store(ITEM_A, "resolved")).unwrap();

    assert_eq!(added.props().unwrap().id, ITEM_A);
    assert!(!package.is_signed());
    assert_eq!(
        checksum_at(&package, "/word/document.xml").as_bytes(),
        expected.as_bytes()
    );
    let bindings = package.custom_xml_bindings().unwrap();
    assert_eq!(bindings.len(), 1);
    assert_eq!(bindings[0].occurrence, 0);
    assert_eq!(bindings[0].control_id, None);

    mark_signed(&mut package);
    let parts = package.opc_package().part_count();
    assert!(package.add_custom_xml(store(ITEM_A, "duplicate")).is_err());
    assert!(package.is_signed());
    assert_eq!(package.opc_package().part_count(), parts);
    assert_eq!(package.custom_xml().unwrap().len(), 1);
}

#[test]
fn same_order_is_signed_noop_and_changed_order_is_managed() {
    let mut package = Package::new().unwrap();
    package.add_custom_xml(store(ITEM_A, "first")).unwrap();
    package.add_custom_xml(store(ITEM_B, "second")).unwrap();
    package
        .edit_opc(|opc| {
            let document = opc.main_document_part()?.partname().clone();
            opc.get_part_mut(&document)?.rels_mut().add_relationship(
                "urn:test:opaque".to_owned(),
                "https://example.test/keep?a=1&b=2".to_owned(),
                "rIdOpaque".to_owned(),
                true,
            );
            Ok(())
        })
        .unwrap();
    let relationship_snapshot = main_relationships(&package);
    mark_signed(&mut package);

    package
        .order_custom_xml(&[ITEM_A.to_owned(), ITEM_B.to_owned()])
        .unwrap();
    assert!(package.is_signed());
    let after_noop = main_relationships(&package);
    assert_eq!(after_noop, relationship_snapshot);

    package
        .order_custom_xml(&[ITEM_B.to_owned(), ITEM_A.to_owned()])
        .unwrap_err();
    assert!(package.is_signed());
    let items = package.custom_xml().unwrap();
    assert_eq!(items[0].props().unwrap().id, ITEM_A);
    assert_eq!(items[1].props().unwrap().id, ITEM_B);

    package.unsign();
    package
        .order_custom_xml(&[ITEM_B.to_owned(), ITEM_A.to_owned()])
        .unwrap();
    let items = package.custom_xml().unwrap();
    assert_eq!(items[0].props().unwrap().id, ITEM_B);
    assert_eq!(items[1].props().unwrap().id, ITEM_A);
    let relationships = main_relationships(&package);
    let before_ids = relationship_snapshot
        .iter()
        .map(|(id, _, _, _)| id)
        .collect::<Vec<_>>();
    let after_ids = relationships
        .iter()
        .map(|(id, _, _, _)| id)
        .collect::<Vec<_>>();
    assert_eq!(after_ids, before_ids);
    assert!(relationships.iter().any(|relationship| {
        relationship.0 == "rIdOpaque"
            && relationship.1 == "urn:test:opaque"
            && relationship.2 == "https://example.test/keep?a=1&b=2"
            && relationship.3
    }));
}

#[test]
fn dual_core_and_word_2012_bindings_refresh_each_exact_checksum() {
    const W15: &str = "http://schemas.microsoft.com/office/word/2012/wordml";
    let mut package = Package::new().unwrap();
    package.add_custom_xml(store(ITEM_A, "old")).unwrap();
    let document = format!(
        r#"<w:document xmlns:w="{W}" xmlns:w15="{W15}" xmlns:mc="{MCE}" xmlns:h="{HASH}" mc:Ignorable="w15 h"><w:body><w:sdt><w:sdtPr><w:dataBinding w:xpath="/root/value" w:storeItemID="{ITEM_A}" h:storeItemChecksum="AAAAAA=="/><w15:dataBinding w:xpath="/root/value" w:storeItemID="{ITEM_A}" h:storeItemChecksum="AQAAAA=="/></w:sdtPr><w:sdtContent/></w:sdt></w:body></w:document>"#
    );
    package
        .edit_opc(|opc| {
            let main = opc.main_document_part()?.partname().clone();
            opc.get_part_mut(&main)?.set_blob(document.into_bytes());
            Ok(())
        })
        .unwrap();
    let updated = b"<root xmlns=\"urn:test\"><value>dual</value></root>".to_vec();
    let expected = Checksum::compute(&updated, &Limits::default()).unwrap();

    package.set_custom_xml(ITEM_A, updated).unwrap();

    let main = package.opc_package().main_document_part().unwrap();
    let snapshot = Snapshot::from_xml(main.blob().to_vec()).unwrap();
    let bindings = snapshot.inventory().occurrences()[0]
        .control()
        .data_bindings();
    assert_eq!(bindings.len(), 2);
    assert_eq!(bindings[0].flavor(), BindingFlavor::Core);
    assert_eq!(bindings[1].flavor(), BindingFlavor::Word2012);
    assert!(bindings.iter().all(|binding| {
        binding
            .checksum()
            .is_some_and(|checksum| checksum.as_bytes() == expected.as_bytes())
    }));
    let projected = package.custom_xml_bindings().unwrap();
    assert_eq!(projected.len(), 2);
    assert_eq!(projected[0].flavor, BindingFlavor::Core);
    assert_eq!(projected[1].flavor, BindingFlavor::Word2012);
    mark_signed(&mut package);
    assert!(package.remove_custom_xml(ITEM_A).is_err());
    assert!(package.is_signed());
    package.unsign();
    assert!(package.remove_custom_xml(ITEM_A).is_err());
    assert!(package.custom_xml_by_id(ITEM_A).unwrap().is_some());
}

#[test]
fn removal_is_managed_while_missing_store_is_an_exact_signed_noop() {
    let mut package = Package::new().unwrap();
    package.add_custom_xml(store(ITEM_A, "remove")).unwrap();
    mark_signed(&mut package);

    assert!(!package.remove_custom_xml(ITEM_B).unwrap());
    assert!(package.is_signed());
    assert!(package.remove_custom_xml(ITEM_A).is_err());
    assert!(package.is_signed());
    assert!(package.custom_xml_by_id(ITEM_A).unwrap().is_some());

    package.unsign();
    assert!(package.remove_custom_xml(ITEM_A).unwrap());
    assert!(package.custom_xml_by_id(ITEM_A).unwrap().is_none());
}

#[test]
fn changed_replace_requires_explicit_unsign() {
    let mut package = Package::new().unwrap();
    package.add_custom_xml(store(ITEM_A, "old")).unwrap();
    mark_signed(&mut package);

    assert!(
        package
            .replace_custom_xml(ITEM_A, store(ITEM_A, "new"))
            .is_err()
    );
    assert!(package.is_signed());
    assert_eq!(
        package.custom_xml_by_id(ITEM_A).unwrap().unwrap().xml(),
        store(ITEM_A, "old").xml
    );

    package.unsign();
    package
        .replace_custom_xml(ITEM_A, store(ITEM_A, "new"))
        .unwrap();
    assert_eq!(
        package.custom_xml_by_id(ITEM_A).unwrap().unwrap().xml(),
        store(ITEM_A, "new").xml
    );
}
