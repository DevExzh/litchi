use litchi_ooxml::custom_xml_data::{CustomXmlConformance, TRANSITIONAL_CUSTOM_XML_RELATIONSHIP};
use litchi_ooxml::docx::{NewCustomXmlDataStore, Package};
use litchi_opc::constants::content_type as ct;
use litchi_opc::packuri::PackURI;
use litchi_opc::part::{BlobPart, Part};

const ITEM_A: &str = "{11111111-1111-4111-8111-111111111111}";
const ITEM_B: &str = "{22222222-2222-4222-8222-222222222222}";
const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

fn store(item_id: &str, value: &str) -> NewCustomXmlDataStore {
    NewCustomXmlDataStore {
        xml: format!(r#"<root xmlns="urn:test"><value>{value}</value></root>"#).into_bytes(),
        content_type: "application/xml".to_string(),
        item_id: item_id.to_string(),
        schema_references: vec!["urn:test:schema".to_string()],
        conformance: CustomXmlConformance::Transitional,
    }
}

#[test]
fn generated_add_find_update_replace_reorder_remove_round_trip() {
    let directory = tempfile::tempdir().unwrap();
    let path = directory.path().join("custom-xml.docx");
    let mut package = Package::new().unwrap();
    package
        .add_custom_xml_data_store(store(ITEM_A, "a"))
        .unwrap();
    package
        .add_custom_xml_data_store(store(ITEM_B, "b"))
        .unwrap();
    assert_eq!(package.custom_xml_data_stores().unwrap().len(), 2);

    package
        .update_custom_xml_data_store(ITEM_A, b"<updated/>".to_vec())
        .unwrap();
    let mut replacement = store(ITEM_B, "replacement");
    replacement.content_type = "application/vnd.example.data+xml".to_string();
    replacement.schema_references.push("urn:second".to_string());
    package
        .replace_custom_xml_data_store(ITEM_B, replacement)
        .unwrap();
    package
        .reorder_custom_xml_data_stores(&[ITEM_B.to_string(), ITEM_A.to_string()])
        .unwrap();
    let items = package.custom_xml_data_stores().unwrap();
    assert_eq!(items[0].properties.as_ref().unwrap().item_id, ITEM_B);
    assert_eq!(items[0].content_type, "application/vnd.example.data+xml");
    assert_eq!(items[1].xml, b"<updated/>");
    package.save(&path).unwrap();

    let mut reopened = Package::open(&path).unwrap();
    assert_eq!(reopened.custom_xml_data_stores().unwrap().len(), 2);
    assert!(reopened.remove_custom_xml_data_store(ITEM_A).unwrap());
    assert!(
        reopened
            .find_custom_xml_data_store(ITEM_A)
            .unwrap()
            .is_none()
    );
    assert!(!reopened.remove_custom_xml_data_store(ITEM_A).unwrap());
}

#[test]
fn binding_integrity_scans_word_containers_without_executing_xpath() {
    let mut package = Package::new().unwrap();
    package
        .add_custom_xml_data_store(store(ITEM_A, "a"))
        .unwrap();
    let header_xml = format!(
        r#"<w:hdr xmlns:w="{W}"><w:sdt><w:sdtPr><w:id w:val="17"/><w:dataBinding w:prefixMappings="xmlns:x='urn:test'" w:xpath="/x:root/x:value" w:storeItemID="{ITEM_A}"/></w:sdtPr><w:sdtContent/></w:sdt></w:hdr>"#
    );
    package.opc_package_mut().add_part(Box::new(BlobPart::new(
        PackURI::new("/word/header42.xml").unwrap(),
        ct::WML_HEADER.to_string(),
        header_xml.into_bytes(),
    )));
    let bindings = package.custom_xml_bindings().unwrap();
    assert_eq!(bindings.len(), 1);
    assert_eq!(bindings[0].content_control_id, 17);
    package.validate_custom_xml_binding_integrity().unwrap();
    assert!(package.remove_custom_xml_data_store(ITEM_A).is_err());
    assert!(
        package
            .find_custom_xml_data_store(ITEM_A)
            .unwrap()
            .is_some()
    );
}

#[test]
fn malformed_binding_and_replacement_fail_without_mutation() {
    let mut package = Package::new().unwrap();
    package
        .add_custom_xml_data_store(store(ITEM_A, "original"))
        .unwrap();
    let before = package.find_custom_xml_data_store(ITEM_A).unwrap().unwrap();
    let bad_header = format!(
        r#"<w:hdr xmlns:w="{W}"><w:sdtPr><w:id w:val="1"/><w:dataBinding w:prefixMappings="xmlns:x=urn:test" w:xpath="/x" w:storeItemID="{ITEM_A}"/></w:sdtPr></w:hdr>"#
    );
    package.opc_package_mut().add_part(Box::new(BlobPart::new(
        PackURI::new("/word/headerBad.xml").unwrap(),
        ct::WML_HEADER.to_string(),
        bad_header.into_bytes(),
    )));
    assert!(package.custom_xml_bindings().is_err());
    assert!(package.remove_custom_xml_data_store(ITEM_A).is_err());

    let mut invalid = store(ITEM_A, "bad");
    invalid.xml = b"<!DOCTYPE root><root/>".to_vec();
    assert!(
        package
            .replace_custom_xml_data_store(ITEM_A, invalid)
            .is_err()
    );
    let after = package.find_custom_xml_data_store(ITEM_A).unwrap().unwrap();
    assert_eq!(after.xml, before.xml);
    assert_eq!(after.properties, before.properties);
}

#[test]
fn removal_preserves_a_data_part_with_an_unrelated_shared_reference() {
    let mut package = Package::new().unwrap();
    let item = package
        .add_custom_xml_data_store(store(ITEM_A, "shared"))
        .unwrap();
    let footer_uri = PackURI::new("/word/footerShared.xml").unwrap();
    let mut footer = BlobPart::new(
        footer_uri.clone(),
        ct::WML_FOOTER.to_string(),
        format!(r#"<w:ftr xmlns:w="{W}"/>"#).into_bytes(),
    );
    footer.rels_mut().add_relationship(
        "urn:test:shared".to_string(),
        item.data_part_name.relative_ref(footer_uri.base_uri()),
        "rIdShared".to_string(),
        false,
    );
    package.opc_package_mut().add_part(Box::new(footer));
    assert!(package.remove_custom_xml_data_store(ITEM_A).unwrap());
    assert!(package.opc_package().get_part(&item.data_part_name).is_ok());
    assert!(
        package
            .opc_package()
            .get_part(item.properties_part_name.as_ref().unwrap())
            .is_ok()
    );
}

#[test]
fn malformed_external_data_relationship_is_rejected_before_crud() {
    let mut package = Package::new().unwrap();
    let item = package
        .add_custom_xml_data_store(store(ITEM_A, "a"))
        .unwrap();
    let source = package
        .opc_package_mut()
        .get_part_mut(&item.source_part_name)
        .unwrap();
    source.rels_mut().remove(&item.relationship_id);
    source.rels_mut().add_relationship(
        TRANSITIONAL_CUSTOM_XML_RELATIONSHIP.to_string(),
        "https://example.invalid/data.xml".to_string(),
        item.relationship_id,
        true,
    );
    let part_count = package.opc_package().part_count();
    assert!(package.custom_xml_data_stores().is_err());
    assert!(package.remove_custom_xml_data_store(ITEM_A).is_err());
    assert_eq!(package.opc_package().part_count(), part_count);
}
