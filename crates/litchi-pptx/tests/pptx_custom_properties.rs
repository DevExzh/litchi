#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]

use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::part::BlobPart;
use litchi_opc::{OpcPackage, PackURI};
use litchi_pptx::Package;
use litchi_pptx::custom::{Host, Props, Value};

const PRESENTATION_XML: &[u8] =
    include_bytes!("../../../test-data/ooxml/pptx/vba-project/presentation.xml");
const CUSTOM_PROPERTIES_XML: &[u8] = br#"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>
<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/custom-properties\" xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\"><property fmtid=\"{D5CDD505-2E9C-101B-9397-08002B2CF9AE}\" pid=\"2\" name=\"Owner\"><vt:lpwstr>Alice</vt:lpwstr></property></Properties>"#;

#[test]
fn custom_properties_are_absent_round_trip_and_remove_idempotently() {
    let mut package = Package::new().unwrap();
    assert!(package.custom_props().unwrap().is_empty());

    let mut props = Props::new();
    props.insert("Owner", "Alice").unwrap();
    props.insert("Retries", 3_i32).unwrap();
    package.put_custom_props(props).unwrap();

    let bytes = package.to_bytes().unwrap();
    let mut reopened = Package::from_bytes(&bytes).unwrap();
    let props = reopened.custom_props().unwrap();
    assert_eq!(props.get("owner"), Some(&Value::Text("Alice".to_owned())));
    assert_eq!(props.get("RETRIES"), Some(&Value::I32(3)));

    reopened.remove_custom_props().unwrap();
    assert!(reopened.custom_props().unwrap().is_empty());
    reopened.remove_custom_props().unwrap();
    assert!(reopened.custom_props().unwrap().is_empty());
}

#[test]
fn custom_properties_reject_malformed_parts_and_invalid_relationship_graphs() {
    let mut malformed = base_package();
    add_custom_part(&mut malformed, "/docProps/custom.xml", b"<Properties>");
    malformed.relate_to("docProps/custom.xml", rt::CUSTOM_PROPERTIES);
    assert!(
        Package::from_opc_package(malformed)
            .unwrap()
            .custom_props()
            .is_err()
    );

    let mut external = base_package();
    external.rels_mut().add_relationship(
        rt::CUSTOM_PROPERTIES.to_owned(),
        "https://example.invalid/custom.xml".to_owned(),
        "rIdCustom".to_owned(),
        true,
    );
    assert!(
        Package::from_opc_package(external)
            .unwrap()
            .custom_props()
            .is_err()
    );

    let mut duplicate = base_package();
    add_custom_part(
        &mut duplicate,
        "/docProps/custom.xml",
        CUSTOM_PROPERTIES_XML,
    );
    add_custom_part(
        &mut duplicate,
        "/docProps/other-custom.xml",
        CUSTOM_PROPERTIES_XML,
    );
    duplicate.relate_to("docProps/custom.xml", rt::CUSTOM_PROPERTIES);
    assert!(
        Package::from_opc_package(duplicate)
            .unwrap()
            .custom_props()
            .is_err()
    );
}

#[test]
fn custom_properties_reject_word_only_reserved_metadata() {
    let mut package = Package::new().unwrap();
    let mut props = Props::new();
    props
        .insert("ClassificationContentMarkingHeaderShapeIds", "1,2")
        .unwrap();

    assert!(props.validate_for(Host::PowerPoint).is_err());
    assert!(package.put_custom_props(props).is_err());
    assert!(package.custom_props().unwrap().is_empty());
}

fn base_package() -> OpcPackage {
    let presentation_name = PackURI::new("/ppt/presentation.xml").unwrap();
    let mut package = OpcPackage::new();
    package.add_part(Box::new(BlobPart::new(
        presentation_name,
        ct::PML_PRES_MACRO_MAIN.to_owned(),
        PRESENTATION_XML.to_vec(),
    )));
    package.relate_to("ppt/presentation.xml", rt::OFFICE_DOCUMENT);
    package
}

fn add_custom_part(package: &mut OpcPackage, name: &str, xml: &[u8]) {
    package.add_part(Box::new(BlobPart::new(
        PackURI::new(name).unwrap(),
        ct::OFC_CUSTOM_PROPERTIES.to_owned(),
        xml.to_vec(),
    )));
}
