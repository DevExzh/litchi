use litchi_ooxml::docx::Package;
use litchi_ooxml_common::custom::Value;
use litchi_opc::OpcPackage;
use litchi_opc::PackURI;
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::part::BlobPart;

#[test]
fn custom_props_round_trip_and_clear_remove_the_package_graph() {
    let directory = tempfile::tempdir().expect("temporary directory");
    let path = directory.path().join("custom-props.docx");
    let mut package = Package::new().expect("new DOCX package");
    package
        .custom_props_mut()
        .insert("Project", "Litchi")
        .expect("valid custom property");
    package
        .custom_props_mut()
        .insert("Version", Value::I32(38))
        .expect("valid custom property");
    package.save(&path).expect("save custom properties");

    let mut reopened = Package::open(&path).expect("reopen custom properties");
    assert_eq!(
        reopened.custom_props().get("project"),
        Some(&Value::Text("Litchi".to_string()))
    );
    assert_eq!(
        reopened.custom_props().get("VERSION"),
        Some(&Value::I32(38))
    );
    reopened.custom_props_mut().clear();
    reopened.save(&path).expect("save cleared properties");

    let graph = OpcPackage::open(&path).expect("open cleared OPC graph");
    assert!(
        graph
            .rels()
            .iter()
            .all(|relationship| relationship.reltype() != rt::CUSTOM_PROPERTIES)
    );
    assert!(
        graph
            .iter_parts()
            .all(|part| part.content_type() != ct::OFC_CUSTOM_PROPERTIES)
    );
    assert!(
        Package::open(&path)
            .expect("reopen cleared package")
            .custom_props()
            .is_empty()
    );
}

#[test]
fn malformed_custom_props_are_not_treated_as_absent() {
    let directory = tempfile::tempdir().expect("temporary directory");
    let path = directory.path().join("malformed-custom-props.docx");
    let mut package = Package::new().expect("new DOCX package");
    package.save(&path).expect("save base package");

    let mut opc = OpcPackage::open(&path).expect("open base OPC package");
    let part = PackURI::new("/docProps/custom.xml").expect("static package URI");
    opc.add_part(Box::new(BlobPart::new(
        part,
        ct::OFC_CUSTOM_PROPERTIES.to_string(),
        b"<not-properties/>".to_vec(),
    )));
    opc.relate_to("docProps/custom.xml", rt::CUSTOM_PROPERTIES);

    assert!(Package::from_opc_package(opc).is_err());
}
