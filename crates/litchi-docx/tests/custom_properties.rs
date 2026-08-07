use litchi_docx::Package;
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

#[test]
fn word_reserved_custom_properties_are_host_scoped_and_transactional() {
    let directory = tempfile::tempdir().expect("temporary directory");
    let valid_path = directory.path().join("word-reserved-props.docx");
    let mut valid = Package::new().expect("new DOCX package");
    valid
        .custom_props_mut()
        .insert(
            "ClassificationContentMarkingHeaderFontProps",
            "#ffFF00,23,Calibri",
        )
        .expect("valid Word header font properties");
    valid
        .save(&valid_path)
        .expect("save Word reserved properties");
    assert_eq!(
        Package::open(&valid_path)
            .expect("reopen Word reserved properties")
            .custom_props()
            .get("ClassificationContentMarkingHeaderFontProps"),
        Some(&Value::Text("#ffFF00,23,Calibri".to_owned()))
    );

    let invalid_path = directory.path().join("non-word-reserved-props.docx");
    let mut invalid = Package::new().expect("new DOCX package");
    invalid
        .custom_props_mut()
        .insert(
            "ClassificationContentMarkingHeaderLocations",
            r"Office\:Suite:10",
        )
        .expect("valid PowerPoint location syntax");
    assert!(invalid.save(&invalid_path).is_err());
    assert!(
        invalid
            .opc_package()
            .iter_parts()
            .all(|part| part.content_type() != ct::OFC_CUSTOM_PROPERTIES)
    );
}
