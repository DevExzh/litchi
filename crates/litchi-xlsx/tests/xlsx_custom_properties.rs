use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::part::BlobPart;
use litchi_opc::{OpcPackage, PackURI, TargetMode};
use litchi_xlsx::custom::{Host, Props, Value};
use litchi_xlsx::{Error, Package};

const CUSTOM_PART: &str = "/docProps/custom.xml";

fn package() -> OpcPackage {
    Package::create().expect("create XLSX package").into()
}

fn install_custom_part(package: &mut OpcPackage, xml: &[u8]) {
    package.add_part(Box::new(BlobPart::new(
        PackURI::new(CUSTOM_PART).expect("static custom-properties URI"),
        ct::OFC_CUSTOM_PROPERTIES.to_owned(),
        xml.to_vec(),
    )));
    package.relate_to("docProps/custom.xml", rt::CUSTOM_PROPERTIES);
}

fn signed_package() -> Package {
    let mut raw = package();
    raw.rels_mut()
        .try_add_relationship(
            rt::DIGITAL_SIGNATURE_ORIGIN.to_owned(),
            "_xmlsignatures/origin.sigs".to_owned(),
            "rIdSignature".to_owned(),
            TargetMode::Internal,
        )
        .expect("signature-origin relationship");
    Package::from_opc(raw).expect("signed XLSX package")
}

#[test]
fn custom_properties_absence_is_explicit() {
    let mut package = Package::create().expect("create XLSX package");

    assert!(
        package
            .custom_props()
            .expect("read absent custom properties")
            .is_empty()
    );
    package
        .remove_custom_props()
        .expect("remove absent custom properties");
}

#[test]
fn custom_properties_round_trip_through_the_xlsx_facade() {
    let mut props: Props = Props::new();
    props
        .insert("Project", "Litchi")
        .expect("insert text property");
    props
        .insert("Version", Value::I32(38))
        .expect("insert integer property");

    let mut package = Package::create().expect("create XLSX package");
    package
        .put_custom_props(props)
        .expect("store custom properties");
    let bytes = package.to_bytes().expect("serialize XLSX package");
    let reopened = Package::from_bytes(bytes).expect("reopen XLSX package");
    let props = reopened.custom_props().expect("read custom properties");

    assert_eq!(
        props.get("project"),
        Some(&Value::Text("Litchi".to_owned()))
    );
    assert_eq!(props.get("VERSION"), Some(&Value::I32(38)));
}

#[test]
fn custom_properties_are_nameable_through_the_xlsx_context() {
    let mut props = Props::new();
    props
        .insert("Approved", Value::Bool(true))
        .expect("insert contextual custom property");
    props
        .validate_for(Host::Excel)
        .expect("validate contextual Excel custom property");

    let mut package = Package::create().expect("create XLSX package");
    package
        .put_custom_props(props)
        .expect("store contextual custom property");
    assert_eq!(
        package.custom_props().unwrap().get("approved"),
        Some(&Value::Bool(true))
    );
}

#[test]
fn custom_properties_reject_reserved_non_excel_host_metadata() {
    let mut props = Props::new();
    props
        .insert("ClassificationContentMarkingHeaderText", "Header")
        .expect("valid Word or PowerPoint metadata");

    let mut package = Package::create().expect("create XLSX package");
    assert!(matches!(
        package.put_custom_props(props),
        Err(Error::Common(_))
    ));
    assert!(package.custom_props().unwrap().is_empty());
}

#[test]
fn removing_custom_properties_removes_the_complete_package_graph() {
    let mut props = Props::new();
    props.insert("Project", "Litchi").unwrap();

    let mut package = Package::create().expect("create XLSX package");
    package.put_custom_props(props).unwrap();
    package.remove_custom_props().unwrap();
    assert!(package.custom_props().unwrap().is_empty());

    let raw: OpcPackage = package.into();
    assert!(
        raw.rels()
            .iter()
            .all(|relationship| relationship.reltype() != rt::CUSTOM_PROPERTIES)
    );
    assert!(
        raw.iter_parts()
            .all(|part| part.content_type() != ct::OFC_CUSTOM_PROPERTIES)
    );
}

#[test]
fn malformed_external_and_duplicate_custom_property_graphs_map_to_common_errors() {
    let malformed = {
        let mut raw = package();
        install_custom_part(&mut raw, b"<not-properties/>");
        raw
    };
    let external = {
        let mut raw = package();
        raw.rels_mut()
            .try_add_relationship(
                rt::CUSTOM_PROPERTIES.to_owned(),
                "https://example.invalid/custom.xml".to_owned(),
                "rIdCustom".to_owned(),
                TargetMode::External,
            )
            .expect("external custom-properties relationship");
        raw
    };
    let duplicate = {
        let mut raw = package();
        install_custom_part(&mut raw, br#"<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"/>"#);
        raw.rels_mut()
            .try_add_relationship(
                rt::CUSTOM_PROPERTIES.to_owned(),
                "docProps/custom.xml".to_owned(),
                "rIdCustomDuplicate".to_owned(),
                TargetMode::Internal,
            )
            .expect("duplicate custom-properties relationship");
        raw
    };

    for raw in [malformed, external, duplicate] {
        let package = Package::from_opc(raw).expect("valid SpreadsheetML graph");
        assert!(matches!(package.custom_props(), Err(Error::Common(_))));
    }
}

#[test]
fn custom_property_noops_preserve_and_changes_invalidate_signatures() {
    let mut package = signed_package();
    assert!(package.custom_props().unwrap().is_empty());
    package.put_custom_props(Props::new()).unwrap();
    assert!(OpcPackage::from(package.clone()).is_signed());

    let mut props = Props::new();
    props.insert("Project", "Litchi").unwrap();
    package.put_custom_props(props).unwrap();
    assert!(!OpcPackage::from(package).is_signed());
}
