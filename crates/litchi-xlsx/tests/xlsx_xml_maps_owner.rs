use litchi_opc::OpcPackage;
use litchi_xlsx::Package;
use litchi_xlsx::xml_maps::{
    EditOptions, XmlMap, XmlMapConformance, XmlMapInfo, XmlMapSchema, load_from_package,
};

fn fixture_info() -> XmlMapInfo {
    XmlMapInfo {
        selection_namespaces: String::new(),
        schemas: vec![XmlMapSchema {
            id: "schema-1".into(),
            schema_reference: None,
            namespace: Some("urn:litchi:example".into()),
            payload_xml: None,
        }],
        maps: vec![XmlMap {
            id: 1,
            name: "Example map".into(),
            root_element: "example".into(),
            schema_id: "schema-1".into(),
            show_import_export_validation_errors: true,
            auto_fit: true,
            append: false,
            preserve_sort_auto_filter_layout: true,
            preserve_format: true,
            data_binding: None,
        }],
    }
}

fn package() -> Package {
    Package::new().expect("create workbook package")
}

fn store(package: &Package, value: XmlMapInfo) -> Package {
    let mut edit = package
        .edit_xml_maps_with_options(EditOptions::default().conformance(XmlMapConformance::Strict))
        .unwrap();
    assert!(edit.set(value).unwrap());
    edit.commit().unwrap().into_package()
}

#[test]
fn xml_maps_survive_direct_package_publication() {
    let value = fixture_info();
    let package = store(&package(), value.clone()).into_plain_opc();
    assert_eq!(load_from_package(&package).unwrap(), Some(value.clone()));

    let directory = tempfile::tempdir().unwrap();
    let path = directory.path().join("xml-maps.xlsx");
    package.save(&path).unwrap();
    let reopened = OpcPackage::open(&path, litchi_opc::SourceLimits::default()).unwrap();
    assert_eq!(load_from_package(&reopened).unwrap(), Some(value));
}

#[test]
fn xml_maps_removal_is_idempotent() {
    let package = store(&package(), fixture_info());
    let mut edit = package.edit_xml_maps().unwrap();
    assert!(edit.remove().unwrap().is_some());
    let package = edit.commit().unwrap().into_package();
    assert_eq!(
        load_from_package(&package.clone().into_plain_opc()).unwrap(),
        None
    );
    let mut edit = package.edit_xml_maps().unwrap();
    assert!(edit.remove().unwrap().is_none());
    let commit = edit.commit().unwrap();
    assert!(!commit.diagnostics().changed());
    assert!(commit.patch().is_noop());
}
