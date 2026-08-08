use litchi_opc::OpcPackage;
use litchi_xlsx::Package;
use litchi_xlsx::xml_maps::{
    XmlMap, XmlMapConformance, XmlMapInfo, XmlMapSchema, load_from_package, remove_from_package,
    store_in_package,
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

fn package() -> OpcPackage {
    Package::create()
        .expect("create workbook package")
        .into_plain_opc()
}

#[test]
fn xml_maps_survive_direct_package_publication() {
    let mut package = package();
    let value = fixture_info();
    store_in_package(&mut package, &value, XmlMapConformance::Strict).unwrap();
    assert_eq!(load_from_package(&package).unwrap(), Some(value.clone()));

    let directory = tempfile::tempdir().unwrap();
    let path = directory.path().join("xml-maps.xlsx");
    package.save(&path).unwrap();
    let reopened = OpcPackage::open(&path).unwrap();
    assert_eq!(load_from_package(&reopened).unwrap(), Some(value));
}

#[test]
fn xml_maps_removal_is_idempotent() {
    let mut package = package();
    store_in_package(&mut package, &fixture_info(), XmlMapConformance::Strict).unwrap();
    assert!(remove_from_package(&mut package).unwrap());
    assert_eq!(load_from_package(&package).unwrap(), None);
    assert!(!remove_from_package(&mut package).unwrap());
}
