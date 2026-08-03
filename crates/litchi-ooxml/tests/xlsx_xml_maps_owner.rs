use litchi_ooxml::xlsx::{
    Workbook, XmlMap, XmlMapConformance, XmlMapInfo, XmlMapSchema, load_xml_maps_from_package,
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

#[test]
fn legacy_host_preserves_xml_maps_through_writer_materialization() {
    let mut workbook = Workbook::create().expect("create workbook");
    let value = fixture_info();
    workbook
        .set_xml_maps(&value, XmlMapConformance::Strict)
        .expect("set XML Maps");
    assert_eq!(
        workbook.xml_maps().expect("read XML Maps"),
        Some((value.clone(), XmlMapConformance::Strict))
    );
    assert_eq!(
        load_xml_maps_from_package(workbook.opc_package()).expect("read forwarded XML Maps"),
        Some(value.clone())
    );

    workbook
        .worksheet_mut(0)
        .expect("first worksheet")
        .set_cell_value(1, 1, "materialized");
    let directory = tempfile::tempdir().expect("temporary directory");
    let path = directory.path().join("materialized-xml-maps.xlsx");
    workbook.save(&path).expect("save workbook");
    let reopened = Workbook::open(&path).expect("reopen workbook");
    assert_eq!(
        reopened.xml_maps().expect("read saved XML Maps"),
        Some((value.clone(), XmlMapConformance::Strict))
    );

    let mut reopened = reopened;
    assert!(reopened.remove_xml_maps().expect("remove XML Maps"));
    assert_eq!(reopened.xml_maps().expect("read removed XML Maps"), None);
    assert!(
        !reopened
            .remove_xml_maps()
            .expect("idempotent XML Maps removal")
    );
}
