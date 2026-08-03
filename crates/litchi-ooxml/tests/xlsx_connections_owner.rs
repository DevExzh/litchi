use litchi_ooxml::xlsx::connections::{
    Connection, Connections, CredentialsMethod, load_from_package, store_in_package,
};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::part::BlobPart;
use litchi_opc::{OpcPackage, PackURI};

fn connection(id: u32) -> Connection {
    Connection {
        id,
        source_file: None,
        odc_file: None,
        keep_alive: None,
        interval: None,
        name: Some("canonical".into()),
        description: None,
        connection_type: Some(1),
        reconnection_method: None,
        refreshed_version: 7,
        min_refreshable_version: None,
        save_password: Some(false),
        new_connection: None,
        deleted: None,
        only_use_connection_file: None,
        background: None,
        refresh_on_load: Some(false),
        save_data: Some(true),
        credentials: Some(CredentialsMethod::None),
        single_sign_on_id: None,
        database: None,
        olap: None,
        web: None,
        text: None,
        parameters: None,
        extension_xml: None,
    }
}

fn package() -> OpcPackage {
    let mut package = OpcPackage::new();
    let workbook = PackURI::new("/xl/workbook.xml").unwrap();
    package.add_part(Box::new(BlobPart::new(
        workbook,
        ct::SML_SHEET_MAIN.into(),
        br#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>"#
            .to_vec(),
    )));
    package.relate_to("xl/workbook.xml", rt::OFFICE_DOCUMENT);
    package
}

#[test]
fn host_connections_api_forwards_the_canonical_owner() {
    let value = Connections {
        connections: vec![connection(7)],
    };
    let mut package = package();

    fn accepts_canonical_owner(_: &litchi_xlsx::connections::Connections) {}
    accepts_canonical_owner(&value);

    store_in_package(&mut package, &value, false).unwrap();
    let forwarded = load_from_package(&package).unwrap().unwrap();
    let canonical = litchi_xlsx::connections::load_from_package(&package)
        .unwrap()
        .unwrap();
    assert_eq!(forwarded, canonical);
    assert_eq!(forwarded.connections[0].id, 7);
}

#[test]
fn host_adapter_retains_query_table_reference_validation() {
    let mut package = package();
    package.add_part(Box::new(BlobPart::new(
        PackURI::new("/xl/queryTables/queryTable1.xml").unwrap(),
        litchi_ooxml::xlsx::query_table::QUERY_TABLE_CONTENT_TYPE.into(),
        br#"<queryTable xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="Q" connectionId="99"/>"#.to_vec(),
    )));

    let error = store_in_package(
        &mut package,
        &Connections {
            connections: vec![connection(7)],
        },
        false,
    )
    .unwrap_err();
    assert!(error.to_string().contains("missing connection ID 99"));
    assert!(load_from_package(&package).unwrap().is_none());
}
