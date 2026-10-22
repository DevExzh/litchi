use litchi_opc::part::BlobPart;
use litchi_opc::{OpcPackage, PackURI};
use litchi_xlsx::Package;
use litchi_xlsx::connections::{Connection, Connections, CredentialsMethod, load_from_package};

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
    Package::new().unwrap().into_plain_opc()
}

fn load(package: &OpcPackage) -> litchi_xlsx::Result<Option<Connections>> {
    let limits = package.read_limits();
    load_from_package(package, &limits)
}

fn store(package: OpcPackage, value: Connections) -> litchi_xlsx::Result<OpcPackage> {
    let mut edit = Package::from_opc(package)?.edit_connections()?;
    edit.replace(Some(value))?;
    Ok(edit.commit()?.into_package().into_plain_opc())
}

#[test]
fn standalone_connections_api_uses_the_canonical_owner() {
    let value = Connections {
        connections: vec![connection(7)],
    };
    fn accepts_canonical_owner(_: &Connections) {}
    accepts_canonical_owner(&value);

    let package = store(package(), value).unwrap();
    let forwarded = load(&package).unwrap().unwrap();
    let canonical = load(&package).unwrap().unwrap();
    assert_eq!(forwarded, canonical);
    assert_eq!(forwarded.connections[0].id, 7);
}

#[test]
fn standalone_adapter_retains_query_table_reference_validation() {
    let mut package = package();
    package.add_part(Box::new(BlobPart::new(
        PackURI::new("/xl/queryTables/queryTable1.xml").unwrap(),
        litchi_xlsx::query_table::QUERY_TABLE_CONTENT_TYPE.into(),
        br#"<queryTable xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="Q" connectionId="99"/>"#.to_vec(),
    )));

    let error = store(
        package.clone(),
        Connections {
            connections: vec![connection(7)],
        },
    )
    .unwrap_err();
    // The edit starts from a validated graph, so the orphan query table is
    // refused before the staged catalog's reference check runs.
    assert!(
        error
            .to_string()
            .contains("query-table parts require a workbook connections part"),
        "{error}"
    );
    assert!(load(&package).unwrap().is_none());
}
