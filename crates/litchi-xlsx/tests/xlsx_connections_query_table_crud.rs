use litchi_opc::part::BlobPart;
use litchi_opc::{OpcPackage, PackURI, Part};
use litchi_xlsx::Package;
use litchi_xlsx::connections::{Connection, Connections, CredentialsMethod, load_from_package};
use litchi_xlsx::query_table::{Conformance, Field, Refresh, Table};
use litchi_xlsx::{find_worksheet_query_table, load_worksheet_query_tables};

fn connection(id: u32, name: &str) -> Connection {
    Connection {
        id,
        source_file: None,
        odc_file: None,
        keep_alive: None,
        interval: None,
        name: Some(name.into()),
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

/// A workbook whose first sheet is `/xl/worksheets/sheet1.xml`, with
/// `connections` stored.
fn package(connections: Connections) -> (Package, PackURI) {
    let mut edit = Package::new().unwrap().edit_connections().unwrap();
    edit.replace(Some(connections)).unwrap();
    let package = edit.commit().unwrap().into_package();
    (package, PackURI::new("/xl/worksheets/sheet1.xml").unwrap())
}

fn plain(package: &Package) -> OpcPackage {
    package.clone().into_plain_opc()
}

fn load(package: &Package) -> litchi_xlsx::Result<Option<Connections>> {
    let plain = plain(package);
    let limits = plain.read_limits();
    load_from_package(&plain, &limits)
}

fn table(name: &str, connection_id: u32) -> Table {
    let mut refresh = Refresh::new();
    let mut field = Field::new(1);
    field.set_name(Some("Value".into()));
    field.set_row_numbers(Some(false));
    field.set_fill_formulas(Some(true));
    refresh.add_field(field);
    refresh.add_deleted_field("Removed".into());
    let mut table = Table::new(name, connection_id);
    table.set_headers(Some(true));
    table.set_row_numbers(Some(false));
    table.set_fill_formulas(Some(true));
    table.set_disable_refresh(Some(true));
    table.set_refresh_on_load(Some(false));
    table.set_refresh(Some(refresh));
    table
}

#[test]
fn generated_connections_and_query_tables_round_trip_inertly() {
    let mut connections = Connections {
        connections: vec![connection(1, "first"), connection(2, "second")],
    };
    connections.reorder(&[2, 1]).unwrap();
    let (package, worksheet) = package(connections);
    assert_eq!(load(&package).unwrap().unwrap().connections[0].id, 2);

    let mut edit = package.edit_query_tables(0_usize).unwrap();
    let first = edit
        .add(table("Query A", 1), Conformance::Transitional)
        .unwrap();
    let second = edit.add(table("Query B", 2), Conformance::Strict).unwrap();
    let package = edit.commit().unwrap().into_package();
    assert_eq!(
        load_worksheet_query_tables(&plain(&package), &worksheet)
            .unwrap()
            .len(),
        2
    );
    assert!(
        find_worksheet_query_table(&plain(&package), &worksheet, first.relationship_id())
            .unwrap()
            .is_some()
    );

    let mut edit = package.edit_query_tables(0_usize).unwrap();
    edit.replace(
        first.relationship_id(),
        table("Updated", 1),
        Conformance::Transitional,
    )
    .unwrap();
    let reordered = edit
        .reorder(&[
            second.relationship_id().into(),
            first.relationship_id().into(),
        ])
        .unwrap();
    assert_eq!(reordered[0].query_table().name(), "Query B");
    assert!(edit.remove(reordered[1].relationship_id()).unwrap());
    let package = edit.commit().unwrap().into_package();
    assert_eq!(
        load_worksheet_query_tables(&plain(&package), &worksheet)
            .unwrap()
            .len(),
        1
    );
}

#[test]
fn missing_connection_and_bad_reorder_are_atomic() {
    let (package, worksheet) = package(Connections {
        connections: vec![connection(1, "only")],
    });
    let mut edit = package.edit_query_tables(0_usize).unwrap();
    assert!(
        edit.add(table("Missing", 99), Conformance::Transitional)
            .is_err()
    );
    assert!(edit.query_tables().unwrap().is_empty());

    let added = edit
        .add(table("Valid", 1), Conformance::Transitional)
        .unwrap();
    assert!(edit.reorder(&["missing".into()]).is_err());
    assert!(edit.query_table(added.relationship_id()).unwrap().is_some());
    let package = edit.commit().unwrap().into_package();
    assert!(
        find_worksheet_query_table(&plain(&package), &worksheet, added.relationship_id())
            .unwrap()
            .is_some()
    );
}

#[test]
fn removal_preserves_a_query_table_target_shared_by_another_part() {
    let (package, _) = package(Connections {
        connections: vec![connection(1, "only")],
    });
    let mut edit = package.edit_query_tables(0_usize).unwrap();
    let added = edit
        .add(table("Shared", 1), Conformance::Transitional)
        .unwrap();
    let mut package = plain(&edit.commit().unwrap().into_package());
    let owner_name = PackURI::new("/xl/customXml/shared.xml").unwrap();
    let mut owner = BlobPart::new(owner_name, "application/xml".into(), b"<shared/>".to_vec());
    owner.relate_to(
        &PackURI::new(added.part_name())
            .unwrap()
            .relative_ref("/xl/customXml/"),
        "urn:test:shared-query-table",
    );
    package.add_part(Box::new(owner));
    let target = PackURI::new(added.part_name()).unwrap();
    let mut edit = Package::from_opc(package)
        .unwrap()
        .edit_query_tables(0_usize)
        .unwrap();
    assert!(edit.remove(added.relationship_id()).unwrap());
    let package = plain(&edit.commit().unwrap().into_package());
    assert!(package.get_part(&target).is_ok());
}
