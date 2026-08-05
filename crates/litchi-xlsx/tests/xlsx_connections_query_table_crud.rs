use litchi_ooxml::xlsx::connections::{
    Connection, Connections, CredentialsMethod, load_from_package, store_in_package,
};
use litchi_ooxml::xlsx::query_table::{Conformance, Field, Refresh, Table};
use litchi_ooxml::xlsx::{
    add_worksheet_query_table, find_worksheet_query_table, load_worksheet_query_tables,
    remove_worksheet_query_table, reorder_worksheet_query_tables, update_worksheet_query_table,
};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::part::BlobPart;
use litchi_opc::{OpcPackage, PackURI, Part};

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

fn package() -> (OpcPackage, PackURI) {
    let mut package = OpcPackage::new();
    let workbook_name = PackURI::new("/xl/workbook.xml").unwrap();
    let worksheet_name = PackURI::new("/xl/worksheets/sheet1.xml").unwrap();
    let mut workbook = BlobPart::new(
        workbook_name.clone(),
        ct::SML_SHEET_MAIN.into(),
        br#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>"#
            .to_vec(),
    );
    workbook.relate_to("worksheets/sheet1.xml", rt::WORKSHEET);
    package.add_part(Box::new(workbook));
    package.add_part(Box::new(BlobPart::new(
        worksheet_name.clone(),
        ct::SML_WORKSHEET.into(),
        br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>"#.to_vec(),
    )));
    package.relate_to("xl/workbook.xml", rt::OFFICE_DOCUMENT);
    (package, worksheet_name)
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
    let (mut package, worksheet) = package();
    let mut connections = Connections {
        connections: vec![connection(1, "first"), connection(2, "second")],
    };
    connections.reorder(&[2, 1]).unwrap();
    store_in_package(&mut package, &connections, false).unwrap();
    assert_eq!(
        load_from_package(&package).unwrap().unwrap().connections[0].id,
        2
    );

    let first = add_worksheet_query_table(
        &mut package,
        &worksheet,
        table("Query A", 1),
        Conformance::Transitional,
    )
    .unwrap();
    let second = add_worksheet_query_table(
        &mut package,
        &worksheet,
        table("Query B", 2),
        Conformance::Strict,
    )
    .unwrap();
    assert_eq!(
        load_worksheet_query_tables(&package, &worksheet)
            .unwrap()
            .len(),
        2
    );
    assert!(
        find_worksheet_query_table(&package, &worksheet, first.relationship_id())
            .unwrap()
            .is_some()
    );

    update_worksheet_query_table(
        &mut package,
        &worksheet,
        first.relationship_id(),
        table("Updated", 1),
        Conformance::Transitional,
    )
    .unwrap();
    let reordered = reorder_worksheet_query_tables(
        &mut package,
        &worksheet,
        &[
            second.relationship_id().into(),
            first.relationship_id().into(),
        ],
    )
    .unwrap();
    assert_eq!(reordered[0].query_table().name(), "Query B");
    assert!(
        remove_worksheet_query_table(&mut package, &worksheet, reordered[1].relationship_id())
            .unwrap()
    );
}

#[test]
fn missing_connection_and_bad_reorder_are_atomic() {
    let (mut package, worksheet) = package();
    store_in_package(
        &mut package,
        &Connections {
            connections: vec![connection(1, "only")],
        },
        false,
    )
    .unwrap();
    assert!(
        add_worksheet_query_table(
            &mut package,
            &worksheet,
            table("Missing", 99),
            Conformance::Transitional,
        )
        .is_err()
    );
    assert!(
        load_worksheet_query_tables(&package, &worksheet)
            .unwrap()
            .is_empty()
    );

    let added = add_worksheet_query_table(
        &mut package,
        &worksheet,
        table("Valid", 1),
        Conformance::Transitional,
    )
    .unwrap();
    assert!(reorder_worksheet_query_tables(&mut package, &worksheet, &["missing".into()]).is_err());
    assert!(
        find_worksheet_query_table(&package, &worksheet, added.relationship_id())
            .unwrap()
            .is_some()
    );
}

#[test]
fn removal_preserves_a_query_table_target_shared_by_another_part() {
    let (mut package, worksheet) = package();
    store_in_package(
        &mut package,
        &Connections {
            connections: vec![connection(1, "only")],
        },
        false,
    )
    .unwrap();
    let added = add_worksheet_query_table(
        &mut package,
        &worksheet,
        table("Shared", 1),
        Conformance::Transitional,
    )
    .unwrap();
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
    assert!(
        remove_worksheet_query_table(&mut package, &worksheet, added.relationship_id()).unwrap()
    );
    assert!(package.get_part(&target).is_ok());
}
