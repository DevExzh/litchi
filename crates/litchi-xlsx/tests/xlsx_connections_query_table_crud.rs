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

fn mixed_query_owners() -> (Package, String, String) {
    let (base, worksheet) = package(Connections {
        connections: vec![connection(1, "connection")],
    });
    let mut edit = base.edit_query_tables(0_usize).unwrap();
    let direct = edit
        .add(table("direct", 1), Conformance::Transitional)
        .unwrap();
    let indirect = edit
        .add(table("indirect", 1), Conformance::Transitional)
        .unwrap();
    let mut opc = edit.commit().unwrap().into_package().into_plain_opc();
    opc.get_part_mut(&worksheet)
        .unwrap()
        .rels_mut()
        .remove(indirect.relationship_id());
    let owner = PackURI::new("/xl/tables/table1.xml").unwrap();
    let mut table=BlobPart::new(owner.clone(),litchi_opc::constants::content_type::SML_TABLE.into(),br#"<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Linked" displayName="Linked" ref="A1:A2" tableType="queryTable"><tableColumns count="1"><tableColumn id="1" name="Value" queryTableFieldId="1"/></tableColumns></table>"#.to_vec());
    table.rels_mut().add_relationship(
        litchi_xlsx::query_table::QUERY_TABLE_RELATIONSHIP_TYPE.into(),
        "../queryTables/queryTable2.xml".into(),
        direct.relationship_id().into(),
        false,
    );
    opc.try_add_part(Box::new(table)).unwrap();
    opc.get_part_mut(&worksheet)
        .unwrap()
        .rels_mut()
        .add_relationship(
            litchi_opc::constants::relationship_type::TABLE.into(),
            "../tables/table1.xml".into(),
            "rIdTable".into(),
            false,
        );
    (
        Package::from_opc(opc).unwrap(),
        direct.relationship_id().into(),
        indirect.part_name().into(),
    )
}

#[test]
fn table_owned_query_identity_disambiguates_local_ids_and_refuses_wrong_owner_removal() {
    let (package, id, indirect) = mixed_query_owners();
    let mut edit = package.edit_query_tables(0_usize).unwrap();
    let original = edit.query_tables().unwrap();
    assert_eq!(original.len(), 2);
    assert_ne!(original[0].owner_part_name(), original[1].owner_part_name());
    assert!(edit.query_table(&id).is_err());
    assert!(
        edit.replace(&id, table("wrong", 1), Conformance::Transitional)
            .is_err()
    );
    assert!(edit.remove(&id).is_err());
    assert_eq!(edit.query_tables().unwrap(), original);
    let mut unbound = original
        .iter()
        .find(|part| part.part_name() == indirect)
        .unwrap()
        .query_table()
        .clone();
    unbound.set_refresh(None);
    assert!(matches!(
        edit.replace_part(&indirect, unbound, Conformance::Transitional),
        Err(litchi_xlsx::Error::Unsupported { .. })
    ));
    assert_eq!(edit.query_tables().unwrap(), original);

    edit.replace_part(&indirect, table("selected", 1), Conformance::Transitional)
        .unwrap();
    let after = edit.commit().unwrap().into_package();
    let catalog = after
        .workbook()
        .sheet(0_usize)
        .unwrap()
        .unwrap()
        .query_tables()
        .unwrap();
    assert_eq!(
        catalog
            .iter()
            .find(|p| p.part_name() == indirect)
            .unwrap()
            .query_table()
            .name(),
        "selected"
    );
    assert_eq!(
        catalog
            .iter()
            .find(|p| p.part_name() != indirect)
            .unwrap()
            .query_table()
            .name(),
        "direct"
    );
    let mut opc = after.into_plain_opc();
    opc.get_part_mut(&PackURI::new("/xl/worksheets/sheet1.xml").unwrap())
        .unwrap()
        .rels_mut()
        .remove(&id);
    opc.remove_part(&PackURI::new("/xl/queryTables/queryTable1.xml").unwrap());
    let only = Package::from_opc(opc).unwrap();
    let mut edit = only.edit_query_tables(0_usize).unwrap();
    assert!(matches!(
        edit.remove(&id),
        Err(litchi_xlsx::Error::Unsupported { .. })
    ));
    assert!(matches!(
        edit.reorder(&[id]),
        Err(litchi_xlsx::Error::Unsupported { .. })
    ));
    assert_eq!(edit.query_tables().unwrap().len(), 1);
}

#[test]
fn table_query_owners_reject_external_duplicate_and_orphan_relationships() {
    let (base, _, _) = mixed_query_owners();
    let worksheet = PackURI::new("/xl/worksheets/sheet1.xml").unwrap();
    let owner = PackURI::new("/xl/tables/table1.xml").unwrap();
    for variant in 0..4 {
        let mut opc = base.clone().into_plain_opc();
        match variant {
            0 => {
                opc.get_part_mut(&worksheet)
                    .unwrap()
                    .rels_mut()
                    .remove("rIdTable");
            },
            1 => {
                opc.get_part_mut(&owner)
                    .unwrap()
                    .rels_mut()
                    .add_relationship(
                        litchi_xlsx::query_table::QUERY_TABLE_RELATIONSHIP_TYPE.into(),
                        "../queryTables/queryTable1.xml".into(),
                        "extra".into(),
                        false,
                    );
            },
            2 => {
                opc.get_part_mut(&owner)
                    .unwrap()
                    .rels_mut()
                    .add_relationship(
                        litchi_xlsx::query_table::QUERY_TABLE_RELATIONSHIP_TYPE.into(),
                        "https://invalid.example/query.xml".into(),
                        "external".into(),
                        true,
                    );
            },
            _ => {
                opc.get_part_mut(&worksheet)
                    .unwrap()
                    .rels_mut()
                    .add_relationship(
                        litchi_xlsx::query_table::QUERY_TABLE_RELATIONSHIP_TYPE.into(),
                        "../queryTables/queryTable2.xml".into(),
                        "duplicate".into(),
                        false,
                    );
            },
        }
        assert!(
            litchi_xlsx::connections::Snapshot::load(&opc, &opc.read_limits()).is_err(),
            "variant {variant}"
        );
        if variant != 0 {
            assert!(load_worksheet_query_tables(&opc, &worksheet).is_err());
        }
    }
}

#[test]
fn query_rename_preserves_bom_quotes_entities_and_unknown_metadata() {
    let (base, _) = package(Connections {
        connections: vec![connection(1, "connection")],
    });
    let mut edit = base.edit_query_tables(0_usize).unwrap();
    let part = edit
        .add(table("before", 1), Conformance::Transitional)
        .unwrap();
    let mut opc = edit.commit().unwrap().into_package().into_plain_opc();
    let original=b"\xef\xbb\xbf<?xml version='1.0'?>\r\n<q:queryTable xmlns:q='http://schemas.openxmlformats.org/spreadsheetml/2006/main' xmlns:mc='http://schemas.openxmlformats.org/markup-compatibility/2006' xmlns:v='urn:future' mc:Ignorable='v' name = 'b&#101;fore' connectionId='1' v:uid='opaque'><v:future>before</v:future></q:queryTable>";
    opc.get_part_mut(&PackURI::new(part.part_name()).unwrap())
        .unwrap()
        .set_blob(original.to_vec());
    let package = Package::from_opc(opc).unwrap();
    let mut edit = package.edit_query_tables(0_usize).unwrap();
    let mut value = edit.query_tables().unwrap()[0].query_table().clone();
    value.set_name("after & 'quoted' \"value\"".into());
    edit.replace_part(part.part_name(), value.clone(), Conformance::Transitional)
        .unwrap();
    let result = edit.commit().unwrap().into_package();
    let raw = plain(&result)
        .get_part(&PackURI::new(part.part_name()).unwrap())
        .unwrap()
        .blob()
        .to_vec();
    let expected = String::from_utf8(original.to_vec()).unwrap().replace(
        "b&#101;fore",
        "after &amp; &apos;quoted&apos; &quot;value&quot;",
    );
    assert!(raw == expected.as_bytes());
    assert_eq!(
        result
            .workbook()
            .sheet(0_usize)
            .unwrap()
            .unwrap()
            .query_tables()
            .unwrap()[0]
            .query_table(),
        &value
    );
}
