//! Public Data Model and External Data Connections graph-boundary tests.
//!
//! These fixtures keep the opaque model bytes inert and exercise the public
//! package/workbook APIs around their OPC ownership and dependency closure.

use std::io::Cursor;

use litchi_opc::{BlobPart, OpcPackage, PackURI};
use litchi_xlsb::data_model::{Definition, Model, Table};
use litchi_xlsb::package::connections::{Connection, Connections, SourceType};
use litchi_xlsb::writer::{MutableWorksheet, WorkbookWriter};
use litchi_xlsb::{Package, Workbook};

const WORKBOOK_PART_NAME: &str = "/xl/workbook.bin";
const CONNECTIONS_CONTENT_TYPE: &str = "application/vnd.ms-excel.connections";
const CONNECTIONS_PART_NAME: &str = "/xl/connections.bin";
const CONNECTIONS_RELATIONSHIP_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/connections";

fn definition(connection: &str) -> Definition {
    Definition {
        min_version_load: 5,
        tables: vec![Table {
            id: "table-1".to_owned(),
            name: "Sales".to_owned(),
            connection: connection.to_owned(),
        }],
        relationships: Vec::new(),
        time_groupings: Vec::new(),
    }
}

fn connection(id: u32, name: &str) -> Connection {
    Connection {
        connection_id: id,
        source_type: SourceType::Odbc,
        name: name.to_owned(),
        ..Connection::default()
    }
}

fn package_with_connections(names: &[&str]) -> Package {
    package_with_model_connection(names, names[0])
}

fn package_with_model_connection(names: &[&str], model_connection: &str) -> Package {
    assert!(!names.is_empty());
    let mut writer = WorkbookWriter::new();
    writer.add_worksheet(MutableWorksheet::new("Sheet1"));
    writer
        .set_connections(Connections {
            connections: names
                .iter()
                .enumerate()
                .map(|(index, name)| connection(index as u32 + 1, name))
                .collect(),
        })
        .expect("connections");
    writer
        .set_data_model(
            Model::from_bytes(definition(model_connection), vec![1, 2, 3, 4]).expect("model"),
        )
        .expect("data model");

    let mut output = Cursor::new(Vec::new());
    writer.save(&mut output).expect("save");
    Package::from_bytes(output.into_inner()).expect("package")
}

fn workbook_uri() -> PackURI {
    PackURI::new(WORKBOOK_PART_NAME).expect("workbook URI")
}

fn connections_uri() -> PackURI {
    PackURI::new(CONNECTIONS_PART_NAME).expect("connections URI")
}

fn remove_connections_graph(package: &mut OpcPackage) {
    let workbook = workbook_uri();
    let relationship_id = package
        .get_part(&workbook)
        .expect("workbook")
        .rels()
        .iter()
        .find(|relationship| relationship.reltype() == CONNECTIONS_RELATIONSHIP_TYPE)
        .expect("connections relationship")
        .r_id()
        .to_owned();
    package
        .get_part_mut(&workbook)
        .expect("workbook")
        .rels_mut()
        .remove(&relationship_id);
    assert!(package.remove_part(&connections_uri()));
}

fn connections_payload(package: &Package) -> Vec<u8> {
    package
        .opc_package()
        .get_part(&connections_uri())
        .expect("connections part")
        .blob()
        .to_vec()
}

fn replace_wide_once(bytes: &mut [u8], from: &str, to: &str) {
    let from = from
        .encode_utf16()
        .flat_map(u16::to_le_bytes)
        .collect::<Vec<_>>();
    let to = to
        .encode_utf16()
        .flat_map(u16::to_le_bytes)
        .collect::<Vec<_>>();
    assert_eq!(
        from.len(),
        to.len(),
        "fixture replacement must be same-sized"
    );
    let positions = bytes
        .windows(from.len())
        .enumerate()
        .filter_map(|(index, value)| (value == from.as_slice()).then_some(index))
        .collect::<Vec<_>>();
    assert_eq!(positions.len(), 1, "connection name should occur once");
    bytes[positions[0]..positions[0] + to.len()].copy_from_slice(&to);
}

fn add_duplicate_connections_relationship(package: &mut OpcPackage) {
    let workbook = workbook_uri();
    let target = package
        .get_part(&workbook)
        .expect("workbook")
        .rels()
        .iter()
        .find(|relationship| relationship.reltype() == CONNECTIONS_RELATIONSHIP_TYPE)
        .expect("connections relationship")
        .target_ref()
        .to_owned();
    package
        .get_part_mut(&workbook)
        .expect("workbook")
        .rels_mut()
        .add_relationship(
            CONNECTIONS_RELATIONSHIP_TYPE.to_owned(),
            target,
            "rIdDuplicateConnections".to_owned(),
            false,
        );
}

fn replace_connections_relationship_with_external(package: &mut OpcPackage) {
    let workbook = workbook_uri();
    let relationship_id = package
        .get_part(&workbook)
        .expect("workbook")
        .rels()
        .iter()
        .find(|relationship| relationship.reltype() == CONNECTIONS_RELATIONSHIP_TYPE)
        .expect("connections relationship")
        .r_id()
        .to_owned();
    let workbook_part = package.get_part_mut(&workbook).expect("workbook");
    workbook_part.rels_mut().remove(&relationship_id);
    workbook_part.rels_mut().add_relationship(
        CONNECTIONS_RELATIONSHIP_TYPE.to_owned(),
        "https://example.invalid/connections.bin".to_owned(),
        "rIdExternalConnections".to_owned(),
        true,
    );
}

fn add_connections_outbound_relationship(package: &mut OpcPackage) {
    package
        .get_part_mut(&connections_uri())
        .expect("connections part")
        .rels_mut()
        .add_relationship(
            "urn:litchi:test:forbidden-child".to_owned(),
            "workbook.bin".to_owned(),
            "rIdForbiddenChild".to_owned(),
            false,
        );
}

fn saved_workbook(workbook: &Workbook) -> Vec<u8> {
    let mut output = Cursor::new(Vec::new());
    workbook.save(&mut output).expect("workbook bytes");
    output.into_inner()
}

fn assert_rejected_edit_preserves_bytes(
    mut workbook: Workbook,
    mutate: impl FnOnce(&mut OpcPackage),
) {
    let before = saved_workbook(&workbook);
    assert!(
        workbook
            .edit_opc(|package| {
                mutate(package);
                Ok(())
            })
            .is_err()
    );
    assert_eq!(saved_workbook(&workbook), before);
}

#[test]
fn valid_model_and_connections_owner_form_a_readable_closure() {
    let package = package_with_connections(&["Alpha"]);
    let snapshot = package.data_model().expect("Data Model snapshot");
    assert_eq!(
        snapshot.definition().expect("definition").tables[0].connection,
        "Alpha"
    );
    assert_eq!(snapshot.connection_names(), Some(&["Alpha".to_owned()][..]));

    let workbook = package.workbook().expect("workbook");
    assert_eq!(
        workbook.connections().expect("connections").connections[0].name,
        "Alpha"
    );
}

#[test]
fn model_connection_binding_is_case_insensitive() {
    let package = package_with_model_connection(&["Alpha"], "alpha");
    let snapshot = package.data_model().expect("Data Model snapshot");
    assert_eq!(
        snapshot.definition().expect("definition").tables[0].connection,
        "alpha"
    );
    assert_eq!(snapshot.connection_names(), Some(&["Alpha".to_owned()][..]));
}

#[test]
fn case_variant_connection_names_are_rejected_as_ambiguous_owners() {
    let mut writer = WorkbookWriter::new();
    writer.add_worksheet(MutableWorksheet::new("Sheet1"));
    writer
        .set_data_model(Model::from_bytes(definition("alpha"), vec![1, 2, 3, 4]).expect("model"))
        .expect("stage model before owner");
    let error = writer
        .set_connections(Connections {
            connections: vec![connection(1, "Alpha"), connection(2, "alpha")],
        })
        .expect_err("case-folded duplicate owner names");
    assert!(error.to_string().contains("duplicate connection name"));
    assert!(writer.data_model().is_some());
}

#[test]
fn missing_connections_owner_rejects_model_snapshot_without_mutating_bytes() {
    let mut package = package_with_connections(&["Alpha"]).into_opc();
    remove_connections_graph(&mut package);
    let package = Package::from_opc(package).expect("package without optional owner");
    let before = package.to_bytes().expect("invalid package bytes");
    let error = package
        .data_model()
        .expect_err("table must have a connections owner");
    assert!(error.to_string().contains("External Data Connections"));
    assert_eq!(
        package.to_bytes().expect("bytes after rejected read"),
        before
    );
}

#[test]
fn duplicate_connections_relationship_is_rejected() {
    let mut package = package_with_connections(&["Alpha"]).into_opc();
    add_duplicate_connections_relationship(&mut package);
    let error = Package::from_opc(package).expect_err("duplicate owner relationship");
    assert!(
        error
            .to_string()
            .contains("multiple connections relationships")
    );
}

#[test]
fn orphan_connections_part_is_rejected() {
    let package = package_with_connections(&["Alpha"]);
    let payload = connections_payload(&package);
    let mut package = package.into_opc();
    package.add_part(Box::new(BlobPart::new(
        PackURI::new("/xl/orphan-connections.bin").expect("orphan URI"),
        CONNECTIONS_CONTENT_TYPE.to_owned(),
        payload,
    )));
    let error = Package::from_opc(package).expect_err("orphan connections owner");
    assert!(
        error
            .to_string()
            .contains("orphan or additional connections")
    );
}

#[test]
fn connections_relationship_requires_the_connections_content_type() {
    let mut package = package_with_connections(&["Alpha"]).into_opc();
    package
        .get_part_mut(&connections_uri())
        .expect("connections part")
        .set_content_type("application/octet-stream".to_owned())
        .expect("content type mutation");
    let error = Package::from_opc(package).expect_err("wrong connections content type");
    assert!(error.to_string().contains("connections part"));
    assert!(error.to_string().contains("content type"));
}

#[test]
fn external_connections_relationship_is_rejected() {
    let mut package = package_with_connections(&["Alpha"]).into_opc();
    replace_connections_relationship_with_external(&mut package);
    let error = Package::from_opc(package).expect_err("external connections relationship");
    assert!(
        error
            .to_string()
            .contains("connections relationship cannot be external")
    );
}

#[test]
fn connections_part_rejects_outbound_relationships() {
    let mut package = package_with_connections(&["Alpha"]).into_opc();
    add_connections_outbound_relationship(&mut package);
    let error = Package::from_opc(package).expect_err("connections outbound relationship");
    assert!(
        error
            .to_string()
            .contains("connections part must not have relationships")
    );
}

#[test]
fn missing_model_table_connection_name_is_rejected() {
    let package = package_with_connections(&["Alpha"]);
    let mut payload = connections_payload(&package);
    replace_wide_once(&mut payload, "Alpha", "Bravo");
    let mut package = package.into_opc();
    package
        .get_part_mut(&connections_uri())
        .expect("connections part")
        .set_blob(payload);
    let package = Package::from_opc(package).expect("graph owner remains valid");
    let error = package
        .data_model()
        .expect_err("unknown table connection name");
    assert!(error.to_string().contains("unknown workbook connection"));
}

#[test]
fn ambiguous_model_table_connection_names_are_rejected() {
    let package = package_with_connections(&["Alpha", "Bravo"]);
    let mut payload = connections_payload(&package);
    replace_wide_once(&mut payload, "Bravo", "Alpha");
    let mut package = package.into_opc();
    package
        .get_part_mut(&connections_uri())
        .expect("connections part")
        .set_blob(payload);
    let error = Package::from_opc(package).expect_err("ambiguous connection owner");
    assert!(error.to_string().contains("duplicate connection name"));
}

#[test]
fn rejected_connection_graph_edits_preserve_workbook_bytes() {
    let package = package_with_connections(&["Alpha"]);

    assert_rejected_edit_preserves_bytes(
        package.clone().into_workbook().expect("workbook"),
        |package| {
            add_duplicate_connections_relationship(package);
        },
    );
    assert_rejected_edit_preserves_bytes(
        package.clone().into_workbook().expect("workbook"),
        |package| {
            add_connections_outbound_relationship(package);
        },
    );
    assert_rejected_edit_preserves_bytes(package.into_workbook().expect("workbook"), |package| {
        replace_connections_relationship_with_external(package);
    });
}
