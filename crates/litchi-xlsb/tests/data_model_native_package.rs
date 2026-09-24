//! Native-package integration for synthetic XLSB Data Model ownership.
//!
//! The native fixture deliberately has no Data Model. Every model and
//! connection payload in these tests is authored by the public API; this is
//! package preservation coverage, not native Data Model producer compatibility.

use std::collections::BTreeMap;
use std::io::Cursor;
use std::path::Path;

use litchi_xlsb::Package;
use litchi_xlsb::data_model::{Definition, Model, Table};
use litchi_xlsb::package::connections::{Connection, Connections, SourceType};

const FIXTURE: &str = "test-data/ooxml/xlsb/date.xlsb";
// `sha256sum test-data/ooxml/xlsb/date.xlsb`:
// fbb969989aabed2e057b2ddfc4cfd10ef1f6ebe9b7697c3755283515158d5b09

fn fixture_bytes() -> Vec<u8> {
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../")
        .join(FIXTURE);
    std::fs::read(path).expect("native XLSB fixture")
}

fn native_package() -> Package {
    Package::from_bytes(fixture_bytes()).expect("native XLSB package")
}

fn connections() -> Connections {
    Connections {
        connections: vec![Connection {
            connection_id: 1,
            source_type: SourceType::Odbc,
            name: "Native synthetic connection".to_owned(),
            ..Connection::default()
        }],
    }
}

fn definition() -> Definition {
    Definition {
        min_version_load: 5,
        tables: vec![Table {
            id: "native-synthetic-table".to_owned(),
            name: "NativeSyntheticSales".to_owned(),
            connection: "Native synthetic connection".to_owned(),
        }],
        relationships: Vec::new(),
        time_groupings: Vec::new(),
    }
}

fn attach_connections(package: Package) -> Package {
    let mut workbook = package.into_workbook().expect("native workbook");
    workbook
        .set_connections(connections())
        .expect("attach synthetic connections");
    let mut output = Cursor::new(Vec::new());
    workbook
        .save(&mut output)
        .expect("save synthetic connections");
    Package::from_bytes(output.into_inner()).expect("reopen synthetic connections")
}

fn native_part_bytes(package: &Package) -> BTreeMap<String, Vec<u8>> {
    package
        .opc_package()
        .try_iter_parts()
        .map(|part| part.expect("part payload decodes"))
        .filter(|part| {
            !matches!(
                part.partname().as_str(),
                "/xl/workbook.bin" | "/xl/connections.bin" | "/xl/model/item.data"
            )
        })
        .map(|part| (part.partname().as_str().to_owned(), part.blob().to_vec()))
        .collect()
}

fn part_bytes(package: &Package, name: &str) -> Vec<u8> {
    package
        .opc_package()
        .get_part(&litchi_opc::PackURI::new(name).expect("part URI"))
        .expect("part")
        .blob()
        .to_vec()
}

#[test]
fn native_fixture_is_model_free_before_synthetic_attachment() {
    let source = fixture_bytes();
    assert!(!source.is_empty());

    let package = native_package();
    let snapshot = package.data_model().expect("native Data Model snapshot");
    assert!(!snapshot.is_present());
    assert!(snapshot.connection_names().is_none());

    let before = package.to_bytes().expect("native package bytes");
    let noop = snapshot.edit().commit().expect("native Data Model no-op");
    assert!(!noop.changed());
    let applied = package
        .apply_data_model(&noop)
        .expect("apply native Data Model no-op");
    assert_eq!(applied.to_bytes().expect("native no-op bytes"), before);
}

#[test]
fn native_package_preserves_members_through_synthetic_model_lifecycle() {
    let source = native_package();
    let native_members = native_part_bytes(&source);
    assert!(!native_members.is_empty());

    let with_connections = attach_connections(source);
    let with_connections_workbook = part_bytes(&with_connections, "/xl/workbook.bin");
    assert_eq!(native_part_bytes(&with_connections), native_members);
    assert_eq!(
        with_connections
            .workbook()
            .expect("workbook")
            .connections()
            .expect("synthetic connections")
            .connections[0]
            .name,
        "Native synthetic connection"
    );
    assert!(
        !with_connections
            .data_model()
            .expect("model-free snapshot")
            .is_present()
    );

    let before = with_connections.data_model().expect("before snapshot");
    let mut create = before.edit();
    create
        .replace_model(Some(
            Model::from_bytes(definition(), vec![1, 2, 3, 4]).expect("synthetic model"),
        ))
        .expect("stage synthetic model");
    let create_commit = create.commit().expect("commit synthetic model");
    let attached = with_connections
        .apply_data_model(&create_commit)
        .expect("publish synthetic model");
    assert_eq!(native_part_bytes(&attached), native_members);

    let attached_bytes = attached.to_bytes().expect("synthetic model bytes");
    let attached_reopened = Package::from_bytes(attached_bytes).expect("reopen synthetic model");
    let attached_snapshot = attached_reopened.data_model().expect("attached snapshot");
    assert_eq!(attached_snapshot.definition(), Some(&definition()));
    assert_eq!(
        attached_snapshot.part().expect("model part").bytes(),
        &[1, 2, 3, 4]
    );
    assert_eq!(native_part_bytes(&attached_reopened), native_members);

    let mut version_edit = attached_snapshot.edit();
    version_edit
        .edit_definition(|value| {
            value.min_version_load = 6;
            Ok(())
        })
        .expect("stage version edit");
    let version_commit = version_edit.commit().expect("commit version edit");
    let changed = attached_reopened
        .apply_data_model(&version_commit)
        .expect("publish version edit");
    assert_eq!(
        changed
            .data_model()
            .expect("changed snapshot")
            .definition()
            .expect("changed definition")
            .min_version_load,
        6
    );
    let changed_reopened = Package::from_bytes(changed.to_bytes().expect("changed bytes"))
        .expect("reopen changed model");
    assert_eq!(native_part_bytes(&changed_reopened), native_members);

    let restored = changed_reopened
        .apply_data_model_patch(&version_commit.patch().inverse())
        .expect("inverse version edit");
    assert_eq!(
        restored.data_model().expect("restored snapshot"),
        attached_snapshot
    );
    assert_eq!(native_part_bytes(&restored), native_members);

    let mut remove = restored.data_model().expect("model for removal").edit();
    remove.remove_model().expect("stage model removal");
    let remove_commit = remove.commit().expect("commit model removal");
    let removed = restored
        .apply_data_model(&remove_commit)
        .expect("publish model removal");
    assert!(!removed.data_model().expect("removed snapshot").is_present());
    assert_eq!(native_part_bytes(&removed), native_members);
    assert_eq!(
        part_bytes(&removed, "/xl/workbook.bin"),
        with_connections_workbook
    );

    let removed_reopened = Package::from_bytes(removed.to_bytes().expect("removed bytes"))
        .expect("reopen removed model");
    assert!(
        !removed_reopened
            .data_model()
            .expect("removed reopened snapshot")
            .is_present()
    );
    assert_eq!(native_part_bytes(&removed_reopened), native_members);
    assert_eq!(
        part_bytes(&removed_reopened, "/xl/workbook.bin"),
        with_connections_workbook
    );
}
