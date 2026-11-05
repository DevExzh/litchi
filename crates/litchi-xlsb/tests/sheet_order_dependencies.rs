#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "dependency-boundary fixtures deliberately fail at the assertion site"
)]

//! Complete-catalog reorder dependency boundaries.
//!
//! A tab reorder changes the workbook display order, so every modeled owner
//! whose payload stores a catalog ordinal must either be remapped as part of
//! the same atomic replacement or refuse the operation before publication.
//! These fixtures exercise both sides of that contract.

use litchi_opc::constants::{content_type, relationship_type};
use litchi_opc::{BlobPart, PackURI, TargetMode};
use litchi_xlsb::Package;
use litchi_xlsb::builder::{self, Builder};
use litchi_xlsb::data_model::{Definition, Model, Table};
use litchi_xlsb::package::PackageError;
use litchi_xlsb::package::connections::{Connection, Connections, SourceType};
use litchi_xlsb::raw::{Kind, Records, kind};
use litchi_xlsb::shared_workbook::{
    Catalog, Guid, Header, Info, RawRecord, RevisionHeaders, RevisionLog, ShortDateTime, User,
    UserNames,
};
use litchi_xlsb::volatile_dependencies::{
    CONTENT_TYPE as VOLATILE_CONTENT_TYPE, CachedValue, CellReference, DEFAULT_PART_NAME,
    Dependencies, DependencyKind, MainTopic, RELATIONSHIP_TYPE as VOLATILE_RELATIONSHIP_TYPE,
    Topic, VolatileType,
};

fn worksheets(names: &[&str]) -> Package {
    let mut builder = Builder::new();
    for name in names {
        builder.add_worksheet(builder::Worksheet::new(*name));
    }
    builder.build().expect("worksheet package")
}

fn selectors<'a>(names: &[&'a str]) -> Vec<litchi_xlsb::sheet::Selector<'a>> {
    names.iter().copied().map(Into::into).collect()
}

fn owner_uri() -> PackURI {
    PackURI::new(DEFAULT_PART_NAME).expect("volatile owner URI")
}

fn owner_blob(package: &Package) -> Vec<u8> {
    package
        .opc_package()
        .get_part(&owner_uri())
        .expect("volatile owner")
        .blob()
        .to_vec()
}

fn volatile_dependencies() -> Dependencies {
    Dependencies::new(vec![VolatileType {
        kind: DependencyKind::Rtd,
        mains: vec![MainTopic {
            first: "server".to_owned(),
            topics: vec![Topic {
                subtopics: vec!["topic".to_owned()],
                value: CachedValue::Number(42.5),
                references: vec![
                    CellReference::new(1, 2, 0).expect("first reference"),
                    CellReference::new(3, 4, 2).expect("third-sheet reference"),
                ],
            }],
        }],
    }])
}

fn package_with_volatile_dependencies() -> Package {
    let package = worksheets(&["A", "B", "C"]);
    let mut edit = package
        .edit_volatile_dependencies()
        .expect("volatile transaction");
    edit.set_dependencies(volatile_dependencies())
        .expect("stage volatile dependencies");
    let commit = edit.commit().expect("commit volatile dependencies");
    package
        .apply_volatile_dependencies(commit.patch())
        .map(litchi_core::edit::Commit::into_package)
        .expect("publish volatile dependencies")
}

fn records(bytes: &[u8]) -> Vec<(Kind, Vec<u8>)> {
    Records::new(bytes, litchi_xlsb::raw::Limits::DEFAULT)
        .map(|record| {
            let record = record.expect("BIFF12 record");
            (record.kind(), record.payload().to_vec())
        })
        .collect()
}

fn assert_volatile_only_ish_changed(before: &[u8], after: &[u8], expected: &[u32]) {
    let before = records(before);
    let after = records(after);
    assert_eq!(before.len(), after.len());
    let mut volatile_index = 0usize;
    for ((before_kind, before_payload), (after_kind, after_payload)) in before.iter().zip(after) {
        assert_eq!(*before_kind, after_kind);
        assert_eq!(before_payload.len(), after_payload.len());
        if *before_kind == kind::VOL_REF {
            assert_eq!(before_payload[..8], after_payload[..8]);
            assert_eq!(
                &after_payload[8..12],
                &expected
                    .get(volatile_index)
                    .expect("expected volatile reference")
                    .to_le_bytes()
            );
            volatile_index = volatile_index.saturating_add(1);
        } else {
            assert_eq!(before_payload, &after_payload);
        }
    }
    assert_eq!(volatile_index, expected.len());
}

fn add_part_with_relationship(
    package: Package,
    part_name: &str,
    part_content_type: &str,
    relationship_type: &str,
    payload: &[u8],
    relationship_id: &str,
) -> Package {
    let mut opc = package.into_opc();
    let workbook = opc
        .main_document_part()
        .expect("workbook part")
        .partname()
        .clone();
    let target = PackURI::new(part_name).expect("dependency part URI");
    opc.add_part(Box::new(BlobPart::new(
        target.clone(),
        part_content_type.to_owned(),
        payload.to_vec(),
    )));
    opc.get_part_mut(&workbook)
        .expect("workbook part")
        .rels_mut()
        .try_add_relationship(
            relationship_type.to_owned(),
            target.relative_ref(workbook.base_uri()),
            relationship_id.to_owned(),
            TargetMode::Internal,
        )
        .expect("dependency relationship");
    Package::from_opc(opc).expect("dependency package")
}

fn package_with_calc_chain(relationship_type: &str) -> Package {
    add_part_with_relationship(
        worksheets(&["A", "B", "C"]),
        "/xl/calcChain.bin",
        "application/vnd.ms-excel.calcChain",
        relationship_type,
        &[],
        "rIdCalcChain",
    )
}

fn package_with_vba_project() -> Package {
    add_part_with_relationship(
        worksheets(&["A", "B", "C"]),
        "/xl/vbaProject.bin",
        content_type::OFC_VBA_PROJECT,
        relationship_type::VBA_PROJECT,
        b"opaque-vba-project",
        "rIdVbaProject",
    )
}

fn package_with_connections() -> Package {
    let mut builder = Builder::new();
    for name in ["A", "B", "C"] {
        builder.add_worksheet(builder::Worksheet::new(name));
    }
    builder
        .set_connections(Connections {
            connections: vec![Connection {
                connection_id: 1,
                source_type: SourceType::Odbc,
                name: "Connection".to_owned(),
                ..Connection::default()
            }],
        })
        .expect("connections");
    builder.build().expect("connections package")
}

fn package_with_data_model_and_connections() -> Package {
    let mut builder = Builder::new();
    for name in ["A", "B", "C"] {
        builder.add_worksheet(builder::Worksheet::new(name));
    }
    builder
        .set_connections(Connections {
            connections: vec![Connection {
                connection_id: 1,
                source_type: SourceType::Odbc,
                name: "Connection".to_owned(),
                ..Connection::default()
            }],
        })
        .expect("connections");
    builder
        .set_data_model(
            Model::from_bytes(
                Definition {
                    min_version_load: 5,
                    tables: vec![Table {
                        id: "table-1".to_owned(),
                        name: "Sales".to_owned(),
                        connection: "Connection".to_owned(),
                    }],
                    relationships: Vec::new(),
                    time_groupings: Vec::new(),
                },
                vec![1, 2, 3, 4],
                litchi_xlsb::data_model::SemanticLimits::DEFAULT,
            )
            .expect("model"),
        )
        .expect("data model");
    builder.build().expect("data-model package")
}

fn package_with_data_model() -> Package {
    let mut builder = Builder::new();
    for name in ["A", "B", "C"] {
        builder.add_worksheet(builder::Worksheet::new(name));
    }
    builder
        .set_data_model(
            Model::from_bytes(
                Definition::default(),
                vec![1, 2, 3, 4],
                litchi_xlsb::data_model::SemanticLimits::DEFAULT,
            )
            .expect("model"),
        )
        .expect("data model");
    builder.build().expect("data-model package")
}

fn package_with_shared_workbook() -> Package {
    let source = worksheets(&["A", "B", "C"]);
    let guid = Guid::from_bytes([7; 16]);
    let date = ShortDateTime {
        year: 2024,
        month: 1,
        day: 1,
        hour: 9,
        minute: 30,
        second: 0,
        weekday: 1,
    };
    let catalog = Catalog {
        users: Some(UserNames::new(vec![User {
            id: 1,
            guid,
            opened_at: date,
            name: "Alice".to_owned(),
        }])),
        headers: Some(RevisionHeaders::new(
            Info {
                guid,
                root_guid: guid,
                revision_id: 0,
                version: 1,
                has_revisions: false,
                no_revision_history: true,
                protected: false,
                revision_history_interval: 0,
            },
            vec![Header {
                guid,
                saved_at: date,
                next_sheet_id: 4,
                revision_min: 0,
                revision_max: 0,
                user_name: "Alice".to_owned(),
                relationship_id: "rIdLog".to_owned(),
                sheet_ids: vec![1, 2, 3],
                reviewed: Vec::new(),
            }],
        )),
        logs: vec![RevisionLog::new(vec![RawRecord::new(
            0x3fff,
            b"opaque-log".to_vec(),
        )])],
    };
    let mut edit = source.edit_shared_workbook().unwrap();
    edit.replace(catalog).unwrap();
    edit.commit().unwrap().into_package()
}

fn assert_reorder_guard(package: Package, label: &str) {
    let before = package.to_bytes().expect("source bytes");

    // Identity order checks precede dependency guards, so callers can safely
    // publish an exact no-op even when an opaque owner is present.
    let mut noop = package.edit().expect("identity edit");
    assert!(
        !noop
            .reorder_sheets(&selectors(&["A", "B", "C"]))
            .expect("identity reorder")
    );
    let noop = noop.commit().expect("identity commit");
    assert!(noop.patch().is_noop(), "{label}: identity should be exact");
    assert_eq!(noop.package().to_bytes().expect("identity bytes"), before);

    let mut edit = package.edit().expect("reorder edit");
    edit.reorder_sheets(&selectors(&["C", "A", "B"]))
        .expect("stage reorder");
    let error = edit.commit().expect_err("dependency guard");
    assert!(
        matches!(error, PackageError::Refused(_)),
        "{label}: expected typed unsupported error, got {error:?}"
    );
    assert_eq!(package.to_bytes().expect("unchanged source"), before);
}

#[test]
fn volatile_reference_ordinals_remap_and_inverse_restores_every_byte() {
    let source = package_with_volatile_dependencies();
    let before = source.to_bytes().expect("source bytes");
    let owner_before = owner_blob(&source);
    let mut edit = source.edit().expect("forward source edit");
    edit.reorder_sheets(&selectors(&["B", "C", "A"]))
        .expect("reorder with volatile owner");
    let commit = edit.commit().expect("forward reorder commit");
    let target = commit.package();
    let owner_after = owner_blob(target);

    // Old catalog positions A=0, B=1, C=2 map to new positions A=2, B=0,
    // C=1. The owner stream itself remains byte-identical except for each
    // BrtVolRef.ish field at payload offset 8.
    assert_volatile_only_ish_changed(&owner_before, &owner_after, &[2, 1]);

    // The structural patch inverse is the stronger exact-image assertion: it
    // restores the same source ZIP, including OPC lexical sources.
    assert_eq!(
        commit
            .package()
            .apply(&commit.patch().inverse())
            .map(litchi_xlsb::cell_values::structure::Commit::into_package)
            .expect("structural inverse")
            .to_bytes()
            .expect("structural inverse bytes"),
        before
    );
}

#[test]
fn opaque_reorder_dependencies_are_typed_guards_and_leave_sources_unchanged() {
    let cases = [
        (
            "calculation chain",
            package_with_calc_chain(relationship_type::CALC_CHAIN),
        ),
        (
            "strict calculation chain",
            package_with_calc_chain(relationship_type::STRICT_CALC_CHAIN),
        ),
        ("VBA project", package_with_vba_project()),
        ("shared workbook", package_with_shared_workbook()),
        ("connections", package_with_connections()),
        ("data model", package_with_data_model()),
        (
            "data model and connections",
            package_with_data_model_and_connections(),
        ),
    ];
    for (label, package) in cases {
        assert_reorder_guard(package, label);
    }
}

#[test]
fn volatile_owner_content_type_is_present_for_dependency_remapping() {
    let package = package_with_volatile_dependencies();
    let part = package
        .opc_package()
        .get_part(&owner_uri())
        .expect("volatile owner");
    assert_eq!(part.content_type(), VOLATILE_CONTENT_TYPE);
    let workbook = package
        .opc_package()
        .main_document_part()
        .expect("workbook part");
    assert!(
        workbook
            .rels()
            .iter()
            .any(|relationship| relationship.reltype() == VOLATILE_RELATIONSHIP_TYPE)
    );
}

#[test]
fn opaque_volatile_record_blocks_reordering_but_retains_exact_noop() {
    let source = package_with_volatile_dependencies();
    let mut opc = source.opc_package().clone();
    let before = owner_blob(&source);
    let mut output = Vec::new();
    let mut writer = litchi_xlsb::raw::Encoder::new(&mut output, litchi_xlsb::raw::Limits::DEFAULT);
    for record in Records::new(&before, litchi_xlsb::raw::Limits::DEFAULT) {
        let record = record.unwrap();
        writer
            .write_record(record.kind(), record.payload())
            .unwrap();
    }
    writer
        .write_record(Kind::new(0x3ffe).unwrap(), &[0, 1])
        .unwrap();
    opc.get_part_mut(&owner_uri()).unwrap().set_blob(output);
    assert_reorder_guard(Package::from_opc(opc).unwrap(), "opaque volatile record");
}

#[test]
fn orphan_volatile_part_is_not_rewritten() {
    let source = package_with_volatile_dependencies();
    let mut opc = source.opc_package().clone();
    opc.try_add_part(Box::new(BlobPart::new(
        PackURI::new("/xl/orphanVolatile.bin").unwrap(),
        VOLATILE_CONTENT_TYPE.to_owned(),
        owner_blob(&source),
    )))
    .unwrap();
    let source = Package::from_opc(opc).unwrap();
    let before = source.to_bytes().unwrap();
    let mut edit = source.edit().unwrap();
    edit.reorder_sheets(&selectors(&["C", "A", "B"])).unwrap();
    assert!(edit.commit().is_err());
    assert_eq!(source.to_bytes().unwrap(), before);
}

#[test]
fn orphan_ordinal_dependency_parts_are_guarded() {
    for content_type in [
        litchi_xlsb::slicer::package::CACHE_CONTENT_TYPE,
        litchi_xlsb::slicer::package::VIEWS_CONTENT_TYPE,
        "application/vnd.ms-excel.TimelineCache+xml",
        "application/vnd.ms-excel.Timeline+xml",
        "application/vnd.ms-excel.revisionLog",
    ] {
        let package = worksheets(&["A", "B", "C"]);
        let mut opc = package.opc_package().clone();
        opc.try_add_part(Box::new(BlobPart::new(
            PackURI::new("/xl/opaque.bin").unwrap(),
            content_type.to_owned(),
            b"<opaque/>".to_vec(),
        )))
        .unwrap();
        assert_reorder_guard(Package::from_opc(opc).unwrap(), content_type);
    }
}
