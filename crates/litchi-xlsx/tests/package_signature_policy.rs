#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    clippy::too_many_lines,
    reason = "integration tests use panic-on-failure extraction"
)]

//! The package-level families refuse a change to a signed package instead of
//! removing its signature (ADR 0006; unified-ops record 0014b).
//!
//! Each family is checked three ways: a change to a signed package is refused
//! with [`Error::Signed`] and leaves the package as it was (it still publishes
//! its signed source byte for byte); an exact no-op keeps the signature byte
//! for byte; and after the explicit removal (`OpcPackage::unsign` or
//! `Package::unsign`) the same change publishes a package without signature
//! parts. Changes to the producer-signed fixtures are stores; removals and
//! patches run against a package that already holds the family's content and
//! is signed in the test.

use std::path::PathBuf;

use litchi_opc::constants::relationship_type as rt;
use litchi_opc::{BlobPart, OpcPackage, PackURI, PackageWriter, Part as _, TargetMode};
use litchi_xlsx::Error;

/// The producer-signed fixtures: one signature by `xmlsec`, one by Office 2010.
const SIGNED: [&str; 2] = ["hello-world-signed.xlsx", "ms-office-2010-signed.xlsx"];

fn read(path: &str) -> Vec<u8> {
    std::fs::read(
        PathBuf::from(env!("CARGO_MANIFEST_DIR"))
            .join("../..")
            .join(path),
    )
    .expect(path)
}

fn signed_fixture(name: &str) -> Vec<u8> {
    read(&format!("test-data/poi/test-data/xmldsign/{name}"))
}

/// An owned open, so an untouched package republishes its exact source bytes.
fn open(bytes: &[u8]) -> OpcPackage {
    OpcPackage::from_vec(bytes.to_vec()).unwrap()
}

fn publish(package: &OpcPackage) -> Vec<u8> {
    PackageWriter::to_bytes(package).unwrap()
}

/// `plain` with an OPC signature graph (origin part, one XML signature and
/// their relationships), written out so the signature is an ingressed source
/// signature, as in a file a producer signed.
fn sign(plain: &OpcPackage) -> Vec<u8> {
    let mut package = OpcPackage::new();
    for relationship in plain.rels().iter() {
        package
            .rels_mut()
            .try_add_relationship(
                relationship.reltype().to_owned(),
                relationship.target_ref().to_owned(),
                relationship.r_id().to_owned(),
                relationship.target_mode(),
            )
            .unwrap();
    }
    let names: Vec<_> = plain
        .iter_parts()
        .map(|part| part.partname().clone())
        .collect();
    for name in names {
        let part = plain.get_part(&name).unwrap();
        let mut copy = BlobPart::new(
            part.partname().clone(),
            part.content_type().to_owned(),
            part.blob().to_vec(),
        );
        for relationship in part.rels().iter() {
            copy.rels_mut()
                .try_add_relationship(
                    relationship.reltype().to_owned(),
                    relationship.target_ref().to_owned(),
                    relationship.r_id().to_owned(),
                    relationship.target_mode(),
                )
                .unwrap();
        }
        package.try_add_part(Box::new(copy)).unwrap();
    }
    let origin = PackURI::new("/_xmlsignatures/origin.sigs").unwrap();
    let signature = PackURI::new("/_xmlsignatures/sig1.xml").unwrap();
    let mut origin_part = BlobPart::new(
        origin.clone(),
        litchi_opc::constants::content_type::OPC_DIGITAL_SIGNATURE_ORIGIN.to_owned(),
        Vec::new(),
    );
    origin_part
        .rels_mut()
        .try_add_relationship(
            "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/signature"
                .to_owned(),
            signature.relative_ref(origin.base_uri()),
            "rIdSignature".to_owned(),
            TargetMode::Internal,
        )
        .unwrap();
    package.try_add_part(Box::new(origin_part)).unwrap();
    package
        .try_add_part(Box::new(BlobPart::new(
            signature,
            litchi_opc::constants::content_type::OPC_DIGITAL_SIGNATURE_XMLSIGNATURE.to_owned(),
            b"<Signature xmlns=\"http://www.w3.org/2000/09/xmldsig#\"/>".to_vec(),
        )))
        .unwrap();
    package
        .rels_mut()
        .try_add_relationship(
            rt::DIGITAL_SIGNATURE_ORIGIN.to_owned(),
            "_xmlsignatures/origin.sigs".to_owned(),
            "rIdSignatureOrigin".to_owned(),
            TargetMode::Internal,
        )
        .unwrap();
    let bytes = publish(&package);
    assert!(open(&bytes).is_signed());
    bytes
}

/// `signed` with its signature graph deleted part by part instead of removed
/// with `unsign`. Deleting signature parts does not authorize a change: the
/// OPC edit-policy requirement survives, so the families still refuse. The
/// package root keeps the relationships of the unsigned package `signed` was
/// made from, so a patch planned on that package is not stale here.
fn signature_deleted_by_hand(signed: &[u8]) -> OpcPackage {
    let mut package = open(signed);
    assert!(package.rels_mut().remove("rIdSignatureOrigin").is_some());
    for name in ["/_xmlsignatures/sig1.xml", "/_xmlsignatures/origin.sigs"] {
        assert!(package.remove_part(&PackURI::new(name).unwrap()));
    }
    assert!(!package.is_signed());
    assert!(package.requires_signature_edit_policy());
    package
}

/// An unsigned copy of a producer-signed fixture, removed explicitly.
fn unsigned_fixture(name: &str) -> OpcPackage {
    let mut package = open(&signed_fixture(name));
    package.unsign();
    package
}

/// Whether `bytes` still carries any OPC signature infrastructure.
fn carries_signature(bytes: &[u8]) -> bool {
    let package = open(bytes);
    package.is_signed()
        || package
            .iter_parts()
            .any(|part| part.partname().as_str().starts_with("/_xmlsignatures/"))
}

fn first_worksheet(package: &OpcPackage) -> PackURI {
    let workbook = package.main_document_part().unwrap();
    workbook
        .rels()
        .iter()
        .find(|relationship| relationship.reltype() == rt::WORKSHEET)
        .expect("a first worksheet")
        .target_partname()
        .unwrap()
}

#[track_caller]
fn assert_signed<T: std::fmt::Debug>(result: litchi_xlsx::Result<T>) {
    let error = result.expect_err("a change to a signed package must be refused");
    assert!(matches!(error, Error::Signed), "{error:?}");
}

#[track_caller]
fn assert_boxed_signed<T: std::fmt::Debug>(
    result: Result<T, Box<dyn std::error::Error + Send + Sync>>,
) {
    let error = result.expect_err("a change to a signed package must be refused");
    assert!(
        matches!(error.downcast_ref::<Error>(), Some(Error::Signed)),
        "{error:?}"
    );
}

mod connections_family {
    use super::*;
    use litchi_xlsx::connections::{self, Connection, Connections, CredentialsMethod, Transaction};

    pub(super) fn connection(id: u32, name: &str) -> Connection {
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

    pub(super) fn catalog() -> Connections {
        Connections {
            connections: vec![connection(1, "Inert"), connection(2, "Second")],
        }
    }

    #[test]
    fn a_signed_fixture_refuses_a_store_and_keeps_exact_noops() {
        for name in SIGNED {
            let original = signed_fixture(name);
            let mut package = open(&original);

            assert!(!connections::remove_from_package(&mut package).unwrap());
            let commit = Transaction::new(&mut package).unwrap().commit().unwrap();
            assert!(!commit.changed());
            assert!(commit.patch().is_empty());
            assert_eq!(publish(&package), original, "{name}: exact no-ops");

            assert_boxed_signed(connections::store_in_package(
                &mut package,
                &catalog(),
                false,
            ));
            let mut transaction = Transaction::new(&mut package).unwrap();
            assert!(transaction.set(connection(3, "Staged")).unwrap());
            assert_boxed_signed(transaction.commit());
            assert_eq!(publish(&package), original, "{name}: refusals");

            package.unsign();
            connections::store_in_package(&mut package, &catalog(), false).unwrap();
            let output = publish(&package);
            assert!(!carries_signature(&output), "{name}");
            assert_eq!(
                connections::load_from_package(&open(&output)).unwrap(),
                Some(catalog())
            );
        }
    }

    #[test]
    fn a_signed_package_refuses_a_removal_and_a_patch() {
        let mut plain = unsigned_fixture(SIGNED[0]);
        connections::store_in_package(&mut plain, &catalog(), false).unwrap();
        let signed = sign(&plain);
        let mut package = open(&signed);

        assert_boxed_signed(connections::remove_from_package(&mut package));
        assert_eq!(publish(&package), signed);

        // The snapshot binds the package root, so the patch targets the same
        // graph with its signature parts deleted by hand.
        let patch = {
            let mut transaction = Transaction::new(&mut plain).unwrap();
            assert!(transaction.remove(1).unwrap().is_some());
            transaction.commit().unwrap().patch().clone()
        };
        let mut package = signature_deleted_by_hand(&signed);
        assert_boxed_signed(patch.apply(&mut package));
        assert!(package.requires_signature_edit_policy());
        assert_eq!(
            connections::load_from_package(&package)
                .unwrap()
                .unwrap()
                .connections
                .len(),
            2
        );

        package.unsign();
        patch.apply(&mut package).unwrap();
        let output = publish(&package);
        assert!(!carries_signature(&output));
        assert_eq!(
            connections::load_from_package(&open(&output))
                .unwrap()
                .unwrap()
                .connections
                .len(),
            1
        );
    }
}

mod query_table_family {
    use super::connections_family::catalog;
    use super::*;
    use litchi_xlsx::connections;
    use litchi_xlsx::query_table::{Conformance, Table};
    use litchi_xlsx::{
        add_worksheet_query_table, load_worksheet_query_tables, remove_worksheet_query_table,
        reorder_worksheet_query_tables, replace_worksheet_query_table,
    };

    fn plain_with_connections() -> OpcPackage {
        let mut plain = unsigned_fixture(SIGNED[0]);
        connections::store_in_package(&mut plain, &catalog(), false).unwrap();
        plain
    }

    #[test]
    fn a_signed_package_refuses_query_table_changes_and_keeps_exact_noops() {
        let mut plain = plain_with_connections();
        let worksheet = first_worksheet(&plain);
        let signed = sign(&plain);
        let mut package = open(&signed);

        // No query tables: removing and reordering publish nothing.
        assert!(!remove_worksheet_query_table(&mut package, &worksheet, "rIdQueryTable1").unwrap());
        assert!(
            reorder_worksheet_query_tables(&mut package, &worksheet, &[])
                .unwrap()
                .is_empty()
        );
        assert_eq!(publish(&package), signed, "exact no-ops");

        assert_signed(add_worksheet_query_table(
            &mut package,
            &worksheet,
            Table::new("Query", 1),
            Conformance::Transitional,
        ));
        assert_eq!(publish(&package), signed, "refused add");

        // The same content, signed with two query tables in canonical order.
        add_worksheet_query_table(
            &mut plain,
            &worksheet,
            Table::new("First", 1),
            Conformance::Transitional,
        )
        .unwrap();
        add_worksheet_query_table(
            &mut plain,
            &worksheet,
            Table::new("Second", 2),
            Conformance::Transitional,
        )
        .unwrap();
        let ids = reorder_worksheet_query_tables(
            &mut plain,
            &worksheet,
            &["rIdQueryTable1".into(), "rIdQueryTable2".into()],
        )
        .unwrap()
        .iter()
        .map(|part| part.relationship_id().to_owned())
        .collect::<Vec<_>>();
        assert_eq!(ids, ["rIdQueryTable1", "rIdQueryTable2"]);
        let signed = sign(&plain);
        let mut package = open(&signed);

        let unchanged = reorder_worksheet_query_tables(&mut package, &worksheet, &ids).unwrap();
        assert_eq!(unchanged.len(), 2);
        assert_eq!(publish(&package), signed, "reorder into the current order");

        let swapped = [ids[1].clone(), ids[0].clone()];
        assert_signed(reorder_worksheet_query_tables(
            &mut package,
            &worksheet,
            &swapped,
        ));
        assert_signed(replace_worksheet_query_table(
            &mut package,
            &worksheet,
            &ids[0],
            Table::new("Replaced", 1),
            Conformance::Transitional,
        ));
        assert_signed(remove_worksheet_query_table(
            &mut package,
            &worksheet,
            &ids[0],
        ));
        assert_eq!(publish(&package), signed, "refused changes");

        package.unsign();
        let reordered = reorder_worksheet_query_tables(&mut package, &worksheet, &swapped).unwrap();
        assert_eq!(reordered[0].query_table().name(), "Second");
        assert!(remove_worksheet_query_table(&mut package, &worksheet, &ids[1]).unwrap());
        let output = publish(&package);
        assert!(!carries_signature(&output));
        assert_eq!(
            load_worksheet_query_tables(&open(&output), &worksheet)
                .unwrap()
                .len(),
            1
        );
    }
}

mod xml_maps_family {
    use super::*;
    use litchi_xlsx::xml_maps::{
        self, Transaction, XmlMap, XmlMapConformance, XmlMapInfo, XmlMapSchema,
    };

    fn info(name: &str) -> XmlMapInfo {
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
                name: name.into(),
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
    fn a_signed_fixture_refuses_a_store_and_keeps_exact_noops() {
        for name in SIGNED {
            let original = signed_fixture(name);
            let mut package = open(&original);

            assert!(!xml_maps::remove_from_package(&mut package).unwrap());
            let commit = Transaction::new(&mut package).unwrap().commit().unwrap();
            assert!(!commit.changed());
            assert_eq!(publish(&package), original, "{name}: exact no-ops");

            assert_boxed_signed(xml_maps::store_in_package(
                &mut package,
                &info("Map"),
                XmlMapConformance::Transitional,
            ));
            assert_eq!(publish(&package), original, "{name}: refused store");

            package.unsign();
            xml_maps::store_in_package(&mut package, &info("Map"), XmlMapConformance::Transitional)
                .unwrap();
            let output = publish(&package);
            assert!(!carries_signature(&output), "{name}");
            assert_eq!(
                xml_maps::load_from_package(&open(&output)).unwrap(),
                Some(info("Map"))
            );
        }
    }

    #[test]
    fn a_signed_package_keeps_an_equal_store_and_refuses_changes() {
        let mut plain = unsigned_fixture(SIGNED[1]);
        xml_maps::store_in_package(&mut plain, &info("Map"), XmlMapConformance::Transitional)
            .unwrap();
        let signed = sign(&plain);
        let mut package = open(&signed);

        xml_maps::store_in_package(&mut package, &info("Map"), XmlMapConformance::Transitional)
            .unwrap();
        assert_eq!(publish(&package), signed, "an equal store");

        assert_boxed_signed(xml_maps::store_in_package(
            &mut package,
            &info("Renamed"),
            XmlMapConformance::Transitional,
        ));
        assert_boxed_signed(xml_maps::remove_from_package(&mut package));
        let mut transaction = Transaction::new(&mut package).unwrap();
        assert!(transaction.replace(None).unwrap());
        assert_boxed_signed(transaction.commit());
        assert_eq!(publish(&package), signed, "refused changes");

        // The snapshot binds the package root, so the patch targets the same
        // graph with its signature parts deleted by hand.
        let patch = {
            let mut transaction = Transaction::new(&mut plain).unwrap();
            assert!(transaction.replace(Some(info("Patched"))).unwrap());
            transaction.commit().unwrap().patch().clone()
        };
        let mut package = signature_deleted_by_hand(&signed);
        assert_boxed_signed(patch.apply(&mut package));
        assert_eq!(
            xml_maps::load_from_package(&package).unwrap(),
            Some(info("Map"))
        );

        package.unsign();
        patch.apply(&mut package).unwrap();
        let output = publish(&package);
        assert!(!carries_signature(&output));
        assert_eq!(
            xml_maps::load_from_package(&open(&output)).unwrap(),
            Some(info("Patched"))
        );
    }
}

mod volatile_dependencies_family {
    use super::*;
    use litchi_xlsx::volatile_dependencies::{
        self, VolatileDependencies, VolatileDependenciesConformance,
    };

    fn value() -> VolatileDependencies {
        VolatileDependencies::parse(
            br#"<volTypes xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><volType type="realTimeData"><main first="server.id"><tp t="s"><v>ready</v><tr r="A1" s="0"/></tp></main></volType></volTypes>"#,
        )
        .unwrap()
    }

    #[test]
    fn a_signed_package_refuses_changes_and_keeps_exact_noops() {
        for name in SIGNED {
            let original = signed_fixture(name);
            let mut package = open(&original);
            assert!(!volatile_dependencies::remove_from_package(&mut package).unwrap());
            assert_boxed_signed(volatile_dependencies::store_in_package(
                &mut package,
                &value(),
                VolatileDependenciesConformance::Transitional,
            ));
            assert_eq!(publish(&package), original, "{name}");

            package.unsign();
            volatile_dependencies::store_in_package(
                &mut package,
                &value(),
                VolatileDependenciesConformance::Transitional,
            )
            .unwrap();
            let output = publish(&package);
            assert!(!carries_signature(&output), "{name}");
            assert_eq!(
                volatile_dependencies::load_from_package(&open(&output)).unwrap(),
                Some(value())
            );
        }

        let mut plain = unsigned_fixture(SIGNED[0]);
        volatile_dependencies::store_in_package(
            &mut plain,
            &value(),
            VolatileDependenciesConformance::Transitional,
        )
        .unwrap();
        let signed = sign(&plain);
        let mut package = open(&signed);
        assert_boxed_signed(volatile_dependencies::remove_from_package(&mut package));
        assert_eq!(publish(&package), signed);
        package.unsign();
        assert!(volatile_dependencies::remove_from_package(&mut package).unwrap());
        assert!(!carries_signature(&publish(&package)));
    }
}

mod comments_family {
    use super::*;
    use litchi_xlsx::workbook::comments::{
        self, Comments, Transaction, load_from_worksheet, remove_from_worksheet, store_on_worksheet,
    };

    fn note(package: &OpcPackage, worksheet: &PackURI, text: &str) -> Comments {
        let mut staged = package.clone();
        let mut transaction = Transaction::new(&mut staged, worksheet).unwrap();
        transaction.set("B2", "Reviewer", text).unwrap();
        transaction.comments().unwrap().clone()
    }

    #[test]
    fn a_signed_package_refuses_changes_and_keeps_exact_noops() {
        for name in SIGNED {
            let original = signed_fixture(name);
            let mut package = open(&original);
            let worksheet = first_worksheet(&package);
            let value = note(&package, &worksheet, "Checked");

            assert!(!remove_from_worksheet(&mut package, &worksheet).unwrap());
            let commit = Transaction::new(&mut package, &worksheet)
                .unwrap()
                .commit()
                .unwrap();
            assert!(!commit.changed());
            assert_eq!(publish(&package), original, "{name}: exact no-ops");

            assert_signed(store_on_worksheet(&mut package, &worksheet, &value));
            let mut transaction = Transaction::new(&mut package, &worksheet).unwrap();
            assert!(transaction.set("C3", "Reviewer", "Staged").unwrap());
            assert_signed(transaction.commit());
            assert_eq!(publish(&package), original, "{name}: refusals");

            package.unsign();
            store_on_worksheet(&mut package, &worksheet, &value).unwrap();
            let output = publish(&package);
            assert!(!carries_signature(&output), "{name}");
            assert_eq!(
                load_from_worksheet(&open(&output), &worksheet)
                    .unwrap()
                    .unwrap()
                    .comments,
                value
            );
        }
    }

    #[test]
    fn a_signed_package_refuses_a_removal_and_a_patch() {
        let mut plain = unsigned_fixture(SIGNED[1]);
        let worksheet = first_worksheet(&plain);
        let value = note(&plain, &worksheet, "Signed note");
        store_on_worksheet(&mut plain, &worksheet, &value).unwrap();
        let signed = sign(&plain);
        let mut package = open(&signed);

        assert_signed(remove_from_worksheet(&mut package, &worksheet));
        let patch = {
            let mut transaction = Transaction::new(&mut plain, &worksheet).unwrap();
            assert!(transaction.clear());
            transaction.commit().unwrap().patch().clone()
        };
        assert_signed(patch.apply(&mut package));
        assert_eq!(publish(&package), signed);
        comments::validate_graph(&package).unwrap();

        package.unsign();
        patch.apply(&mut package).unwrap();
        let output = publish(&package);
        assert!(!carries_signature(&output));
        assert!(
            load_from_worksheet(&open(&output), &worksheet)
                .unwrap()
                .is_none()
        );
    }
}

mod revisions_family {
    use super::*;
    use litchi_xlsx::revisions::{
        RevisionConformance, RevisionUser, Transaction, load_workbook_revisions,
        remove_workbook_revisions, store_workbook_revisions,
    };

    const TRACKED: &str =
        "test-data/libreoffice-core/sc/qa/unit/data/xlsx/track-changes/simple-cell-changes.xlsx";

    fn user(name: &str) -> RevisionUser {
        RevisionUser {
            guid: "{00000000-0000-0000-0000-000000000009}".into(),
            name: name.into(),
            id: 7,
            date_time: "2026-07-17T12:00:00Z".into(),
            extension_elements: vec![],
        }
    }

    #[test]
    fn a_signed_fixture_refuses_a_store_and_keeps_exact_noops() {
        let mut value = load_workbook_revisions(&open(&read(TRACKED)))
            .unwrap()
            .expect("tracked revisions");
        value.users_relationship_id = "rIdRevisionUsers".into();
        value.headers_relationship_id = "rIdRevisionHeaders".into();
        for name in SIGNED {
            let original = signed_fixture(name);
            let mut package = open(&original);
            assert!(!remove_workbook_revisions(&mut package).unwrap());
            let commit = Transaction::new(&mut package).unwrap().commit().unwrap();
            assert!(!commit.changed());
            assert_eq!(publish(&package), original, "{name}: exact no-ops");

            assert_signed(store_workbook_revisions(
                &mut package,
                &value,
                RevisionConformance::Transitional,
            ));
            assert_eq!(publish(&package), original, "{name}: refused store");

            package.unsign();
            store_workbook_revisions(&mut package, &value, RevisionConformance::Transitional)
                .unwrap();
            let output = publish(&package);
            assert!(!carries_signature(&output), "{name}");
            let stored = load_workbook_revisions(&open(&output)).unwrap().unwrap();
            assert_eq!(stored.headers, value.headers, "{name}");
            assert_eq!(stored.logs.len(), value.logs.len(), "{name}");
        }
    }

    #[test]
    fn a_signed_package_refuses_edits_removal_and_patches() {
        let mut plain = open(&read(TRACKED));
        let signed = sign(&plain);
        let mut package = open(&signed);

        let commit = Transaction::new(&mut package).unwrap().commit().unwrap();
        assert!(!commit.changed());
        assert_eq!(publish(&package), signed, "an exact no-op");

        let mut transaction = Transaction::new(&mut package).unwrap();
        assert!(transaction.set_user(user("Alice")).unwrap());
        assert_signed(transaction.commit());
        assert_signed(remove_workbook_revisions(&mut package));
        let patch = {
            let mut transaction = Transaction::new(&mut plain).unwrap();
            assert!(transaction.set_user(user("Patched")).unwrap());
            transaction.commit().unwrap().patch().clone()
        };
        assert_signed(patch.apply(&mut package));
        assert_eq!(publish(&package), signed, "refused changes");

        package.unsign();
        patch.apply(&mut package).unwrap();
        let output = publish(&package);
        assert!(!carries_signature(&output));
        let users = load_workbook_revisions(&open(&output))
            .unwrap()
            .unwrap()
            .users
            .users;
        assert!(users.iter().any(|user| user.name == "Patched"));
    }
}

mod named_sheet_view_family {
    use super::*;
    use litchi_xlsx::named_sheet_view::{
        Guid, View, Views, load_worksheet_named_sheet_views, remove_worksheet_named_sheet_views,
        store_worksheet_named_sheet_views,
    };

    fn views(name: &str) -> Views {
        Views::new(
            View::with_id(
                name,
                Guid::new("{01234567-89AB-CDEF-0123-456789ABCDEF}").unwrap(),
            )
            .unwrap(),
        )
    }

    #[test]
    fn a_signed_package_refuses_changes_and_keeps_exact_noops() {
        for name in SIGNED {
            let original = signed_fixture(name);
            let mut package = open(&original);
            let worksheet = first_worksheet(&package);
            assert!(!remove_worksheet_named_sheet_views(&mut package, &worksheet).unwrap());
            assert_signed(store_worksheet_named_sheet_views(
                &mut package,
                &worksheet,
                &views("Data"),
            ));
            assert_eq!(publish(&package), original, "{name}");

            package.unsign();
            store_worksheet_named_sheet_views(&mut package, &worksheet, &views("Data")).unwrap();
            let output = publish(&package);
            assert!(!carries_signature(&output), "{name}");
            assert!(
                load_worksheet_named_sheet_views(&open(&output), &worksheet)
                    .unwrap()
                    .is_some()
            );
        }

        let mut plain = unsigned_fixture(SIGNED[1]);
        let worksheet = first_worksheet(&plain);
        store_worksheet_named_sheet_views(&mut plain, &worksheet, &views("Data")).unwrap();
        let signed = sign(&plain);
        let mut package = open(&signed);
        assert_signed(store_worksheet_named_sheet_views(
            &mut package,
            &worksheet,
            &views("Renamed"),
        ));
        assert_signed(remove_worksheet_named_sheet_views(&mut package, &worksheet));
        assert_eq!(publish(&package), signed);
        package.unsign();
        assert!(remove_worksheet_named_sheet_views(&mut package, &worksheet).unwrap());
        assert!(!carries_signature(&publish(&package)));
    }
}

mod active_x_family {
    use super::*;
    use litchi_xlsx::{
        Binary, Control, ControlSet, Descriptor, LoadedControl, Persistence, load_from_worksheet,
        remove_from_worksheet, replace_on_worksheet, store_on_worksheet,
    };

    fn controls(name: &str) -> ControlSet {
        ControlSet {
            controls: vec![LoadedControl {
                control: Control {
                    shape_id: 42,
                    relationship_id: "rIdControlSigned".into(),
                    name: Some(name.into()),
                    properties: None,
                },
                descriptor_uri: PackURI::new("/xl/activeX/signedControl.xml").unwrap(),
                descriptor: Descriptor {
                    class_id: "{00000000-0000-0000-0000-000000000000}".into(),
                    license: None,
                    persistence: Persistence::StreamInit,
                    relationship_id: Some("rIdBinary".into()),
                    properties: Vec::new(),
                },
                binaries: vec![Binary {
                    relationship_id: "rIdBinary".into(),
                    part_uri: PackURI::new("/xl/activeX/signedControl.bin").unwrap(),
                    bytes: vec![0, 1, 2, 0xff],
                }],
                preview: None,
            }],
        }
    }

    #[test]
    fn a_signed_package_refuses_changes_and_keeps_exact_noops() {
        for name in SIGNED {
            let original = signed_fixture(name);
            let mut package = open(&original);
            let worksheet = first_worksheet(&package);
            assert!(!remove_from_worksheet(&mut package, &worksheet).unwrap());
            assert_signed(store_on_worksheet(
                &mut package,
                &worksheet,
                &controls("Button"),
            ));
            assert_signed(replace_on_worksheet(
                &mut package,
                &worksheet,
                &controls("Button"),
            ));
            assert_eq!(publish(&package), original, "{name}");

            package.unsign();
            store_on_worksheet(&mut package, &worksheet, &controls("Button")).unwrap();
            let output = publish(&package);
            assert!(!carries_signature(&output), "{name}");
            assert_eq!(
                load_from_worksheet(&open(&output), &worksheet).unwrap(),
                controls("Button")
            );
        }

        let mut plain = unsigned_fixture(SIGNED[0]);
        let worksheet = first_worksheet(&plain);
        store_on_worksheet(&mut plain, &worksheet, &controls("Button")).unwrap();
        let signed = sign(&plain);
        let mut package = open(&signed);
        assert_signed(remove_from_worksheet(&mut package, &worksheet));
        assert_signed(replace_on_worksheet(
            &mut package,
            &worksheet,
            &controls("Renamed"),
        ));
        assert_eq!(publish(&package), signed);
        package.unsign();
        assert!(remove_from_worksheet(&mut package, &worksheet).unwrap());
        assert!(!carries_signature(&publish(&package)));
    }
}

mod smart_tags_family {
    use super::*;
    use litchi_xlsx::Package;
    use litchi_xlsx::smart_tags::{Cell, Collection, Tag};

    fn cell() -> Cell {
        Cell::new("A1", vec![Tag::new(7).unwrap()]).unwrap()
    }

    #[test]
    fn a_signed_workbook_refuses_changes_and_keeps_exact_noops() {
        for name in SIGNED {
            let original = signed_fixture(name);
            let mut package = Package::from_bytes(&original).unwrap();
            assert!(package.is_signed());

            let mut transaction = package.edit_smart_tags(0_usize).unwrap();
            assert!(transaction.remove("A1").unwrap().is_none());
            assert!(!transaction.clear().unwrap());
            transaction.replace(Collection::default()).unwrap();
            assert_signed(transaction.set(cell()));
            package = transaction.commit().unwrap().into_package();
            assert_eq!(package.to_bytes().unwrap(), original, "{name}");

            package = package.unsign().unwrap();
            assert!(!package.is_signed());
            let mut transaction = package.edit_smart_tags(0_usize).unwrap();
            transaction.set(cell()).unwrap();
            package = transaction.commit().unwrap().into_package();
            let output = package.to_bytes().unwrap();
            assert!(!carries_signature(&output), "{name}");
            let reopened = Package::from_bytes(&output).unwrap();
            assert_eq!(reopened.smart_tags(0_usize).unwrap().unwrap().len(), 1);
        }

        let mut plain = Package::from_bytes(&signed_fixture(SIGNED[1])).unwrap();
        plain = plain.unsign().unwrap();
        let mut transaction = plain.edit_smart_tags(0_usize).unwrap();
        transaction.set(cell()).unwrap();
        plain = transaction.commit().unwrap().into_package();
        let signed = sign(&open(&plain.to_bytes().unwrap()));
        let mut package = Package::from_bytes(&signed).unwrap();
        let mut transaction = package.edit_smart_tags(0_usize).unwrap();
        assert_signed(transaction.remove("A1"));
        assert_signed(transaction.clear());
        package = transaction.commit().unwrap().into_package();
        assert_eq!(package.to_bytes().unwrap(), signed);
    }
}

mod slicer_and_timeline_family {
    use super::*;
    use litchi_xlsx::Package;
    use litchi_xlsx::slicer::{self, Definition, Slicer};
    use litchi_xlsx::timeline::{self, CacheDefinition, FilterType, Level, Range, State, View};

    const X14: &str = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";

    fn slicer_cache(name: &str) -> Definition {
        slicer::read(
            format!(r#"<slicerCacheDefinition xmlns="{X14}" name="{name}" sourceName="State"><data><tabular pivotCacheId="5"><items count="2"><i x="1"/><i x="0" s="1"/></items></tabular></data></slicerCacheDefinition>"#).as_bytes(),
        )
        .unwrap()
    }

    fn timeline_cache(name: &str) -> CacheDefinition {
        CacheDefinition {
            name: name.into(),
            uid: None,
            source_name: "Date".into(),
            pivot_tables: Vec::new(),
            state: State {
                selection: Some(
                    Range::new("2026-01-01T00:00:00Z", "2026-01-31T23:59:59Z").unwrap(),
                ),
                bounds: Range::new("2026-01-01T00:00:00Z", "2026-12-31T23:59:59Z").unwrap(),
                extension_list: None,
                single_range_filter_state: Some(true),
                minimal_refresh_version: 0,
                last_refresh_version: 1,
                pivot_cache_id: 1,
                filter_type: FilterType::DateBetween,
            },
            timeline_pivot_filter: None,
            extension_list: None,
        }
    }

    fn timeline_view(name: &str, cache: &str) -> View {
        View {
            name: name.into(),
            uid: None,
            cache: cache.into(),
            caption: Some(name.into()),
            show_header: Some(true),
            show_selection_label: Some(true),
            show_time_level: Some(true),
            show_horizontal_scrollbar: Some(true),
            level: Level::Month,
            selection_level: Level::Day,
            scroll_position: Some("2026-01-01T00:00:00Z".into()),
            style: Some("TimelineStyleLight1".into()),
            extension_list: None,
        }
    }

    #[test]
    fn a_signed_workbook_refuses_slicer_and_timeline_changes() {
        for name in SIGNED {
            let original = signed_fixture(name);
            let mut package = Package::from_bytes(&original).unwrap();
            let worksheet = first_worksheet(&open(&original));

            let mut slicers = package.edit_slicers().unwrap();
            assert!(!slicers.remove_cache("Absent").unwrap());
            assert!(!slicers.remove_view(&worksheet, "Absent").unwrap());
            assert!(!slicers.update_cache("Absent", |_| {}).unwrap());
            // A worksheet without slicers has nothing to reorder.
            assert!(slicers.reorder_views(&worksheet, &[]).unwrap().is_empty());
            assert_signed(slicers.add_cache(slicer_cache("Cache_A")));
            package = slicers.commit().unwrap().into_package();

            let mut timelines = package.edit_timelines().unwrap();
            assert!(!timelines.remove_cache("Absent").unwrap());
            assert!(!timelines.remove_view(&worksheet, "Absent").unwrap());
            assert!(timelines.reorder_views(&worksheet, &[]).unwrap().is_empty());
            assert_signed(timelines.add_cache(timeline_cache("Timeline_A")));
            package = timelines.commit().unwrap().into_package();
            assert_eq!(package.to_bytes().unwrap(), original, "{name}");

            package = package.unsign().unwrap();
            let mut slicers = package.edit_slicers().unwrap();
            slicers.add_cache(slicer_cache("Cache_A")).unwrap();
            slicers
                .add_view(&worksheet, Slicer::new("View", "Cache_A", 228_600))
                .unwrap();
            package = slicers.commit().unwrap().into_package();
            let mut timelines = package.edit_timelines().unwrap();
            timelines.add_cache(timeline_cache("Timeline_A")).unwrap();
            timelines
                .add_view(&worksheet, timeline_view("Timeline", "Timeline_A"))
                .unwrap();
            package = timelines.commit().unwrap().into_package();
            let output = package.to_bytes().unwrap();
            assert!(!carries_signature(&output), "{name}");
            let reopened = open(&output);
            assert_eq!(slicer::load_caches(&reopened).unwrap().len(), 1);
            let workbook = reopened.main_document_part().unwrap().partname().clone();
            assert_eq!(
                timeline::load_caches(&reopened, &workbook).unwrap().len(),
                1
            );
        }
    }

    #[test]
    fn a_signed_workbook_with_slicers_and_timelines_refuses_their_changes() {
        let mut plain = Package::from_bytes(&signed_fixture(SIGNED[0])).unwrap();
        plain = plain.unsign().unwrap();
        let worksheet = first_worksheet(&open(&plain.to_bytes().unwrap()));
        let mut slicers = plain.edit_slicers().unwrap();
        slicers.add_cache(slicer_cache("Cache_A")).unwrap();
        slicers.add_cache(slicer_cache("Cache_B")).unwrap();
        slicers
            .add_view(&worksheet, Slicer::new("View_A", "Cache_A", 228_600))
            .unwrap();
        slicers
            .add_view(&worksheet, Slicer::new("View_B", "Cache_B", 228_600))
            .unwrap();
        plain = slicers.commit().unwrap().into_package();
        let mut timelines = plain.edit_timelines().unwrap();
        timelines.add_cache(timeline_cache("Timeline_A")).unwrap();
        timelines
            .add_view(&worksheet, timeline_view("Timeline", "Timeline_A"))
            .unwrap();
        plain = timelines.commit().unwrap().into_package();
        let signed = sign(&open(&plain.to_bytes().unwrap()));
        let mut package = Package::from_bytes(&signed).unwrap();

        let mut slicers = package.edit_slicers().unwrap();
        // Library-authored parts reserialize to their own bytes, so a reorder
        // into the current order publishes nothing.
        let order = [String::from("View_A"), String::from("View_B")];
        assert_eq!(slicers.reorder_views(&worksheet, &order).unwrap().len(), 2);
        let caches = [String::from("Cache_A"), String::from("Cache_B")];
        assert_eq!(slicers.reorder_caches(&caches).unwrap().len(), 2);
        assert_signed(slicers.reorder_caches(&[caches[1].clone(), caches[0].clone()]));
        assert_signed(slicers.reorder_views(&worksheet, &[order[1].clone(), order[0].clone()]));
        assert_signed(slicers.update_view(&worksheet, "View_A", |view| {
            view.caption = Some("Changed".into());
        }));
        assert_signed(slicers.remove_view(&worksheet, "View_B"));
        assert_signed(slicers.update_cache("Cache_A", |cache| {
            cache.source_name = "City".into();
        }));
        assert_signed(slicers.add_view(&worksheet, Slicer::new("View_C", "Cache_A", 228_600)));
        package = slicers.commit().unwrap().into_package();

        let mut timelines = package.edit_timelines().unwrap();
        let order = [String::from("Timeline")];
        assert_eq!(
            timelines.reorder_views(&worksheet, &order).unwrap().len(),
            1
        );
        // The store wrote the cache extension with another serializer than
        // the reorder's, so even the current order changes bytes: a change,
        // refused (record 0014b, F1c).
        assert_signed(timelines.reorder_caches(&[String::from("Timeline_A")]));
        assert_signed(timelines.update_view(&worksheet, "Timeline", |view| {
            view.caption = Some("Changed".into());
        }));
        assert_signed(timelines.remove_view(&worksheet, "Timeline"));
        assert_signed(timelines.update_cache("Timeline_A", |cache| {
            cache.source_name = "Shipped".into();
        }));
        package = timelines.commit().unwrap().into_package();
        assert_eq!(package.to_bytes().unwrap(), signed);

        package = package.unsign().unwrap();
        let mut slicers = package.edit_slicers().unwrap();
        assert!(slicers.remove_view(&worksheet, "View_B").unwrap());
        package = slicers.commit().unwrap().into_package();
        let mut timelines = package.edit_timelines().unwrap();
        assert!(timelines.remove_view(&worksheet, "Timeline").unwrap());
        package = timelines.commit().unwrap().into_package();
        assert!(!carries_signature(&package.to_bytes().unwrap()));
    }
}

#[test]
fn package_unsign_is_explicit_and_leaves_an_unsigned_package_untouched() {
    for name in SIGNED {
        let original = signed_fixture(name);
        let package = litchi_xlsx::Package::from_bytes(&original).unwrap();
        assert!(package.is_signed());
        assert_eq!(package.to_bytes().unwrap(), original);
        let stripped = package.unsign().unwrap();
        assert!(!stripped.is_signed());
        // The source package keeps its signature.
        assert!(package.is_signed());
        assert_eq!(package.to_bytes().unwrap(), original);
        let output = stripped.to_bytes().unwrap();
        assert!(!carries_signature(&output), "{name}");

        let unsigned = litchi_xlsx::Package::from_bytes(&output)
            .unwrap()
            .unsign()
            .unwrap();
        assert_eq!(unsigned.to_bytes().unwrap(), output, "{name}");
    }
}

/// The shared custom-properties and VBA owners of `litchi-ooxml-common`
/// refuse a change to a signed package with this crate's typed
/// [`Error::Signed`] instead of removing the signature (unified-ops record
/// 0014a, F1b).
mod shared_owners {
    use super::*;
    use litchi_xlsx::custom::{Host, Props};

    fn props() -> Props {
        let mut props = Props::new();
        props.insert("Project", "Litchi").unwrap();
        props
    }

    /// Stage `props` as the package's custom properties (empty removes them)
    /// through the metadata family and commit.
    fn set_custom_props(
        package: &litchi_xlsx::Package,
        props: Props,
    ) -> Result<litchi_xlsx::Package, litchi_xlsx::metadata::Error> {
        let mut edit = package.edit_metadata()?;
        edit.set_custom_props(props);
        Ok(edit.commit()?.into_package())
    }

    #[track_caller]
    fn assert_metadata_signed<T: std::fmt::Debug>(result: Result<T, litchi_xlsx::metadata::Error>) {
        let error = result.expect_err("a change to a signed package must be refused");
        assert!(
            matches!(error, litchi_xlsx::metadata::Error::Signed),
            "{error:?}"
        );
    }

    #[test]
    fn custom_properties_refuse_a_signed_fixture_and_keep_exact_noops() {
        for name in SIGNED {
            let original = signed_fixture(name);
            let package = litchi_xlsx::Package::from_bytes(&original).unwrap();
            let package = set_custom_props(&package, Props::new()).unwrap();
            assert_eq!(package.to_bytes().unwrap(), original, "{name}: no-ops");

            assert_metadata_signed(set_custom_props(&package, props()));
            assert!(package.is_signed());
            assert_eq!(package.to_bytes().unwrap(), original, "{name}: refusal");

            let package = set_custom_props(&package.unsign().unwrap(), props()).unwrap();
            let output = package.to_bytes().unwrap();
            assert!(!carries_signature(&output), "{name}");
            let reopened = litchi_xlsx::Package::from_bytes(&output).unwrap();
            assert!(reopened.custom_props().unwrap().get("Project").is_some());
        }
    }

    #[test]
    fn a_signed_package_keeps_equal_custom_properties_and_refuses_their_removal() {
        let mut plain = unsigned_fixture(SIGNED[0]);
        props().write_for(&mut plain, Host::Excel).unwrap();
        let signed = sign(&plain);
        let package = litchi_xlsx::Package::from_bytes(&signed).unwrap();

        let package = set_custom_props(&package, props()).unwrap();
        assert_eq!(package.to_bytes().unwrap(), signed, "an equal store");
        assert_metadata_signed(set_custom_props(&package, Props::new()));
        assert_eq!(package.to_bytes().unwrap(), signed, "the refusal");

        let package = set_custom_props(&package.unsign().unwrap(), Props::new()).unwrap();
        assert!(package.custom_props().unwrap().is_empty());
    }

    #[cfg(feature = "vba-inspection")]
    #[test]
    fn vba_projects_refuse_a_signed_package() {
        use litchi_vba::{Limits, build};
        use litchi_xlsx::active_x::vba;

        let payload = || {
            build::Project::new("ExcelProject")
                .module(build::Module::standard(
                    "Module1",
                    "Public Sub Hello()\r\nEnd Sub\r\n",
                ))
                .finish(&Limits::default())
                .unwrap()
        };
        for name in SIGNED {
            let original = signed_fixture(name);
            let mut package = open(&original);
            let workbook = package.main_document_part().unwrap().partname().clone();
            assert!(!vba::remove(&mut package, &workbook).unwrap());
            assert_eq!(publish(&package), original, "{name}: no-op");

            assert_signed(vba::store(&mut package, &workbook, payload()));
            assert_eq!(publish(&package), original, "{name}: refusal");

            package.unsign();
            vba::store(&mut package, &workbook, payload()).unwrap();
            assert!(!carries_signature(&publish(&package)), "{name}");
        }

        let mut plain = unsigned_fixture(SIGNED[0]);
        let workbook = plain.main_document_part().unwrap().partname().clone();
        vba::store(&mut plain, &workbook, payload()).unwrap();
        let signed = sign(&plain);
        let mut package = open(&signed);
        assert_signed(vba::remove(&mut package, &workbook));
        assert_eq!(publish(&package), signed, "the refusal");
        package.unsign();
        assert!(vba::remove(&mut package, &workbook).unwrap());
    }
}
