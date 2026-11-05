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
//! for byte; and after the explicit removal (`Package::unsign`) the same
//! change publishes a package without signature parts. Changes to the producer-signed fixtures are stores; removals and
//! patches run against a package that already holds the family's content and
//! is signed in the test.

use std::path::PathBuf;

use litchi_opc::constants::relationship_type as rt;
use litchi_opc::{
    BlobPart, OpcPackage, PackURI, PackageWriter, Part as _, SourceLimits, TargetMode,
};
use litchi_xlsx::{Error, Package};

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
    OpcPackage::from_vec(bytes.to_vec(), SourceLimits::default()).unwrap()
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

/// A package adopted from `bytes`; an untouched package republishes them.
fn adopt(bytes: &[u8]) -> Package {
    Package::from_bytes(bytes).unwrap()
}

fn bytes(package: &Package) -> Vec<u8> {
    package.to_bytes().unwrap()
}

/// An unsigned copy of a producer-signed fixture, removed explicitly.
fn unsigned_package(name: &str) -> Package {
    adopt(&signed_fixture(name)).unsign().unwrap()
}

/// `plain` signed in the test.
fn sign_package(plain: &Package) -> Vec<u8> {
    sign(&open(&bytes(plain)))
}

mod connections_family {
    use super::*;
    use litchi_xlsx::connections::{self, Connection, Connections, CredentialsMethod};

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

    pub(super) fn with_catalog(package: &Package) -> Package {
        let mut edit = package.edit_connections().unwrap();
        assert!(edit.replace(Some(catalog())).unwrap());
        edit.commit().unwrap().into_package()
    }

    fn load(package: &OpcPackage) -> litchi_xlsx::Result<Option<Connections>> {
        let limits = package.read_limits();
        connections::load_from_package(package, &limits)
    }

    #[test]
    fn a_signed_fixture_refuses_a_store_and_keeps_exact_noops() {
        for name in SIGNED {
            let original = signed_fixture(name);
            let package = adopt(&original);

            let mut edit = package.edit_connections().unwrap();
            assert!(!edit.replace(None).unwrap());
            let commit = edit.commit().unwrap();
            assert!(!commit.diagnostics().changed());
            assert!(commit.patch().is_noop());
            assert_eq!(bytes(commit.package()), original, "{name}: exact no-ops");

            let mut edit = package.edit_connections().unwrap();
            assert!(edit.replace(Some(catalog())).unwrap());
            assert_signed(edit.commit());
            let mut edit = package.edit_connections().unwrap();
            assert!(edit.set(connection(3, "Staged")).unwrap());
            assert_signed(edit.commit());
            assert_eq!(bytes(&package), original, "{name}: refusals");

            let output = bytes(&with_catalog(&package.unsign().unwrap()));
            assert!(!carries_signature(&output), "{name}");
            assert_eq!(load(&open(&output)).unwrap(), Some(catalog()));
        }
    }

    #[test]
    fn a_signed_package_refuses_a_removal_and_a_patch() {
        let plain = with_catalog(&unsigned_package(SIGNED[0]));
        let signed = sign_package(&plain);
        let package = adopt(&signed);

        let mut edit = package.edit_connections().unwrap();
        assert!(edit.replace(None).unwrap());
        assert_signed(edit.commit());
        assert_eq!(bytes(&package), signed);

        // A changed patch is refused as signed before its source state is
        // checked; after the explicit removal it targets the plain graph.
        let patch = {
            let mut edit = plain.edit_connections().unwrap();
            assert!(edit.remove(1).unwrap().is_some());
            edit.commit().unwrap().patch().clone()
        };
        assert_signed(package.apply_connections(&patch));
        assert_eq!(
            load(&open(&bytes(&package)))
                .unwrap()
                .unwrap()
                .connections
                .len(),
            2
        );

        let output = bytes(
            package
                .unsign()
                .unwrap()
                .apply_connections(&patch)
                .unwrap()
                .package(),
        );
        assert!(!carries_signature(&output));
        assert_eq!(load(&open(&output)).unwrap().unwrap().connections.len(), 1);
    }
}

mod query_table_family {
    use super::connections_family::with_catalog;
    use super::*;
    use litchi_xlsx::query_table::{Conformance, Table};

    #[test]
    fn a_signed_package_refuses_query_table_changes_and_keeps_exact_noops() {
        let plain = with_catalog(&unsigned_package(SIGNED[0]));
        let signed = sign_package(&plain);
        let package = adopt(&signed);

        // No query tables: removing and reordering publish nothing.
        let mut edit = package.edit_query_tables(0_usize).unwrap();
        assert!(!edit.remove("rIdQueryTable1").unwrap());
        assert!(edit.reorder(&[]).unwrap().is_empty());
        assert_signed(edit.add(Table::new("Query", 1), Conformance::Transitional));
        assert_eq!(
            bytes(&edit.commit().unwrap().into_package()),
            signed,
            "exact no-ops and a refused add"
        );

        // The same content, signed with two query tables in canonical order.
        let mut edit = plain.edit_query_tables(0_usize).unwrap();
        edit.add(Table::new("First", 1), Conformance::Transitional)
            .unwrap();
        edit.add(Table::new("Second", 2), Conformance::Transitional)
            .unwrap();
        let ids = edit
            .reorder(&["rIdQueryTable1".into(), "rIdQueryTable2".into()])
            .unwrap()
            .iter()
            .map(|part| part.relationship_id().to_owned())
            .collect::<Vec<_>>();
        assert_eq!(ids, ["rIdQueryTable1", "rIdQueryTable2"]);
        let plain = edit.commit().unwrap().into_package();
        let signed = sign_package(&plain);
        let package = adopt(&signed);

        let mut edit = package.edit_query_tables(0_usize).unwrap();
        assert_eq!(edit.reorder(&ids).unwrap().len(), 2);
        let swapped = [ids[1].clone(), ids[0].clone()];
        assert_signed(edit.reorder(&swapped));
        assert_signed(edit.replace(
            &ids[0],
            Table::new("Replaced", 1),
            Conformance::Transitional,
        ));
        assert_signed(edit.remove(&ids[0]));
        assert_eq!(
            bytes(&edit.commit().unwrap().into_package()),
            signed,
            "reorder into the current order and refused changes"
        );

        let mut edit = package
            .unsign()
            .unwrap()
            .edit_query_tables(0_usize)
            .unwrap();
        let reordered = edit.reorder(&swapped).unwrap();
        assert_eq!(reordered[0].query_table().name(), "Second");
        assert!(edit.remove(&ids[1]).unwrap());
        let output = bytes(&edit.commit().unwrap().into_package());
        assert!(!carries_signature(&output));
        assert_eq!(
            adopt(&output)
                .edit_query_tables(0_usize)
                .unwrap()
                .query_tables()
                .unwrap()
                .len(),
            1
        );
    }
}

mod xml_maps_family {
    use super::*;
    use litchi_xlsx::xml_maps::{self, XmlMap, XmlMapInfo, XmlMapSchema};

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

    fn store(package: &Package, value: XmlMapInfo) -> litchi_xlsx::Result<Package> {
        let mut edit = package.edit_xml_maps()?;
        edit.set(value)?;
        Ok(edit.commit()?.into_package())
    }

    #[test]
    fn a_signed_fixture_refuses_a_store_and_keeps_exact_noops() {
        for name in SIGNED {
            let original = signed_fixture(name);
            let package = adopt(&original);

            let mut edit = package.edit_xml_maps().unwrap();
            assert!(edit.remove().unwrap().is_none());
            let commit = edit.commit().unwrap();
            assert!(!commit.diagnostics().changed());
            assert_eq!(bytes(commit.package()), original, "{name}: exact no-ops");

            assert_signed(store(&package, info("Map")));
            assert_eq!(bytes(&package), original, "{name}: refused store");

            let output = bytes(&store(&package.unsign().unwrap(), info("Map")).unwrap());
            assert!(!carries_signature(&output), "{name}");
            assert_eq!(
                xml_maps::load_from_package(&open(&output)).unwrap(),
                Some(info("Map"))
            );
        }
    }

    #[test]
    fn a_signed_package_keeps_an_equal_store_and_refuses_changes() {
        let plain = store(&unsigned_package(SIGNED[1]), info("Map")).unwrap();
        let signed = sign_package(&plain);
        let package = adopt(&signed);

        assert_eq!(
            bytes(&store(&package, info("Map")).unwrap()),
            signed,
            "an equal store"
        );
        assert_signed(store(&package, info("Renamed")));
        let mut edit = package.edit_xml_maps().unwrap();
        assert!(edit.replace(None).unwrap());
        assert_signed(edit.commit());
        assert_eq!(bytes(&package), signed, "refused changes");

        // A changed patch is refused as signed before its source state is
        // checked; after the explicit removal it targets the plain graph.
        let patch = {
            let mut edit = plain.edit_xml_maps().unwrap();
            assert!(edit.replace(Some(info("Patched"))).unwrap());
            edit.commit().unwrap().patch().clone()
        };
        assert_signed(package.apply_xml_maps(&patch));
        assert_eq!(
            xml_maps::load_from_package(&open(&bytes(&package))).unwrap(),
            Some(info("Map"))
        );

        let output = bytes(
            package
                .unsign()
                .unwrap()
                .apply_xml_maps(&patch)
                .unwrap()
                .package(),
        );
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

    fn store(package: &Package) -> Package {
        let mut edit = package.edit_volatile_dependencies().unwrap();
        edit.set(&value(), VolatileDependenciesConformance::Transitional)
            .unwrap();
        edit.commit().unwrap().into_package()
    }

    #[test]
    fn a_signed_package_refuses_changes_and_keeps_exact_noops() {
        for name in SIGNED {
            let original = signed_fixture(name);
            let package = adopt(&original);
            let mut edit = package.edit_volatile_dependencies().unwrap();
            assert!(!edit.remove().unwrap());
            assert_signed(edit.set(&value(), VolatileDependenciesConformance::Transitional));
            assert_eq!(
                bytes(&edit.commit().unwrap().into_package()),
                original,
                "{name}"
            );

            let output = bytes(&store(&package.unsign().unwrap()));
            assert!(!carries_signature(&output), "{name}");
            assert_eq!(
                volatile_dependencies::load_from_package(&open(&output)).unwrap(),
                Some(value())
            );
        }

        let plain = store(&unsigned_package(SIGNED[0]));
        let signed = sign_package(&plain);
        let package = adopt(&signed);
        let mut edit = package.edit_volatile_dependencies().unwrap();
        assert_signed(edit.remove());
        assert_eq!(bytes(&edit.commit().unwrap().into_package()), signed);
        let mut edit = package
            .unsign()
            .unwrap()
            .edit_volatile_dependencies()
            .unwrap();
        assert!(edit.remove().unwrap());
        assert!(!carries_signature(&bytes(
            &edit.commit().unwrap().into_package()
        )));
    }
}

mod comments_family {
    use super::*;
    use litchi_xlsx::workbook::comments::{self, Comment, Comments};

    fn note(text: &str) -> Comments {
        let mut value = Comments {
            authors: vec!["Reviewer".into()],
            ..Comments::default()
        };
        value.comments.insert(
            "B2".into(),
            Comment {
                cell_ref: "B2".into(),
                author: "Reviewer".into(),
                author_id: 0,
                text: text.into(),
                guid: None,
                shape_id: None,
                runs: Vec::new(),
            },
        );
        value
    }

    fn store(package: &Package, value: &Comments) -> Package {
        let mut edit = package.edit_comments().unwrap();
        assert!(edit.replace(0, Some(value.clone())).unwrap());
        edit.commit().unwrap().into_package()
    }

    #[test]
    fn a_signed_package_refuses_changes_and_keeps_exact_noops() {
        for name in SIGNED {
            let original = signed_fixture(name);
            let package = adopt(&original);
            let value = note("Checked");

            let mut edit = package.edit_comments().unwrap();
            assert!(!edit.clear(0).unwrap());
            let commit = edit.commit().unwrap();
            assert!(!commit.diagnostics().changed());
            assert_eq!(bytes(commit.package()), original, "{name}: exact no-ops");

            let mut edit = package.edit_comments().unwrap();
            assert_signed(edit.replace(0, Some(value.clone())));
            assert_signed(edit.set(0, "C3", "Reviewer", "Staged"));
            assert_eq!(
                bytes(&edit.commit().unwrap().into_package()),
                original,
                "{name}: refusals"
            );

            let output = bytes(&store(&package.unsign().unwrap(), &value));
            assert!(!carries_signature(&output), "{name}");
            assert_eq!(
                adopt(&output).edit_comments().unwrap().comments(0).unwrap(),
                Some(value)
            );
        }
    }

    #[test]
    fn a_signed_package_refuses_a_removal_and_a_patch() {
        let plain = store(&unsigned_package(SIGNED[1]), &note("Signed note"));
        let signed = sign_package(&plain);
        let package = adopt(&signed);

        let mut edit = package.edit_comments().unwrap();
        assert_signed(edit.clear(0));
        let patch = {
            let mut edit = plain.edit_comments().unwrap();
            assert!(edit.clear(0).unwrap());
            edit.commit().unwrap().patch().clone()
        };
        assert_signed(package.apply_comments(&patch));
        assert_eq!(bytes(&package), signed);
        comments::validate_graph(&open(&signed)).unwrap();

        let output = bytes(
            package
                .unsign()
                .unwrap()
                .apply_comments(&patch)
                .unwrap()
                .package(),
        );
        assert!(!carries_signature(&output));
        assert!(
            adopt(&output)
                .edit_comments()
                .unwrap()
                .comments(0)
                .unwrap()
                .is_none()
        );
    }
}

mod revisions_family {
    use super::*;
    use litchi_xlsx::revisions::{
        EditOptions, RevisionConformance, RevisionUser, load_workbook_revisions,
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
        let mut value = load_workbook_revisions(&open(&read(TRACKED)), &SourceLimits::default())
            .unwrap()
            .expect("tracked revisions");
        value.users_relationship_id = "rIdRevisionUsers".into();
        value.headers_relationship_id = "rIdRevisionHeaders".into();
        let options = EditOptions::default().conformance(RevisionConformance::Transitional);
        for name in SIGNED {
            let original = signed_fixture(name);
            let package = adopt(&original);
            let mut edit = package.edit_revisions().unwrap();
            assert!(!edit.replace(None).unwrap());
            let commit = edit.commit().unwrap();
            assert!(!commit.diagnostics().changed());
            assert_eq!(bytes(commit.package()), original, "{name}: exact no-ops");

            let mut edit = package.edit_revisions_with_options(options).unwrap();
            assert!(edit.replace(Some(value.clone())).unwrap());
            assert_signed(edit.commit());
            assert_eq!(bytes(&package), original, "{name}: refused store");

            let mut edit = package
                .unsign()
                .unwrap()
                .edit_revisions_with_options(options)
                .unwrap();
            assert!(edit.replace(Some(value.clone())).unwrap());
            let output = bytes(&edit.commit().unwrap().into_package());
            assert!(!carries_signature(&output), "{name}");
            let stored = load_workbook_revisions(&open(&output), &SourceLimits::default())
                .unwrap()
                .unwrap();
            assert_eq!(stored.headers, value.headers, "{name}");
            assert_eq!(stored.logs.len(), value.logs.len(), "{name}");
        }
    }

    #[test]
    fn a_signed_package_refuses_edits_removal_and_patches() {
        let plain = adopt(&read(TRACKED));
        let signed = sign(&open(&read(TRACKED)));
        let package = adopt(&signed);

        let commit = package.edit_revisions().unwrap().commit().unwrap();
        assert!(!commit.diagnostics().changed());
        assert_eq!(bytes(commit.package()), signed, "an exact no-op");

        let mut edit = package.edit_revisions().unwrap();
        assert!(edit.set_user(user("Alice")).unwrap());
        assert_signed(edit.commit());
        let mut edit = package.edit_revisions().unwrap();
        assert!(edit.replace(None).unwrap());
        assert_signed(edit.commit());
        let patch = {
            let mut edit = plain.edit_revisions().unwrap();
            assert!(edit.set_user(user("Patched")).unwrap());
            edit.commit().unwrap().patch().clone()
        };
        assert_signed(package.apply_revisions(&patch));
        assert_eq!(bytes(&package), signed, "refused changes");

        let output = bytes(
            package
                .unsign()
                .unwrap()
                .apply_revisions(&patch)
                .unwrap()
                .package(),
        );
        assert!(!carries_signature(&output));
        let users = load_workbook_revisions(&open(&output), &SourceLimits::default())
            .unwrap()
            .unwrap()
            .users
            .users;
        assert!(users.iter().any(|user| user.name == "Patched"));
    }
}

mod named_sheet_view_family {
    use super::*;
    use litchi_xlsx::named_sheet_view::{Guid, View, Views};

    fn views(name: &str) -> Views {
        Views::new(
            View::with_id(
                name,
                Guid::new("{01234567-89AB-CDEF-0123-456789ABCDEF}").unwrap(),
            )
            .unwrap(),
        )
    }

    fn store(package: &Package, value: &Views) -> Package {
        let mut edit = package.edit_named_sheet_views(0_usize).unwrap();
        edit.set(value).unwrap();
        edit.commit().unwrap().into_package()
    }

    #[test]
    fn a_signed_package_refuses_changes_and_keeps_exact_noops() {
        for name in SIGNED {
            let original = signed_fixture(name);
            let package = adopt(&original);
            let mut edit = package.edit_named_sheet_views(0_usize).unwrap();
            assert!(!edit.remove().unwrap());
            assert_signed(edit.set(&views("Data")));
            assert_eq!(
                bytes(&edit.commit().unwrap().into_package()),
                original,
                "{name}"
            );

            let output = bytes(&store(&package.unsign().unwrap(), &views("Data")));
            assert!(!carries_signature(&output), "{name}");
            assert!(
                adopt(&output)
                    .edit_named_sheet_views(0_usize)
                    .unwrap()
                    .views()
                    .unwrap()
                    .is_some()
            );
        }

        let plain = store(&unsigned_package(SIGNED[1]), &views("Data"));
        let signed = sign_package(&plain);
        let package = adopt(&signed);
        let mut edit = package.edit_named_sheet_views(0_usize).unwrap();
        assert_signed(edit.set(&views("Renamed")));
        assert_signed(edit.remove());
        assert_eq!(bytes(&edit.commit().unwrap().into_package()), signed);
        let mut edit = package
            .unsign()
            .unwrap()
            .edit_named_sheet_views(0_usize)
            .unwrap();
        assert!(edit.remove().unwrap());
        assert!(!carries_signature(&bytes(
            &edit.commit().unwrap().into_package()
        )));
    }
}

mod active_x_family {
    use super::*;
    use litchi_xlsx::{Binary, Control, ControlSet, Descriptor, LoadedControl, Persistence};

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

    fn store(package: &Package, value: &ControlSet) -> Package {
        let mut edit = package.edit_active_x(0_usize).unwrap();
        edit.set(value).unwrap();
        edit.commit().unwrap().into_package()
    }

    #[test]
    fn a_signed_package_refuses_changes_and_keeps_exact_noops() {
        for name in SIGNED {
            let original = signed_fixture(name);
            let package = adopt(&original);
            let mut edit = package.edit_active_x(0_usize).unwrap();
            assert!(!edit.remove().unwrap());
            assert_signed(edit.set(&controls("Button")));
            assert_signed(edit.replace(&controls("Button")));
            assert_eq!(
                bytes(&edit.commit().unwrap().into_package()),
                original,
                "{name}"
            );

            let output = bytes(&store(&package.unsign().unwrap(), &controls("Button")));
            assert!(!carries_signature(&output), "{name}");
            assert_eq!(
                adopt(&output)
                    .edit_active_x(0_usize)
                    .unwrap()
                    .controls()
                    .unwrap(),
                controls("Button")
            );
        }

        let plain = store(&unsigned_package(SIGNED[0]), &controls("Button"));
        let signed = sign_package(&plain);
        let package = adopt(&signed);
        let mut edit = package.edit_active_x(0_usize).unwrap();
        assert_signed(edit.remove());
        assert_signed(edit.replace(&controls("Renamed")));
        assert_eq!(bytes(&edit.commit().unwrap().into_package()), signed);
        let mut edit = package.unsign().unwrap().edit_active_x(0_usize).unwrap();
        assert!(edit.remove().unwrap());
        assert!(!carries_signature(&bytes(
            &edit.commit().unwrap().into_package()
        )));
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
        let package = Package::from_bytes(&original).unwrap();
        assert!(package.is_signed());
        assert_eq!(package.to_bytes().unwrap(), original);
        let stripped = package.unsign().unwrap();
        assert!(!stripped.is_signed());
        // The source package keeps its signature.
        assert!(package.is_signed());
        assert_eq!(package.to_bytes().unwrap(), original);
        let output = stripped.to_bytes().unwrap();
        assert!(!carries_signature(&output), "{name}");

        let unsigned = Package::from_bytes(&output).unwrap().unsign().unwrap();
        assert_eq!(unsigned.to_bytes().unwrap(), output, "{name}");
    }
}

/// The shared custom-properties and VBA owners of `litchi-ooxml-common`
/// refuse a change to a signed package with this crate's typed
/// [`Error::Signed`] instead of removing the signature (unified-ops record
/// 0014a, F1b).
mod shared_owners {
    use super::*;
    use litchi_xlsx::custom::Props;

    fn props() -> Props {
        let mut props = Props::new();
        props.insert("Project", "Litchi").unwrap();
        props
    }

    /// Stage `props` as the package's custom properties through the metadata
    /// family and commit.
    fn set_custom_props(
        package: &Package,
        props: Props,
    ) -> Result<Package, litchi_xlsx::metadata::Error> {
        let mut edit = package.edit_metadata()?;
        edit.set_custom_props(props)?;
        Ok(edit.commit()?.into_package())
    }

    fn remove_custom_props(package: &Package) -> Result<Package, litchi_xlsx::metadata::Error> {
        let mut edit = package.edit_metadata()?;
        edit.remove_custom_props()?;
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
            let package = Package::from_bytes(&original).unwrap();
            let current = package.custom_props().unwrap().clone();
            let package = set_custom_props(&package, current).unwrap();
            assert_eq!(package.to_bytes().unwrap(), original, "{name}: no-ops");

            assert_metadata_signed(set_custom_props(&package, props()));
            assert!(package.is_signed());
            assert_eq!(package.to_bytes().unwrap(), original, "{name}: refusal");

            let package = set_custom_props(&package.unsign().unwrap(), props()).unwrap();
            let output = package.to_bytes().unwrap();
            assert!(!carries_signature(&output), "{name}");
            let reopened = Package::from_bytes(&output).unwrap();
            assert!(reopened.custom_props().unwrap().get("Project").is_some());
        }
    }

    #[test]
    fn a_signed_package_keeps_equal_custom_properties_and_refuses_their_removal() {
        let plain = Package::from_opc(unsigned_fixture(SIGNED[0])).unwrap();
        let plain = set_custom_props(&plain, props()).unwrap();
        let signed = sign(&plain.into_plain_opc());
        let package = Package::from_bytes(&signed).unwrap();

        let package = set_custom_props(&package, props()).unwrap();
        assert_eq!(package.to_bytes().unwrap(), signed, "an equal store");
        assert_metadata_signed(remove_custom_props(&package));
        assert_eq!(package.to_bytes().unwrap(), signed, "the refusal");

        let package = remove_custom_props(&package.unsign().unwrap()).unwrap();
        assert!(package.custom_props().unwrap().is_empty());
    }

    #[cfg(feature = "vba-inspection")]
    #[test]
    fn vba_projects_refuse_a_signed_package() {
        use litchi_vba::{Limits, build};

        let payload = || {
            build::Project::new("ExcelProject")
                .module(build::Module::standard(
                    "Module1",
                    "Public Sub Hello()\r\nEnd Sub\r\n",
                ))
                .finish(&Limits::default())
                .unwrap()
        };
        let store = |package: &Package| {
            let mut edit = package.edit_vba().unwrap();
            edit.store(payload()).unwrap();
            edit.commit().unwrap().into_package()
        };
        for name in SIGNED {
            let original = signed_fixture(name);
            let package = adopt(&original);
            let mut edit = package.edit_vba().unwrap();
            assert!(!edit.remove().unwrap());
            assert_signed(edit.store(payload()));
            assert_eq!(
                bytes(&edit.commit().unwrap().into_package()),
                original,
                "{name}: no-op and refusal"
            );

            let output = bytes(&store(&package.unsign().unwrap()));
            assert!(!carries_signature(&output), "{name}");
        }

        let plain = store(&unsigned_package(SIGNED[0]));
        let signed = sign_package(&plain);
        let package = adopt(&signed);
        let mut edit = package.edit_vba().unwrap();
        assert_signed(edit.remove());
        assert_eq!(
            bytes(&edit.commit().unwrap().into_package()),
            signed,
            "the refusal"
        );
        let mut edit = package.unsign().unwrap().edit_vba().unwrap();
        assert!(edit.remove().unwrap());
    }
}
