#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "unit tests panic on fixture failure"
)]

use litchi_core::metadata::{Field, Metadata, Timestamp};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::part::BlobPart;
use litchi_opc::{OpcPackage, PackURI};

use super::{Core, Error, Family, Keywords, Snapshot, app, write_core};
use crate::custom::{self, Host, Value};
use crate::leniency::ToleranceReport;

const APP: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>Producer</Application><TitlesOfParts><vt:vector size="1" baseType="lpstr"><vt:lpstr>Sheet1</vt:lpstr></vt:vector></TitlesOfParts><Company>Old</Company></Properties>"#;

/// A package holding all three families, as a producer wrote them.
fn full() -> OpcPackage {
    let mut package = OpcPackage::new();
    let core = Core {
        keywords: Some(Keywords::new().text("alpha ").value("beta")),
        ..Core::new().title("Before").creator("Ada")
    };
    write_core(&mut package, core).unwrap();
    package.add_part(Box::new(BlobPart::new(
        PackURI::new("/docProps/app.xml").unwrap(),
        ct::OFC_EXTENDED_PROPERTIES.to_owned(),
        APP.as_bytes().to_vec(),
    )));
    package.relate_to("docProps/app.xml", rt::EXTENDED_PROPERTIES);
    let mut props = custom::Props::new();
    props.insert("Project", "Litchi").unwrap();
    props.write_for(&mut package, Host::Word).unwrap();
    package
}

fn sign(package: &mut OpcPackage) {
    package.add_part(Box::new(BlobPart::new(
        PackURI::new("/_xmlsignatures/origin.sigs").unwrap(),
        ct::OPC_DIGITAL_SIGNATURE_ORIGIN.to_owned(),
        Vec::new(),
    )));
    package.relate_to("_xmlsignatures/origin.sigs", rt::DIGITAL_SIGNATURE_ORIGIN);
    assert!(package.is_signed());
}

fn bytes(package: &OpcPackage, name: &str) -> Option<Vec<u8>> {
    package
        .get_part(&PackURI::new(name).unwrap())
        .ok()
        .map(|part| part.blob().to_vec())
}

#[test]
fn snapshot_projects_core_and_app_fields_into_one_metadata() {
    let snapshot = Snapshot::read(&full(), Host::Word).unwrap();
    let metadata = snapshot.metadata().unwrap();
    assert_eq!(metadata.title.as_deref(), Some("Before"));
    assert_eq!(metadata.author.as_deref(), Some("Ada"));
    assert_eq!(metadata.keywords.as_deref(), Some("alpha beta"));
    assert_eq!(metadata.company.as_deref(), Some("Old"));
    assert_eq!(metadata.application.as_deref(), Some("Producer"));
    assert_eq!(
        snapshot.custom_props().get("Project"),
        Some(&Value::Text("Litchi".into()))
    );
}

#[test]
fn metadata_edit_routes_fields_to_their_families_and_inverts_exactly() {
    let mut package = full();
    let core_before = bytes(&package, "/docProps/core.xml").unwrap();
    let mut edit = Snapshot::read(&package, Host::Word).unwrap().edit();
    edit.update_metadata(|metadata| {
        metadata.title = Some("After".into());
        metadata.manager = Some("Lin".into());
        metadata.created = Some(Timestamp::parse("2024-05-06T07:08:09Z").unwrap());
    })
    .unwrap();
    let commit = edit.commit().unwrap();
    assert_eq!(commit.diagnostics().families(), [Family::Core, Family::App]);
    assert_eq!(
        commit.diagnostics().fields(),
        [Field::Title, Field::Created, Field::Manager]
    );
    let patch = commit.patch().clone();
    assert!(patch.apply(&mut package).unwrap());

    let after = Snapshot::read(&package, Host::Word).unwrap();
    let metadata = after.metadata().unwrap();
    assert_eq!(metadata.title.as_deref(), Some("After"));
    assert_eq!(metadata.manager.as_deref(), Some("Lin"));
    assert_eq!(metadata.company.as_deref(), Some("Old"));
    // The unchanged keyword mixed content kept its structure.
    let core_xml = String::from_utf8(bytes(&package, "/docProps/core.xml").unwrap()).unwrap();
    assert!(core_xml.contains("<cp:value>beta</cp:value>"), "{core_xml}");
    assert!(core_xml.contains("2024-05-06T07:08:09Z"), "{core_xml}");
    // The application part kept its vector and every unchanged byte.
    assert_eq!(
        String::from_utf8(bytes(&package, "/docProps/app.xml").unwrap()).unwrap(),
        APP.replace("</Properties>", "<Manager>Lin</Manager></Properties>")
    );

    assert!(patch.inverse().apply(&mut package).unwrap());
    assert_eq!(bytes(&package, "/docProps/core.xml").unwrap(), core_before);
    assert_eq!(
        bytes(&package, "/docProps/app.xml").unwrap(),
        APP.as_bytes()
    );
}

#[test]
fn equal_values_commit_nothing_even_on_a_signed_package() {
    let mut package = full();
    sign(&mut package);
    let snapshot = Snapshot::read(&package, Host::Word).unwrap();
    let mut edit = snapshot.edit();
    edit.set_metadata(snapshot.metadata().unwrap());
    let commit = edit.commit().unwrap();
    assert!(!commit.diagnostics().changed());
    assert!(commit.patch().is_noop());
    assert!(!commit.patch().apply(&mut package).unwrap());

    let mut edit = snapshot.edit();
    edit.update_metadata(|metadata| metadata.title = Some("Changed".into()))
        .unwrap();
    let patch = edit.commit().unwrap().patch().clone();
    assert!(matches!(patch.apply(&mut package), Err(Error::Signed)));
}

#[test]
fn fields_no_family_stores_are_refused_not_dropped() {
    let snapshot = Snapshot::read(&full(), Host::Word).unwrap();
    let updates: [(Field, fn(&mut Metadata)); 3] = [
        (Field::ContentType, |metadata| {
            metadata.content_type = Some("x".into());
        }),
        (Field::Codepage, |metadata| metadata.codepage = Some(65001)),
        (Field::LastBackup, |metadata| {
            metadata.last_backup = Some(Timestamp::parse("2024-01-01T00:00:00").unwrap());
        }),
    ];
    for (field, update) in updates {
        let mut edit = snapshot.edit();
        edit.update_metadata(update).unwrap();
        assert!(
            matches!(edit.commit(), Err(Error::Unsupported { field: refused }) if refused == field)
        );
    }
    let mut edit = snapshot.edit();
    edit.update_metadata(|metadata| metadata.page_count = Some(u32::MAX))
        .unwrap();
    assert!(matches!(edit.commit(), Err(Error::Invalid(_))));
}

#[test]
fn an_empty_package_gains_canonical_parts() {
    let mut package = OpcPackage::new();
    let mut edit = Snapshot::read(&package, Host::Excel).unwrap().edit();
    edit.set_metadata(Metadata {
        author: Some("Ada".into()),
        company: Some("Litchi".into()),
        ..Metadata::default()
    });
    let mut custom = custom::Props::new();
    custom.insert("Reviewed", Value::Bool(true)).unwrap();
    edit.set_custom_props(custom);
    let patch = edit.commit().unwrap().patch().clone();
    assert!(
        patch
            .changes()
            .iter()
            .all(|change| change.before().is_none() && change.after().is_some())
    );
    assert!(patch.apply(&mut package).unwrap());
    let metadata = super::read_with_report(&package, &mut ToleranceReport::default())
        .unwrap()
        .unwrap();
    assert_eq!(metadata.author.as_deref(), Some("Ada"));
    assert_eq!(
        app::read(&package).unwrap().unwrap().company.as_deref(),
        Some("Litchi")
    );
    assert!(patch.inverse().apply(&mut package).unwrap());
    assert_eq!(package.iter_parts().count(), 0);
    assert_eq!(package.rels().iter().count(), 0);
}

#[test]
fn a_stale_patch_conflicts_without_mutation() {
    let mut package = full();
    let mut edit = Snapshot::read(&package, Host::Word).unwrap().edit();
    edit.update_metadata(|metadata| metadata.title = Some("One".into()))
        .unwrap();
    let first = edit.commit().unwrap().patch().clone();
    let mut edit = Snapshot::read(&package, Host::Word).unwrap().edit();
    edit.update_metadata(|metadata| metadata.title = Some("Two".into()))
        .unwrap();
    let second = edit.commit().unwrap().patch().clone();
    assert!(first.apply(&mut package).unwrap());
    let before = bytes(&package, "/docProps/core.xml");
    assert!(matches!(
        second.apply(&mut package),
        Err(Error::Conflict(_))
    ));
    assert_eq!(bytes(&package, "/docProps/core.xml"), before);
}
