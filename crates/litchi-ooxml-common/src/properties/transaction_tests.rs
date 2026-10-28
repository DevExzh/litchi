#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "unit tests panic on fixture failure"
)]

use std::sync::Arc;
use std::sync::atomic::{AtomicBool, Ordering};

use litchi_core::metadata::{Field, Metadata, Timestamp};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::part::{BlobPart, Part};
use litchi_opc::{OpcError, OpcPackage, PackURI, PackageWriter, Relationships, TargetMode};

use super::{
    Core, Edit, EditOptions, Error, Family, Keywords, Snapshot, app, publish_core_metadata,
    write_core,
};
use crate::custom::{self, Host, Value};
use crate::leniency::ToleranceReport;

const APP: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>Producer</Application><TitlesOfParts><vt:vector size="1" baseType="lpstr"><vt:lpstr>Sheet1</vt:lpstr></vt:vector></TitlesOfParts><Company>Old</Company></Properties>"#;

/// A deliberately unsealed part whose observable state can change without a
/// package mutation seam. The transaction fast path must therefore reject a
/// snapshot from a package containing it and use exact patch authentication.
#[derive(Clone)]
struct InteriorMutablePart {
    name: PackURI,
    old_blob: Vec<u8>,
    new_blob: Vec<u8>,
    old_content_type: String,
    new_content_type: String,
    old_rels: Relationships,
    new_rels: Relationships,
    changed: Arc<AtomicBool>,
    flip_after_blob_arc: Arc<AtomicBool>,
}

impl Part for InteriorMutablePart {
    fn blob(&self) -> &[u8] {
        if self.changed.load(Ordering::Relaxed) {
            &self.new_blob
        } else {
            &self.old_blob
        }
    }

    fn blob_arc(&self) -> Arc<Vec<u8>> {
        let blob = Arc::new(self.blob().to_vec());
        if self.flip_after_blob_arc.swap(false, Ordering::Relaxed) {
            self.changed.store(true, Ordering::Relaxed);
        }
        blob
    }

    fn content_type(&self) -> &str {
        if self.changed.load(Ordering::Relaxed) {
            &self.new_content_type
        } else {
            &self.old_content_type
        }
    }

    fn partname(&self) -> &PackURI {
        &self.name
    }

    fn rels(&self) -> &Relationships {
        if self.changed.load(Ordering::Relaxed) {
            &self.new_rels
        } else {
            &self.old_rels
        }
    }

    fn rels_mut(&mut self) -> &mut Relationships {
        if self.changed.load(Ordering::Relaxed) {
            &mut self.new_rels
        } else {
            &mut self.old_rels
        }
    }

    fn set_blob(&mut self, blob: Vec<u8>) {
        self.old_blob = blob.clone();
        self.new_blob = blob;
    }
}

fn package_with_interior_mutable_core() -> (OpcPackage, Arc<AtomicBool>) {
    package_with_interior_mutable_core_and(false)
}

fn package_with_interior_mutable_core_and(
    flip_after_blob_arc: bool,
) -> (OpcPackage, Arc<AtomicBool>) {
    let mut package = full();
    let core_name = PackURI::new("/docProps/core.xml").unwrap();
    let source = package.get_part(&core_name).unwrap();
    let old_blob = source.blob().to_vec();
    let new_blob = String::from_utf8(old_blob.clone())
        .unwrap()
        .replace("Before", "Changed")
        .into_bytes();
    let old_content_type = source.content_type().to_owned();
    let old_rels = source.rels().clone();
    package.remove_part(&core_name);
    let changed = Arc::new(AtomicBool::new(false));
    package.add_part(Box::new(InteriorMutablePart {
        name: core_name,
        old_blob,
        new_blob,
        old_content_type: old_content_type.clone(),
        new_content_type: old_content_type,
        old_rels: old_rels.clone(),
        new_rels: old_rels,
        changed: Arc::clone(&changed),
        flip_after_blob_arc: Arc::new(AtomicBool::new(flip_after_blob_arc)),
    }));
    (package, changed)
}

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
    let mut edit = Edit::new(&package, Host::Word).unwrap();
    edit.set_custom_props(props);
    edit.commit().unwrap().into_package()
}

fn app_only() -> OpcPackage {
    let mut package = OpcPackage::new();
    let core = Core::new().title("Before").creator("Ada");
    write_core(&mut package, core).unwrap();
    package.add_part(Box::new(BlobPart::new(
        PackURI::new("/docProps/app.xml").unwrap(),
        ct::OFC_EXTENDED_PROPERTIES.to_owned(),
        APP.as_bytes().to_vec(),
    )));
    package.relate_to("docProps/app.xml", rt::EXTENDED_PROPERTIES);
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
    assert_eq!(metadata.keywords, ["alpha", "beta"]);
    assert_eq!(metadata.company.as_deref(), Some("Old"));
    assert_eq!(metadata.application.as_deref(), Some("Producer"));
    assert_eq!(
        snapshot.custom_props().get("Project"),
        Some(&Value::Text("Litchi".into()))
    );
}

#[test]
fn core_only_staging_preserves_application_properties() {
    let package = full();
    let app_before = bytes(&package, "/docProps/app.xml").unwrap();
    let mut edit = Edit::new(&package, Host::Word).unwrap();
    edit.set_core_metadata(Metadata {
        title: Some("After".to_owned()),
        ..Metadata::default()
    })
    .unwrap();
    let after = edit.commit().unwrap().into_package();

    assert_eq!(bytes(&after, "/docProps/app.xml").unwrap(), app_before);
    let snapshot = Snapshot::read(&after, Host::Word).unwrap();
    let metadata = snapshot.metadata().unwrap();
    assert_eq!(metadata.title.as_deref(), Some("After"));
    assert_eq!(metadata.application.as_deref(), Some("Producer"));
    assert_eq!(metadata.company.as_deref(), Some("Old"));
}

#[test]
fn core_only_publication_replays_exactly_and_preserves_unread_families() {
    let package = full();
    let core_before = bytes(&package, "/docProps/core.xml").unwrap();
    let app_before = bytes(&package, "/docProps/app.xml").unwrap();
    let custom_before = bytes(&package, "/docProps/custom.xml").unwrap();
    let target = Metadata {
        title: Some("After".to_owned()),
        author: Some("Ada".to_owned()),
        keywords: ["alpha", "beta"].map(str::to_owned).to_vec(),
        ..Metadata::default()
    };

    let (commit, published) =
        publish_core_metadata(&package, Host::Word, Some(target.clone())).unwrap();
    assert_eq!(
        published.as_ref().and_then(|value| value.title.as_deref()),
        Some("After")
    );
    assert_eq!(
        bytes(commit.package(), "/docProps/app.xml").unwrap(),
        app_before
    );
    assert_eq!(
        bytes(commit.package(), "/docProps/custom.xml").unwrap(),
        custom_before
    );
    assert_ne!(
        bytes(commit.package(), "/docProps/core.xml").unwrap(),
        core_before
    );

    let readback = Snapshot::read(commit.package(), Host::Word)
        .unwrap()
        .metadata()
        .unwrap();
    assert_eq!(readback.title, target.title);
    assert_eq!(readback.author, target.author);
    assert_eq!(readback.keywords, target.keywords);
    assert_eq!(readback.application.as_deref(), Some("Producer"));
    assert_eq!(readback.company.as_deref(), Some("Old"));

    let replayed = commit.patch().apply(&package).unwrap().into_package();
    assert_eq!(
        bytes(&replayed, "/docProps/core.xml"),
        bytes(commit.package(), "/docProps/core.xml")
    );
    assert_eq!(bytes(&replayed, "/docProps/app.xml").unwrap(), app_before);
    assert_eq!(
        bytes(&replayed, "/docProps/custom.xml").unwrap(),
        custom_before
    );

    let restored = commit
        .patch()
        .inverse()
        .apply(commit.package())
        .unwrap()
        .into_package();
    assert_eq!(bytes(&restored, "/docProps/core.xml").unwrap(), core_before);
    assert_eq!(bytes(&restored, "/docProps/app.xml").unwrap(), app_before);
    assert_eq!(
        bytes(&restored, "/docProps/custom.xml").unwrap(),
        custom_before
    );
}

#[test]
fn snapshot_constructor_reauthenticates_a_mismatched_package() {
    let source = full();
    let snapshot = Snapshot::read(&source, Host::Word).unwrap();

    let mut other_edit = Edit::new(&full(), Host::Word).unwrap();
    other_edit.set_metadata(Metadata {
        title: Some("Other source".to_owned()),
        ..Metadata::default()
    });
    let other = other_edit.commit().unwrap().into_package();

    let mut edit = Edit::new_with_snapshot(&other, snapshot, EditOptions::default());
    edit.set_metadata(Metadata {
        title: Some("Candidate".to_owned()),
        ..Metadata::default()
    });
    assert!(matches!(edit.commit(), Err(Error::PatchConflict)));
}

#[test]
fn snapshot_constructor_reauthenticates_a_same_handle_change() {
    let mut package = full();
    let snapshot = Snapshot::read(&package, Host::Word).unwrap();
    let mut changed_edit = Edit::new(&package, Host::Word).unwrap();
    changed_edit.set_metadata(Metadata {
        title: Some("Changed after capture".to_owned()),
        ..Metadata::default()
    });
    let changed = changed_edit.commit().unwrap().into_package();
    let core = PackURI::new("/docProps/core.xml").unwrap();
    package
        .get_part_mut(&core)
        .unwrap()
        .set_blob(bytes(&changed, "/docProps/core.xml").unwrap());

    let mut edit = Edit::new_with_snapshot(&package, snapshot, EditOptions::default());
    edit.set_metadata(Metadata {
        title: Some("Candidate".to_owned()),
        ..Metadata::default()
    });
    assert!(matches!(edit.commit(), Err(Error::PatchConflict)));
}

#[test]
fn snapshot_constructor_reauthenticates_a_content_type_change() {
    let mut package = full();
    let snapshot = Snapshot::read(&package, Host::Word).unwrap();
    let core = PackURI::new("/docProps/core.xml").unwrap();
    package
        .get_part_mut(&core)
        .unwrap()
        .set_content_type(ct::OFC_EXTENDED_PROPERTIES.to_owned())
        .unwrap();

    let mut edit = Edit::new_with_snapshot(&package, snapshot, EditOptions::default());
    edit.set_metadata(Metadata {
        title: Some("Candidate".to_owned()),
        ..Metadata::default()
    });
    assert!(edit.commit().is_err());
}

#[test]
fn custom_part_interior_mutation_never_uses_the_own_snapshot_fast_path() {
    let (package, changed) = package_with_interior_mutable_core();
    let mut edit = Edit::new(&package, Host::Word).unwrap();
    changed.store(true, Ordering::Relaxed);
    edit.set_metadata(Metadata {
        title: Some("Candidate".to_owned()),
        ..Metadata::default()
    });
    assert!(matches!(edit.commit(), Err(Error::PatchConflict)));

    let (package, changed) = package_with_interior_mutable_core();
    let snapshot = Snapshot::read(&package, Host::Word).unwrap();
    let mut edit = Edit::new_with_snapshot(&package, snapshot, EditOptions::default());
    changed.store(true, Ordering::Relaxed);
    edit.set_metadata(Metadata {
        title: Some("Candidate".to_owned()),
        ..Metadata::default()
    });
    assert!(matches!(edit.commit(), Err(Error::PatchConflict)));
}

#[test]
fn core_only_publication_reauthenticates_custom_parts_after_snapshot() {
    let (package, _) = package_with_interior_mutable_core_and(true);
    let result = publish_core_metadata(
        &package,
        Host::Word,
        Some(Metadata {
            title: Some("Candidate".to_owned()),
            ..Metadata::default()
        }),
    );
    assert!(matches!(result, Err(Error::PatchConflict)));
}

#[test]
fn repaired_core_content_is_refused_only_when_core_is_rewritten() {
    let mut package = full();
    let mut dirty = String::from_utf8(bytes(&package, "/docProps/core.xml").unwrap()).unwrap();
    let closing = dirty.find("</cp:coreProperties>").unwrap();
    dirty.insert_str(closing, "<cp:opaque>preserve me</cp:opaque>");
    let dirty = dirty.into_bytes();
    package
        .get_part_mut(&PackURI::new("/docProps/core.xml").unwrap())
        .unwrap()
        .set_blob(dirty.clone());

    let snapshot = Snapshot::read(&package, Host::Word).unwrap();
    assert!(!snapshot.core_tolerance_report().is_clean());

    let mut core_edit = Edit::new(&package, Host::Word).unwrap();
    core_edit
        .update_metadata(|metadata| metadata.title = Some("rewritten".to_owned()))
        .unwrap();
    assert!(matches!(
        core_edit.commit(),
        Err(Error::Refused(refusal))
            if refusal.kind() == litchi_core::edit::RefusalKind::UnsupportedContent
    ));

    let mut app_edit = Edit::new(&package, Host::Word).unwrap();
    app_edit
        .update_metadata(|metadata| metadata.company = Some("New".to_owned()))
        .unwrap();
    let after = app_edit.commit().unwrap().into_package();
    assert_eq!(bytes(&after, "/docProps/core.xml").unwrap(), dirty);
    assert_eq!(
        Snapshot::read(&after, Host::Word)
            .unwrap()
            .metadata()
            .unwrap()
            .company
            .as_deref(),
        Some("New")
    );
}

#[test]
fn core_only_staging_refuses_application_fields_as_properties() {
    let package = full();
    let mut edit = Edit::new(&package, Host::Word).unwrap();
    assert!(matches!(
        edit.set_core_metadata(Metadata {
            company: Some("not core".to_owned()),
            ..Metadata::default()
        }),
        Err(Error::Refused(refusal))
            if refusal.kind() == litchi_core::edit::RefusalKind::UnsupportedProperty
    ));
}

#[test]
fn setting_empty_custom_properties_preserves_an_existing_empty_part() {
    const EMPTY_CUSTOM: &[u8] = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"></Properties>"#;
    let mut package = full();
    package
        .get_part_mut(&PackURI::new("/docProps/custom.xml").unwrap())
        .unwrap()
        .set_blob(EMPTY_CUSTOM.to_vec());
    let raw_before = bytes(&package, "/docProps/custom.xml").unwrap();

    let mut edit = Edit::new(&package, Host::Word).unwrap();
    edit.set_custom_props(custom::Props::new());
    let commit = edit.commit().unwrap();
    assert!(!commit.diagnostics().changed());
    assert!(commit.patch().is_noop());
    let preserved = commit.into_package();
    assert_eq!(
        bytes(&preserved, "/docProps/custom.xml").unwrap(),
        raw_before
    );

    let mut signed_noop = package.clone();
    sign(&mut signed_noop);
    let mut edit = Edit::new(&signed_noop, Host::Word).unwrap();
    edit.set_custom_props(custom::Props::new());
    let kept = edit.commit().unwrap();
    assert!(!kept.diagnostics().changed());
    assert!(kept.package().is_signed());

    let mut edit = Edit::new(&package, Host::Word).unwrap();
    edit.remove_custom_props();
    let commit = edit.commit().unwrap();
    assert!(commit.diagnostics().changed());
    let cleared = commit.into_package();
    assert!(
        cleared
            .get_part(&PackURI::new("/docProps/custom.xml").unwrap())
            .is_err()
    );

    let mut signed = package;
    sign(&mut signed);
    let mut edit = Edit::new(&signed, Host::Word).unwrap();
    edit.remove_custom_props();
    assert!(matches!(edit.commit(), Err(Error::Signed)));
}

#[test]
fn explicit_custom_removal_keeps_dirty_raw_graphs_removable() {
    const DIRTY_CUSTOM: &[u8] = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="Keep"><vt:lpwstr>value</vt:lpwstr></property><property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="3" name="Skip"><vt:date>2026-01-01</vt:date></property></Properties>"#;
    let mut package = full();
    package
        .get_part_mut(&PackURI::new("/docProps/custom.xml").unwrap())
        .unwrap()
        .set_blob(DIRTY_CUSTOM.to_vec());

    let snapshot = Snapshot::read(&package, Host::Word).unwrap();
    assert!(!snapshot.custom_props().tolerance_report().is_clean());
    let mut edit = Edit::new(&package, Host::Word).unwrap();
    edit.remove_custom_props();
    let commit = edit.commit().unwrap();
    assert!(commit.diagnostics().changed());
    assert!(
        commit
            .into_package()
            .get_part(&PackURI::new("/docProps/custom.xml").unwrap())
            .is_err()
    );
}

#[test]
fn rewriting_tolerated_custom_content_is_a_typed_refusal() {
    const DIRTY_CUSTOM: &[u8] = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="Keep"><vt:lpwstr>value</vt:lpwstr></property><property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="3" name="Skip"><vt:date>2026-01-01</vt:date></property></Properties>"#;
    let mut package = full();
    package
        .get_part_mut(&PackURI::new("/docProps/custom.xml").unwrap())
        .unwrap()
        .set_blob(DIRTY_CUSTOM.to_vec());
    let snapshot = Snapshot::read(&package, Host::Word).unwrap();
    let mut plan = snapshot.plan();
    let mut target = plan.custom_props().clone();
    target.insert("Keep", "changed").unwrap();
    plan.set_custom_props(target);
    assert!(matches!(
        plan.finish(),
        Err(Error::Refused(refusal))
            if refusal.kind() == litchi_core::edit::RefusalKind::UnsupportedContent
    ));
}

#[test]
fn shared_metadata_removal_is_a_typed_refusal() {
    let mut package = full();
    package
        .rels_mut()
        .try_add_relationship(
            "urn:litchi:test".to_owned(),
            "docProps/custom.xml".to_owned(),
            "rIdExtra".to_owned(),
            TargetMode::Internal,
        )
        .unwrap();
    let snapshot = Snapshot::read(&package, Host::Word).unwrap();
    let mut plan = snapshot.plan();
    plan.remove_custom_props();
    assert!(matches!(
        plan.finish(),
        Err(Error::Refused(refusal))
            if refusal.kind() == litchi_core::edit::RefusalKind::UnsupportedContent
    ));
}

#[test]
fn clearing_core_keeps_the_app_projection_staged() {
    let package = full();
    let mut edit = Edit::new(&package, Host::Word).unwrap();
    edit.clear_metadata();
    let staged = edit.metadata().unwrap();
    assert_eq!(staged.title, None);
    assert_eq!(staged.company.as_deref(), Some("Old"));
    edit.update_metadata(|metadata| metadata.company = Some("New".to_owned()))
        .unwrap();
    let cleared = edit.commit().unwrap().into_package();

    assert!(
        Snapshot::read(&cleared, Host::Word)
            .unwrap()
            .metadata()
            .unwrap()
            .title
            .is_none()
    );
    assert_eq!(
        Snapshot::read(&cleared, Host::Word)
            .unwrap()
            .metadata()
            .unwrap()
            .company
            .as_deref(),
        Some("New")
    );
}

#[test]
fn metadata_commit_admits_exact_escaped_patch_size_before_capture() {
    let package = full();
    let mut edit = Edit::new(&package, Host::Word).unwrap();
    edit.update_metadata(|metadata| {
        metadata.title = Some("A & <property> \r\n B".to_owned());
    })
    .unwrap();
    let expected = edit.clone().commit().unwrap().patch().retained_bytes();

    let committed = Edit::new_with_options(
        &package,
        Host::Word,
        EditOptions::new().max_retained_bytes(expected),
    )
    .and_then(|mut edit| {
        edit.update_metadata(|metadata| {
            metadata.title = Some("A & <property> \r\n B".to_owned());
        })?;
        edit.commit()
    })
    .unwrap();
    assert_eq!(committed.patch().retained_bytes(), expected);

    let mut under = Edit::new_with_options(
        &package,
        Host::Word,
        EditOptions::new().max_retained_bytes(expected - 1),
    )
    .unwrap();
    under
        .update_metadata(|metadata| {
            metadata.title = Some("A & <property> \r\n B".to_owned());
        })
        .unwrap();
    assert!(matches!(
        under.commit(),
        Err(Error::Limit {
            kind: super::LimitKind::RetainedBytes,
            ..
        })
    ));
}

#[test]
fn metadata_commit_admits_exact_app_and_custom_escape_sizes() {
    let package = full();
    let mut edit = Edit::new(&package, Host::Word).unwrap();
    edit.update_metadata(|metadata| {
        metadata.company = Some("A & <application>".to_owned());
    })
    .unwrap();
    let mut custom = edit.custom_props().clone();
    custom.insert("Project", "A & <custom>").unwrap();
    edit.set_custom_props(custom);
    let expected = edit.clone().commit().unwrap().patch().retained_bytes();

    let committed = Edit::new_with_options(
        &package,
        Host::Word,
        EditOptions::new().max_retained_bytes(expected),
    )
    .and_then(|mut edit| {
        edit.update_metadata(|metadata| {
            metadata.company = Some("A & <application>".to_owned());
        })?;
        let mut custom = edit.custom_props().clone();
        custom.insert("Project", "A & <custom>")?;
        edit.set_custom_props(custom);
        edit.commit()
    })
    .unwrap();
    assert_eq!(committed.patch().retained_bytes(), expected);

    let mut under = Edit::new_with_options(
        &package,
        Host::Word,
        EditOptions::new().max_retained_bytes(expected - 1),
    )
    .unwrap();
    under
        .update_metadata(|metadata| {
            metadata.company = Some("A & <application>".to_owned());
        })
        .unwrap();
    let mut custom = under.custom_props().clone();
    custom.insert("Project", "A & <custom>").unwrap();
    under.set_custom_props(custom);
    assert!(matches!(
        under.commit(),
        Err(Error::Limit {
            kind: super::LimitKind::RetainedBytes,
            ..
        })
    ));
}

#[test]
fn keyword_lists_round_trip_the_source_bytes_until_they_change() {
    let mut package = OpcPackage::new();
    let core = Core {
        keywords: Some(Keywords::new().text("a; b,c ").value("d")),
        ..Core::new().title("Keywords")
    };
    write_core(&mut package, core).unwrap();
    let before = bytes(&package, "/docProps/core.xml").unwrap();
    let snapshot = Snapshot::read(&package, Host::Word).unwrap();
    assert_eq!(snapshot.metadata().unwrap().keywords, ["a", "b", "c", "d"]);

    // Setting the list it already holds writes nothing.
    let mut edit = Edit::new(&package, Host::Word).unwrap();
    edit.update_metadata(|metadata| {
        metadata.keywords = ["a", "b", "c", "d"].map(str::to_owned).to_vec();
    })
    .unwrap();
    let commit = edit.commit().unwrap();
    assert!(!commit.diagnostics().changed());
    assert!(commit.patch().is_noop());
    assert_eq!(
        bytes(commit.package(), "/docProps/core.xml").unwrap(),
        before
    );
    assert_eq!(bytes(&package, "/docProps/core.xml").unwrap(), before);

    // A changed list is joined with `, `; a keyword holding a separator
    // becomes a `cp:value`, so it reads back whole.
    let mut edit = Edit::new(&package, Host::Word).unwrap();
    edit.update_metadata(|metadata| {
        metadata.keywords = ["a", "b", "x, y"].map(str::to_owned).to_vec();
    })
    .unwrap();
    let commit = edit.commit().unwrap();
    assert_eq!(commit.diagnostics().fields(), [Field::Keywords]);
    assert!(commit.patch().apply_in_place(&mut package).unwrap());
    assert_eq!(
        bytes(commit.package(), "/docProps/core.xml"),
        bytes(&package, "/docProps/core.xml")
    );
    let core_xml = String::from_utf8(bytes(&package, "/docProps/core.xml").unwrap()).unwrap();
    assert!(
        core_xml.contains("<cp:keywords>a, b, <cp:value>x, y</cp:value></cp:keywords>"),
        "{core_xml}"
    );
    let read = Snapshot::read(&package, Host::Word).unwrap();
    assert_eq!(read.metadata().unwrap().keywords, ["a", "b", "x, y"]);
}

#[test]
fn metadata_edit_routes_fields_to_their_families_and_inverts_exactly() {
    let mut package = full();
    let core_before = bytes(&package, "/docProps/core.xml").unwrap();
    let mut edit = Edit::new(&package, Host::Word).unwrap();
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
    assert!(patch.apply_in_place(&mut package).unwrap());

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

    assert!(patch.inverse().apply_in_place(&mut package).unwrap());
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
    let mut edit = Edit::new(&package, Host::Word).unwrap();
    edit.set_metadata(snapshot.metadata().unwrap());
    let commit = edit.commit().unwrap();
    assert!(!commit.diagnostics().changed());
    assert!(commit.patch().is_noop());
    assert!(commit.package().is_signed());
    assert!(!commit.patch().apply_in_place(&mut package).unwrap());

    // A changed edit of a signed package is refused at commit; its planned
    // patch is refused by the replay.
    let mut edit = Edit::new(&package, Host::Word).unwrap();
    edit.update_metadata(|metadata| metadata.title = Some("Changed".into()))
        .unwrap();
    assert!(matches!(edit.commit(), Err(Error::Signed)));
    let mut plan = snapshot.plan();
    plan.update_metadata(|metadata| metadata.title = Some("Changed".into()))
        .unwrap();
    let (_, patch, _) = plan.finish().unwrap();
    assert!(matches!(
        patch.apply_in_place(&mut package),
        Err(Error::Signed)
    ));
}

#[test]
fn an_exact_no_op_patch_rejects_a_different_property_owner() {
    let source = full();
    let snapshot = Snapshot::read(&source, Host::Word).unwrap();
    let (_, no_op, diagnostics) = snapshot.plan().finish().unwrap();
    assert!(!diagnostics.changed());
    assert!(no_op.is_noop());

    let mut changed_plan = snapshot.plan();
    changed_plan
        .update_metadata(|metadata| metadata.title = Some("Different".into()))
        .unwrap();
    let (_, changed, diagnostics) = changed_plan.finish().unwrap();
    assert!(diagnostics.changed());
    let mut different = source.clone();
    changed.apply_in_place(&mut different).unwrap();
    assert!(matches!(
        no_op.apply_in_place(&mut different),
        Err(Error::PatchConflict)
    ));

    let mut exact = source.clone();
    assert!(!no_op.apply_in_place(&mut exact).unwrap());
}

#[test]
fn composed_no_op_scopes_are_checked_instead_of_dropped() {
    let source = full();
    let snapshot = Snapshot::read(&source, Host::Word).unwrap();
    let (_, no_op, _) = snapshot.plan().finish().unwrap();
    let (changed_state, changed) = edited(&snapshot, "Changed");
    let (_, changed_no_op, _) = changed_state.plan().finish().unwrap();

    // A generated no-op authenticates all three property families. The
    // composed changed patch retains that complete before scope instead of
    // reducing authentication to the changed family.
    let no_op_then_changed = no_op.compose(&changed).unwrap();
    let mut direct = source.clone();
    no_op_then_changed.apply_in_place(&mut direct).unwrap();
    let mut sequential = source.clone();
    changed.apply_in_place(&mut sequential).unwrap();
    assert_eq!(
        bytes(&direct, "/docProps/core.xml"),
        bytes(&sequential, "/docProps/core.xml")
    );

    let changed_then_no_op = changed.compose(&changed_no_op).unwrap();
    let mut direct = source.clone();
    changed_then_no_op.apply_in_place(&mut direct).unwrap();
    assert_eq!(
        bytes(&direct, "/docProps/core.xml"),
        bytes(&sequential, "/docProps/core.xml")
    );

    // The untouched application family remains authenticated by the
    // composed scope.
    let mut app_plan = snapshot.plan();
    app_plan.set_app_props(Some(app::Props {
        company: Some("Other".into()),
        ..app::Props::default()
    }));
    let (_app_state, app_patch, _) = app_plan.finish().unwrap();
    let mut different = source.clone();
    app_patch.apply_in_place(&mut different).unwrap();
    assert!(matches!(
        no_op_then_changed.apply_in_place(&mut different),
        Err(Error::PatchConflict)
    ));

    // The same exact owner remains composable, while a different property
    // state is rejected even though both patches are semantic no-ops.
    assert!(no_op.compose(&no_op).unwrap().is_noop());
    assert!(matches!(
        no_op.compose(&changed_no_op),
        Err(Error::PatchConflict)
    ));

    // The inverse composition of a changed patch has no complete scope, so
    // it retains the affected before image and remains an authenticated
    // no-op on the original owner only.
    let round_trip = changed.compose(&changed.inverse()).unwrap();
    let mut exact = source.clone();
    assert!(!round_trip.apply_in_place(&mut exact).unwrap());
    let mut after = source.clone();
    changed.apply_in_place(&mut after).unwrap();
    assert!(matches!(
        round_trip.apply_in_place(&mut after),
        Err(Error::PatchConflict)
    ));

    // A partial cancellation scope remains useful when followed by a
    // disjoint changed family. Both the retained core guard and the app
    // before image must survive composition.
    let mut app_plan = snapshot.plan();
    app_plan.set_app_props(Some(app::Props {
        company: Some("Other".into()),
        ..app::Props::default()
    }));
    let (_, app_patch, _) = app_plan.finish().unwrap();
    let disjoint = round_trip.compose(&app_patch).unwrap();
    let mut composed = source.clone();
    disjoint.apply_in_place(&mut composed).unwrap();
    let mut expected = source.clone();
    app_patch.apply_in_place(&mut expected).unwrap();
    assert_eq!(
        bytes(&composed, "/docProps/app.xml"),
        bytes(&expected, "/docProps/app.xml")
    );
    let mut wrong = source.clone();
    changed.apply_in_place(&mut wrong).unwrap();
    assert!(matches!(
        disjoint.apply_in_place(&mut wrong),
        Err(Error::PatchConflict)
    ));
}

#[test]
fn composed_no_op_scope_authenticates_absent_property_families() {
    let empty = OpcPackage::new();
    let snapshot = Snapshot::read(&empty, Host::Excel).unwrap();
    let (_, no_op, _) = snapshot.plan().finish().unwrap();
    assert!(no_op.compose(&no_op).unwrap().is_noop());
    let mut exact = empty.clone();
    assert!(!no_op.apply_in_place(&mut exact).unwrap());

    let (changed_state, changed) = edited(&snapshot, "Created");
    let (_, changed_no_op, _) = changed_state.plan().finish().unwrap();
    let composed = no_op.compose(&changed).unwrap();
    let mut composed_package = empty.clone();
    composed.apply_in_place(&mut composed_package).unwrap();
    let mut expected = empty.clone();
    changed.apply_in_place(&mut expected).unwrap();
    assert_eq!(
        bytes(&composed_package, "/docProps/core.xml"),
        bytes(&expected, "/docProps/core.xml")
    );
    // The untouched absent families retain diagnostic placeholders, but replay
    // authenticates their actual absence independently of those IDs.
    assert!(
        composed
            .inverse()
            .apply_in_place(&mut composed_package)
            .unwrap()
    );
    assert_eq!(composed_package.iter_parts().count(), 0);

    // The same wildcard absence guard composes with an application-only or
    // custom-only creation, whichever free relationship ID the plan chooses.
    let mut app_plan = snapshot.plan();
    app_plan.set_app_props(Some(app::Props {
        company: Some("Litchi".into()),
        ..app::Props::default()
    }));
    let (_, app_creation, _) = app_plan.finish().unwrap();
    let app_composed = no_op.compose(&app_creation).unwrap();
    let mut app_package = empty.clone();
    app_composed.apply_in_place(&mut app_package).unwrap();
    app_composed
        .inverse()
        .apply_in_place(&mut app_package)
        .unwrap();
    assert_eq!(app_package.iter_parts().count(), 0);

    let mut custom_plan = snapshot.plan();
    let mut custom_props = custom::Props::new();
    custom_props.insert("Created", "yes").unwrap();
    custom_plan.set_custom_props(custom_props);
    let (_, custom_creation, _) = custom_plan.finish().unwrap();
    let custom_composed = no_op.compose(&custom_creation).unwrap();
    let mut custom_package = empty.clone();
    custom_composed.apply_in_place(&mut custom_package).unwrap();
    custom_composed
        .inverse()
        .apply_in_place(&mut custom_package)
        .unwrap();
    assert_eq!(custom_package.iter_parts().count(), 0);

    // A creation followed by an independently captured exact-after no-op
    // keeps its creation relationship and still composes to one patch.
    let mut created = empty.clone();
    changed.apply_in_place(&mut created).unwrap();
    let after_snapshot = Snapshot::read(&created, Host::Excel).unwrap();
    let (_, after_no_op, _) = after_snapshot.plan().finish().unwrap();
    let creation_then_no_op = changed.compose(&after_no_op).unwrap();
    let mut recreated = empty.clone();
    creation_then_no_op.apply_in_place(&mut recreated).unwrap();
    creation_then_no_op
        .inverse()
        .apply_in_place(&mut recreated)
        .unwrap();
    assert_eq!(recreated.iter_parts().count(), 0);
    assert!(matches!(
        no_op.compose(&changed_no_op),
        Err(Error::PatchConflict)
    ));
}

#[test]
fn composed_creation_authenticates_its_actual_relationship_id() {
    let empty = OpcPackage::new();
    let snapshot = Snapshot::read(&empty, Host::Excel).unwrap();
    let (_, no_op, _) = snapshot.plan().finish().unwrap();
    let (_, creation) = edited(&snapshot, "Created");
    let composed = no_op.compose(&creation).unwrap();

    // The no-op's absent-family scope uses wildcard relationship IDs. The
    // creation itself still has to reject a recipient that occupied the
    // concrete ID selected by its after package.
    let mut occupied = empty.clone();
    occupied.add_part(Box::new(BlobPart::new(
        PackURI::new("/dummy.xml").unwrap(),
        "application/xml".to_owned(),
        b"<dummy/>".to_vec(),
    )));
    occupied
        .rels_mut()
        .try_add_relationship(
            "urn:litchi:test".to_owned(),
            "dummy.xml".to_owned(),
            "rId1".to_owned(),
            TargetMode::Internal,
        )
        .unwrap();
    assert!(matches!(
        composed.apply_in_place(&mut occupied),
        Err(Error::PatchConflict)
    ));

    let occupied_source = source_backed(&occupied);
    assert!(matches!(
        composed.prepare_source_backed(&occupied_source),
        Err(Error::PatchConflict)
    ));
}

#[test]
fn removal_composes_only_with_a_no_op_from_its_exact_after_graph() {
    let source = full();
    let snapshot = Snapshot::read(&source, Host::Word).unwrap();
    let mut remove_plan = snapshot.plan();
    remove_plan.remove_custom_props();
    let (_, removal, _) = remove_plan.finish().unwrap();

    let mut removed = source.clone();
    assert!(removal.apply_in_place(&mut removed).unwrap());
    let removed_snapshot = Snapshot::read(&removed, Host::Word).unwrap();
    let (_, exact_no_op, _) = removed_snapshot.plan().finish().unwrap();
    let composed = removal.compose(&exact_no_op).unwrap();
    let mut round_trip = source.clone();
    assert!(composed.apply_in_place(&mut round_trip).unwrap());
    assert!(composed.inverse().apply_in_place(&mut round_trip).unwrap());
    assert_eq!(
        bytes(&round_trip, "/docProps/custom.xml"),
        bytes(&source, "/docProps/custom.xml")
    );

    // A separately captured no-op may choose a different free ID because an
    // unrelated relationship occupies the old gap. That phantom ID is not
    // part of the property state, so the valid composition must remain
    // supported.
    let removed_custom_id = source
        .rels()
        .iter()
        .find(|relationship| relationship.reltype() == rt::CUSTOM_PROPERTIES)
        .unwrap()
        .r_id()
        .to_owned();
    let mut unrelated = removed.clone();
    unrelated.add_part(Box::new(BlobPart::new(
        PackURI::new("/dummy.xml").unwrap(),
        "application/octet-stream".to_owned(),
        Vec::new(),
    )));
    unrelated
        .rels_mut()
        .try_add_relationship(
            "urn:litchi:test".to_owned(),
            "dummy.xml".to_owned(),
            removed_custom_id,
            TargetMode::Internal,
        )
        .unwrap();
    let unrelated_snapshot = Snapshot::read(&unrelated, Host::Word).unwrap();
    let (_, different_no_op, _) = unrelated_snapshot.plan().finish().unwrap();
    let composed_with_different_gap = removal.compose(&different_no_op).unwrap();
    let mut different_gap_result = source.clone();
    composed_with_different_gap
        .apply_in_place(&mut different_gap_result)
        .unwrap();
    composed_with_different_gap
        .inverse()
        .apply_in_place(&mut different_gap_result)
        .unwrap();
    assert_eq!(
        bytes(&different_gap_result, "/docProps/custom.xml"),
        bytes(&source, "/docProps/custom.xml")
    );
}

#[test]
fn one_plan_reserves_removed_relationship_ids_for_later_creations() {
    let source = app_only();
    let before = PackageWriter::to_bytes(&source).unwrap();
    let snapshot = Snapshot::read(&source, Host::Word).unwrap();
    let mut plan = snapshot.plan();
    plan.set_app_props(None);
    let mut custom_props = custom::Props::new();
    custom_props.insert("Created", "yes").unwrap();
    plan.set_custom_props(custom_props);
    let (_, patch, _) = plan.finish().unwrap();

    let mut changed = source.clone();
    assert!(patch.apply_in_place(&mut changed).unwrap());
    let changed_snapshot = Snapshot::read(&changed, Host::Word).unwrap();
    assert!(changed_snapshot.app_props().unwrap().is_none());
    assert_eq!(
        changed_snapshot.custom_props().get("Created"),
        Some(&Value::Text("yes".into()))
    );

    let mut restored = changed.clone();
    assert!(patch.inverse().apply_in_place(&mut restored).unwrap());
    assert_eq!(
        PackageWriter::to_bytes(&restored).unwrap(),
        before,
        "a delete/create plan must replay and invert byte-for-byte"
    );
}

#[test]
fn composed_removal_authenticates_all_inbound_relationships() {
    let source = full();
    let snapshot = Snapshot::read(&source, Host::Word).unwrap();
    let mut remove_plan = snapshot.plan();
    remove_plan.remove_custom_props();
    let (_, removal, _) = remove_plan.finish().unwrap();

    let mut removed = source.clone();
    removal.apply_in_place(&mut removed).unwrap();
    let removed_snapshot = Snapshot::read(&removed, Host::Word).unwrap();
    let (_, exact_no_op, _) = removed_snapshot.plan().finish().unwrap();
    let composed = removal.compose(&exact_no_op).unwrap();

    // The supplemental state scope authenticates the custom part itself, but
    // the actual removal must still prove that no relationship besides the
    // recorded owner targets it.
    let mut extra_inbound = source.clone();
    extra_inbound
        .rels_mut()
        .try_add_relationship(
            "urn:litchi:test".to_owned(),
            "docProps/custom.xml".to_owned(),
            "rIdExtra".to_owned(),
            TargetMode::Internal,
        )
        .unwrap();
    assert!(matches!(
        composed.apply_in_place(&mut extra_inbound),
        Err(Error::Package(crate::Error::Relationship(_)))
    ));

    let extra_source = source_backed(&extra_inbound);
    assert!(matches!(
        composed.prepare_source_backed(&extra_source),
        Err(Error::Package(crate::Error::Relationship(_)))
    ));
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
    for (_field, update) in updates {
        let mut plan = snapshot.plan();
        plan.update_metadata(update).unwrap();
        assert!(matches!(
            plan.finish(),
            Err(Error::Refused(refusal))
                if refusal.kind() == litchi_core::edit::RefusalKind::UnsupportedProperty
        ));
    }
    let mut plan = snapshot.plan();
    plan.update_metadata(|metadata| metadata.page_count = Some(u32::MAX))
        .unwrap();
    assert!(matches!(plan.finish(), Err(Error::Invalid(_))));
}

#[test]
fn metadata_refusals_are_typed_and_content_free() {
    let snapshot = Snapshot::read(&full(), Host::Word).unwrap();
    let mut plan = snapshot.plan();
    plan.update_metadata(|metadata| {
        metadata.content_type = Some("authored/type".into());
    })
    .unwrap();
    let Error::Refused(refusal) = plan.finish().unwrap_err() else {
        panic!("unsupported metadata fields must be typed refusals");
    };
    assert_eq!(
        refusal.kind(),
        litchi_core::edit::RefusalKind::UnsupportedProperty
    );
    assert_eq!(refusal.operation(), "set_metadata");
    assert!(!refusal.reason().contains("contentType"));
}

#[test]
fn nested_opc_source_changes_are_promoted() {
    let expected = litchi_core::SourceVersion::new(7, 11);
    let observed = litchi_core::SourceVersion::new(7, 12);
    let error = Error::from(OpcError::IncompleteOutput {
        written: 1,
        source: Box::new(OpcError::SourceChanged {
            expected,
            actual: observed,
        }),
    });
    assert!(matches!(
        error,
        Error::SourceChanged {
            expected: got_expected,
            observed: got_observed,
        } if got_expected == expected && got_observed == observed
    ));
}

#[test]
fn direct_common_semantic_errors_keep_their_metadata_classification() {
    use litchi_core::edit::{Refusal, RefusalKind};

    let refusal = Refusal::new(
        RefusalKind::UnsupportedContent,
        "set_metadata",
        "unsupported metadata content cannot be rewritten",
    );
    assert!(matches!(
        Error::from(crate::Error::Refused(refusal)),
        Error::Refused(value)
            if value.kind() == RefusalKind::UnsupportedContent
                && value.reason() == "unsupported metadata content cannot be rewritten"
    ));
    assert!(matches!(
        Error::from(crate::Error::PatchConflict),
        Error::PatchConflict
    ));
    assert!(matches!(Error::from(crate::Error::Signed), Error::Signed));
    let expected = litchi_core::SourceVersion::new(8, 1);
    let observed = litchi_core::SourceVersion::new(8, 2);
    assert!(matches!(
        Error::from(crate::Error::SourceChanged { expected, observed }),
        Error::SourceChanged { expected: before, observed: after }
            if before == expected && after == observed
    ));
    assert!(matches!(
        Error::from(crate::Error::Internal("prepared metadata invariant failed")),
        Error::Internal("prepared metadata invariant failed")
    ));
}

#[test]
fn malformed_source_property_xml_is_package_corruption() {
    let mut package = full();
    package
        .get_part_mut(&PackURI::new("/docProps/core.xml").unwrap())
        .unwrap()
        .set_blob(b"<cp:coreProperties xmlns:cp=\"broken\">".to_vec());
    assert!(matches!(
        Snapshot::read(&package, Host::Word),
        Err(Error::Package(_))
    ));

    let mut package = full();
    package
        .get_part_mut(&PackURI::new("/docProps/app.xml").unwrap())
        .unwrap()
        .set_blob(b"<Properties>".to_vec());
    let snapshot = Snapshot::read(&package, Host::Word).unwrap();
    assert!(matches!(snapshot.metadata(), Err(Error::Package(_))));
}

#[test]
fn authored_application_xml_text_remains_an_invalid_value() {
    let package = full();
    let mut edit = Edit::new(&package, Host::Word).unwrap();
    edit.update_metadata(|metadata| metadata.company = Some("bad\0value".to_owned()))
        .unwrap();
    assert!(matches!(
        edit.commit(),
        Err(Error::Invalid(message)) if message.contains("XML 1.0-forbidden character")
    ));
}

#[test]
fn source_xml_controls_are_corruption_but_authored_core_controls_are_invalid_values() {
    let mut package = full();
    let app = String::from_utf8(bytes(&package, "/docProps/app.xml").unwrap())
        .unwrap()
        .replace("<Company>Old</Company>", "<Company>&#x1;</Company>")
        .into_bytes();
    package
        .get_part_mut(&PackURI::new("/docProps/app.xml").unwrap())
        .unwrap()
        .set_blob(app);
    let snapshot = Snapshot::read(&package, Host::Word).unwrap();
    assert!(matches!(
        snapshot.metadata(),
        Err(Error::Package(crate::Error::Invalid(message)))
            if message.contains("U+0001")
    ));

    let mut package = full();
    let core = String::from_utf8(bytes(&package, "/docProps/core.xml").unwrap())
        .unwrap()
        .replace("<dc:title>Before</dc:title>", "<dc:title>&#x1;</dc:title>")
        .into_bytes();
    package
        .get_part_mut(&PackURI::new("/docProps/core.xml").unwrap())
        .unwrap()
        .set_blob(core);
    assert!(matches!(
        Snapshot::read(&package, Host::Word),
        Err(Error::Package(crate::Error::Invalid(message)))
            if message.contains("U+0001")
    ));

    let mut package = full();
    let core = String::from_utf8(bytes(&package, "/docProps/core.xml").unwrap())
        .unwrap()
        .replace(
            "<cp:coreProperties",
            "<cp:coreProperties xmlns:opaque=\"urn:opaque&#x1;\"",
        )
        .into_bytes();
    package
        .get_part_mut(&PackURI::new("/docProps/core.xml").unwrap())
        .unwrap()
        .set_blob(core);
    assert!(matches!(
        Snapshot::read(&package, Host::Word),
        Err(Error::Package(crate::Error::Invalid(message)))
            if message.contains("U+0001")
    ));

    let mut package = full();
    let core = String::from_utf8(bytes(&package, "/docProps/core.xml").unwrap())
        .unwrap()
        .replace(
            "</cp:coreProperties>",
            "<cp:opaque><cp:nested>&#x1;</cp:nested></cp:opaque></cp:coreProperties>",
        )
        .into_bytes();
    package
        .get_part_mut(&PackURI::new("/docProps/core.xml").unwrap())
        .unwrap()
        .set_blob(core);
    assert!(matches!(
        Snapshot::read(&package, Host::Word),
        Err(Error::Package(crate::Error::Invalid(message)))
            if message.contains("U+0001")
    ));

    let mut package = full();
    let core = String::from_utf8(bytes(&package, "/docProps/core.xml").unwrap())
        .unwrap()
        .replace(
            "</cp:coreProperties>",
            "<cp:opaque><!--\u{1}--></cp:opaque></cp:coreProperties>",
        )
        .into_bytes();
    package
        .get_part_mut(&PackURI::new("/docProps/core.xml").unwrap())
        .unwrap()
        .set_blob(core);
    assert!(matches!(
        Snapshot::read(&package, Host::Word),
        Err(Error::Package(crate::Error::Invalid(message)))
            if message.contains("U+0001")
    ));

    let mut package = full();
    let core = String::from_utf8(bytes(&package, "/docProps/core.xml").unwrap())
        .unwrap()
        .replace(
            "</cp:coreProperties>",
            "<cp:opaque><?keep \u{1}?></cp:opaque></cp:coreProperties>",
        )
        .into_bytes();
    package
        .get_part_mut(&PackURI::new("/docProps/core.xml").unwrap())
        .unwrap()
        .set_blob(core);
    assert!(matches!(
        Snapshot::read(&package, Host::Word),
        Err(Error::Package(crate::Error::Invalid(message)))
            if message.contains("U+0001")
    ));

    let package = full();
    let mut edit = Edit::new(&package, Host::Word).unwrap();
    edit.update_metadata(|metadata| metadata.title = Some("bad\0value".to_owned()))
        .unwrap();
    assert!(matches!(
        edit.commit(),
        Err(Error::Invalid(message)) if message.contains("U+0000")
    ));
}

#[test]
fn an_empty_package_gains_canonical_parts() {
    let mut package = OpcPackage::new();
    let mut edit = Edit::new(&package, Host::Excel).unwrap();
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
    assert!(patch.apply_in_place(&mut package).unwrap());
    let metadata = super::read_with_report(&package, &mut ToleranceReport::default())
        .unwrap()
        .unwrap();
    assert_eq!(metadata.author.as_deref(), Some("Ada"));
    assert_eq!(
        app::read(&package).unwrap().unwrap().company.as_deref(),
        Some("Litchi")
    );
    assert!(patch.inverse().apply_in_place(&mut package).unwrap());
    assert_eq!(package.iter_parts().count(), 0);
    assert_eq!(package.rels().iter().count(), 0);
}

#[test]
fn a_stale_patch_conflicts_without_mutation() {
    let mut package = full();
    let mut edit = Edit::new(&package, Host::Word).unwrap();
    edit.update_metadata(|metadata| metadata.title = Some("One".into()))
        .unwrap();
    let first = edit.commit().unwrap().patch().clone();
    let mut edit = Edit::new(&package, Host::Word).unwrap();
    edit.update_metadata(|metadata| metadata.title = Some("Two".into()))
        .unwrap();
    let second = edit.commit().unwrap().patch().clone();
    assert!(first.apply_in_place(&mut package).unwrap());
    let before = bytes(&package, "/docProps/core.xml");
    assert!(matches!(
        second.apply_in_place(&mut package),
        Err(Error::PatchConflict)
    ));
    assert_eq!(bytes(&package, "/docProps/core.xml"), before);
}

fn source_backed(package: &OpcPackage) -> litchi_opc::SourceBackedPackage {
    let bytes = PackageWriter::to_bytes(package).unwrap();
    litchi_opc::SourceBackedPackage::from_vec(
        bytes,
        litchi_opc::SourceLimits::default(),
        litchi_opc::SourceCacheLimits::default(),
        None,
    )
    .unwrap()
}

fn edited(snapshot: &Snapshot, title: &str) -> (Snapshot, super::Patch) {
    let mut plan = snapshot.plan();
    plan.update_metadata(|metadata| metadata.title = Some(title.into()))
        .unwrap();
    let (after, patch, _) = plan.finish().unwrap();
    (after, patch)
}

#[test]
fn composed_patches_apply_as_their_sequence_and_an_inverse_composes_to_a_no_op() {
    let package = full();
    let snapshot = Snapshot::read(&package, Host::Word).unwrap();
    let (once, first) = edited(&snapshot, "One");
    let (_, second) = edited(&once, "Two");
    let composed = first.compose(&second).unwrap();
    let mut sequential = package.clone();
    first.apply_in_place(&mut sequential).unwrap();
    second.apply_in_place(&mut sequential).unwrap();
    let mut direct = package.clone();
    composed.apply_in_place(&mut direct).unwrap();
    assert_eq!(
        bytes(&direct, "/docProps/core.xml"),
        bytes(&sequential, "/docProps/core.xml")
    );
    assert!(first.compose(&first.inverse()).unwrap().is_noop());
    // A patch that does not continue the first one is refused.
    assert!(matches!(first.compose(&first), Err(Error::PatchConflict)));
}

#[test]
fn a_patch_prepares_over_a_source_backed_package_without_materializing_it() {
    let package = full();
    let source = source_backed(&package);
    let snapshot = Snapshot::read_source_backed(&source, Host::Word).unwrap();
    let mut plan = snapshot.plan();
    plan.update_metadata(|metadata| {
        metadata.title = Some("After".into());
        metadata.manager = Some("Lin".into());
    })
    .unwrap();
    let mut custom = plan.custom_props().clone();
    custom.insert("Phase", "two").unwrap();
    plan.set_custom_props(custom);
    let (after, patch, _) = plan.finish().unwrap();
    let topology = patch.prepare_source_backed(&source).unwrap();
    let mut published = Vec::new();
    topology.publish_to_stream(&mut published).unwrap();
    let reopened = OpcPackage::from_bytes(&published, litchi_opc::SourceLimits::default()).unwrap();
    let read = Snapshot::read(&reopened, Host::Word).unwrap();
    assert_eq!(read.metadata().unwrap(), after.metadata().unwrap());
    assert_eq!(
        read.custom_props().get("Phase"),
        Some(&Value::Text("two".into()))
    );

    // The same patch on the owning package stores the same parts.
    let mut owning = package.clone();
    patch.apply_in_place(&mut owning).unwrap();
    for name in [
        "/docProps/core.xml",
        "/docProps/app.xml",
        "/docProps/custom.xml",
    ] {
        assert_eq!(bytes(&reopened, name), bytes(&owning, name), "{name}");
    }

    // A stale patch is refused, and an empty patch is the exact source.
    let changed = source_backed(&owning);
    assert!(matches!(
        patch.prepare_source_backed(&changed),
        Err(Error::PatchConflict)
    ));
    let mut exact = Vec::new();
    super::Patch::default()
        .prepare_source_backed(&source)
        .unwrap()
        .publish_to_stream(&mut exact)
        .unwrap();
    assert_eq!(exact, PackageWriter::to_bytes(&package).unwrap());
}

#[test]
fn a_changed_source_backed_patch_refuses_a_signed_package() {
    let mut package = full();
    let snapshot = Snapshot::read(&package, Host::Word).unwrap();
    let (_, patch) = edited(&snapshot, "Changed");
    sign(&mut package);
    let source = source_backed(&package);
    assert!(matches!(
        patch.prepare_source_backed(&source),
        Err(Error::Signed)
    ));
}

#[test]
fn an_edit_publishes_a_new_package_and_leaves_its_source_unchanged() {
    let package = full();
    let core_before = bytes(&package, "/docProps/core.xml").unwrap();
    let mut edit = Edit::new(&package, Host::Word).unwrap();
    edit.update_metadata(|metadata| metadata.title = Some("Staged".into()))
        .unwrap();
    edit.reset();
    let unchanged = edit.clone().commit().unwrap();
    assert!(!unchanged.diagnostics().changed());
    assert!(unchanged.patch().is_noop());

    edit.update_metadata(|metadata| metadata.title = Some("After".into()))
        .unwrap();
    let commit = edit.commit().unwrap();
    assert!(commit.diagnostics().changed());
    assert_eq!(bytes(&package, "/docProps/core.xml").unwrap(), core_before);
    let after = Snapshot::read(commit.package(), Host::Word).unwrap();
    assert_eq!(after.metadata().unwrap().title.as_deref(), Some("After"));
    let (published, patch) = commit.into_parts();
    let mut restored = published.clone();
    assert!(patch.inverse().apply_in_place(&mut restored).unwrap());
    assert_eq!(bytes(&restored, "/docProps/core.xml").unwrap(), core_before);
}
