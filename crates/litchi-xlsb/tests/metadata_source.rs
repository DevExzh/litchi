#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "integration tests use panic-on-failure extraction"
)]

//! Document-property edits of a source-backed XLSB owner (record 0261).

use std::io::Cursor;
use std::path::PathBuf;

use litchi_opc::{BlobPart, OpcPackage, PackURI, PackageWriter, Part as _, TargetMode};
use litchi_xlsb::metadata::Error;
use litchi_xlsb::{Package, SourceBackedWorkbook};

fn fixture(name: &str) -> Vec<u8> {
    std::fs::read(
        PathBuf::from(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/ooxml/xlsb")
            .join(name),
    )
    .expect("XLSB fixture")
}

fn source(bytes: &[u8]) -> SourceBackedWorkbook {
    SourceBackedWorkbook::from_reader(Cursor::new(bytes.to_vec())).expect("source owner")
}

fn published(workbook: &SourceBackedWorkbook) -> Vec<u8> {
    let mut bytes = Vec::new();
    workbook.write_to(&mut bytes).expect("publication");
    bytes
}

/// Every part outside `docProps/` with its payload.
fn workbook_parts(bytes: &[u8]) -> Vec<(String, Vec<u8>)> {
    let package =
        OpcPackage::from_bytes(bytes, litchi_opc::SourceLimits::default()).expect("package");
    let mut parts: Vec<_> = package
        .iter_parts()
        .filter(|part| !part.partname().as_str().starts_with("/docProps/"))
        .map(|part| {
            let name = part.partname().clone();
            let blob = package.get_part(&name).expect("part").blob().to_vec();
            (name.as_str().to_owned(), blob)
        })
        .collect();
    parts.sort();
    parts
}

#[test]
fn a_source_edit_publishes_the_property_parts_over_the_source() {
    let bytes = fixture("Simple.xlsb");
    let workbook = source(&bytes);
    let original = workbook.metadata().unwrap();
    let mut edit = workbook.edit_metadata().unwrap();
    edit.update(|metadata| {
        metadata.title = Some("Budget".to_owned());
        metadata.company = Some("Litchi".to_owned());
    })
    .unwrap();
    let commit = edit.commit().unwrap();
    assert!(commit.diagnostics().changed());
    assert!(!commit.patch().is_noop());
    let edited = commit.into_package();
    assert!(edited.has_metadata_edits());
    assert!(!workbook.has_metadata_edits());

    // Reads observe the edit; the source owner is unchanged.
    assert_eq!(edited.metadata().unwrap().title.as_deref(), Some("Budget"));
    assert_eq!(
        edited.app_props().unwrap().unwrap().company.as_deref(),
        Some("Litchi")
    );
    assert_eq!(workbook.metadata().unwrap(), original);
    assert_eq!(published(&workbook), bytes);

    // Publication changes only the property parts.
    let output = published(&edited);
    assert_eq!(workbook_parts(&output), workbook_parts(&bytes));
    let reopened = Package::from_bytes(&output).unwrap();
    assert_eq!(
        reopened.metadata().unwrap().title.as_deref(),
        Some("Budget")
    );
    assert_eq!(
        reopened.metadata().unwrap(),
        edited.metadata().unwrap(),
        "the published parts are the edited snapshot"
    );

    // The same edit through the eager owner stores the same property bytes.
    let eager = Package::from_bytes(&bytes).unwrap();
    let mut eager_edit = eager.edit_metadata().unwrap();
    eager_edit
        .update(|metadata| {
            metadata.title = Some("Budget".to_owned());
            metadata.company = Some("Litchi".to_owned());
        })
        .unwrap();
    let eager_patch = eager_edit.commit().unwrap().into_parts().1;
    let source_patch = {
        let mut edit = workbook.edit_metadata().unwrap();
        edit.update(|metadata| {
            metadata.title = Some("Budget".to_owned());
            metadata.company = Some("Litchi".to_owned());
        })
        .unwrap();
        edit.commit().unwrap().into_parts().1
    };
    assert_eq!(eager_patch, source_patch);

    // Materialization keeps the edit.
    let materialized = edited.to_workbook().unwrap();
    assert_eq!(
        litchi_ooxml_common::properties::read_with_report(
            materialized.opc_package(),
            &mut Default::default()
        )
        .unwrap()
        .unwrap()
        .title
        .as_deref(),
        Some("Budget")
    );
}

#[test]
fn an_exact_no_op_keeps_the_source_owner_and_its_bytes() {
    let bytes = fixture("Simple.xlsb");
    let workbook = source(&bytes);
    let current = workbook.metadata().unwrap();
    let mut edit = workbook.edit_metadata().unwrap();
    edit.set(current);
    let commit = edit.commit().unwrap();
    assert!(!commit.diagnostics().changed());
    assert!(commit.patch().is_noop());
    let after = commit.into_package();
    assert!(!after.has_metadata_edits());
    assert_eq!(published(&after), bytes);
}

#[test]
fn source_noop_replay_checks_equal_and_foreign_recorded_bases() {
    let bytes = fixture("Simple.xlsb");
    let workbook = source(&bytes);
    let equal = workbook.edit_metadata().unwrap().commit().unwrap();
    let replayed = workbook.apply_metadata(equal.patch()).unwrap();
    assert!(!replayed.diagnostics().changed());
    assert_eq!(published(&replayed.into_package()), bytes);

    let mut edit = workbook.edit_metadata().unwrap();
    edit.update(|metadata| metadata.title = Some("current".to_owned()))
        .unwrap();
    let current = edit.commit().unwrap().into_package();
    let current_noop = current.edit_metadata().unwrap().commit().unwrap();
    let replayed = current.apply_metadata(current_noop.patch()).unwrap();
    assert_eq!(
        published(replayed.package()),
        published(&current),
        "a no-op from a revised source owner must authenticate its current state"
    );

    let eager = Package::from_bytes(&bytes).unwrap();
    let mut changed = eager.edit_metadata().unwrap();
    changed
        .update(|metadata| metadata.title = Some("foreign".to_owned()))
        .unwrap();
    let foreign = changed.commit().unwrap().into_package();
    let foreign_noop = foreign.edit_metadata().unwrap().commit().unwrap();
    assert!(matches!(
        workbook.apply_metadata(foreign_noop.patch()),
        Err(Error::PatchConflict)
    ));
}

#[test]
fn sequential_edits_compose_and_an_undo_restores_the_source_values() {
    let bytes = fixture("Simple.xlsb");
    let workbook = source(&bytes);
    let mut first = workbook.edit_metadata().unwrap();
    first
        .update(|metadata| metadata.title = Some("One".to_owned()))
        .unwrap();
    let (once, first_patch) = first.commit().unwrap().into_parts();
    let mut second = once.edit_metadata().unwrap();
    second
        .update(|metadata| metadata.subject = Some("Two".to_owned()))
        .unwrap();
    let (twice, second_patch) = second.commit().unwrap().into_parts();
    let reopened = Package::from_bytes(&published(&twice)).unwrap();
    let metadata = reopened.metadata().unwrap();
    assert_eq!(metadata.title.as_deref(), Some("One"));
    assert_eq!(metadata.subject.as_deref(), Some("Two"));
    assert_eq!(workbook_parts(&published(&twice)), workbook_parts(&bytes));

    // Each step's patch replays on the eager owner of its before state.
    let eager = Package::from_bytes(&bytes).unwrap();
    let eager_twice = eager
        .apply_metadata(&first_patch)
        .unwrap()
        .into_package()
        .apply_metadata(&second_patch)
        .unwrap()
        .into_package();
    assert_eq!(eager_twice.metadata().unwrap(), metadata);

    // Undoing both edits restores the source values over the source parts.
    let mut undo = twice.edit_metadata().unwrap();
    undo.set(workbook.metadata().unwrap());
    let restored = undo.commit().unwrap().into_package();
    let output = published(&restored);
    assert_eq!(
        Package::from_bytes(&output).unwrap().metadata().unwrap(),
        workbook.metadata().unwrap()
    );
    assert_eq!(workbook_parts(&output), workbook_parts(&bytes));
}

#[test]
fn source_metadata_patches_replay_against_the_content_base() {
    let bytes = fixture("Simple.xlsb");
    let eager = Package::from_bytes(&bytes).unwrap();
    let mut first = eager.edit_metadata().unwrap();
    first
        .update(|metadata| metadata.title = Some("One".to_owned()))
        .unwrap();
    let (once, first_patch) = first.commit().unwrap().into_parts();
    let mut second = once.edit_metadata().unwrap();
    second
        .update(|metadata| metadata.subject = Some("Two".to_owned()))
        .unwrap();
    let (_, second_patch) = second.commit().unwrap().into_parts();

    let source_owner = source(&bytes);
    let once = source_owner
        .apply_metadata(&first_patch)
        .unwrap()
        .into_package();
    let twice = once.apply_metadata(&second_patch).unwrap().into_package();
    assert_eq!(twice.metadata().unwrap().title.as_deref(), Some("One"));
    assert_eq!(twice.metadata().unwrap().subject.as_deref(), Some("Two"));

    let restored = twice
        .apply_metadata(&second_patch.inverse())
        .unwrap()
        .into_package()
        .apply_metadata(&first_patch.inverse())
        .unwrap()
        .into_package();
    assert_eq!(published(&restored), bytes);
    assert_eq!(published(&source_owner), bytes);
}

#[test]
fn unsupported_fields_are_refused_and_change_nothing() {
    let bytes = fixture("Simple.xlsb");
    let workbook = source(&bytes);
    let mut edit = workbook.edit_metadata().unwrap();
    edit.update(|metadata| metadata.content_type = Some("text/plain".to_owned()))
        .unwrap();
    assert!(matches!(
        edit.commit(),
        Err(Error::Refused(refusal))
            if refusal.kind() == litchi_core::edit::RefusalKind::UnsupportedProperty
    ));
    assert_eq!(published(&workbook), bytes);
}

#[test]
fn a_signed_source_refuses_a_changed_edit_and_keeps_a_no_op() {
    let bytes = signed(&fixture("Simple.xlsb"));
    let workbook = source(&bytes);
    let current = workbook.metadata().unwrap();
    let mut no_op = workbook.edit_metadata().unwrap();
    no_op.set(current);
    let kept = no_op.commit().unwrap().into_package();
    assert_eq!(published(&kept), bytes);

    let mut edit = workbook.edit_metadata().unwrap();
    edit.update(|metadata| metadata.title = Some("Changed".to_owned()))
        .unwrap();
    assert!(matches!(edit.commit(), Err(Error::Signed)));
    assert_eq!(published(&workbook), bytes);
}

fn signed(plain: &[u8]) -> Vec<u8> {
    let source = OpcPackage::from_bytes(plain, litchi_opc::SourceLimits::default()).unwrap();
    let mut package = OpcPackage::new();
    for relationship in source.rels().iter() {
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
    let names: Vec<_> = source
        .iter_parts()
        .map(|part| part.partname().clone())
        .collect();
    for name in names {
        let part = source.get_part(&name).unwrap();
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
            litchi_opc::constants::relationship_type::DIGITAL_SIGNATURE_ORIGIN.to_owned(),
            "_xmlsignatures/origin.sigs".to_owned(),
            "rIdSignatureOrigin".to_owned(),
            TargetMode::Internal,
        )
        .unwrap();
    let bytes = PackageWriter::to_bytes(&package).unwrap();
    assert!(
        OpcPackage::from_bytes(&bytes, litchi_opc::SourceLimits::default())
            .unwrap()
            .is_signed()
    );
    bytes
}
