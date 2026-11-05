#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    clippy::shadow_unrelated,
    reason = "integration tests favor concise panic-driven assertions"
)]

//! PPT metadata: fresh authoring through `Builder::set_metadata` and the
//! opened-package `metadata` transaction family (ADR 0034 §7).

use std::io::Cursor;

use litchi_cfb::OleFile;
use litchi_core::Metadata;
use litchi_ole_common::property_set::{
    Binding, CodePage, PropertySetReader, Value, document_summary,
};
use litchi_ppt::package::metadata::Error;
use litchi_ppt::{Builder, Package};

fn metadata() -> Metadata {
    Metadata {
        title: Some("Roadmap".to_owned()),
        author: Some("Ada Lovelace".to_owned()),
        company: Some("Litchi Project".to_owned()),
        created: Some("2026-10-05T09:30:00Z".parse().unwrap()),
        ..Metadata::default()
    }
}

fn fresh() -> Builder {
    let mut builder = Builder::new();
    let slide = builder.add_slide().unwrap();
    builder
        .add_textbox(slide, 0, 0, 1_000_000, 500_000, "Slide text")
        .unwrap();
    builder
}

fn root_clsid(bytes: &[u8]) -> String {
    OleFile::open(Cursor::new(bytes), litchi_cfb::OleFileLimits::default())
        .unwrap()
        .root_entry()
        .unwrap()
        .clsid
        .clone()
}

fn assert_metadata(package: &Package) {
    let read = package.metadata().unwrap();
    let expected = metadata();
    assert_eq!(read.title, expected.title);
    assert_eq!(read.author, expected.author);
    assert_eq!(read.company, expected.company);
    assert_eq!(read.created, expected.created);
}

#[test]
fn builder_metadata_reads_back() {
    let mut builder = fresh();
    builder.set_metadata(metadata()).unwrap();
    builder
        .summary_properties_mut()
        .set_user_defined("Reviewed", Value::Bool(true))
        .unwrap();
    let package = builder.build().unwrap();
    assert_metadata(&package);
    assert!(
        package
            .presentation()
            .unwrap()
            .text()
            .unwrap()
            .contains("Slide text")
    );
    let mut ole = OleFile::open(
        Cursor::new(package.to_bytes().unwrap()),
        litchi_cfb::OleFileLimits::default(),
    )
    .unwrap();
    let custom = ole.get_metadata().unwrap().custom_properties;
    assert_eq!(custom.get("Reviewed"), Some(&Value::Bool(true)));
}

#[test]
fn output_without_metadata_keeps_the_fixed_streams() {
    let plain = fresh().build().unwrap().to_bytes().unwrap();
    let mut cleared = fresh();
    cleared.set_metadata(metadata()).unwrap();
    cleared.set_metadata(Metadata::default()).unwrap();
    assert_eq!(cleared.build().unwrap().to_bytes().unwrap(), plain);

    let mut first = fresh();
    first.set_metadata(metadata()).unwrap();
    let mut second = fresh();
    second.set_metadata(metadata()).unwrap();
    let first = first.build().unwrap().to_bytes().unwrap();
    assert_eq!(first, second.build().unwrap().to_bytes().unwrap());
    assert_eq!(root_clsid(&first), root_clsid(&plain));
}

#[cfg(feature = "encryption")]
#[test]
fn encryption_with_metadata_is_refused() {
    let mut builder = fresh();
    builder.set_metadata(metadata()).unwrap();
    builder
        .set_password(
            "secret",
            litchi_ppt::EncryptionProfile::CryptoApiRc4 { key_bits: 128 },
        )
        .unwrap();
    builder.set_encryption_entropy(litchi_ppt::EncryptionEntropy::from_bytes([1; 16], [2; 16]));
    assert!(matches!(
        builder.build(),
        Err(litchi_ppt::Error::Authoring(
            litchi_ppt::WriteError::EncryptedMetadata
        ))
    ));
}

#[test]
fn unrepresentable_metadata_is_refused_without_change() {
    let mut builder = fresh();
    builder.set_metadata(metadata()).unwrap();
    let before = builder.summary_properties().clone();
    assert!(matches!(
        builder.set_metadata(Metadata {
            identifier: Some("urn:x".to_owned()),
            ..Metadata::default()
        }),
        Err(litchi_ppt::WriteError::Metadata(_))
    ));
    assert_eq!(builder.summary_properties(), &before);
}

#[test]
fn metadata_transaction_fills_the_fixed_blank_streams() {
    let package = fresh().build().unwrap();
    let source = package.to_bytes().unwrap();

    let commit = package.edit_metadata().unwrap().commit().unwrap();
    assert!(!commit.diagnostics().changed());
    assert_eq!(commit.package().to_bytes().unwrap(), source);

    let mut edit = package.edit_metadata().unwrap();
    assert!(edit.merge_metadata(&metadata()).unwrap());
    let (edited, patch) = edit.commit().unwrap().into_parts();
    assert_metadata(&edited);
    let bytes = edited.to_bytes().unwrap();
    assert_eq!(root_clsid(&bytes), root_clsid(&source));
    let mut ole = OleFile::open(
        Cursor::new(bytes.clone()),
        litchi_cfb::OleFileLimits::default(),
    )
    .unwrap();
    assert_eq!(
        document_summary::Snapshot::from_stream(
            &ole.property_set(Binding::DocumentSummaryInformation)
                .unwrap()
        )
        .unwrap()
        .codepage(),
        Some(CodePage::WINDOWS_1252)
    );
    assert_eq!(
        edited.presentation().unwrap().text().unwrap(),
        package.presentation().unwrap().text().unwrap()
    );
    assert_eq!(
        package
            .apply_metadata(&patch)
            .unwrap()
            .into_package()
            .to_bytes()
            .unwrap(),
        bytes
    );
    assert!(matches!(
        edited.apply_metadata(&patch),
        Err(Error::PatchConflict)
    ));
    assert_eq!(
        edited
            .apply_metadata(&patch.inverse())
            .unwrap()
            .into_package()
            .to_bytes()
            .unwrap(),
        source
    );
}

#[test]
fn powerpoint_files_keep_their_other_streams() {
    let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/ole/ppt/SampleShow.ppt");
    let source = std::fs::read(path).unwrap();
    let package = Package::from_bytes(&source).unwrap();
    let mut edit = package.edit_metadata().unwrap();
    edit.merge_metadata(&Metadata {
        title: Some("Retitled".to_owned()),
        ..Metadata::default()
    })
    .unwrap();
    let edited = edit.commit().unwrap().into_package();
    assert_eq!(
        edited.metadata().unwrap().title.as_deref(),
        Some("Retitled")
    );
    assert_eq!(
        edited.metadata().unwrap().author,
        package.metadata().unwrap().author
    );
    let bytes = edited.to_bytes().unwrap();
    assert_eq!(root_clsid(&bytes), root_clsid(&source));
}
