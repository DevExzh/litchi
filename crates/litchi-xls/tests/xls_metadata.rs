#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    clippy::shadow_unrelated,
    reason = "integration tests favor concise panic-driven assertions"
)]

//! XLS metadata: fresh authoring through `Builder::set_metadata` and the
//! opened-package `metadata` transaction family (ADR 0034 §7).

use std::io::Cursor;

use litchi_cfb::{OleFile, OleWriter};
use litchi_core::Metadata;
use litchi_ole_common::property_set::{
    Binding, CodePage, PropertySetReader, Value, document_summary,
};
use litchi_xls::package::Error as PackageError;
use litchi_xls::package::metadata::Error;
use litchi_xls::{Builder, EncryptionEntropy, EncryptionProfile, OpenOptions, Package};

fn metadata() -> Metadata {
    Metadata {
        title: Some("Budget".to_owned()),
        author: Some("Ada Lovelace".to_owned()),
        company: Some("Litchi Project".to_owned()),
        created: Some("2026-10-05T09:30:00Z".parse().unwrap()),
        ..Metadata::default()
    }
}

fn fresh() -> Builder {
    let mut builder = Builder::new();
    let sheet = builder.add_worksheet("Sheet1").unwrap();
    builder.write_number(sheet, 0, 0, 42.0).unwrap();
    builder
}

fn signed_package() -> Package {
    let source = fresh().build().unwrap().to_bytes().unwrap();
    let mut ole = OleFile::open(Cursor::new(source), litchi_cfb::OleFileLimits::default()).unwrap();
    let workbook = ole.open_stream(&["Workbook"]).unwrap();
    let mut writer = OleWriter::new();
    writer.create_stream(&["Workbook"], &workbook).unwrap();
    writer
        .create_stream(&["DigitalSignature"], b"opaque signature")
        .unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    Package::from_bytes(&output.into_inner()).unwrap()
}

fn stream_names(bytes: &[u8]) -> Vec<String> {
    OleFile::open(Cursor::new(bytes), litchi_cfb::OleFileLimits::default())
        .unwrap()
        .list_streams()
        .into_iter()
        .map(|path| path.join("/"))
        .collect()
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
    let bytes = package.to_bytes().unwrap();
    let mut ole = OleFile::open(Cursor::new(bytes), litchi_cfb::OleFileLimits::default()).unwrap();
    let custom = ole.get_metadata().unwrap().custom_properties;
    assert_eq!(custom.get("Reviewed"), Some(&Value::Bool(true)));
}

#[test]
fn output_without_metadata_is_byte_identical() {
    let plain = fresh().build().unwrap().to_bytes().unwrap();
    assert!(
        !stream_names(&plain)
            .iter()
            .any(|name| name.starts_with('\u{5}'))
    );
    let mut cleared = fresh();
    cleared.set_metadata(metadata()).unwrap();
    cleared.set_metadata(Metadata::default()).unwrap();
    assert_eq!(cleared.build().unwrap().to_bytes().unwrap(), plain);

    let mut first = fresh();
    first.set_metadata(metadata()).unwrap();
    let mut second = fresh();
    second.set_metadata(metadata()).unwrap();
    assert_eq!(
        first.build().unwrap().to_bytes().unwrap(),
        second.build().unwrap().to_bytes().unwrap()
    );
}

#[test]
fn encryption_policies() {
    let mut cryptoapi = fresh();
    cryptoapi.set_metadata(metadata()).unwrap();
    cryptoapi
        .set_password("secret", EncryptionProfile::CryptoApiRc4 { key_bits: 128 })
        .unwrap();
    cryptoapi.set_encryption_entropy(EncryptionEntropy::from_bytes([1; 16], [2; 16]));
    assert!(matches!(
        cryptoapi.build(),
        Err(litchi_xls::package::Error::Reader(
            litchi_xls::Error::EncryptedMetadata
        ))
    ));

    let mut rc4 = fresh();
    rc4.set_metadata(metadata()).unwrap();
    rc4.set_password("secret", EncryptionProfile::OfficeBinaryRc4)
        .unwrap();
    rc4.set_encryption_entropy(EncryptionEntropy::from_bytes([1; 16], [2; 16]));
    let package = rc4.build().unwrap();
    assert!(package.is_encrypted());
    let reopened = Package::from_bytes_with_options(
        &package.to_encrypted_bytes().unwrap(),
        OpenOptions::default().password(Some("secret")),
    )
    .unwrap();
    assert_eq!(
        reopened.metadata().unwrap().title.as_deref(),
        Some("Budget")
    );
    assert!(matches!(reopened.edit_metadata(), Err(Error::Protected)));
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
        Err(litchi_xls::Error::Metadata(_))
    ));
    assert_eq!(builder.summary_properties(), &before);
}

#[test]
fn metadata_transaction_round_trip() {
    let package = fresh().build().unwrap();
    let source = package.to_bytes().unwrap();

    let commit = package.edit_metadata().unwrap().commit().unwrap();
    assert!(!commit.diagnostics().changed());
    assert_eq!(commit.package().to_bytes().unwrap(), source);

    let mut edit = package.edit_metadata().unwrap();
    assert!(edit.merge_metadata(&metadata()).unwrap());
    edit.properties_mut()
        .unwrap()
        .set("Department", "Research")
        .unwrap();
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
        edited.workbook().worksheet_names().unwrap(),
        package.workbook().worksheet_names().unwrap()
    );

    assert_eq!(
        package
            .apply_metadata(&patch)
            .unwrap()
            .package()
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
            .package()
            .to_bytes()
            .unwrap(),
        source
    );
}

#[test]
fn signed_metadata_and_property_set_noops_preserve_bytes_and_changes_refuse() {
    let package = signed_package();
    let source = package.to_bytes().unwrap();

    let mut metadata = package.edit_metadata().unwrap();
    assert!(metadata.properties().unwrap().is_empty());
    let noop = metadata.commit().unwrap();
    assert!(!noop.diagnostics().changed());
    assert_eq!(noop.package().to_bytes().unwrap(), source);
    assert_eq!(
        Package::from_bytes(&source)
            .unwrap()
            .apply_metadata(noop.patch())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap(),
        source
    );

    let mut changed = package.edit_metadata().unwrap();
    changed
        .properties_mut()
        .unwrap()
        .set("title", "changed")
        .unwrap();
    assert!(matches!(changed.commit(), Err(Error::Protected)));

    let mut property_sets = package.edit_property_sets().unwrap();
    assert_eq!(property_sets.before().unwrap().bytes(), source.as_slice());
    let noop = property_sets.commit().unwrap();
    assert!(!noop.diagnostics().changed());
    assert_eq!(noop.package().to_bytes().unwrap(), source);

    let mut changed = package.edit_property_sets().unwrap();
    assert!(matches!(
        changed.edit_summary_information(|edit| edit.set_title("changed")),
        Err(PackageError::Reader(litchi_xls::Error::Signed))
    ));
}

#[test]
fn excel_workbooks_keep_their_other_streams() {
    let path =
        std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../test-data/ole/xls/Simple.xls");
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
