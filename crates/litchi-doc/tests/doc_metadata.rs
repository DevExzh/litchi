#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    clippy::shadow_unrelated,
    reason = "integration tests favor concise panic-driven assertions"
)]

//! DOC metadata: fresh authoring through `Builder::set_metadata` and the
//! opened-package `metadata` transaction family (ADR 0034 §7).

use std::io::Cursor;

use litchi_cfb::OleFile;
use litchi_core::Metadata;
use litchi_doc::package::metadata::Error;
use litchi_doc::writer::{EncryptionProfile, WriteError};
use litchi_doc::{Builder, OpenOptions, Package, Password};
use litchi_ole_common::property_set::{
    Binding, CodePage, PropertySetReader, Value, document_summary, summary_information,
};

fn time(text: &str) -> chrono::DateTime<chrono::Utc> {
    text.parse().unwrap()
}

fn metadata() -> Metadata {
    Metadata {
        title: Some("Quarterly report".to_owned()),
        author: Some("Ada Lovelace".to_owned()),
        subject: Some("Revenue".to_owned()),
        keywords: vec!["q3".to_owned(), "revenue".to_owned()],
        description: Some("Draft for review".to_owned()),
        company: Some("Litchi Project".to_owned()),
        created: Some(time("2026-10-05T09:30:00Z").into()),
        modified: Some(time("2026-10-05T10:45:00Z").into()),
        ..Metadata::default()
    }
}

fn fresh() -> Builder {
    let mut builder = Builder::new();
    builder.add_paragraph("Body text").unwrap();
    builder
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
    let read = package.metadata();
    let expected = metadata();
    assert_eq!(read.title, expected.title);
    assert_eq!(read.author, expected.author);
    assert_eq!(read.subject, expected.subject);
    assert_eq!(read.keywords, expected.keywords);
    assert_eq!(read.description, expected.description);
    assert_eq!(read.company, expected.company);
    assert_eq!(read.created, expected.created);
    assert_eq!(read.modified, expected.modified);
    assert_eq!(package.document().text().unwrap(), "Body text\r");
}

#[test]
fn builder_metadata_reads_back_through_every_reader() {
    let mut builder = fresh();
    builder.set_metadata(metadata()).unwrap();
    builder
        .summary_properties_mut()
        .set_user_defined("Pages reviewed", Value::I4(12))
        .unwrap();
    let package = builder.build().unwrap();
    assert_metadata(&package);
    let bytes = package.to_bytes().unwrap();

    let mut ole = OleFile::open(
        Cursor::new(bytes.clone()),
        litchi_cfb::OleFileLimits::default(),
    )
    .unwrap();
    let summary = summary_information::Snapshot::from_stream(
        &ole.property_set(Binding::SummaryInformation).unwrap(),
    )
    .unwrap();
    assert_eq!(summary.title(), Some("Quarterly report"));
    assert_eq!(summary.codepage(), Some(CodePage::WINDOWS_1252));
    let document_summary = ole
        .property_set(Binding::DocumentSummaryInformation)
        .unwrap();
    assert_eq!(
        document_summary::Snapshot::from_stream(&document_summary)
            .unwrap()
            .company(),
        Some("Litchi Project")
    );
    let custom = ole.get_metadata().unwrap().custom_properties;
    assert_eq!(custom.get("Pages reviewed"), Some(&Value::I4(12)));

    // The builder route keeps Word's root class identifier.
    let plain = fresh().build().unwrap().to_bytes().unwrap();
    assert_eq!(root_clsid(&bytes), root_clsid(&plain));
    assert_ne!(
        root_clsid(&plain),
        root_clsid(&{
            let mut writer = litchi_cfb::OleWriter::new();
            writer.create_stream(&["x"], b"x").unwrap();
            let mut output = Cursor::new(Vec::new());
            writer.write_to(&mut output).unwrap();
            output.into_inner()
        })
    );
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
fn unrepresentable_metadata_is_refused_without_change() {
    let mut builder = fresh();
    builder.set_metadata(metadata()).unwrap();
    let before = builder.summary_properties().clone();
    let error = builder
        .set_metadata(Metadata {
            title: Some("other".to_owned()),
            identifier: Some("urn:x".to_owned()),
            ..Metadata::default()
        })
        .unwrap_err();
    assert!(matches!(error, WriteError::Metadata(_)));
    assert_eq!(builder.summary_properties(), &before);
}

#[test]
fn cryptoapi_encryption_with_metadata_is_refused() {
    let mut builder = fresh();
    builder.set_metadata(metadata()).unwrap();
    builder
        .set_password("secret", EncryptionProfile::CryptoApiRc4 { key_bits: 128 })
        .unwrap();
    assert!(matches!(
        builder.build(),
        Err(WriteError::EncryptedMetadata)
    ));

    // Without metadata the same profile builds as before.
    let mut plain = fresh();
    plain
        .set_password("secret", EncryptionProfile::CryptoApiRc4 { key_bits: 128 })
        .unwrap();
    assert!(plain.build().is_ok());

    // XOR obfuscation has no encrypted form for the property streams; they
    // are written in plaintext and read back after decryption.
    let mut xor = fresh();
    xor.set_metadata(metadata()).unwrap();
    xor.set_password("secret", EncryptionProfile::WordXorObfuscation)
        .unwrap();
    let bytes = xor.build().unwrap().to_bytes().unwrap();
    let reopened = Package::from_bytes_with_options(
        &bytes,
        OpenOptions::default().password(Password::new("secret".to_owned())),
    )
    .unwrap();
    assert_eq!(
        reopened.metadata().title.as_deref(),
        Some("Quarterly report")
    );

    // Construction is O(1); the encrypted source is refused when the
    // deferred property snapshot is first requested.
    let mut edit = reopened
        .edit_metadata()
        .expect("metadata edit construction is deferred");
    assert!(matches!(edit.properties(), Err(Error::Encrypted)));
}

#[test]
fn metadata_transaction_creates_missing_streams_and_keeps_the_clsid() {
    let package = fresh().build().unwrap();
    let source = package.to_bytes().unwrap();
    assert!(package.metadata().title.is_none());

    // An exact no-op commits the same package.
    let commit = package.edit_metadata().unwrap().commit().unwrap();
    assert!(!commit.diagnostics().changed());
    assert!(commit.patch().is_noop());
    assert_eq!(commit.package().to_bytes().unwrap(), source);

    let mut edit = package.edit_metadata().unwrap();
    assert!(edit.merge_metadata(&metadata()).unwrap());
    edit.properties_mut()
        .expect("metadata properties")
        .set("Department", "Research")
        .unwrap();
    let commit = edit.commit().unwrap();
    assert!(commit.diagnostics().changed());
    let (edited, patch) = commit.into_parts();
    assert_metadata(&edited);
    let bytes = edited.to_bytes().unwrap();
    assert_eq!(root_clsid(&bytes), root_clsid(&source));

    let mut ole = OleFile::open(
        Cursor::new(bytes.clone()),
        litchi_cfb::OleFileLimits::default(),
    )
    .unwrap();
    let document_summary = document_summary::Snapshot::from_stream(
        &ole.property_set(Binding::DocumentSummaryInformation)
            .unwrap(),
    )
    .unwrap();
    assert_eq!(document_summary.codepage(), Some(CodePage::WINDOWS_1252));

    // Replay, conflict and revert.
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

    // A second edit removes a property and leaves the others.
    let mut edit = edited.edit_metadata().unwrap();
    assert!(
        edit.properties_mut()
            .expect("metadata properties")
            .remove("title")
            .unwrap()
    );
    let removed = edit.commit().unwrap().into_package();
    assert!(removed.metadata().title.is_none());
    assert_eq!(removed.metadata().author.as_deref(), Some("Ada Lovelace"));
}

#[test]
fn refused_metadata_edits_leave_the_edit_unchanged() {
    let package = fresh().build().unwrap();
    let mut edit = package.edit_metadata().unwrap();
    edit.merge_metadata(&Metadata {
        title: Some("kept".to_owned()),
        ..Metadata::default()
    })
    .unwrap();
    let staged = edit.properties().expect("metadata properties").clone();
    for refused in [
        Metadata {
            hyperlink_base: Some("https://example.invalid/".to_owned()),
            ..Metadata::default()
        },
        Metadata {
            revision: Some("seven".to_owned()),
            ..Metadata::default()
        },
    ] {
        assert!(matches!(
            edit.merge_metadata(&refused),
            Err(Error::Metadata(_))
        ));
        assert_eq!(edit.properties().expect("metadata properties"), &staged);
    }
    assert!(
        edit.properties_mut()
            .expect("metadata properties")
            .set("_PID_LINKBASE", "x")
            .is_err()
    );
}

#[test]
fn word_documents_keep_their_other_streams() {
    let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/ole/doc/FancyFoot.doc");
    let source = std::fs::read(path).unwrap();
    let package = Package::from_bytes(&source).unwrap();
    let mut edit = package.edit_metadata().unwrap();
    edit.merge_metadata(&Metadata {
        title: Some("Retitled".to_owned()),
        ..Metadata::default()
    })
    .unwrap();
    let edited = edit.commit().unwrap().into_package();
    assert_eq!(edited.metadata().title.as_deref(), Some("Retitled"));
    assert_eq!(edited.metadata().author, package.metadata().author);
    assert_eq!(
        edited.document().text().unwrap(),
        package.document().text().unwrap()
    );
    let bytes = edited.to_bytes().unwrap();
    assert_eq!(stream_names(&bytes).len(), stream_names(&source).len());
    assert_eq!(root_clsid(&bytes), root_clsid(&source));
}
