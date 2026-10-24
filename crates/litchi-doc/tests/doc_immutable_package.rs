#![allow(clippy::unwrap_used, clippy::expect_used, reason = "test assertions")]

use litchi_doc::{Builder, Document, Error, OpenOptions, Package, SourceLimitKind, SourceLimits};
use std::io::Cursor;

fn corpus() -> Vec<u8> {
    let mut builder = Builder::new();
    builder.add_paragraph("retained projection").unwrap();
    builder.build().unwrap().to_bytes().unwrap()
}

#[test]
fn detached_document_is_a_cheap_shared_projection() {
    fn shared<T: Clone + Send + Sync>() {}
    shared::<Package>();
    shared::<Document>();
    assert_eq!(size_of::<Document>(), size_of::<usize>());
    let bytes = corpus();
    let package = Package::from_bytes(&bytes).unwrap();
    let document = package.document();
    let clone = document.clone();
    // Body text borrows the same immutable parse after both owners are gone.
    let original = document
        .body_paragraph_text(litchi_core::Position::new(0))
        .unwrap();
    let cloned = clone
        .body_paragraph_text(litchi_core::Position::new(0))
        .unwrap();
    assert_eq!(original, cloned);
    assert!(std::ptr::eq(
        original.unwrap().as_ptr(),
        cloned.unwrap().as_ptr()
    ));
    drop(package);
    drop(bytes);
    assert!(clone.text().unwrap().contains("retained projection"));
}

#[test]
fn canonical_ingress_and_publication_preserve_exact_bytes() {
    let bytes = corpus();
    let package = Package::from_reader(Cursor::new(&bytes)).unwrap();
    let sibling = package.clone();
    assert_eq!(package.to_bytes().unwrap(), bytes);
    let mut output = Vec::new();
    sibling.write_to(&mut output).unwrap();
    assert_eq!(output, bytes);
    assert_eq!(
        package.document().text().unwrap(),
        sibling.document().text().unwrap()
    );
}

#[test]
fn merged_options_enforce_package_and_stream_limits_at_ingress() {
    let bytes = corpus();
    let package_limit = SourceLimits::builder()
        .max_input_bytes(1)
        .unwrap()
        .build()
        .unwrap();
    for error in [
        Package::from_bytes_with_options(
            &bytes,
            OpenOptions::default().source_limits(package_limit),
        )
        .unwrap_err(),
        Package::from_reader_with_options(
            Cursor::new(&bytes),
            OpenOptions::default().source_limits(package_limit),
        )
        .unwrap_err(),
    ] {
        assert!(matches!(
            error,
            Error::Limit {
                kind: SourceLimitKind::InputBytes,
                ..
            }
        ));
    }
    let stream_limit = SourceLimits::builder()
        .max_entry_bytes(1)
        .unwrap()
        .build()
        .unwrap();
    let error = Package::from_bytes_with_options(
        &bytes,
        OpenOptions::default().source_limits(stream_limit),
    )
    .unwrap_err();
    assert!(matches!(
        error,
        Error::Limit {
            kind: SourceLimitKind::EntryBytes,
            ..
        }
    ));
}

#[test]
fn encrypted_builder_and_ingress_preserve_ciphertext() {
    let mut builder = Builder::new();
    builder.add_paragraph("encrypted projection").unwrap();
    builder
        .set_password(
            "secret",
            litchi_doc::writer::EncryptionProfile::OfficeBinaryRc4,
        )
        .unwrap();
    let authored = builder.build().unwrap();
    assert!(
        authored
            .document()
            .text()
            .unwrap()
            .contains("encrypted projection")
    );
    let ciphertext = authored.to_bytes().unwrap();
    assert!(matches!(
        Package::from_bytes(&ciphertext),
        Err(Error::PasswordRequired)
    ));
    assert!(matches!(
        Package::from_bytes_with_options(
            &ciphertext,
            OpenOptions::default().password("wrong".to_owned().into())
        ),
        Err(Error::InvalidPassword)
    ));
    let opened = Package::from_bytes_with_options(
        &ciphertext,
        OpenOptions::default().password("secret".to_owned().into()),
    )
    .unwrap();
    assert!(
        opened
            .document()
            .text()
            .unwrap()
            .contains("encrypted projection")
    );
    assert_eq!(opened.to_bytes().unwrap(), ciphertext);
}

#[test]
fn path_open_seals_an_owned_snapshot_and_checks_input_limit() {
    let path = std::env::temp_dir().join(format!("litchi-doc-owner-{}.doc", std::process::id()));
    let bytes = corpus();
    let mut file = std::fs::OpenOptions::new()
        .write(true)
        .create_new(true)
        .open(&path)
        .unwrap();
    std::io::Write::write_all(&mut file, &bytes).unwrap();
    drop(file);
    let result = (|| -> Result<(), Box<dyn std::error::Error>> {
        let limits = SourceLimits::builder()
            .max_input_bytes(bytes.len() - 1)?
            .build()?;
        assert!(matches!(
            Package::open_with_options(&path, OpenOptions::default().source_limits(limits)),
            Err(Error::Limit {
                kind: SourceLimitKind::InputBytes,
                ..
            })
        ));
        let package = Package::open(&path)?;
        let version = package.source_version()?;
        let sibling = package.clone();
        std::fs::write(&path, b"changed after open")?;
        assert_eq!(sibling.source_version()?, version);
        assert_eq!(package.to_bytes()?, bytes);
        Ok(())
    })();
    std::fs::remove_file(path).unwrap();
    result.unwrap();
}

#[test]
fn lenient_read_owner_does_not_bypass_strict_transaction_validation() {
    let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/ole/doc/duplicate-style-names.doc");
    let bytes = std::fs::read(path).unwrap();
    let package = Package::from_bytes_with_options(
        &bytes,
        OpenOptions::default().leniency(litchi_doc::Leniency::TolerateStylesheetDefects),
    )
    .unwrap();
    assert_eq!(
        package.leniency(),
        litchi_doc::Leniency::TolerateStylesheetDefects
    );
    assert!(!package.document().tolerance_report().defects().is_empty());
    let dropped = package
        .edit()
        .expect("malformed body is not read during edit construction");
    assert_eq!(dropped.retained_render_bytes(), None);
    drop(dropped);
    let mut edit = package
        .edit()
        .expect("edit construction is deferred until the first body operation");
    assert!(
        edit.changes().is_err(),
        "the body engine still requires strict validation"
    );
    assert_eq!(package.to_bytes().unwrap(), bytes);
}
