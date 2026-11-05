#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "integration tests use panic-on-failure extraction"
)]

//! A changed opened-presentation commit refuses a signed package (ADR 0006;
//! unified-ops record 0013). Removing the signature is explicit
//! (`Package::unsign`), and an exact no-op keeps the source bytes, signature
//! included.

use std::path::PathBuf;

use litchi_opc::{OpcError, OpcPackage};
use litchi_pptx::{Error, Package};

/// The producer-signed fixtures: one signature by `xmlsec`, one by Office 2010.
const SIGNED: [&str; 2] = ["hello-world-signed.pptx", "ms-office-2010-signed.pptx"];

fn fixture(name: &str) -> Vec<u8> {
    std::fs::read(
        PathBuf::from(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/poi/test-data/xmldsign")
            .join(name),
    )
    .expect("signed PPTX fixture")
}

/// Whether `bytes` still carries any OPC signature infrastructure.
fn carries_signature(bytes: &[u8]) -> bool {
    let package = OpcPackage::from_bytes(bytes, litchi_opc::SourceLimits::default()).unwrap();
    package.is_signed()
        || package
            .iter_parts()
            .any(|part| part.partname().as_str().starts_with("/_xmlsignatures/"))
}

#[test]
fn changed_commit_on_a_signed_package_is_refused() {
    for name in SIGNED {
        let bytes = fixture(name);
        let package = Package::from_reader(std::io::Cursor::new(bytes.clone())).unwrap();
        assert!(package.opc().unwrap().is_signed(), "{name}");
        let mut transaction = package.edit().unwrap();
        transaction
            .add_text_box(0_usize, "edited", (0, 0, 914_400, 914_400))
            .unwrap();
        let error = transaction.commit().unwrap_err();
        assert!(
            matches!(
                error,
                Error::Opc(OpcError::SignedSourceRequiresExplicitPolicy)
            ),
            "{name}: {error:?}"
        );
        // Nothing was staged into the package: it still publishes its source.
        assert_eq!(package.to_bytes().unwrap(), bytes, "{name}");
    }
}

#[test]
fn exact_noop_commit_keeps_the_signature_bytes() {
    for name in SIGNED {
        let bytes = fixture(name);
        let mut package = Package::from_reader(std::io::Cursor::new(bytes.clone())).unwrap();
        let commit = package.edit().unwrap().commit().unwrap();
        assert!(!commit.diagnostics().changed(), "{name}");
        package = commit.into_package();
        assert_eq!(package.to_bytes().unwrap(), bytes, "{name}");
    }
}

#[test]
fn explicit_unsign_authorizes_a_changed_commit() {
    for name in SIGNED {
        let mut package = Package::from_reader(std::io::Cursor::new(fixture(name))).unwrap();
        package = package.unsign().unwrap();
        assert!(!package.opc().unwrap().is_signed(), "{name}");
        let mut transaction = package.edit().unwrap();
        transaction
            .add_text_box(0_usize, "edited after unsign", (0, 0, 914_400, 914_400))
            .unwrap();
        let commit = transaction.commit().unwrap();
        assert!(commit.diagnostics().changed(), "{name}");
        package = commit.into_package();
        let output = package.to_bytes().unwrap();
        assert!(!carries_signature(&output), "{name}");
        let reopened = Package::from_reader(std::io::Cursor::new(output)).unwrap();
        let text = reopened.presentation().unwrap().slides().unwrap()[0]
            .text()
            .unwrap();
        assert!(text.contains("edited after unsign"), "{name}: {text}");
    }
}

#[test]
fn unsigning_an_unsigned_package_changes_nothing() {
    let mut builder = litchi_pptx::builder::Builder::new().unwrap();
    builder.presentation().unwrap().add_slide().unwrap();
    let authored = builder.build().unwrap();
    let bytes = authored.to_bytes().unwrap();
    let mut package = Package::from_reader(std::io::Cursor::new(bytes.clone())).unwrap();
    package = package.unsign().unwrap();
    assert!(!package.opc().unwrap().is_signed());
    let commit = package.edit().unwrap().commit().unwrap();
    assert!(!commit.diagnostics().changed());
    assert_eq!(package.to_bytes().unwrap(), bytes);
}

#[test]
fn insertion_refuses_signed_sources_before_layout_resolution() {
    for name in SIGNED {
        let source = Package::from_reader(std::io::Cursor::new(fixture(name))).unwrap();
        let before = source.to_bytes().unwrap();
        let mut edit = source.edit().unwrap();
        assert!(matches!(
            edit.insert_slide(litchi_core::Position::new(0), "unused layout selector"),
            Err(Error::Opc(OpcError::SignedSourceRequiresExplicitPolicy))
        ));
        assert!(edit.commit().unwrap().patch().is_noop());
        assert_eq!(source.to_bytes().unwrap(), before);
    }
}
