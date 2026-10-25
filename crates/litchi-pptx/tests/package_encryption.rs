#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]
#![cfg(feature = "encryption")]

mod in_place;
use in_place::InPlace as _;

use std::io::Cursor;

use litchi_opc::constants::relationship_type as rt;
use litchi_opc::{BlobPart, OpcPackage, PackURI, PackageWriter, Part, Relationships};
use litchi_pptx::encryption::{self, Kind, Mode};
use litchi_pptx::{EncryptionLimits, Error, Package, SourceLimits};

fn plain() -> Vec<u8> {
    Package::new().unwrap().to_bytes().unwrap()
}

fn encrypted(mode: Mode, password: &str) -> Vec<u8> {
    encryption::encrypt(plain(), password, mode, &encryption::Limits::default()).unwrap()
}

fn spreadsheet_package() -> Vec<u8> {
    let mut package = OpcPackage::new();
    package.relate_to(
        "xl/workbook.xml",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
    );
    package.add_part(Box::new(BlobPart::new(
        PackURI::new("/xl/workbook.xml").unwrap(),
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml".into(),
        br#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>"#
            .to_vec(),
    )));
    PackageWriter::to_bytes(&package).unwrap()
}

const SIGNATURE_REL: &str =
    "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/signature";
const ORIGIN_BYTES: &[u8] = b"litchi inert signature origin sentinel";
const SIGNATURE_BYTES: &[u8] = br#"<Signature xmlns="http://www.w3.org/2000/09/xmldsig#"><Object Id="litchi-sentinel"/></Signature>"#;

#[derive(Debug, PartialEq, Eq)]
struct SignatureGraph {
    package_relationships: Vec<(String, String, String, bool)>,
    origin_content_type: String,
    origin_bytes: Vec<u8>,
    origin_relationships: Vec<(String, String, String, bool)>,
    signature_content_type: String,
    signature_bytes: Vec<u8>,
    signature_relationships: Vec<(String, String, String, bool)>,
}

fn install_inert_signature(opc: &mut OpcPackage) {
    let mut origin = BlobPart::new(
        PackURI::new("/_xmlsignatures/origin.sigs").unwrap(),
        "application/vnd.openxmlformats-package.digital-signature-origin".into(),
        ORIGIN_BYTES.to_vec(),
    );
    origin.relate_to("sig1.xml", SIGNATURE_REL);
    let signature = BlobPart::new(
        PackURI::new("/_xmlsignatures/sig1.xml").unwrap(),
        "application/vnd.openxmlformats-package.digital-signature-xmlsignature+xml".into(),
        SIGNATURE_BYTES.to_vec(),
    );
    opc.relate_to("_xmlsignatures/origin.sigs", rt::DIGITAL_SIGNATURE_ORIGIN);
    opc.add_part(Box::new(origin));
    opc.add_part(Box::new(signature));
}

fn relationship_snapshot(
    relationships: &Relationships,
    filter: Option<&str>,
) -> Vec<(String, String, String, bool)> {
    let mut values = relationships
        .iter()
        .filter(|relationship| filter.is_none_or(|value| relationship.reltype() == value))
        .map(|relationship| {
            (
                relationship.r_id().to_owned(),
                relationship.reltype().to_owned(),
                relationship.target_ref().to_owned(),
                relationship.is_external(),
            )
        })
        .collect::<Vec<_>>();
    values.sort();
    values
}

fn signature_graph(opc: &OpcPackage) -> SignatureGraph {
    let origin = opc
        .get_part(&PackURI::new("/_xmlsignatures/origin.sigs").unwrap())
        .unwrap();
    let signature = opc
        .get_part(&PackURI::new("/_xmlsignatures/sig1.xml").unwrap())
        .unwrap();
    SignatureGraph {
        package_relationships: relationship_snapshot(
            opc.rels(),
            Some(rt::DIGITAL_SIGNATURE_ORIGIN),
        ),
        origin_content_type: origin.content_type().to_owned(),
        origin_bytes: origin.blob().to_vec(),
        origin_relationships: relationship_snapshot(origin.rels(), None),
        signature_content_type: signature.content_type().to_owned(),
        signature_bytes: signature.blob().to_vec(),
        signature_relationships: relationship_snapshot(signature.rels(), None),
    }
}

#[test]
fn password_aware_ingress_accepts_plain_zip_without_provenance() {
    let package =
        Package::from_reader_with_password(Cursor::new(plain()), "unused password").unwrap();
    assert_eq!(package.encryption(), None);
}

#[test]
fn encrypted_wrong_family_is_rejected_after_decryption() {
    let encrypted = encryption::encrypt(
        spreadsheet_package(),
        "secret",
        Mode::Agile,
        &encryption::Limits::default(),
    )
    .unwrap();
    assert!(matches!(
        Package::from_reader_with_password(Cursor::new(encrypted), "secret"),
        Err(Error::ContentType { .. })
    ));
}

#[test]
fn standard_and_agile_ingress_retain_mode_and_reencrypt() {
    for mode in [Mode::Standard, Mode::Agile] {
        let bytes = encrypted(mode, "old password");
        let package =
            Package::from_reader_with_password(Cursor::new(bytes), "old password").unwrap();
        assert_eq!(package.encryption(), Some(mode));
        assert!(matches!(
            package.to_bytes(),
            Err(Error::EncryptionPolicy { .. })
        ));

        let rekeyed = package.to_reencrypted("new password").unwrap();
        assert_eq!(
            encryption::inspect(&rekeyed, &encryption::Limits::default()).unwrap(),
            Kind::Encrypted(mode)
        );
        assert!(matches!(
            Package::from_reader_with_password(Cursor::new(rekeyed.clone()), "old password"),
            Err(Error::Encryption(encryption::Error::Password))
        ));
        Package::from_reader_with_password(Cursor::new(rekeyed), "new password").unwrap();
    }
}

#[test]
fn password_aware_ingress_reports_password_and_independent_limits() {
    let bytes = encrypted(Mode::Agile, "secret");
    assert!(matches!(
        Package::from_reader_with_password(Cursor::new(bytes.clone()), "wrong"),
        Err(Error::Encryption(encryption::Error::Password))
    ));

    let crypto_limits = EncryptionLimits {
        max_input_bytes: bytes.len() - 1,
        ..EncryptionLimits::default()
    };
    assert!(matches!(
        Package::from_reader_with_password_and_options(
            Cursor::new(bytes.clone()),
            "secret",
            litchi_pptx::OpenOptions::default()
                .encryption_limits(crypto_limits)
                .source_limits(SourceLimits::default())
        ),
        Err(Error::Encryption(encryption::Error::Limit {
            resource: "input",
            ..
        }))
    ));

    let opc_limits = SourceLimits::builder()
        .max_input_bytes(1)
        .unwrap()
        .build()
        .unwrap();
    assert!(matches!(
        Package::from_reader_with_password_and_options(
            Cursor::new(bytes),
            "secret",
            litchi_pptx::OpenOptions::default()
                .encryption_limits(EncryptionLimits::default())
                .source_limits(opc_limits)
        ),
        Err(Error::Opc(_))
    ));
}

#[test]
fn downgrade_and_mutation_are_explicit() {
    let mut package = Package::from_reader_with_password(
        Cursor::new(encrypted(Mode::Standard, "secret")),
        "secret",
    )
    .unwrap();

    assert!(matches!(
        package.put_fonts_in_place(Default::default()),
        Err(Error::EncryptionPolicy { .. })
    ));

    let clear = package.to_plain_bytes().unwrap();
    assert_eq!(package.encryption(), Some(Mode::Standard));
    assert_eq!(
        encryption::inspect(&clear, &encryption::Limits::default()).unwrap(),
        Kind::Plain
    );
    assert!(matches!(
        package.to_bytes(),
        Err(Error::EncryptionPolicy { .. })
    ));
}

#[test]
fn selected_mode_saves_leave_the_retained_mode_unchanged() {
    let package = Package::from_reader_with_password(
        Cursor::new(encrypted(Mode::Standard, "secret")),
        "secret",
    )
    .unwrap();
    let directory = tempfile::tempdir().unwrap();

    let agile_bytes = package.to_encrypted("byte password", Mode::Agile).unwrap();
    assert_eq!(package.encryption(), Some(Mode::Standard));
    assert_eq!(
        encryption::inspect(&agile_bytes, &encryption::Limits::default()).unwrap(),
        Kind::Encrypted(Mode::Agile)
    );

    assert!(
        package
            .save_encrypted(directory.path(), "new password", Mode::Agile)
            .is_err()
    );
    assert_eq!(package.encryption(), Some(Mode::Standard));

    let output = directory.path().join("rekeyed.pptx");
    package
        .save_encrypted(&output, "new password", Mode::Agile)
        .unwrap();
    // Output never changes the package: it keeps the mode it was opened with.
    assert_eq!(package.encryption(), Some(Mode::Standard));
    let saved = std::fs::read(output).unwrap();
    assert_eq!(
        encryption::inspect(&saved, &encryption::Limits::default()).unwrap(),
        Kind::Encrypted(Mode::Agile)
    );
    Package::from_reader_with_password(Cursor::new(saved), "new password").unwrap();
}

#[cfg(unix)]
#[test]
fn atomic_encrypted_save_preserves_permissions_and_refuses_symlinks() {
    use std::os::unix::fs::{PermissionsExt, symlink};

    let package = Package::new().unwrap();
    let directory = tempfile::tempdir().unwrap();
    let output = directory.path().join("encrypted.pptx");
    std::fs::write(&output, b"old destination").unwrap();
    std::fs::set_permissions(&output, std::fs::Permissions::from_mode(0o640)).unwrap();

    package
        .save_encrypted(&output, "password", Mode::Agile)
        .unwrap();
    assert_eq!(package.encryption(), None);
    assert_eq!(
        std::fs::metadata(&output).unwrap().permissions().mode() & 0o777,
        0o640
    );

    let target = directory.path().join("target.pptx");
    let link = directory.path().join("linked.pptx");
    std::fs::write(&target, b"original target").unwrap();
    symlink(&target, &link).unwrap();
    assert!(matches!(
        package.save_encrypted(&link, "password", Mode::Standard),
        Err(Error::Opc(_))
    ));
    assert_eq!(package.encryption(), None);
    assert_eq!(std::fs::read(&target).unwrap(), b"original target");
    assert!(
        std::fs::symlink_metadata(&link)
            .unwrap()
            .file_type()
            .is_symlink()
    );
}

#[test]
fn retained_mode_is_required_and_raw_edit_is_explicit_declassification() {
    let plain = Package::new().unwrap();
    assert!(matches!(
        plain.to_reencrypted("password"),
        Err(Error::EncryptionPolicy { .. })
    ));

    let mut retained =
        Package::from_reader_with_password(Cursor::new(encrypted(Mode::Agile, "secret")), "secret")
            .unwrap();
    retained.edit_opc_in_place(|_| Ok(())).unwrap();
    assert_eq!(retained.encryption(), None);
}

#[test]
fn unchanged_inner_signature_survives_encrypt_open_and_reencrypt() {
    let mut signed = Package::new().unwrap();
    signed
        .edit_opc_in_place(|opc| {
            install_inert_signature(opc);
            Ok(())
        })
        .unwrap();
    assert!(signed.opc().unwrap().is_signed());
    let expected = signature_graph(signed.opc().unwrap());

    let encrypted = signed.to_encrypted("first", Mode::Agile).unwrap();
    assert_eq!(signature_graph(signed.opc().unwrap()), expected);
    let opened = Package::from_reader_with_password(Cursor::new(encrypted), "first").unwrap();
    assert_eq!(signature_graph(opened.opc().unwrap()), expected);

    let reencrypted = opened.to_reencrypted("second").unwrap();
    assert_eq!(signature_graph(opened.opc().unwrap()), expected);
    let reopened = Package::from_reader_with_password(Cursor::new(reencrypted), "second").unwrap();
    assert_eq!(signature_graph(reopened.opc().unwrap()), expected);
}

#[test]
fn raw_opc_noop_preserves_signature_and_changed_edit_requires_explicit_unsign() {
    let mut package = Package::new().unwrap();
    package
        .edit_opc_in_place(|opc| {
            install_inert_signature(opc);
            Ok(())
        })
        .unwrap();
    let signed = package.to_bytes().unwrap();

    package.edit_opc_in_place(|_| Ok(())).unwrap();
    assert!(package.opc().unwrap().is_signed());
    assert_eq!(package.to_bytes().unwrap(), signed);

    let add_raw_part = |opc: &mut OpcPackage| {
        opc.add_part(Box::new(BlobPart::new(
            PackURI::new("/custom/raw.bin").unwrap(),
            "application/octet-stream".into(),
            vec![1, 2, 3],
        )));
    };
    // A changed raw edit refuses the signed package and rolls back (ADR
    // 0006); removing the signature inside the closure authorizes it.
    let error = package
        .edit_opc_in_place(|opc| {
            add_raw_part(opc);
            Ok(())
        })
        .unwrap_err();
    assert!(
        matches!(
            error,
            Error::Opc(litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
        ),
        "{error:?}"
    );
    assert!(package.opc().unwrap().is_signed());
    assert_eq!(package.to_bytes().unwrap(), signed);

    package
        .edit_opc_in_place(|opc| {
            add_raw_part(opc);
            opc.unsign();
            Ok(())
        })
        .unwrap();
    assert!(!package.opc().unwrap().is_signed());
    assert!(
        package
            .opc()
            .unwrap()
            .contains_part(&PackURI::new("/custom/raw.bin").unwrap())
    );
}

#[test]
fn immutable_reencryption_preserves_profile_content_and_protection() {
    for mode in [Mode::Standard, Mode::Agile] {
        let input = encrypted(mode, "old");
        let package = Package::from_reader_with_password(Cursor::new(input), "old").unwrap();
        let mut output = Vec::new();
        package.write_reencrypted(&mut output, "new").unwrap();
        let clear = encryption::open(output, "new", &encryption::Limits::default()).unwrap();
        assert_eq!(clear.mode(), Some(mode));
        assert_eq!(clear.into_bytes(), plain());
        assert_eq!(package.encryption(), Some(mode));
        assert!(matches!(
            package.write_to(Vec::new()),
            Err(Error::EncryptionPolicy { .. })
        ));
    }
}

#[test]
fn immutable_reencryption_refuses_before_emitting_ciphertext_on_staging_failure() {
    // The encryption limits are the ones the package was opened with; a
    // metadata edit (which keeps the provenance) grows the clear package
    // past them.
    let input = encrypted(Mode::Standard, "old");
    let clear_len = Package::from_reader_with_password(Cursor::new(input.clone()), "old")
        .unwrap()
        .to_plain_bytes()
        .unwrap()
        .len();
    let limits = EncryptionLimits {
        max_plaintext_bytes: clear_len + 256,
        ..EncryptionLimits::default()
    };
    let opened = Package::from_reader_with_password_and_options(
        Cursor::new(input),
        "old",
        litchi_pptx::OpenOptions::default().encryption_limits(limits),
    )
    .unwrap();
    let mut edit = opened.edit_metadata().unwrap();
    // An incompressible title, so the clear ZIP grows past the limit.
    let mut state = 0x2545_f491_u32;
    let title: String = (0..8192)
        .map(|_| {
            state = state.wrapping_mul(1_103_515_245).wrapping_add(12_345);
            char::from(b'a' + u8::try_from((state >> 16) % 26).unwrap())
        })
        .collect();
    edit.update(|metadata| metadata.title = Some(title))
        .unwrap();
    let package = edit.commit().unwrap().into_package();
    let mut sink = Vec::new();
    let error = package.write_reencrypted(&mut sink, "new").unwrap_err();
    assert!(matches!(
        error,
        Error::Encryption(encryption::Error::Limit {
            resource: "plaintext",
            ..
        })
    ));
    assert!(sink.is_empty());
    assert!(package.write_reencrypted(&mut sink, "").is_err());
    assert!(sink.is_empty());
    assert_eq!(package.encryption(), Some(Mode::Standard));
    let authored = Package::new().unwrap();
    assert!(matches!(
        authored.write_reencrypted(&mut sink, "new"),
        Err(Error::EncryptionPolicy { .. })
    ));
    let clear = Package::from_reader(Cursor::new(plain())).unwrap();
    assert!(matches!(
        clear.write_reencrypted(&mut sink, "new"),
        Err(Error::EncryptionPolicy { .. })
    ));
    assert!(sink.is_empty());
}

#[test]
fn immutable_reencryption_keeps_sink_errors_and_owner_protection() {
    struct FailedSink;
    impl std::io::Write for FailedSink {
        fn write(&mut self, _: &[u8]) -> std::io::Result<usize> {
            Err(std::io::Error::from(std::io::ErrorKind::BrokenPipe))
        }
        fn flush(&mut self) -> std::io::Result<()> {
            Ok(())
        }
    }
    struct FlushFailure;
    impl std::io::Write for FlushFailure {
        fn write(&mut self, bytes: &[u8]) -> std::io::Result<usize> {
            Ok(bytes.len())
        }
        fn flush(&mut self) -> std::io::Result<()> {
            Err(std::io::Error::from(std::io::ErrorKind::BrokenPipe))
        }
    }
    let package =
        Package::from_reader_with_password(Cursor::new(encrypted(Mode::Standard, "old")), "old")
            .unwrap();
    let error = package.write_reencrypted(FailedSink, "new").unwrap_err();
    assert!(
        matches!(error, Error::Io(ref source) if source.kind() == std::io::ErrorKind::BrokenPipe)
    );
    let error = package.write_reencrypted(FlushFailure, "new").unwrap_err();
    assert!(
        matches!(error, Error::Io(ref source) if source.kind() == std::io::ErrorKind::BrokenPipe)
    );
    assert_eq!(package.encryption(), Some(Mode::Standard));
    assert!(matches!(
        package.write_to(Vec::new()),
        Err(Error::EncryptionPolicy { .. })
    ));
}
