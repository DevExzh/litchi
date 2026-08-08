#![cfg(feature = "encryption")]

use std::io::Cursor;

use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{BlobPart, OpcPackage, PackURI, PackageWriter, Part, Relationship, TargetMode};
use litchi_xlsx::encryption::{Error as CryptoError, Limits, Mode};
use litchi_xlsx::{Error, Package, ReadLimits, Workbook};

const PASSWORD: &str = "Litchi XLSX password 42!";
const NEW_PASSWORD: &str = "Litchi XLSX changed password 7!";

#[test]
fn package_and_workbook_round_trip_managed_encryption() {
    for mode in [Mode::Standard, Mode::Agile] {
        let package = Package::create().unwrap();
        let encrypted = package.to_encrypted(PASSWORD, mode).unwrap();
        assert_eq!(package.encryption(), None);

        let package = Package::from_reader_with_password(Cursor::new(encrypted), PASSWORD).unwrap();
        assert_eq!(package.encryption(), Some(mode));
        let workbook = package.into_workbook().unwrap();
        assert_eq!(workbook.encryption(), Some(mode));
        assert_eq!(workbook.len(), 1);
    }
}

#[test]
fn encrypted_provenance_blocks_implicit_plaintext_and_mutation() {
    let source = Package::create().unwrap();
    let encrypted = source.to_encrypted(PASSWORD, Mode::Standard).unwrap();
    let mut package = Package::from_bytes_with_password(encrypted, PASSWORD).unwrap();

    assert!(matches!(
        package.to_bytes(),
        Err(Error::EncryptionPolicy {
            operation: "to_bytes",
            ..
        })
    ));
    assert!(matches!(
        package.edit_slicers(),
        Err(Error::EncryptionPolicy {
            operation: "edit_slicers",
            ..
        })
    ));
    assert_eq!(&package.to_plain_bytes().unwrap()[..2], b"PK");

    let workbook = package.workbook().unwrap();
    assert!(matches!(
        workbook.edit(),
        Err(Error::EncryptionPolicy {
            operation: "edit",
            ..
        })
    ));
    assert_eq!(&workbook.to_plain_bytes().unwrap()[..2], b"PK");
}

#[test]
fn reencrypted_output_retains_mode_and_atomic_failure_retains_state() {
    let directory = tempfile::tempdir().unwrap();
    let target = directory.path().join("reencrypted.xlsx");
    let package = Package::create().unwrap();
    let encrypted = package.to_encrypted(PASSWORD, Mode::Standard).unwrap();
    let mut package = Package::from_bytes_with_password(encrypted, PASSWORD).unwrap();

    package.save_reencrypted(&target, NEW_PASSWORD).unwrap();
    assert!(matches!(
        Package::open_with_password(&target, PASSWORD),
        Err(Error::Encryption(CryptoError::Password))
    ));
    assert_eq!(
        Package::open_with_password(&target, NEW_PASSWORD)
            .unwrap()
            .encryption(),
        Some(Mode::Standard)
    );

    let before = package.encryption();
    let existing = directory.path().join("existing.xlsx");
    let sentinel = b"existing destination must survive";
    std::fs::write(&existing, sentinel).unwrap();
    assert!(matches!(
        package.save_encrypted(&existing, "", Mode::Agile),
        Err(Error::Encryption(CryptoError::PasswordRequired))
    ));
    assert_eq!(std::fs::read(&existing).unwrap(), sentinel);
    assert_eq!(package.encryption(), before);

    assert!(
        package
            .save_encrypted(directory.path(), PASSWORD, Mode::Agile)
            .is_err()
    );
    assert_eq!(package.encryption(), before);

    let selected = directory.path().join("selected-mode.xlsx");
    package
        .save_encrypted(&selected, NEW_PASSWORD, Mode::Agile)
        .unwrap();
    assert_eq!(package.encryption(), Some(Mode::Agile));
}

#[test]
fn outer_crypto_limit_is_independent_and_typed() {
    let limits = Limits {
        max_input_bytes: 4,
        ..Limits::default()
    };
    let error = Package::from_reader_with_password_and_limits(
        Cursor::new(vec![0u8; 5]),
        PASSWORD,
        &limits,
        ReadLimits::default(),
    )
    .unwrap_err();
    assert!(matches!(
        error,
        Error::Encryption(CryptoError::Limit {
            resource: "input",
            actual: 5,
            maximum: 4,
        })
    ));
}

#[test]
fn byte_generation_does_not_change_source_provenance() {
    let workbook = Workbook::new().unwrap();
    assert!(matches!(
        workbook.to_reencrypted(PASSWORD),
        Err(Error::EncryptionPolicy {
            operation: "to_reencrypted",
            ..
        })
    ));
    let encrypted = workbook.to_encrypted(PASSWORD, Mode::Standard).unwrap();
    let clone = workbook.clone();
    assert_eq!(clone.encryption(), None);
    assert_eq!(
        Workbook::from_slice_with_password(&encrypted, PASSWORD)
            .unwrap()
            .encryption(),
        Some(Mode::Standard)
    );
}

#[test]
fn password_aware_ingress_accepts_plain_zip_without_provenance() {
    let plaintext = Package::create().unwrap().to_plain_bytes().unwrap();
    let package = Package::from_bytes_with_password(plaintext, "irrelevant password").unwrap();

    assert_eq!(package.encryption(), None);
}

#[test]
fn restrictive_inner_opc_limits_reject_valid_encrypted_xlsx() {
    let plaintext = Package::create().unwrap().to_plain_bytes().unwrap();
    let encrypted = litchi_xlsx::encryption::encrypt(plaintext, PASSWORD, Mode::Standard).unwrap();
    let opc_limits = ReadLimits::builder()
        .max_input_bytes(4)
        .unwrap()
        .build()
        .unwrap();

    assert!(matches!(
        Package::from_bytes_with_password_and_limits(
            encrypted,
            PASSWORD,
            &Limits::default(),
            opc_limits,
        ),
        Err(Error::Package(_))
    ));
}

#[test]
fn decrypted_wrong_family_package_is_rejected() {
    let mut presentation = OpcPackage::new();
    let main_uri = PackURI::new("/ppt/presentation.xml").unwrap();
    presentation
        .try_add_part(Box::new(BlobPart::new(
            main_uri,
            ct::PML_PRESENTATION_MAIN.to_owned(),
            br#"<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>"#
                .to_vec(),
        )))
        .unwrap();
    presentation
        .rels_mut()
        .try_add_relationship(
            rt::OFFICE_DOCUMENT.to_owned(),
            "ppt/presentation.xml".to_owned(),
            "rId1".to_owned(),
            TargetMode::Internal,
        )
        .unwrap();
    let plaintext = PackageWriter::to_bytes(&presentation).unwrap();
    let encrypted = litchi_xlsx::encryption::encrypt(plaintext, PASSWORD, Mode::Standard).unwrap();

    assert!(matches!(
        Package::from_bytes_with_password(encrypted, PASSWORD),
        Err(Error::Invalid(_))
    ));
}

const SIGNATURE_REL: &str =
    "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/signature";
const ORIGIN_BYTES: &[u8] = b"Litchi inert signature origin sentinel";
const SIGNATURE_BYTES: &[u8] = br#"<Signature xmlns="http://www.w3.org/2000/09/xmldsig#"><Object>Litchi inert signature XML sentinel</Object></Signature>"#;

#[derive(Debug, PartialEq, Eq)]
struct RelationshipSnapshot {
    id: String,
    kind: String,
    target: String,
    mode: TargetMode,
}

impl From<&Relationship> for RelationshipSnapshot {
    fn from(relationship: &Relationship) -> Self {
        Self {
            id: relationship.r_id().to_owned(),
            kind: relationship.reltype().to_owned(),
            target: relationship.target_ref().to_owned(),
            mode: relationship.target_mode(),
        }
    }
}

#[derive(Debug, PartialEq, Eq)]
struct SignatureSnapshot {
    root_relationship: RelationshipSnapshot,
    origin_content_type: String,
    origin_bytes: Vec<u8>,
    origin_relationship: RelationshipSnapshot,
    signature_content_type: String,
    signature_bytes: Vec<u8>,
}

fn signature_snapshot(package: &Package) -> SignatureSnapshot {
    let raw = package.clone().into_plain_opc();
    let origin_uri = PackURI::new("/_xmlsignatures/origin.sigs").unwrap();
    let signature_uri = PackURI::new("/_xmlsignatures/sig1.xml").unwrap();
    let root_relationship = raw
        .rels()
        .iter()
        .find(|relationship| relationship.reltype() == rt::DIGITAL_SIGNATURE_ORIGIN)
        .unwrap();
    let origin = raw.get_part(&origin_uri).unwrap();
    let origin_relationship = origin
        .rels()
        .iter()
        .find(|relationship| relationship.reltype() == SIGNATURE_REL)
        .unwrap();
    let signature = raw.get_part(&signature_uri).unwrap();

    SignatureSnapshot {
        root_relationship: root_relationship.into(),
        origin_content_type: origin.content_type().to_owned(),
        origin_bytes: origin.blob().to_vec(),
        origin_relationship: origin_relationship.into(),
        signature_content_type: signature.content_type().to_owned(),
        signature_bytes: signature.blob().to_vec(),
    }
}

#[test]
fn signed_inner_opc_survives_encryption_and_reencryption_unchanged() {
    let mut raw = Package::create().unwrap().into_plain_opc();
    let origin_uri = PackURI::new("/_xmlsignatures/origin.sigs").unwrap();
    let signature_uri = PackURI::new("/_xmlsignatures/sig1.xml").unwrap();
    let mut origin = BlobPart::new(
        origin_uri,
        ct::OPC_DIGITAL_SIGNATURE_ORIGIN.to_owned(),
        ORIGIN_BYTES.to_vec(),
    );
    origin
        .rels_mut()
        .try_add_relationship(
            SIGNATURE_REL.to_owned(),
            "sig1.xml".to_owned(),
            "rIdSignatureXml".to_owned(),
            TargetMode::Internal,
        )
        .unwrap();
    raw.try_add_part(Box::new(origin)).unwrap();
    raw.try_add_part(Box::new(BlobPart::new(
        signature_uri,
        ct::OPC_DIGITAL_SIGNATURE_XMLSIGNATURE.to_owned(),
        SIGNATURE_BYTES.to_vec(),
    )))
    .unwrap();
    raw.rels_mut()
        .try_add_relationship(
            rt::DIGITAL_SIGNATURE_ORIGIN.to_owned(),
            "_xmlsignatures/origin.sigs".to_owned(),
            "rIdSignatureOrigin".to_owned(),
            TargetMode::Internal,
        )
        .unwrap();

    let signed = Package::from_opc(raw).unwrap();
    let expected = signature_snapshot(&signed);
    assert_eq!(expected.origin_bytes, ORIGIN_BYTES);
    assert_eq!(expected.signature_bytes, SIGNATURE_BYTES);
    let clear = signed.to_plain_bytes().unwrap();

    let encrypted = signed.to_encrypted(PASSWORD, Mode::Standard).unwrap();
    let opened = Package::from_bytes_with_password(encrypted, PASSWORD).unwrap();
    assert_eq!(signature_snapshot(&opened), expected);
    assert_eq!(opened.to_plain_bytes().unwrap(), clear);

    let reencrypted = opened.to_reencrypted(NEW_PASSWORD).unwrap();
    let reopened = Package::from_bytes_with_password(reencrypted, NEW_PASSWORD).unwrap();
    assert_eq!(signature_snapshot(&reopened), expected);
    assert_eq!(reopened.to_plain_bytes().unwrap(), clear);
}

#[test]
fn encrypted_source_raw_declassification_is_explicit() {
    let plaintext = Package::create().unwrap().to_plain_bytes().unwrap();
    let encrypted = litchi_xlsx::encryption::encrypt(plaintext, PASSWORD, Mode::Standard).unwrap();
    let package = Package::from_bytes_with_password(encrypted, PASSWORD).unwrap();
    assert_eq!(package.encryption(), Some(Mode::Standard));

    let raw = package.into_plain_opc();
    let declassified = PackageWriter::to_bytes(&raw).unwrap();
    assert_eq!(&declassified[..2], b"PK");
}
