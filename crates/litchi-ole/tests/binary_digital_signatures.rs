use litchi_cfb::{OleFile, OleWriter};
use litchi_ole::signature::{
    BinaryOfficeFormat, BinaryOfficeSignatureEditor, BinaryOfficeSignatureError, PackageSigner,
    SignatureVerificationPolicy, VerificationStatus, verify_binary_office_signatures,
};
use p256::ecdsa::SigningKey as EcdsaSigningKey;
use rsa::pkcs8::EncodePublicKey;
use rsa::{RsaPrivateKey, rand_core::OsRng};
use std::fs::File;
use std::io::Cursor;
use std::path::PathBuf;

fn ole(streams: &[(&[&str], &[u8])]) -> Vec<u8> {
    let mut writer = OleWriter::new();
    for (path, data) in streams {
        if path.len() > 1 {
            writer.create_storage(&path[..path.len() - 1]).unwrap();
        }
        writer.create_stream(path, data).unwrap();
    }
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    output.into_inner()
}

fn rsa_signer() -> PackageSigner {
    let key = RsaPrivateKey::new(&mut OsRng, 2048).unwrap();
    let public_key = key.to_public_key().to_public_key_der().unwrap();
    let certificate = fake_certificate(public_key.as_bytes());
    let mut signer = PackageSigner::rsa_sha256(key).unwrap();
    signer
        .set_certificates(vec![certificate, vec![0x30, 0x00]])
        .unwrap()
        .set_signing_time(Some("2026-07-19T12:34:56Z"))
        .unwrap();
    signer
}

fn fake_certificate(spki: &[u8]) -> Vec<u8> {
    let empty_sequence = [0x30, 0x00];
    let serial = [0x02, 0x01, 0x01];
    let mut tbs = Vec::new();
    tbs.extend_from_slice(&serial);
    for _ in 0..4 {
        tbs.extend_from_slice(&empty_sequence);
    }
    tbs.extend_from_slice(spki);
    let tbs = der(0x30, &tbs);
    let mut certificate = tbs;
    certificate.extend_from_slice(&empty_sequence);
    certificate.extend_from_slice(&[0x03, 0x01, 0x00]);
    der(0x30, &certificate)
}

fn der(tag: u8, body: &[u8]) -> Vec<u8> {
    let mut output = vec![tag];
    if body.len() < 128 {
        output.push(body.len() as u8);
    } else {
        let bytes = (body.len() as u32).to_be_bytes();
        let first = bytes.iter().position(|byte| *byte != 0).unwrap();
        output.push(0x80 | (bytes.len() - first) as u8);
        output.extend_from_slice(&bytes[first..]);
    }
    output.extend_from_slice(body);
    output
}

fn verify(bytes: &[u8], format: BinaryOfficeFormat) -> Result<Vec<litchi_ole::signature::BinaryOfficeSignatureVerification>, BinaryOfficeSignatureError> {
    let mut file = OleFile::open(Cursor::new(bytes)).unwrap();
    verify_binary_office_signatures(&mut file, format, &SignatureVerificationPolicy::strict())
}

fn signature_xml(bytes: &[u8]) -> (String, Vec<u8>) {
    let mut file = OleFile::open(Cursor::new(bytes)).unwrap();
    let name = file
        .list_directory_entries(&["_xmlsignatures"])
        .unwrap()[0]
        .name
        .clone();
    let xml = file.open_stream(&["_xmlsignatures", &name]).unwrap();
    (name, xml)
}

#[test]
fn rsa_and_ecdsa_signatures_round_trip_with_chain_time_and_multiple_streams() {
    let original = ole(&[(&["Payload"], b"signed bytes")]);
    let mut editor = BinaryOfficeSignatureEditor::new(original, BinaryOfficeFormat::Doc).unwrap();
    editor.add_signature(&rsa_signer()).unwrap();
    let ecdsa = PackageSigner::ecdsa_p256_sha256(EcdsaSigningKey::random(&mut OsRng));
    editor.add_signature(&ecdsa).unwrap();
    let signed = editor.finish().unwrap();

    let reports = verify(&signed, BinaryOfficeFormat::Doc).unwrap();
    assert_eq!(reports.len(), 2);
    assert!(reports.iter().all(|report| {
        report.package_integrity == VerificationStatus::Valid
            && report.signature_value == VerificationStatus::Valid
    }));
    let rsa = reports.iter().find(|report| report.certificates.len() == 2).unwrap();
    assert_eq!(rsa.signing_time.as_deref(), Some("2026-07-19T12:34:56Z"));
}

#[test]
fn stale_payload_is_reported_without_conflating_certificate_trust() {
    let original = ole(&[(&["Payload"], b"original")]);
    let mut editor = BinaryOfficeSignatureEditor::new(original, BinaryOfficeFormat::Doc).unwrap();
    editor.add_signature(&rsa_signer()).unwrap();
    let signed = editor.finish().unwrap();
    let (name, xml) = signature_xml(&signed);
    let tampered = ole(&[
        (&["Payload"], b"tampered"),
        (&["_xmlsignatures", name.as_str()], xml.as_slice()),
    ]);

    let report = verify(&tampered, BinaryOfficeFormat::Doc).unwrap().remove(0);
    assert_eq!(report.package_integrity, VerificationStatus::Invalid);
    assert_eq!(report.signature_value, VerificationStatus::Valid);
    assert_eq!(report.certificate_trust, litchi_ole::signature::CertificateTrust::NotEvaluated);
}

#[test]
fn no_op_clear_and_resign_are_atomic_and_preserve_payload_streams() {
    let original = ole(&[(&["Payload"], b"preserve exactly")]);
    let editor = BinaryOfficeSignatureEditor::new(original.clone(), BinaryOfficeFormat::Doc).unwrap();
    assert_eq!(editor.finish().unwrap(), original);

    let mut editor = BinaryOfficeSignatureEditor::new(original, BinaryOfficeFormat::Doc).unwrap();
    editor.add_signature(&rsa_signer()).unwrap();
    editor.resign(&rsa_signer()).unwrap();
    let signed = editor.finish().unwrap();
    assert_eq!(verify(&signed, BinaryOfficeFormat::Doc).unwrap().len(), 1);

    let mut editor = BinaryOfficeSignatureEditor::new(signed, BinaryOfficeFormat::Doc).unwrap();
    editor.clear();
    let cleared = editor.finish().unwrap();
    assert!(verify(&cleared, BinaryOfficeFormat::Doc).unwrap().is_empty());
    let mut file = OleFile::open(Cursor::new(cleared)).unwrap();
    assert_eq!(file.open_stream(&["Payload"]).unwrap(), b"preserve exactly");
}

#[test]
fn malformed_legacy_encrypted_and_resource_hostile_containers_are_rejected() {
    let duplicate_numeric = ole(&[
        (&["Payload"], b"x"),
        (&["_xmlsignatures", "1"], b"<Signature/>"),
        (&["_xmlsignatures", "01"], b"<Signature/>"),
    ]);
    assert!(matches!(
        verify(&duplicate_numeric, BinaryOfficeFormat::Doc),
        Err(BinaryOfficeSignatureError::InvalidContainer(_))
    ));

    let legacy = ole(&[(&["Payload"], b"x"), (&["_signatures"], &[0, 0, 0, 0])]);
    assert!(matches!(
        verify(&legacy, BinaryOfficeFormat::Doc),
        Err(BinaryOfficeSignatureError::LegacyCryptoApiUnsupported)
    ));
    let encrypted = ole(&[(&["Payload"], b"x"), (&["EncryptionInfo"], b"encrypted")]);
    assert!(matches!(
        BinaryOfficeSignatureEditor::new(encrypted, BinaryOfficeFormat::Doc),
        Err(BinaryOfficeSignatureError::EncryptedDocument)
    ));

    let original = ole(&[(&["Payload"], b"x")]);
    let mut editor = BinaryOfficeSignatureEditor::new(original, BinaryOfficeFormat::Doc).unwrap();
    editor.add_signature(&rsa_signer()).unwrap();
    let signed = editor.finish().unwrap();
    let mut policy = SignatureVerificationPolicy::strict();
    policy.max_signature_part_bytes = 32;
    let mut file = OleFile::open(Cursor::new(signed)).unwrap();
    assert!(matches!(
        verify_binary_office_signatures(&mut file, BinaryOfficeFormat::Doc, &policy),
        Err(BinaryOfficeSignatureError::ResourceLimit(_))
    ));
}

#[test]
fn producer_fixtures_are_noop_exact_and_facades_discover_unsigned_state() {
    let root = PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../..");
    for (relative, format) in [
        ("test-data/ole/doc/documentProperties.doc", BinaryOfficeFormat::Doc),
        ("test-data/poi/test-data/spreadsheet/Simple.xls", BinaryOfficeFormat::Xls),
        ("test-data/libreoffice-core/sc/qa/unit/data/xls/pivottable_number_grouping.xls", BinaryOfficeFormat::Xls),
        ("test-data/ole/ppt/text-margins.ppt", BinaryOfficeFormat::Ppt),
    ] {
        let bytes = std::fs::read(root.join(relative)).unwrap();
        let editor = BinaryOfficeSignatureEditor::new(bytes.clone(), format).unwrap();
        assert_eq!(editor.finish().unwrap(), bytes);
    }

    let mut doc = litchi_ole::doc::Package::open(root.join("test-data/ole/doc/documentProperties.doc")).unwrap();
    assert!(doc.verify_digital_signatures(&SignatureVerificationPolicy::strict()).unwrap().is_empty());
    let mut xls = litchi_ole::xls::XlsWorkbook::new(File::open(root.join("test-data/poi/test-data/spreadsheet/Simple.xls")).unwrap()).unwrap();
    assert!(xls.verify_digital_signatures(&SignatureVerificationPolicy::strict()).unwrap().is_empty());
    let mut ppt = litchi_ole::ppt::Package::open(root.join("test-data/ole/ppt/text-margins.ppt")).unwrap();
    assert!(ppt.verify_digital_signatures(&SignatureVerificationPolicy::strict()).unwrap().is_empty());
}
