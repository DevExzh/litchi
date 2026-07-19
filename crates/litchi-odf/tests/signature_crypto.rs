use litchi_odf::{
    OdfCanonicalizationAlgorithm, OdfDocumentSigner, OdfEncryptionProfile,
    OdfSignatureAlgorithm, OdfSignatureValidity, OwnedPackage, PackageWriter,
};

const MIME: &str = "application/vnd.oasis.opendocument.text";
const RSA_KEY: &[u8] = include_bytes!("fixtures/signatures/rsa-key.pk8");
const RSA_CERT: &[u8] = include_bytes!("fixtures/signatures/rsa-cert.der");
const EC_KEY: &[u8] = include_bytes!("fixtures/signatures/ec-key.pk8");
const EC_CERT: &[u8] = include_bytes!("fixtures/signatures/ec-cert.der");

fn signer(algorithm: OdfSignatureAlgorithm) -> OdfDocumentSigner {
    let (key, cert) = match algorithm {
        OdfSignatureAlgorithm::RsaSha256 => (RSA_KEY, RSA_CERT),
        OdfSignatureAlgorithm::EcdsaP256Sha256 => (EC_KEY, EC_CERT),
    };
    OdfDocumentSigner::from_pkcs8_der(
        algorithm,
        key,
        vec![cert.to_vec()],
        "2026-07-19T12:00:00Z",
    )
    .unwrap()
}

fn signed_package(
    algorithm: OdfSignatureAlgorithm,
    canonicalization: OdfCanonicalizationAlgorithm,
    encrypted: bool,
) -> Vec<u8> {
    let mut writer = PackageWriter::new();
    writer.set_mimetype(MIME).unwrap();
    writer
        .set_document_signer(signer(algorithm).with_canonicalization(canonicalization))
        .unwrap();
    if encrypted {
        writer
            .set_encryption("encryption-password", OdfEncryptionProfile::compatible())
            .unwrap();
    }
    writer
        .add_file("content.xml", b"<root b=\"2\" a=\"1\"><empty/></root>")
        .unwrap();
    writer
        .add_file("Pictures/a space.bin", b"binary reference")
        .unwrap();
    writer.finish().unwrap()
}

#[test]
fn authors_and_verifies_rsa_and_ecdsa_with_both_canonicalizations() {
    for algorithm in [
        OdfSignatureAlgorithm::RsaSha256,
        OdfSignatureAlgorithm::EcdsaP256Sha256,
    ] {
        for canonicalization in [
            OdfCanonicalizationAlgorithm::InclusiveXml10,
            OdfCanonicalizationAlgorithm::ExclusiveXml10,
        ] {
            let package = OwnedPackage::from_bytes(signed_package(
                algorithm,
                canonicalization,
                false,
            ))
            .unwrap();
            let results = package.verify_document_signatures().unwrap();
            assert_eq!(results.len(), 1);
            assert_eq!(results[0].algorithm, Some(algorithm));
            assert_eq!(results[0].canonicalization, Some(canonicalization));
            assert_eq!(results[0].validity, OdfSignatureValidity::Valid);
            assert_eq!(results[0].signing_certificate_index, Some(0));
        }
    }
}

#[test]
fn signing_composes_with_encrypt_first_packages() {
    let bytes = signed_package(
        OdfSignatureAlgorithm::RsaSha256,
        OdfCanonicalizationAlgorithm::InclusiveXml10,
        true,
    );
    let verification = OwnedPackage::from_bytes(bytes.clone())
        .unwrap()
        .verify_document_signatures()
        .unwrap();
    assert_eq!(verification[0].validity, OdfSignatureValidity::Valid);
    let decrypted = OwnedPackage::from_bytes_with_password(bytes, "encryption-password").unwrap();
    assert_eq!(
        decrypted.get_file("content.xml").unwrap(),
        b"<root b=\"2\" a=\"1\"><empty/></root>"
    );
    assert!(decrypted.get_file("META-INF/documentsignatures.xml").is_ok());
}

#[test]
fn rejects_a_private_key_that_does_not_match_the_leaf_certificate() {
    assert!(OdfDocumentSigner::from_pkcs8_der(
        OdfSignatureAlgorithm::RsaSha256,
        RSA_KEY,
        vec![EC_CERT.to_vec()],
        "2026-07-19T12:00:00Z",
    )
    .is_err());
}

#[test]
fn verifies_the_bundled_libreoffice_xades_signature() {
    let bytes = include_bytes!(
        "../../../3rdparty/libreoffice-core/xmlsecurity/qa/unit/signing/data/signed_with_x509certificate_chain.odt"
    );
    let results = OwnedPackage::from_bytes(bytes.to_vec())
        .unwrap()
        .verify_document_signatures()
        .unwrap();
    assert_eq!(results.len(), 1);
    assert_eq!(results[0].validity, OdfSignatureValidity::Valid, "{results:?}");
}

#[test]
fn signature_implementation_contains_no_unsafe_code() {
    assert!(!include_str!("../src/signature_crypto.rs").contains("unsafe"));
}

fn rewrite_entry(bytes: &[u8], target: &str, replacement: impl FnOnce(Vec<u8>) -> Vec<u8>) -> Vec<u8> {
    let archive = soapberry_zip::office::ArchiveReader::new(bytes).unwrap();
    let mut writer = soapberry_zip::office::StreamingArchiveWriter::new();
    let mut replacement = Some(replacement);
    for name in archive.file_names() {
        let mut content = archive.read(name).unwrap();
        if name == target {
            content = replacement.take().unwrap()(content);
        }
        if name == "mimetype" || archive.is_stored(name).unwrap() {
            writer.write_stored(name, &content).unwrap();
        } else {
            writer.write_deflated(name, &content).unwrap();
        }
    }
    writer.finish_to_bytes().unwrap()
}

#[test]
fn reports_tampering_and_unsupported_algorithms() {
    let bytes = signed_package(
        OdfSignatureAlgorithm::RsaSha256,
        OdfCanonicalizationAlgorithm::InclusiveXml10,
        false,
    );
    let tampered = rewrite_entry(&bytes, "content.xml", |content| {
        String::from_utf8(content)
            .unwrap()
            .replace("empty", "changed")
            .into_bytes()
    });
    let result = OwnedPackage::from_bytes(tampered)
        .unwrap()
        .verify_document_signatures()
        .unwrap();
    assert_eq!(result[0].validity, OdfSignatureValidity::ReferenceDigestMismatch);
    assert_eq!(result[0].failed_reference_uri.as_deref(), Some("content.xml"));

    let unsupported = rewrite_entry(&bytes, "META-INF/documentsignatures.xml", |content| {
        String::from_utf8(content)
            .unwrap()
            .replace(
                "http://www.w3.org/2001/04/xmldsig-more#rsa-sha256",
                "urn:unsupported:signature",
            )
            .into_bytes()
    });
    let result = OwnedPackage::from_bytes(unsupported)
        .unwrap()
        .verify_document_signatures()
        .unwrap();
    assert_eq!(result[0].validity, OdfSignatureValidity::UnsupportedAlgorithm);

    let invalid_signature = rewrite_entry(&bytes, "META-INF/documentsignatures.xml", |mut content| {
        let marker = b"<ds:SignatureValue>";
        let offset = content.windows(marker.len()).position(|value| value == marker).unwrap()
            + marker.len();
        content[offset] = if content[offset] == b'A' { b'B' } else { b'A' };
        content
    });
    let result = OwnedPackage::from_bytes(invalid_signature)
        .unwrap()
        .verify_document_signatures()
        .unwrap();
    assert_eq!(result[0].validity, OdfSignatureValidity::InvalidSignature);
}

#[test]
fn rejects_spoofed_and_duplicate_package_reference_paths() {
    let bytes = signed_package(
        OdfSignatureAlgorithm::RsaSha256,
        OdfCanonicalizationAlgorithm::InclusiveXml10,
        false,
    );
    for replacement in ["../escape", "content.xml"] {
        let malformed = rewrite_entry(&bytes, "META-INF/documentsignatures.xml", |content| {
            String::from_utf8(content)
                .unwrap()
                .replace("Pictures/a%20space.bin", replacement)
                .into_bytes()
        });
        assert!(OwnedPackage::from_bytes(malformed)
            .unwrap()
            .verify_document_signatures()
            .is_err());
    }
}

#[test]
fn late_signing_configuration_failure_is_atomic() {
    let mut writer = PackageWriter::new();
    writer.set_mimetype(MIME).unwrap();
    writer.add_file("content.xml", b"<root/>").unwrap();
    assert!(writer.set_document_signer(signer(OdfSignatureAlgorithm::RsaSha256)).is_err());
    let package = OwnedPackage::from_bytes(writer.finish().unwrap()).unwrap();
    assert!(package.verify_document_signatures().unwrap().is_empty());
}
