//! Safe, trust-neutral ODF document-signature cryptography.

use super::model::{DOCUMENT_SIGNATURE_PATH, DigitalSignature, parse_signature_container};
use base64::{Engine as _, engine::general_purpose::STANDARD as BASE64_STANDARD};
use bergshamra_c14n::{C14nMode, canonicalize};
use litchi_core::{Error, Result, xml::escape_xml};
use p256::ecdsa::{Signature as EcdsaSignature, SigningKey as EcdsaSigningKey, VerifyingKey};
use p256::pkcs8::DecodePrivateKey as _;
use quick_xml::{Reader, events::Event};
use rsa::pkcs1::DecodeRsaPublicKey as _;
use rsa::pkcs1v15::{
    Signature as RsaSignature, SigningKey as RsaSigningKey, VerifyingKey as RsaVerifyingKey,
};
use rsa::sha2::Sha256 as RsaSha256;
use rsa::traits::PublicKeyParts;
use rsa::{RsaPrivateKey, RsaPublicKey};
use sha2::{Digest, Sha256};
use signature::{SignatureEncoding, Signer, Verifier};
use soapberry_zip::office::{ArchiveReader, StreamingArchiveWriter};
use std::collections::HashSet;
use x509_cert::Certificate;
use x509_cert::der::Decode;
use zeroize::Zeroizing;

const DS: &str = "http://www.w3.org/2000/09/xmldsig#";
const DSIG: &str = "urn:oasis:names:tc:opendocument:xmlns:digitalsignature:1.0";
const XADES: &str = "http://uri.etsi.org/01903/v1.3.2#";
const SHA256_URI: &str = "http://www.w3.org/2001/04/xmlenc#sha256";
const RSA_SHA256_URI: &str = "http://www.w3.org/2001/04/xmldsig-more#rsa-sha256";
const ECDSA_SHA256_URI: &str = "http://www.w3.org/2001/04/xmldsig-more#ecdsa-sha256";
const C14N_INCLUSIVE_URI: &str = "http://www.w3.org/TR/2001/REC-xml-c14n-20010315";
const C14N_EXCLUSIVE_URI: &str = "http://www.w3.org/2001/10/xml-exc-c14n#";
const SIGNED_PROPERTIES_TYPE: &str = "http://uri.etsi.org/01903#SignedProperties";

/// Supported XMLDSIG signature algorithms.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SignatureAlgorithm {
    RsaSha256,
    EcdsaP256Sha256,
}

impl SignatureAlgorithm {
    fn uri(self) -> &'static str {
        match self {
            Self::RsaSha256 => RSA_SHA256_URI,
            Self::EcdsaP256Sha256 => ECDSA_SHA256_URI,
        }
    }

    fn parse(value: &str) -> Option<Self> {
        match value {
            RSA_SHA256_URI => Some(Self::RsaSha256),
            ECDSA_SHA256_URI => Some(Self::EcdsaP256Sha256),
            _ => None,
        }
    }
}

/// Supported XML canonicalization modes, always without comments.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum CanonicalizationAlgorithm {
    InclusiveXml10,
    ExclusiveXml10,
}

impl CanonicalizationAlgorithm {
    fn uri(self) -> &'static str {
        match self {
            Self::InclusiveXml10 => C14N_INCLUSIVE_URI,
            Self::ExclusiveXml10 => C14N_EXCLUSIVE_URI,
        }
    }

    fn parse(value: &str) -> Option<Self> {
        match value {
            C14N_INCLUSIVE_URI => Some(Self::InclusiveXml10),
            C14N_EXCLUSIVE_URI => Some(Self::ExclusiveXml10),
            _ => None,
        }
    }

    fn mode(self) -> C14nMode {
        match self {
            Self::InclusiveXml10 => C14nMode::Inclusive,
            Self::ExclusiveXml10 => C14nMode::Exclusive,
        }
    }
}

/// Trust-neutral cryptographic result for one embedded signature.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SignatureVerification {
    pub signature_id: Option<String>,
    pub algorithm: Option<SignatureAlgorithm>,
    pub canonicalization: Option<CanonicalizationAlgorithm>,
    pub validity: SignatureValidity,
    pub signing_time: Option<String>,
    pub certificate_chain_der: Vec<Vec<u8>>,
    /// Certificate whose public key verified the signature, if any.
    pub signing_certificate_index: Option<usize>,
    /// URI whose digest failed, when validity is `ReferenceDigestMismatch`.
    pub failed_reference_uri: Option<String>,
}

/// This reports mathematical/package validity only and never PKI trust.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SignatureValidity {
    Valid,
    ReferenceDigestMismatch,
    InvalidSignature,
    UnsupportedAlgorithm,
    MissingCertificate,
}

/// A validated document signer retaining PKCS#8 key bytes in zeroizing storage.
pub struct DocumentSigner {
    algorithm: SignatureAlgorithm,
    canonicalization: CanonicalizationAlgorithm,
    private_key: Zeroizing<Vec<u8>>,
    certificates: Vec<Vec<u8>>,
    signing_time: String,
}

impl DocumentSigner {
    /// Parse a DER PKCS#8 private key and DER X.509 chain. The leaf certificate must be first.
    ///
    /// # Errors
    ///
    /// Returns an error when the key or certificate chain is malformed, insecure, or does not
    /// match the requested signature algorithm.
    pub fn from_pkcs8_der(
        algorithm: SignatureAlgorithm,
        private_key: impl AsRef<[u8]>,
        certificates: Vec<Vec<u8>>,
        signing_time: impl Into<String>,
    ) -> Result<Self> {
        Self::new(
            algorithm,
            private_key.as_ref().to_vec(),
            certificates,
            signing_time.into(),
        )
    }

    /// Parse an unencrypted PEM PKCS#8 key and a PEM X.509 certificate chain.
    ///
    /// # Errors
    ///
    /// Returns an error when the PEM input does not contain one valid private key and a valid,
    /// matching certificate chain.
    pub fn from_pkcs8_pem(
        algorithm: SignatureAlgorithm,
        private_key_pem: &str,
        certificate_chain_pem: &str,
        signing_time: impl Into<String>,
    ) -> Result<Self> {
        let private_keys = decode_pem_blocks(private_key_pem, "PRIVATE KEY")?;
        if private_keys.len() != 1 {
            return invalid("exactly one PKCS#8 PRIVATE KEY PEM block is required");
        }
        let certificates = decode_pem_blocks(certificate_chain_pem, "CERTIFICATE")?;
        let private_key = private_keys.into_iter().next().ok_or_else(|| {
            Error::InvalidFormat("PKCS#8 key block unexpectedly disappeared".to_string())
        })?;
        Self::new(algorithm, private_key, certificates, signing_time.into())
    }

    #[must_use]
    pub fn with_canonicalization(mut self, value: CanonicalizationAlgorithm) -> Self {
        self.canonicalization = value;
        self
    }

    fn new(
        algorithm: SignatureAlgorithm,
        private_key: Vec<u8>,
        certificates: Vec<Vec<u8>>,
        signing_time: String,
    ) -> Result<Self> {
        if certificates.is_empty() || certificates.len() > 64 {
            return invalid("a bounded non-empty X.509 certificate chain is required");
        }
        if signing_time.is_empty() || signing_time.len() > 128 {
            return invalid("invalid signature signing time");
        }
        for certificate in &certificates {
            Certificate::from_der(certificate)
                .map_err(|error| format_error(format!("invalid DER X.509 certificate: {error}")))?;
        }
        validate_key_matches_certificate(algorithm, &private_key, &certificates[0])?;
        Ok(Self {
            algorithm,
            canonicalization: CanonicalizationAlgorithm::InclusiveXml10,
            private_key: Zeroizing::new(private_key),
            certificates,
            signing_time,
        })
    }
}

pub(crate) fn sign_package(unsigned: &[u8], signer: &DocumentSigner) -> Result<Vec<u8>> {
    let archive = ArchiveReader::new(unsigned)
        .map_err(|error| format_error(format!("invalid ODF ZIP archive while signing: {error}")))?;
    let mut names: Vec<String> = archive
        .file_names()
        .filter(|name| should_reference(name))
        .map(str::to_string)
        .collect();
    names.sort();
    if names.len() > 4096 {
        return invalid("ODF package has too many signature references");
    }

    let signature_id = "ID_litchi_document_signature";
    let properties_id = "idSignedProperties";
    let certificate_digest = BASE64_STANDARD.encode(Sha256::digest(&signer.certificates[0]));
    let signed_properties = format!(
        "<xades:SignedProperties xmlns:xades=\"{XADES}\" xmlns:ds=\"{DS}\" Id=\"{properties_id}\"><xades:SignedSignatureProperties><xades:SigningTime>{}</xades:SigningTime><xades:SigningCertificate><xades:Cert><xades:CertDigest><ds:DigestMethod Algorithm=\"{SHA256_URI}\"></ds:DigestMethod><ds:DigestValue>{certificate_digest}</ds:DigestValue></xades:CertDigest></xades:Cert></xades:SigningCertificate><xades:SignaturePolicyIdentifier><xades:SignaturePolicyImplied></xades:SignaturePolicyImplied></xades:SignaturePolicyIdentifier></xades:SignedSignatureProperties></xades:SignedProperties>",
        escape_xml(&signer.signing_time)
    );
    let properties_digest = digest_canonical_xml(&signed_properties, signer.canonicalization)?;

    let mut signed_info = format!(
        "<ds:SignedInfo xmlns:ds=\"{DS}\"><ds:CanonicalizationMethod Algorithm=\"{}\"></ds:CanonicalizationMethod><ds:SignatureMethod Algorithm=\"{}\"></ds:SignatureMethod>",
        signer.canonicalization.uri(),
        signer.algorithm.uri()
    );
    for name in &names {
        let bytes = archive.read(name).map_err(|error| {
            format_error(format!(
                "unable to read ODF signature reference '{name}': {error}"
            ))
        })?;
        let stored = archive.is_stored(name).map_err(|error| {
            format_error(format!(
                "unable to inspect ODF signature reference '{name}': {error}"
            ))
        })?;
        let transform = if is_xml_path(name) && !stored {
            Some(signer.canonicalization)
        } else {
            None
        };
        let digest = if let Some(canonicalization) = transform {
            digest_canonical_xml_bytes(&bytes, canonicalization)?
        } else {
            Sha256::digest(&bytes).to_vec()
        };
        signed_info.push_str("<ds:Reference URI=\"");
        signed_info.push_str(&escape_xml(&encode_package_uri(name)));
        signed_info.push_str("\">");
        if let Some(canonicalization) = transform {
            signed_info.push_str("<ds:Transforms><ds:Transform Algorithm=\"");
            signed_info.push_str(canonicalization.uri());
            signed_info.push_str("\"></ds:Transform></ds:Transforms>");
        }
        signed_info.push_str("<ds:DigestMethod Algorithm=\"");
        signed_info.push_str(SHA256_URI);
        signed_info.push_str("\"></ds:DigestMethod><ds:DigestValue>");
        signed_info.push_str(&BASE64_STANDARD.encode(digest));
        signed_info.push_str("</ds:DigestValue></ds:Reference>");
    }
    signed_info.push_str("<ds:Reference URI=\"");
    signed_info.push('#');
    signed_info.push_str(properties_id);
    signed_info.push_str("\" Type=\"");
    signed_info.push_str(SIGNED_PROPERTIES_TYPE);
    signed_info.push_str("\"><ds:Transforms><ds:Transform Algorithm=\"");
    signed_info.push_str(signer.canonicalization.uri());
    signed_info.push_str("\"></ds:Transform></ds:Transforms><ds:DigestMethod Algorithm=\"");
    signed_info.push_str(SHA256_URI);
    signed_info.push_str("\"></ds:DigestMethod><ds:DigestValue>");
    signed_info.push_str(&BASE64_STANDARD.encode(properties_digest));
    signed_info.push_str("</ds:DigestValue></ds:Reference></ds:SignedInfo>");

    let canonical_signed_info = canonicalize_fragment(&signed_info, signer.canonicalization)?;
    let signature_value = sign_bytes(signer, &canonical_signed_info)?;
    let mut key_info = String::from("<ds:KeyInfo><ds:X509Data>");
    for certificate in &signer.certificates {
        key_info.push_str("<ds:X509Certificate>");
        key_info.push_str(&BASE64_STANDARD.encode(certificate));
        key_info.push_str("</ds:X509Certificate>");
    }
    key_info.push_str("</ds:X509Data></ds:KeyInfo>");
    let signature_xml = format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\"?><dsig:document-signatures xmlns:dsig=\"{DSIG}\" dsig:version=\"1.3\"><ds:Signature xmlns:ds=\"{DS}\" Id=\"{signature_id}\">{signed_info}<ds:SignatureValue>{}</ds:SignatureValue>{key_info}<ds:Object><xades:QualifyingProperties xmlns:xades=\"{XADES}\" Target=\"#{signature_id}\">{signed_properties}</xades:QualifyingProperties></ds:Object></ds:Signature></dsig:document-signatures>",
        BASE64_STANDARD.encode(signature_value)
    );
    xml_minifier::audit::verify_authored(
        signature_xml.as_bytes(),
        xml_minifier::audit::Limits::default(),
    )
    .map(|_report| ())
    .map_err(|error| {
        format_error(format!(
            "generated ODF signature XML failed publication validation: {error}"
        ))
    })?;

    rebuild_with_signature(&archive, &signature_xml)
}

pub(crate) fn verify_package(data: &[u8]) -> Result<Vec<SignatureVerification>> {
    let archive = ArchiveReader::new(data)
        .map_err(|error| format_error(format!("invalid ODF ZIP archive: {error}")))?;
    if !archive.contains(DOCUMENT_SIGNATURE_PATH) {
        return Ok(Vec::new());
    }
    let signature_xml = archive.read(DOCUMENT_SIGNATURE_PATH).map_err(|error| {
        format_error(format!(
            "unable to read document signature container: {error}"
        ))
    })?;
    let signatures = parse_signature_container(&signature_xml)?;
    let text = std::str::from_utf8(&signature_xml)
        .map_err(|error| format_error(format!("signature XML is not UTF-8: {error}")))?;
    let signed_infos = extract_elements(text, "SignedInfo", None)?;
    if signed_infos.len() != signatures.len() {
        return invalid("signature XML has ambiguous SignedInfo elements");
    }
    signatures
        .iter()
        .zip(signed_infos)
        .map(|(signature, signed_info)| verify_one(&archive, text, signature, &signed_info))
        .collect()
}

fn verify_one(
    archive: &ArchiveReader<'_>,
    signature_xml: &str,
    signature: &DigitalSignature,
    signed_info_xml: &str,
) -> Result<SignatureVerification> {
    let algorithm = SignatureAlgorithm::parse(&signature.signature_method);
    let canonicalization = CanonicalizationAlgorithm::parse(&signature.canonicalization_method);
    let certificates = signature
        .x509_certificates
        .iter()
        .map(|value| {
            BASE64_STANDARD
                .decode(value)
                .map_err(|error| format_error(format!("invalid X.509 certificate base64: {error}")))
        })
        .collect::<Result<Vec<_>>>()?;
    for value in &certificates {
        Certificate::from_der(value)
            .map_err(|error| format_error(format!("invalid DER X.509 certificate: {error}")))?;
    }
    let mut result = SignatureVerification {
        signature_id: signature.id.clone(),
        algorithm,
        canonicalization,
        validity: SignatureValidity::UnsupportedAlgorithm,
        signing_time: signature.signing_time.clone(),
        certificate_chain_der: certificates,
        signing_certificate_index: None,
        failed_reference_uri: None,
    };
    let (Some(verified_algorithm), Some(verified_canonicalization)) = (algorithm, canonicalization)
    else {
        return Ok(result);
    };

    let expected: HashSet<String> = archive
        .file_names()
        .filter(|name| should_reference(name))
        .map(str::to_string)
        .collect();
    let mut referenced = HashSet::new();
    for reference in &signature.references {
        if reference.uri.starts_with('#') {
            continue;
        }
        let path = decode_package_uri(&reference.uri)?;
        if !referenced.insert(path.clone()) {
            return invalid("duplicate ODF package signature reference");
        }
        if !expected.contains(&path) || !archive.contains(&path) {
            return invalid("signature reference points outside the required package set");
        }
    }
    if referenced != expected {
        return invalid("document signature does not reference every required package entry");
    }
    for reference in &signature.references {
        if reference.digest_method != SHA256_URI {
            return Ok(result);
        }
        let expected_digest = BASE64_STANDARD
            .decode(&reference.digest_value)
            .map_err(|error| format_error(format!("invalid reference digest base64: {error}")))?;
        if expected_digest.len() != 32 {
            return invalid("SHA-256 signature digest has the wrong size");
        }
        let bytes = if let Some(id) = reference.uri.strip_prefix('#') {
            if id.is_empty() {
                return invalid("empty same-document signature reference");
            }
            let elements = extract_elements(signature_xml, "", Some(id))?;
            if elements.len() != 1 {
                return invalid("missing or duplicate same-document signature reference");
            }
            if reference.transforms.is_empty() {
                canonicalize_fragment(&elements[0], CanonicalizationAlgorithm::InclusiveXml10)?
            } else {
                apply_transforms(elements[0].as_bytes(), &reference.transforms)?
            }
        } else {
            let path = decode_package_uri(&reference.uri)?;
            let bytes = archive.read(&path).map_err(|error| {
                format_error(format!("unable to read signature reference: {error}"))
            })?;
            apply_transforms(&bytes, &reference.transforms)?
        };
        if Sha256::digest(&bytes).as_slice() != expected_digest.as_slice() {
            result.validity = SignatureValidity::ReferenceDigestMismatch;
            result.failed_reference_uri = Some(reference.uri.clone());
            return Ok(result);
        }
    }
    if result.certificate_chain_der.is_empty() {
        result.validity = SignatureValidity::MissingCertificate;
        return Ok(result);
    }
    let signed_info = canonicalize_fragment(signed_info_xml, verified_canonicalization)?;
    let signature_value = BASE64_STANDARD
        .decode(&signature.signature_value)
        .map_err(|error| format_error(format!("invalid signature value base64: {error}")))?;
    for (index, certificate) in result.certificate_chain_der.iter().enumerate() {
        if matches!(
            verify_bytes(
                verified_algorithm,
                certificate,
                &signed_info,
                &signature_value
            ),
            Ok(true)
        ) {
            result.validity = SignatureValidity::Valid;
            result.signing_certificate_index = Some(index);
            return Ok(result);
        }
    }
    result.validity = SignatureValidity::InvalidSignature;
    Ok(result)
}

fn apply_transforms(bytes: &[u8], transforms: &[String]) -> Result<Vec<u8>> {
    match transforms {
        [] => Ok(bytes.to_vec()),
        [transform] => {
            let mode = CanonicalizationAlgorithm::parse(transform)
                .ok_or_else(|| format_error("unsupported signature reference transform"))?;
            digest_input_canonical_xml(bytes, mode)
        },
        _ => invalid("unsupported signature transform chain"),
    }
}

fn digest_input_canonical_xml(bytes: &[u8], mode: CanonicalizationAlgorithm) -> Result<Vec<u8>> {
    let xml = std::str::from_utf8(bytes)
        .map_err(|error| format_error(format!("C14N input is not UTF-8: {error}")))?;
    canonicalize_fragment(xml, mode)
}

fn digest_canonical_xml(xml: &str, mode: CanonicalizationAlgorithm) -> Result<Vec<u8>> {
    Ok(Sha256::digest(canonicalize_fragment(xml, mode)?).to_vec())
}

fn digest_canonical_xml_bytes(xml: &[u8], mode: CanonicalizationAlgorithm) -> Result<Vec<u8>> {
    Ok(Sha256::digest(digest_input_canonical_xml(xml, mode)?).to_vec())
}

fn canonicalize_fragment(xml: &str, mode: CanonicalizationAlgorithm) -> Result<Vec<u8>> {
    canonicalize::<&str>(xml, mode.mode(), None, &[])
        .map_err(|error| format_error(format!("XML canonicalization failed: {error}")))
}

fn sign_bytes(signer: &DocumentSigner, message: &[u8]) -> Result<Vec<u8>> {
    match signer.algorithm {
        SignatureAlgorithm::RsaSha256 => {
            let key = RsaPrivateKey::from_pkcs8_der(&signer.private_key).map_err(|error| {
                format_error(format!("invalid RSA PKCS#8 private key: {error}"))
            })?;
            if key.n().bits() < 2048 {
                return invalid("RSA signing keys must be at least 2048 bits");
            }
            let signature: RsaSignature = RsaSigningKey::<RsaSha256>::new(key).sign(message);
            Ok(signature.to_vec())
        },
        SignatureAlgorithm::EcdsaP256Sha256 => {
            let key = EcdsaSigningKey::from_pkcs8_der(&signer.private_key).map_err(|error| {
                format_error(format!("invalid P-256 PKCS#8 private key: {error}"))
            })?;
            let signature: EcdsaSignature = key.sign(message);
            Ok(signature.to_bytes().to_vec())
        },
    }
}

fn verify_bytes(
    algorithm: SignatureAlgorithm,
    certificate_der: &[u8],
    message: &[u8],
    signature_bytes: &[u8],
) -> Result<bool> {
    let parsed_certificate = Certificate::from_der(certificate_der)
        .map_err(|error| format_error(format!("invalid DER X.509 certificate: {error}")))?;
    let public = parsed_certificate
        .tbs_certificate
        .subject_public_key_info
        .subject_public_key
        .raw_bytes();
    Ok(match algorithm {
        SignatureAlgorithm::RsaSha256 => {
            let key = RsaPublicKey::from_pkcs1_der(public).map_err(|error| {
                format_error(format!(
                    "certificate does not contain an RSA public key: {error}"
                ))
            })?;
            let rsa_signature = RsaSignature::try_from(signature_bytes).map_err(|error| {
                format_error(format!("invalid RSA signature encoding: {error}"))
            })?;
            RsaVerifyingKey::<RsaSha256>::new(key)
                .verify(message, &rsa_signature)
                .is_ok()
        },
        SignatureAlgorithm::EcdsaP256Sha256 => {
            let key = VerifyingKey::from_sec1_bytes(public).map_err(|error| {
                format_error(format!(
                    "certificate does not contain a P-256 public key: {error}"
                ))
            })?;
            let ecdsa_signature = EcdsaSignature::from_slice(signature_bytes).map_err(|error| {
                format_error(format!("invalid ECDSA signature encoding: {error}"))
            })?;
            key.verify(message, &ecdsa_signature).is_ok()
        },
    })
}

fn validate_key_matches_certificate(
    algorithm: SignatureAlgorithm,
    private_key: &[u8],
    certificate_der: &[u8],
) -> Result<()> {
    let parsed_certificate = Certificate::from_der(certificate_der)
        .map_err(|error| format_error(format!("invalid DER X.509 certificate: {error}")))?;
    let public = parsed_certificate
        .tbs_certificate
        .subject_public_key_info
        .subject_public_key
        .raw_bytes();
    let matches = match algorithm {
        SignatureAlgorithm::RsaSha256 => {
            let private = RsaPrivateKey::from_pkcs8_der(private_key).map_err(|error| {
                format_error(format!("invalid RSA PKCS#8 private key: {error}"))
            })?;
            if private.n().bits() < 2048 {
                return invalid("RSA signing keys must be at least 2048 bits");
            }
            let certificate_key = RsaPublicKey::from_pkcs1_der(public)
                .map_err(|error| format_error(format!("leaf certificate is not RSA: {error}")))?;
            private.to_public_key() == certificate_key
        },
        SignatureAlgorithm::EcdsaP256Sha256 => {
            let private = EcdsaSigningKey::from_pkcs8_der(private_key).map_err(|error| {
                format_error(format!("invalid P-256 PKCS#8 private key: {error}"))
            })?;
            private.verifying_key().to_encoded_point(false).as_bytes() == public
        },
    };
    if !matches {
        return invalid("private key does not match the leaf certificate");
    }
    Ok(())
}

fn rebuild_with_signature(archive: &ArchiveReader<'_>, signature_xml: &str) -> Result<Vec<u8>> {
    let mut writer = StreamingArchiveWriter::new();
    for name in archive.file_names() {
        if name == DOCUMENT_SIGNATURE_PATH {
            continue;
        }
        let bytes = archive.read(name).map_err(|error| {
            format_error(format!("unable to rebuild signed ODF package: {error}"))
        })?;
        let stored = archive.is_stored(name).map_err(|error| {
            format_error(format!(
                "unable to inspect ODF package entry while rebuilding: {error}"
            ))
        })?;
        if name == "mimetype" || stored {
            writer.write_stored(name, &bytes)
        } else {
            writer.write_deflated(name, &bytes)
        }
        .map_err(|error| Error::ZipError(error.to_string()))?;
    }
    writer
        .write_deflated(DOCUMENT_SIGNATURE_PATH, signature_xml.as_bytes())
        .map_err(|error| Error::ZipError(error.to_string()))?;
    writer
        .finish_to_bytes()
        .map_err(|error| Error::ZipError(error.to_string()))
}

fn should_reference(name: &str) -> bool {
    name != DOCUMENT_SIGNATURE_PATH && !name.starts_with("external-data/") && !name.ends_with('/')
}

fn is_xml_path(name: &str) -> bool {
    name.rsplit_once('.').is_some_and(|(_, extension)| {
        extension.eq_ignore_ascii_case("xml") || extension.eq_ignore_ascii_case("rdf")
    })
}

fn encode_package_uri(path: &str) -> String {
    let mut result = String::with_capacity(path.len());
    for byte in path.bytes() {
        if byte.is_ascii_alphanumeric() || matches!(byte, b'-' | b'.' | b'_' | b'~' | b'/') {
            result.push(char::from(byte));
        } else {
            const HEX_DIGITS: &[u8; 16] = b"0123456789ABCDEF";
            result.push('%');
            result.push(char::from(HEX_DIGITS[usize::from(byte >> 4)]));
            result.push(char::from(HEX_DIGITS[usize::from(byte & 0x0f)]));
        }
    }
    result
}

fn decode_package_uri(uri: &str) -> Result<String> {
    if uri.is_empty() || uri.starts_with('/') || uri.contains('?') || uri.contains('#') {
        return invalid("invalid package signature reference URI");
    }
    let bytes = uri.as_bytes();
    let mut decoded = Vec::with_capacity(bytes.len());
    let mut index = 0;
    while index < bytes.len() {
        if bytes[index] == b'%' {
            if index + 2 >= bytes.len() {
                return invalid("invalid percent encoding in signature URI");
            }
            let value = u8::from_str_radix(&uri[index + 1..index + 3], 16).map_err(|error| {
                format_error(format!(
                    "invalid percent encoding in signature URI: {error}"
                ))
            })?;
            decoded.push(value);
            index += 3;
        } else {
            decoded.push(bytes[index]);
            index += 1;
        }
    }
    let path = String::from_utf8(decoded)
        .map_err(|error| format_error(format!("signature URI is not valid UTF-8: {error}")))?;
    if path
        .split('/')
        .any(|segment| matches!(segment, "" | "." | ".."))
        || path.contains('\\')
        || path.contains(':')
    {
        return invalid("non-local package signature reference URI");
    }
    Ok(path)
}

fn extract_elements(xml: &str, local: &str, id: Option<&str>) -> Result<Vec<String>> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().check_end_names = true;
    let mut found = Vec::new();
    let inherited_default_ds = xml.contains(&format!("<Signature xmlns=\"{DS}\""));
    loop {
        let start = usize::try_from(reader.buffer_position())
            .map_err(|error| format_error(format!("signature XML position is invalid: {error}")))?;
        match reader.read_event().map_err(|error| {
            format_error(format!(
                "invalid signature XML while selecting signed nodes: {error}"
            ))
        })? {
            Event::Start(element) => {
                let local_matches =
                    local.is_empty() || element.local_name().as_ref() == local.as_bytes();
                let id_matches = if let Some(expected_id) = id {
                    let mut matched = false;
                    for raw_attribute in element.attributes().with_checks(true) {
                        let parsed_attribute = raw_attribute.map_err(|error| {
                            format_error(format!("invalid signed-node attribute: {error}"))
                        })?;
                        if matches!(parsed_attribute.key.as_ref(), b"Id" | b"ID" | b"id")
                            && parsed_attribute.value.as_ref() == expected_id.as_bytes()
                        {
                            matched = true;
                        }
                    }
                    matched
                } else {
                    true
                };
                if local_matches && id_matches {
                    let mut depth = 1usize;
                    while depth != 0 {
                        match reader.read_event().map_err(|error| {
                            format_error(format!("invalid signed-node XML: {error}"))
                        })? {
                            Event::Start(_) => depth += 1,
                            Event::End(_) => depth -= 1,
                            Event::Eof => return invalid("incomplete signed-node XML"),
                            Event::DocType(_) | Event::GeneralRef(_) | Event::PI(_) => {
                                return invalid("active XML is prohibited in signed nodes");
                            },
                            Event::Empty(_)
                            | Event::Text(_)
                            | Event::CData(_)
                            | Event::Comment(_)
                            | Event::Decl(_) => {},
                        }
                    }
                    let end = usize::try_from(reader.buffer_position()).map_err(|error| {
                        format_error(format!("signature XML position is invalid: {error}"))
                    })?;
                    let mut fragment = xml[start..end].to_string();
                    inject_known_namespaces(&mut fragment, inherited_default_ds);
                    found.push(fragment);
                }
            },
            Event::DocType(_) | Event::GeneralRef(_) | Event::PI(_) => {
                return invalid("active XML is prohibited in signatures");
            },
            Event::Eof => break,
            Event::End(_)
            | Event::Empty(_)
            | Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_) => {},
        }
    }
    Ok(found)
}

fn inject_known_namespaces(fragment: &mut String, inherited_default_ds: bool) {
    let Some(end) = fragment.find('>') else {
        return;
    };
    let start = &fragment[..end];
    let mut declarations = String::new();
    if inherited_default_ds && !start.contains("xmlns=") {
        declarations.push_str(" xmlns=\"");
        declarations.push_str(DS);
        declarations.push('\"');
    }
    if fragment.contains("ds:") && !start.contains("xmlns:ds=") {
        declarations.push_str(" xmlns:ds=\"");
        declarations.push_str(DS);
        declarations.push('\"');
    }
    if fragment.contains("xades:") && !start.contains("xmlns:xades=") {
        declarations.push_str(" xmlns:xades=\"");
        declarations.push_str(XADES);
        declarations.push('\"');
    }
    if fragment.contains("xd:") && !start.contains("xmlns:xd=") {
        declarations.push_str(" xmlns:xd=\"");
        declarations.push_str(XADES);
        declarations.push('\"');
    }
    fragment.insert_str(end, &declarations);
}

fn decode_pem_blocks(value: &str, label: &str) -> Result<Vec<Vec<u8>>> {
    let begin = format!("-----BEGIN {label}-----");
    let end = format!("-----END {label}-----");
    let mut blocks = Vec::new();
    let mut rest = value;
    while let Some(start) = rest.find(&begin) {
        rest = &rest[start + begin.len()..];
        let finish = rest
            .find(&end)
            .ok_or_else(|| format_error("incomplete PEM block"))?;
        let body: String = rest[..finish]
            .chars()
            .filter(|character| !character.is_whitespace())
            .collect();
        blocks.push(
            BASE64_STANDARD
                .decode(body)
                .map_err(|error| format_error(format!("invalid PEM base64: {error}")))?,
        );
        rest = &rest[finish + end.len()..];
    }
    if blocks.is_empty() {
        return invalid(format!("no {label} PEM blocks found"));
    }
    Ok(blocks)
}

fn invalid<T>(message: impl Into<String>) -> Result<T> {
    Err(format_error(message))
}

fn format_error(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}
