//! Safe, trust-neutral ODF document-signature cryptography.

use super::model::{DOCUMENT_SIGNATURE_PATH, OdfDigitalSignature, parse_signature_container};
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
pub enum OdfSignatureAlgorithm {
    RsaSha256,
    EcdsaP256Sha256,
}

impl OdfSignatureAlgorithm {
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
pub enum OdfCanonicalizationAlgorithm {
    InclusiveXml10,
    ExclusiveXml10,
}

impl OdfCanonicalizationAlgorithm {
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
pub struct OdfSignatureVerification {
    pub signature_id: Option<String>,
    pub algorithm: Option<OdfSignatureAlgorithm>,
    pub canonicalization: Option<OdfCanonicalizationAlgorithm>,
    pub validity: OdfSignatureValidity,
    pub signing_time: Option<String>,
    pub certificate_chain_der: Vec<Vec<u8>>,
    /// Certificate whose public key verified the signature, if any.
    pub signing_certificate_index: Option<usize>,
    /// URI whose digest failed, when validity is `ReferenceDigestMismatch`.
    pub failed_reference_uri: Option<String>,
}

/// This reports mathematical/package validity only and never PKI trust.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum OdfSignatureValidity {
    Valid,
    ReferenceDigestMismatch,
    InvalidSignature,
    UnsupportedAlgorithm,
    MissingCertificate,
}

/// A validated document signer retaining PKCS#8 key bytes in zeroizing storage.
pub struct OdfDocumentSigner {
    algorithm: OdfSignatureAlgorithm,
    canonicalization: OdfCanonicalizationAlgorithm,
    private_key: Zeroizing<Vec<u8>>,
    certificates: Vec<Vec<u8>>,
    signing_time: String,
}

impl OdfDocumentSigner {
    /// Parse a DER PKCS#8 private key and DER X.509 chain. The leaf certificate must be first.
    pub fn from_pkcs8_der(
        algorithm: OdfSignatureAlgorithm,
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
    pub fn from_pkcs8_pem(
        algorithm: OdfSignatureAlgorithm,
        private_key_pem: &str,
        certificate_chain_pem: &str,
        signing_time: impl Into<String>,
    ) -> Result<Self> {
        let private_key = decode_pem_blocks(private_key_pem, "PRIVATE KEY")?;
        if private_key.len() != 1 {
            return invalid("exactly one PKCS#8 PRIVATE KEY PEM block is required");
        }
        let certificates = decode_pem_blocks(certificate_chain_pem, "CERTIFICATE")?;
        Self::new(
            algorithm,
            private_key.into_iter().next().expect("length checked"),
            certificates,
            signing_time.into(),
        )
    }

    pub fn with_canonicalization(mut self, value: OdfCanonicalizationAlgorithm) -> Self {
        self.canonicalization = value;
        self
    }

    fn new(
        algorithm: OdfSignatureAlgorithm,
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
                .map_err(|_| format_error("invalid DER X.509 certificate"))?;
        }
        validate_key_matches_certificate(algorithm, &private_key, &certificates[0])?;
        Ok(Self {
            algorithm,
            canonicalization: OdfCanonicalizationAlgorithm::InclusiveXml10,
            private_key: Zeroizing::new(private_key),
            certificates,
            signing_time,
        })
    }
}

pub(crate) fn sign_package(unsigned: &[u8], signer: &OdfDocumentSigner) -> Result<Vec<u8>> {
    let archive = ArchiveReader::new(unsigned)
        .map_err(|_| format_error("invalid ODF ZIP archive while signing"))?;
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
        let bytes = archive.read(name).map_err(|_| {
            format_error(format!("unable to read ODF signature reference '{name}'"))
        })?;
        let transform = if is_xml_path(name) && !archive.is_stored(name).unwrap_or(false) {
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
    signed_info.push_str(&format!(
        "<ds:Reference URI=\"#{properties_id}\" Type=\"{SIGNED_PROPERTIES_TYPE}\"><ds:Transforms><ds:Transform Algorithm=\"{}\"></ds:Transform></ds:Transforms><ds:DigestMethod Algorithm=\"{SHA256_URI}\"></ds:DigestMethod><ds:DigestValue>{}</ds:DigestValue></ds:Reference></ds:SignedInfo>",
        signer.canonicalization.uri(),
        BASE64_STANDARD.encode(properties_digest)
    ));

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

    rebuild_with_signature(&archive, &signature_xml)
}

pub(crate) fn verify_package(data: &[u8]) -> Result<Vec<OdfSignatureVerification>> {
    let archive = ArchiveReader::new(data).map_err(|_| format_error("invalid ODF ZIP archive"))?;
    if !archive.contains(DOCUMENT_SIGNATURE_PATH) {
        return Ok(Vec::new());
    }
    let signature_xml = archive
        .read(DOCUMENT_SIGNATURE_PATH)
        .map_err(|_| format_error("unable to read document signature container"))?;
    let signatures = parse_signature_container(&signature_xml)?;
    let text = std::str::from_utf8(&signature_xml)
        .map_err(|_| format_error("signature XML is not UTF-8"))?;
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
    signature: &OdfDigitalSignature,
    signed_info_xml: &str,
) -> Result<OdfSignatureVerification> {
    let algorithm = OdfSignatureAlgorithm::parse(&signature.signature_method);
    let canonicalization = OdfCanonicalizationAlgorithm::parse(&signature.canonicalization_method);
    let certificates = signature
        .x509_certificates
        .iter()
        .map(|value| {
            BASE64_STANDARD
                .decode(value)
                .map_err(|_| format_error("invalid X.509 certificate base64"))
        })
        .collect::<Result<Vec<_>>>()?;
    for value in &certificates {
        Certificate::from_der(value).map_err(|_| format_error("invalid DER X.509 certificate"))?;
    }
    let mut result = OdfSignatureVerification {
        signature_id: signature.id.clone(),
        algorithm,
        canonicalization,
        validity: OdfSignatureValidity::UnsupportedAlgorithm,
        signing_time: signature.signing_time.clone(),
        certificate_chain_der: certificates,
        signing_certificate_index: None,
        failed_reference_uri: None,
    };
    let (Some(algorithm), Some(canonicalization)) = (algorithm, canonicalization) else {
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
            .map_err(|_| format_error("invalid reference digest base64"))?;
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
                canonicalize_fragment(&elements[0], OdfCanonicalizationAlgorithm::InclusiveXml10)?
            } else {
                apply_transforms(elements[0].as_bytes(), &reference.transforms)?
            }
        } else {
            let path = decode_package_uri(&reference.uri)?;
            let bytes = archive
                .read(&path)
                .map_err(|_| format_error("unable to read signature reference"))?;
            apply_transforms(&bytes, &reference.transforms)?
        };
        if Sha256::digest(&bytes).as_slice() != expected_digest.as_slice() {
            result.validity = OdfSignatureValidity::ReferenceDigestMismatch;
            result.failed_reference_uri = Some(reference.uri.clone());
            return Ok(result);
        }
    }
    if result.certificate_chain_der.is_empty() {
        result.validity = OdfSignatureValidity::MissingCertificate;
        return Ok(result);
    }
    let signed_info = canonicalize_fragment(signed_info_xml, canonicalization)?;
    let signature_value = BASE64_STANDARD
        .decode(&signature.signature_value)
        .map_err(|_| format_error("invalid signature value base64"))?;
    for (index, certificate) in result.certificate_chain_der.iter().enumerate() {
        if verify_bytes(algorithm, certificate, &signed_info, &signature_value).unwrap_or(false) {
            result.validity = OdfSignatureValidity::Valid;
            result.signing_certificate_index = Some(index);
            return Ok(result);
        }
    }
    result.validity = OdfSignatureValidity::InvalidSignature;
    Ok(result)
}

fn apply_transforms(bytes: &[u8], transforms: &[String]) -> Result<Vec<u8>> {
    match transforms {
        [] => Ok(bytes.to_vec()),
        [transform] => {
            let mode = OdfCanonicalizationAlgorithm::parse(transform)
                .ok_or_else(|| format_error("unsupported signature reference transform"))?;
            digest_input_canonical_xml(bytes, mode)
        },
        _ => invalid("unsupported signature transform chain"),
    }
}

fn digest_input_canonical_xml(bytes: &[u8], mode: OdfCanonicalizationAlgorithm) -> Result<Vec<u8>> {
    let xml = std::str::from_utf8(bytes).map_err(|_| format_error("C14N input is not UTF-8"))?;
    canonicalize_fragment(xml, mode)
}

fn digest_canonical_xml(xml: &str, mode: OdfCanonicalizationAlgorithm) -> Result<Vec<u8>> {
    Ok(Sha256::digest(canonicalize_fragment(xml, mode)?).to_vec())
}

fn digest_canonical_xml_bytes(xml: &[u8], mode: OdfCanonicalizationAlgorithm) -> Result<Vec<u8>> {
    Ok(Sha256::digest(digest_input_canonical_xml(xml, mode)?).to_vec())
}

fn canonicalize_fragment(xml: &str, mode: OdfCanonicalizationAlgorithm) -> Result<Vec<u8>> {
    canonicalize::<&str>(xml, mode.mode(), None, &[])
        .map_err(|error| format_error(format!("XML canonicalization failed: {error}")))
}

fn sign_bytes(signer: &OdfDocumentSigner, message: &[u8]) -> Result<Vec<u8>> {
    match signer.algorithm {
        OdfSignatureAlgorithm::RsaSha256 => {
            let key = RsaPrivateKey::from_pkcs8_der(&signer.private_key)
                .map_err(|_| format_error("invalid RSA PKCS#8 private key"))?;
            if key.n().bits() < 2048 {
                return invalid("RSA signing keys must be at least 2048 bits");
            }
            let signature: RsaSignature = RsaSigningKey::<RsaSha256>::new(key).sign(message);
            Ok(signature.to_vec())
        },
        OdfSignatureAlgorithm::EcdsaP256Sha256 => {
            let key = EcdsaSigningKey::from_pkcs8_der(&signer.private_key)
                .map_err(|_| format_error("invalid P-256 PKCS#8 private key"))?;
            let signature: EcdsaSignature = key.sign(message);
            Ok(signature.to_bytes().to_vec())
        },
    }
}

fn verify_bytes(
    algorithm: OdfSignatureAlgorithm,
    certificate_der: &[u8],
    message: &[u8],
    signature: &[u8],
) -> Result<bool> {
    let certificate = Certificate::from_der(certificate_der)
        .map_err(|_| format_error("invalid DER X.509 certificate"))?;
    let public = certificate
        .tbs_certificate
        .subject_public_key_info
        .subject_public_key
        .raw_bytes();
    Ok(match algorithm {
        OdfSignatureAlgorithm::RsaSha256 => {
            let key = RsaPublicKey::from_pkcs1_der(public)
                .map_err(|_| format_error("certificate does not contain an RSA public key"))?;
            let signature = RsaSignature::try_from(signature)
                .map_err(|_| format_error("invalid RSA signature encoding"))?;
            RsaVerifyingKey::<RsaSha256>::new(key)
                .verify(message, &signature)
                .is_ok()
        },
        OdfSignatureAlgorithm::EcdsaP256Sha256 => {
            let key = VerifyingKey::from_sec1_bytes(public)
                .map_err(|_| format_error("certificate does not contain a P-256 public key"))?;
            let signature = EcdsaSignature::from_slice(signature)
                .map_err(|_| format_error("invalid ECDSA signature encoding"))?;
            key.verify(message, &signature).is_ok()
        },
    })
}

fn validate_key_matches_certificate(
    algorithm: OdfSignatureAlgorithm,
    private_key: &[u8],
    certificate: &[u8],
) -> Result<()> {
    let certificate = Certificate::from_der(certificate)
        .map_err(|_| format_error("invalid DER X.509 certificate"))?;
    let public = certificate
        .tbs_certificate
        .subject_public_key_info
        .subject_public_key
        .raw_bytes();
    let matches = match algorithm {
        OdfSignatureAlgorithm::RsaSha256 => {
            let private = RsaPrivateKey::from_pkcs8_der(private_key)
                .map_err(|_| format_error("invalid RSA PKCS#8 private key"))?;
            if private.n().bits() < 2048 {
                return invalid("RSA signing keys must be at least 2048 bits");
            }
            let certificate_key = RsaPublicKey::from_pkcs1_der(public)
                .map_err(|_| format_error("leaf certificate is not RSA"))?;
            private.to_public_key() == certificate_key
        },
        OdfSignatureAlgorithm::EcdsaP256Sha256 => {
            let private = EcdsaSigningKey::from_pkcs8_der(private_key)
                .map_err(|_| format_error("invalid P-256 PKCS#8 private key"))?;
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
        let bytes = archive
            .read(name)
            .map_err(|_| format_error("unable to rebuild signed ODF package"))?;
        if name == "mimetype" || archive.is_stored(name).unwrap_or(false) {
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
    name.ends_with(".xml") || name.ends_with(".rdf")
}

fn encode_package_uri(path: &str) -> String {
    let mut result = String::with_capacity(path.len());
    for byte in path.bytes() {
        if byte.is_ascii_alphanumeric() || matches!(byte, b'-' | b'.' | b'_' | b'~' | b'/') {
            result.push(char::from(byte));
        } else {
            result.push_str(&format!("%{byte:02X}"));
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
            let value = u8::from_str_radix(&uri[index + 1..index + 3], 16)
                .map_err(|_| format_error("invalid percent encoding in signature URI"))?;
            decoded.push(value);
            index += 3;
        } else {
            decoded.push(bytes[index]);
            index += 1;
        }
    }
    let path =
        String::from_utf8(decoded).map_err(|_| format_error("signature URI is not valid UTF-8"))?;
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
        let start = reader.buffer_position() as usize;
        match reader
            .read_event()
            .map_err(|_| format_error("invalid signature XML while selecting signed nodes"))?
        {
            Event::Start(element) => {
                let local_matches =
                    local.is_empty() || element.local_name().as_ref() == local.as_bytes();
                let id_matches = if let Some(id) = id {
                    let mut matched = false;
                    for attribute in element.attributes().with_checks(true) {
                        let attribute =
                            attribute.map_err(|_| format_error("invalid signed-node attribute"))?;
                        if matches!(attribute.key.as_ref(), b"Id" | b"ID" | b"id")
                            && attribute.value.as_ref() == id.as_bytes()
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
                        match reader
                            .read_event()
                            .map_err(|_| format_error("invalid signed-node XML"))?
                        {
                            Event::Start(_) => depth += 1,
                            Event::End(_) => depth -= 1,
                            Event::Eof => return invalid("incomplete signed-node XML"),
                            Event::DocType(_) | Event::GeneralRef(_) | Event::PI(_) => {
                                return invalid("active XML is prohibited in signed nodes");
                            },
                            _ => {},
                        }
                    }
                    let end = reader.buffer_position() as usize;
                    let mut fragment = xml[start..end].to_string();
                    inject_known_namespaces(&mut fragment, inherited_default_ds);
                    found.push(fragment);
                }
            },
            Event::DocType(_) | Event::GeneralRef(_) | Event::PI(_) => {
                return invalid("active XML is prohibited in signatures");
            },
            Event::Eof => break,
            _ => {},
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
        declarations.push_str(&format!(" xmlns=\"{DS}\""));
    }
    if fragment.contains("ds:") && !start.contains("xmlns:ds=") {
        declarations.push_str(&format!(" xmlns:ds=\"{DS}\""));
    }
    if fragment.contains("xades:") && !start.contains("xmlns:xades=") {
        declarations.push_str(&format!(" xmlns:xades=\"{XADES}\""));
    }
    if fragment.contains("xd:") && !start.contains("xmlns:xd=") {
        declarations.push_str(&format!(" xmlns:xd=\"{XADES}\""));
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
            .filter(|value| !value.is_whitespace())
            .collect();
        blocks.push(
            BASE64_STANDARD
                .decode(body)
                .map_err(|_| format_error("invalid PEM base64"))?,
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
