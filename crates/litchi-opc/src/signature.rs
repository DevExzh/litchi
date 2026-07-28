//! Trust-neutral verification and authoring support for OPC digital signatures.

use crate::{OpcPackage, PackURI, Part, Relationships, TargetMode};
use base64::Engine;
use quick_xml::{
    Reader,
    events::{BytesStart, Event},
};
use p256::ecdsa::{
    Signature as EcdsaSignature, SigningKey as EcdsaSigningKey,
    VerifyingKey as EcdsaVerifyingKey,
};
use rsa::{
    BigUint, RsaPrivateKey, RsaPublicKey,
    pkcs1v15::{
        Signature as RsaSignature, SigningKey as RsaSigningKey, VerifyingKey,
    },
    pkcs8::{DecodePrivateKey, DecodePublicKey},
    traits::PublicKeyParts,
};
use sha1_legacy::Sha1;
use sha2_legacy::{Digest, Sha256, Sha384, Sha512};
use signature::{SignatureEncoding, Signer, Verifier};
use std::{
    collections::{BTreeMap, HashMap, HashSet},
    fmt,
    str,
};
use subtle::ConstantTimeEq;
use thiserror::Error;
use zeroize::Zeroizing;

const DS: &str = "http://www.w3.org/2000/09/xmldsig#";
const MDSSI: &str = "http://schemas.openxmlformats.org/package/2006/digital-signature";
const OFFICE_DIGSIG: &str = "http://schemas.microsoft.com/office/2006/digsig";
const REL_NS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const ORIGIN_REL: &str =
    "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/origin";
const SIGNATURE_REL: &str =
    "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/signature";
const CERTIFICATE_REL: &str =
    "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/certificate";
const REL_TRANSFORM: &str = "http://schemas.openxmlformats.org/package/2006/RelationshipTransform";
const C14N: &str = "http://www.w3.org/TR/2001/REC-xml-c14n-20010315";
const C14N_COMMENTS: &str = "http://www.w3.org/TR/2001/REC-xml-c14n-20010315#WithComments";
const EXCLUSIVE_C14N: &str = "http://www.w3.org/2001/10/xml-exc-c14n#";
const EXCLUSIVE_C14N_COMMENTS: &str = "http://www.w3.org/2001/10/xml-exc-c14n#WithComments";
const SHA1: &str = "http://www.w3.org/2000/09/xmldsig#sha1";
const SHA256: &str = "http://www.w3.org/2001/04/xmlenc#sha256";
const SHA384: &str = "http://www.w3.org/2001/04/xmldsig-more#sha384";
const SHA512: &str = "http://www.w3.org/2001/04/xmlenc#sha512";
const RSA_SHA1: &str = "http://www.w3.org/2000/09/xmldsig#rsa-sha1";
const RSA_SHA256: &str = "http://www.w3.org/2001/04/xmldsig-more#rsa-sha256";
const RSA_SHA384: &str = "http://www.w3.org/2001/04/xmldsig-more#rsa-sha384";
const RSA_SHA512: &str = "http://www.w3.org/2001/04/xmldsig-more#rsa-sha512";
const ECDSA_SHA256: &str = "http://www.w3.org/2001/04/xmldsig-more#ecdsa-sha256";
const DSIG11: &str = "http://www.w3.org/2009/xmldsig11#";
const P256_CURVE: &str = "urn:oid:1.2.840.10045.3.1.7";
const XML_NS: &str = "http://www.w3.org/XML/1998/namespace";
const ORIGIN_CONTENT_TYPE: &str = "application/vnd.openxmlformats-package.digital-signature-origin";
const SIGNATURE_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-package.digital-signature-xmlsignature+xml";
const CERTIFICATE_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-package.digital-signature-certificate";

pub type Result<T> = std::result::Result<T, DigitalSignatureError>;

#[derive(Debug, Error)]
pub enum DigitalSignatureError {
    #[error("invalid OPC digital-signature graph: {0}")]
    InvalidGraph(String),
    #[error("invalid or unsafe signature XML: {0}")]
    InvalidXml(String),
    #[error("digital-signature resource limit exceeded: {0}")]
    LimitExceeded(String),
    #[error("unsupported digital-signature algorithm or transform: {0}")]
    UnsupportedAlgorithm(String),
    #[error("SHA-1 is disallowed by the strict verification policy")]
    Sha1Disallowed,
    #[error("invalid RSA verification key: {0}")]
    InvalidKey(String),
    #[error("digital-signature authoring failed: {0}")]
    Signing(String),
}

/// Signature algorithm used for newly authored signatures.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SignatureAlgorithm {
    RsaSha256,
    EcdsaP256Sha256,
}

/// Canonicalization emitted for newly authored signature XML.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum CanonicalizationMethod {
    Inclusive,
    Exclusive,
}

#[derive(Clone)]
enum SigningMaterial {
    Rsa(Box<RsaPrivateKey>),
    Ecdsa(EcdsaSigningKey),
}

/// Private signing material and trust-neutral metadata for an OPC signature.
#[derive(Clone)]
pub struct PackageSigner {
    material: SigningMaterial,
    certificates: Vec<Vec<u8>>,
    signing_time: Option<String>,
    canonicalization: CanonicalizationMethod,
}

impl std::fmt::Debug for PackageSigner {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.debug_struct("PackageSigner")
            .field("algorithm", &self.algorithm())
            .field("certificate_count", &self.certificates.len())
            .field("signing_time", &self.signing_time)
            .field("canonicalization", &self.canonicalization)
            .finish_non_exhaustive()
    }
}

impl PackageSigner {
    /// Decode RSA signing material from a buffer that is wiped when this call
    /// returns. The decoded key remains scoped to the signer and its upstream
    /// cryptographic type controls destruction of the key object.
    pub fn rsa_sha256_pkcs8(key_der: Zeroizing<Vec<u8>>) -> Result<Self> {
        let key = RsaPrivateKey::from_pkcs8_der(&key_der)
            .map_err(|error| DigitalSignatureError::InvalidKey(error.to_string()))?;
        Self::rsa_sha256(key)
    }

    /// Decode a P-256 signing key from a buffer that is wiped when this call
    /// returns.
    pub fn ecdsa_p256_sha256_pkcs8(key_der: Zeroizing<Vec<u8>>) -> Result<Self> {
        let key = EcdsaSigningKey::from_pkcs8_der(&key_der)
            .map_err(|error| DigitalSignatureError::InvalidKey(error.to_string()))?;
        Ok(Self::ecdsa_p256_sha256(key))
    }

    pub fn rsa_sha256(key: RsaPrivateKey) -> Result<Self> {
        if key.n().bits() < 2048 {
            return Err(DigitalSignatureError::InvalidKey(
                "RSA signing keys must be at least 2048 bits".into(),
            ));
        }
        key.validate()
            .map_err(|error| DigitalSignatureError::InvalidKey(error.to_string()))?;
        Ok(Self {
            material: SigningMaterial::Rsa(Box::new(key)),
            certificates: Vec::new(),
            signing_time: None,
            canonicalization: CanonicalizationMethod::Inclusive,
        })
    }

    pub fn ecdsa_p256_sha256(key: EcdsaSigningKey) -> Self {
        Self {
            material: SigningMaterial::Ecdsa(key),
            certificates: Vec::new(),
            signing_time: None,
            canonicalization: CanonicalizationMethod::Inclusive,
        }
    }

    pub fn algorithm(&self) -> SignatureAlgorithm {
        match &self.material {
            SigningMaterial::Rsa(_) => SignatureAlgorithm::RsaSha256,
            SigningMaterial::Ecdsa(_) => SignatureAlgorithm::EcdsaP256Sha256,
        }
    }

    pub fn set_certificates(&mut self, certificates: Vec<Vec<u8>>) -> Result<&mut Self> {
        if certificates.iter().any(Vec::is_empty) {
            return Err(DigitalSignatureError::InvalidKey(
                "certificate DER must not be empty".into(),
            ));
        }
        if let Some(first) = certificates.first() {
            self.validate_leaf_certificate(first)?;
        }
        self.certificates = certificates;
        Ok(self)
    }

    pub fn set_signing_time(&mut self, signing_time: Option<&str>) -> Result<&mut Self> {
        if let Some(value) = signing_time {
            validate_signing_time(value)?;
            self.signing_time = Some(value.to_string());
        } else {
            self.signing_time = None;
        }
        Ok(self)
    }

    pub fn set_canonicalization(&mut self, method: CanonicalizationMethod) -> &mut Self {
        self.canonicalization = method;
        self
    }

    pub fn certificates(&self) -> &[Vec<u8>] {
        &self.certificates
    }

    pub fn signing_time(&self) -> Option<&str> {
        self.signing_time.as_deref()
    }

    fn validate_leaf_certificate(&self, certificate: &[u8]) -> Result<()> {
        let subject_public_key = spki(certificate)?;
        match &self.material {
            SigningMaterial::Rsa(key) => {
                let certificate_key = RsaPublicKey::from_public_key_der(subject_public_key)
                    .map_err(|error| DigitalSignatureError::InvalidKey(error.to_string()))?;
                if certificate_key != key.to_public_key() {
                    return Err(DigitalSignatureError::InvalidKey(
                        "leaf certificate does not match the RSA signing key".into(),
                    ));
                }
            },
            SigningMaterial::Ecdsa(key) => {
                let certificate_key = EcdsaVerifyingKey::from_public_key_der(subject_public_key)
                    .map_err(|error| DigitalSignatureError::InvalidKey(error.to_string()))?;
                if certificate_key != *key.verifying_key() {
                    return Err(DigitalSignatureError::InvalidKey(
                        "leaf certificate does not match the P-256 signing key".into(),
                    ));
                }
            },
        }
        Ok(())
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Sha1Policy {
    AllowWithWarning,
    Reject,
}

#[derive(Debug, Clone)]
pub struct SignatureVerificationPolicy {
    pub sha1: Sha1Policy,
    pub max_signature_part_bytes: usize,
    pub max_xml_depth: usize,
    pub max_xml_elements: usize,
    pub max_attributes_per_element: usize,
    pub max_references: usize,
    pub max_embedded_certificate_bytes: usize,
    pub max_certificates: usize,
    pub max_total_certificate_bytes: usize,
    pub max_rsa_modulus_bits: usize,
}
impl SignatureVerificationPolicy {
    pub fn compatibility() -> Self {
        Self {
            sha1: Sha1Policy::AllowWithWarning,
            max_signature_part_bytes: 8 * 1024 * 1024,
            max_xml_depth: 128,
            max_xml_elements: 100_000,
            max_attributes_per_element: 256,
            max_references: 20_000,
            max_embedded_certificate_bytes: 1024 * 1024,
            max_certificates: 256,
            max_total_certificate_bytes: 16 * 1024 * 1024,
            max_rsa_modulus_bits: 16_384,
        }
    }
    pub fn strict() -> Self {
        Self {
            sha1: Sha1Policy::Reject,
            ..Self::compatibility()
        }
    }
}
impl Default for SignatureVerificationPolicy {
    fn default() -> Self {
        Self::compatibility()
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum VerificationStatus {
    Valid,
    Invalid,
}
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum CertificateTrust {
    NotEvaluated,
}
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct EmbeddedCertificate {
    pub der: Vec<u8>,
}
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ReferenceVerification {
    pub uri: String,
    pub status: VerificationStatus,
}
#[derive(Debug, Clone)]
pub struct DigitalSignatureVerification {
    pub signature_part: PackURI,
    pub package_integrity: VerificationStatus,
    pub signature_value: VerificationStatus,
    pub certificate_trust: CertificateTrust,
    pub references: Vec<ReferenceVerification>,
    pub certificates: Vec<EmbeddedCertificate>,
    pub uses_sha1: bool,
    pub signing_time: Option<String>,
}

/// A caller-resolved byte sequence covered by a detached XMLDSig Manifest.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DetachedSignatureReference {
    pub uri: String,
    pub data: Vec<u8>,
}

/// Trust-neutral result for an XMLDSig document whose references are supplied
/// by its container implementation.
#[derive(Debug, Clone)]
pub struct DetachedDigitalSignatureVerification {
    pub package_integrity: VerificationStatus,
    pub signature_value: VerificationStatus,
    pub certificate_trust: CertificateTrust,
    pub references: Vec<ReferenceVerification>,
    pub certificates: Vec<EmbeddedCertificate>,
    pub uses_sha1: bool,
    pub signing_time: Option<String>,
}

/// Author an Office-compatible detached XMLDSig document.
pub fn author_detached_signature(
    signer: &PackageSigner,
    references: &[DetachedSignatureReference],
) -> Result<Vec<u8>> {
    if references.is_empty() {
        return Err(DigitalSignatureError::Signing(
            "a detached signature must cover at least one reference".into(),
        ));
    }
    let mut seen = HashSet::new();
    let mut ordered = references.to_vec();
    ordered.sort_by(|left, right| left.uri.cmp(&right.uri));
    if ordered.iter().any(|reference| {
        reference.uri.is_empty()
            || reference.uri.contains(['\"', '<', '>', '&'])
            || !seen.insert(reference.uri.clone())
    }) {
        return Err(DigitalSignatureError::Signing(
            "detached reference URIs must be unique XML-safe strings".into(),
        ));
    }
    let manifest = ordered
        .iter()
        .map(|reference| {
            format!(
                "<Reference URI=\"{}\"><DigestMethod Algorithm=\"{SHA256}\"></DigestMethod><DigestValue>{}</DigestValue></Reference>",
                reference.uri,
                encode64(&Sha256::digest(&reference.data))
            )
        })
        .collect::<String>();
    let signing_time = signer.signing_time.as_deref().unwrap_or("1970-01-01T00:00:00Z");
    let package_object = format!(
        "<Object xmlns=\"{DS}\" xmlns:mdssi=\"{MDSSI}\" Id=\"idPackageObject\"><Manifest>{manifest}</Manifest><SignatureProperties><SignatureProperty Id=\"idSignatureTime\" Target=\"#idPackageSignature\"><mdssi:SignatureTime><mdssi:Format>YYYY-MM-DDThh:mm:ssTZD</mdssi:Format><mdssi:Value>{}</mdssi:Value></mdssi:SignatureTime></SignatureProperty></SignatureProperties></Object>",
        escape_text_string(signing_time)
    );
    let office_object = format!(
        "<Object xmlns=\"{DS}\" Id=\"idOfficeObject\"><SignatureProperties><SignatureProperty Id=\"idOfficeSignatureInfo\" Target=\"#idPackageSignature\"><SignatureInfoV1 xmlns=\"{OFFICE_DIGSIG}\"><SetupID></SetupID><SignatureText></SignatureText><SignatureImage></SignatureImage><SignatureComments></SignatureComments><WindowsVersion>1</WindowsVersion><OfficeVersion>12.0</OfficeVersion><ApplicationVersion>12.0</ApplicationVersion><Monitors>1</Monitors><HorizontalResolution>96</HorizontalResolution><VerticalResolution>96</VerticalResolution><ColorDepth>24</ColorDepth><SignatureProviderId></SignatureProviderId><SignatureProviderUrl></SignatureProviderUrl><SignatureProviderDetails>0</SignatureProviderDetails><SignatureType>1</SignatureType><ManifestHashAlgorithm>{SHA256}</ManifestHashAlgorithm></SignatureInfoV1></SignatureProperty></SignatureProperties></Object>"
    );
    let mode = c14n_mode(C14N)?;
    let package_digest = Sha256::digest(canonicalize_authored(&package_object, mode)?);
    let office_digest = Sha256::digest(canonicalize_authored(&office_object, mode)?);
    let signature_algorithm = match signer.algorithm() {
        SignatureAlgorithm::RsaSha256 => RSA_SHA256,
        SignatureAlgorithm::EcdsaP256Sha256 => ECDSA_SHA256,
    };
    let signed_info = format!(
        "<SignedInfo xmlns=\"{DS}\"><CanonicalizationMethod Algorithm=\"{C14N}\"></CanonicalizationMethod><SignatureMethod Algorithm=\"{signature_algorithm}\"></SignatureMethod><Reference URI=\"#idPackageObject\"><Transforms><Transform Algorithm=\"{C14N}\"></Transform></Transforms><DigestMethod Algorithm=\"{SHA256}\"></DigestMethod><DigestValue>{}</DigestValue></Reference><Reference URI=\"#idOfficeObject\"><Transforms><Transform Algorithm=\"{C14N}\"></Transform></Transforms><DigestMethod Algorithm=\"{SHA256}\"></DigestMethod><DigestValue>{}</DigestValue></Reference></SignedInfo>",
        encode64(&package_digest),
        encode64(&office_digest)
    );
    let canonical = canonicalize_authored(&signed_info, mode)?;
    let value = match &signer.material {
        SigningMaterial::Rsa(key) => {
            let signature: RsaSignature = RsaSigningKey::<Sha256>::new(key.as_ref().clone()).sign(&canonical);
            signature.to_vec()
        },
        SigningMaterial::Ecdsa(key) => {
            let signature: EcdsaSignature = key.sign(&canonical);
            signature.to_bytes().to_vec()
        },
    };
    Ok(format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\"?><Signature xmlns=\"{DS}\" Id=\"idPackageSignature\">{signed_info}<SignatureValue>{}</SignatureValue>{}{package_object}{office_object}</Signature>",
        encode64(&value),
        build_key_info(signer)
    )
    .into_bytes())
}

/// Verify an Office-compatible detached XMLDSig document against exactly the
/// references supplied by the container implementation.
pub fn verify_detached_signature(
    xml: &[u8],
    references: &[DetachedSignatureReference],
    policy: &SignatureVerificationPolicy,
) -> Result<DetachedDigitalSignatureVerification> {
    validate_policy(policy)?;
    if xml.len() > policy.max_signature_part_bytes {
        return Err(DigitalSignatureError::LimitExceeded(
            "detached signature XML is too large".into(),
        ));
    }
    let document = Document::parse(xml, policy)?;
    if !document.is(document.root, DS, "Signature") {
        return Err(DigitalSignatureError::InvalidXml(
            "root element must be ds:Signature".into(),
        ));
    }
    let signed_info = document.required_child(document.root, DS, "SignedInfo")?;
    let canonicalization = document.required_child(signed_info, DS, "CanonicalizationMethod")?;
    let mode = c14n_mode(document.attr(canonicalization, "Algorithm")?)?;
    let signed_bytes = document.canonicalize(signed_info, mode);
    let signature_method = document.required_child(signed_info, DS, "SignatureMethod")?;
    let signature_algorithm = document.attr(signature_method, "Algorithm")?;
    sha1_policy(signature_algorithm, policy)?;
    if !matches!(signature_algorithm, RSA_SHA1 | RSA_SHA256 | ECDSA_SHA256) {
        return Err(DigitalSignatureError::UnsupportedAlgorithm(signature_algorithm.into()));
    }
    let manifests = document.descendants(document.root, DS, "Manifest");
    if manifests.len() != 1 {
        return Err(DigitalSignatureError::InvalidXml(
            "detached signature must contain exactly one Manifest".into(),
        ));
    }
    let mut expected = HashMap::new();
    for reference in references {
        if expected.insert(reference.uri.as_str(), reference.data.as_slice()).is_some() {
            return Err(DigitalSignatureError::InvalidGraph(format!(
                "duplicate expected reference {}",
                reference.uri
            )));
        }
    }
    let manifest_references = document.children(manifests[0], DS, "Reference");
    if manifest_references.len() != expected.len()
        || manifest_references.len() > policy.max_references
    {
        return Err(DigitalSignatureError::InvalidGraph(
            "Manifest does not cover the exact detached reference set".into(),
        ));
    }
    let mut reports = Vec::new();
    let mut uses_sha1 = signature_algorithm == RSA_SHA1;
    let mut seen = HashSet::new();
    for reference in manifest_references {
        let uri = document.attr(reference, "URI")?.to_string();
        if !seen.insert(uri.clone()) {
            return Err(DigitalSignatureError::InvalidXml(format!(
                "duplicate Manifest reference {uri}"
            )));
        }
        let data = expected.get(uri.as_str()).ok_or_else(|| {
            DigitalSignatureError::InvalidGraph(format!("unexpected detached reference {uri}"))
        })?;
        let (report, weak) = verify_detached_reference(&document, reference, data, policy)?;
        uses_sha1 |= weak;
        reports.push(report);
    }
    let signed_references = document.children(signed_info, DS, "Reference");
    if signed_references.len() != 2 {
        return Err(DigitalSignatureError::InvalidXml(
            "SignedInfo must reference idPackageObject and idOfficeObject exactly once".into(),
        ));
    }
    for reference in signed_references {
        let uri = document.attr(reference, "URI")?;
        if !matches!(uri, "#idPackageObject" | "#idOfficeObject")
            || !seen.insert(uri.to_string())
        {
            return Err(DigitalSignatureError::InvalidXml(format!(
                "invalid or duplicate SignedInfo reference {uri}"
            )));
        }
        let data = dereference_detached_fragment(&document, reference)?;
        let (report, weak) = verify_detached_reference(&document, reference, &data, policy)?;
        uses_sha1 |= weak;
        reports.push(report);
    }
    let mut certificates = Vec::new();
    for node in document.descendants(document.root, DS, "X509Certificate") {
        let der = decode64(
            &document.text(node)?,
            policy.max_embedded_certificate_bytes,
            "X509Certificate",
        )?;
        push_certificate(&mut certificates, &der, policy)?;
    }
    let encoded_value = document.text(document.required_child(document.root, DS, "SignatureValue")?)?;
    let value = decode64(&encoded_value, policy.max_signature_part_bytes, "SignatureValue")?;
    let signature_valid = match signature_algorithm {
        RSA_SHA1 | RSA_SHA256 => {
            let key = extract_key(&document, &certificates, policy)?;
            let signature = RsaSignature::try_from(value.as_slice())
                .map_err(|error| DigitalSignatureError::InvalidKey(error.to_string()))?;
            if signature_algorithm == RSA_SHA1 {
                VerifyingKey::<Sha1>::new(key).verify(&signed_bytes, &signature).is_ok()
            } else {
                VerifyingKey::<Sha256>::new(key).verify(&signed_bytes, &signature).is_ok()
            }
        },
        ECDSA_SHA256 => {
            let key = extract_ec_key(&document, &certificates)?;
            let signature = EcdsaSignature::from_slice(&value)
                .map_err(|error| DigitalSignatureError::InvalidKey(error.to_string()))?;
            key.verify(&signed_bytes, &signature).is_ok()
        },
        _ => unreachable!(),
    };
    let integrity = reports
        .iter()
        .all(|reference| reference.status == VerificationStatus::Valid);
    Ok(DetachedDigitalSignatureVerification {
        package_integrity: status(integrity),
        signature_value: status(signature_valid),
        certificate_trust: CertificateTrust::NotEvaluated,
        references: reports,
        certificates,
        uses_sha1,
        signing_time: extract_signing_time(&document)?,
    })
}

fn dereference_detached_fragment(document: &Document, reference: usize) -> Result<Vec<u8>> {
    let uri = document.attr(reference, "URI")?;
    let id = uri.strip_prefix('#').ok_or_else(|| {
        DigitalSignatureError::InvalidXml(format!("invalid fragment reference {uri}"))
    })?;
    let node = *document
        .ids
        .get(id)
        .ok_or_else(|| DigitalSignatureError::InvalidXml(format!("unknown Id {id}")))?;
    let transforms = transforms(document, reference)?;
    if transforms.len() > 1 || matches!(transforms.first(), Some(Transform::Relationships(_))) {
        return Err(DigitalSignatureError::UnsupportedAlgorithm(
            "invalid detached fragment transform chain".into(),
        ));
    }
    let mode = match transforms.first() {
        Some(Transform::Canonical(mode)) => *mode,
        None => c14n_mode(C14N)?,
        _ => unreachable!(),
    };
    Ok(document.canonicalize(node, mode))
}

fn verify_detached_reference(
    document: &Document,
    reference: usize,
    data: &[u8],
    policy: &SignatureVerificationPolicy,
) -> Result<(ReferenceVerification, bool)> {
    let uri = document.attr(reference, "URI")?.to_string();
    let method = document.required_child(reference, DS, "DigestMethod")?;
    let algorithm = document.attr(method, "Algorithm")?;
    sha1_policy(algorithm, policy)?;
    let expected = decode64(
        &document.text(document.required_child(reference, DS, "DigestValue")?)?,
        128,
        "DigestValue",
    )?;
    let actual = match algorithm {
        SHA1 => Sha1::digest(data).to_vec(),
        SHA256 => Sha256::digest(data).to_vec(),
        _ => return Err(DigitalSignatureError::UnsupportedAlgorithm(algorithm.into())),
    };
    Ok((
        ReferenceVerification {
            uri,
            status: status(actual.len() == expected.len() && bool::from(actual.ct_eq(&expected))),
        },
        algorithm == SHA1,
    ))
}

pub fn verify_package(
    package: &OpcPackage,
    policy: &SignatureVerificationPolicy,
) -> Result<Vec<DigitalSignatureVerification>> {
    validate_policy(policy)?;
    let origins: Vec<_> = package
        .rels()
        .iter()
        .filter(|r| r.reltype() == ORIGIN_REL)
        .collect();
    if origins.is_empty() {
        return Ok(vec![]);
    }
    if origins.len() != 1 {
        return Err(DigitalSignatureError::InvalidGraph(format!(
            "expected one signature origin relationship, found {}",
            origins.len()
        )));
    }
    internal(origins[0].target_mode(), "signature origin")?;
    let origin_uri = origins[0].target_partname().map_err(graph)?;
    if !origin_uri.as_str().starts_with("/_xmlsignatures/") {
        return Err(DigitalSignatureError::InvalidGraph(
            "signature origin must be under /_xmlsignatures/".into(),
        ));
    }
    let origin = package.get_part(&origin_uri).map_err(graph)?;
    content_type(origin, ORIGIN_CONTENT_TYPE, "signature origin")?;
    if origin
        .rels()
        .iter()
        .any(|relationship| relationship.reltype() != SIGNATURE_REL)
    {
        return Err(DigitalSignatureError::InvalidGraph(
            "signature origin has an unexpected relationship".into(),
        ));
    }
    let mut uris = Vec::new();
    let mut seen = HashSet::new();
    let mut reachable = HashSet::from([origin_uri.clone()]);
    for rel in origin
        .rels()
        .iter()
        .filter(|r| r.reltype() == SIGNATURE_REL)
    {
        internal(rel.target_mode(), "signature")?;
        let uri = rel.target_partname().map_err(graph)?;
        if !uri.as_str().starts_with("/_xmlsignatures/") {
            return Err(DigitalSignatureError::InvalidGraph(
                "signature part must be under /_xmlsignatures/".into(),
            ));
        }
        if !seen.insert(uri.clone()) {
            return Err(DigitalSignatureError::InvalidGraph(format!(
                "duplicate signature target {}",
                uri.as_str()
            )));
        }
        let part = package.get_part(&uri).map_err(graph)?;
        content_type(part, SIGNATURE_CONTENT_TYPE, "signature")?;
        reachable.insert(uri.clone());
        if part
            .rels()
            .iter()
            .any(|relationship| relationship.reltype() != CERTIFICATE_REL)
        {
            return Err(DigitalSignatureError::InvalidGraph(
                "signature part has an unexpected relationship".into(),
            ));
        }
        for cert_rel in part
            .rels()
            .iter()
            .filter(|r| r.reltype() == CERTIFICATE_REL)
        {
            internal(cert_rel.target_mode(), "certificate")?;
            let cert = package
                .get_part(&cert_rel.target_partname().map_err(graph)?)
                .map_err(graph)?;
            content_type(cert, CERTIFICATE_CONTENT_TYPE, "certificate")?;
            reachable.insert(cert.partname().clone());
        }
        uris.push(uri);
    }
    if uris.is_empty() {
        return Err(DigitalSignatureError::InvalidGraph(
            "signature origin has no signature relationships".into(),
        ));
    }
    for part in package.iter_parts().filter(|part| is_signature_infrastructure(*part)) {
        if !reachable.contains(part.partname()) {
            return Err(DigitalSignatureError::InvalidGraph(format!(
                "orphan or spoofed signature infrastructure part {}",
                part.partname().as_str()
            )));
        }
    }
    uris.sort_by(|a, b| a.as_str().cmp(b.as_str()));
    uris.into_iter()
        .map(|u| verify_one(package, u, policy))
        .collect()
}
fn graph<E: std::fmt::Display>(e: E) -> DigitalSignatureError {
    DigitalSignatureError::InvalidGraph(e.to_string())
}
fn internal(mode: TargetMode, what: &str) -> Result<()> {
    if mode == TargetMode::External {
        Err(DigitalSignatureError::InvalidGraph(format!(
            "{what} relationship must be internal"
        )))
    } else {
        Ok(())
    }
}
fn content_type(p: &dyn Part, expected: &str, what: &str) -> Result<()> {
    if p.content_type() != expected {
        Err(DigitalSignatureError::InvalidGraph(format!(
            "{what} part {} has content type {}, expected {expected}",
            p.partname().as_str(),
            p.content_type()
        )))
    } else {
        Ok(())
    }
}
fn validate_policy(p: &SignatureVerificationPolicy) -> Result<()> {
    if p.max_signature_part_bytes == 0
        || p.max_xml_depth == 0
        || p.max_xml_elements == 0
        || p.max_attributes_per_element == 0
        || p.max_references == 0
        || p.max_embedded_certificate_bytes == 0
        || p.max_certificates == 0
        || p.max_total_certificate_bytes == 0
        || p.max_rsa_modulus_bits < 512
    {
        Err(DigitalSignatureError::LimitExceeded(
            "invalid zero or undersized policy limit".into(),
        ))
    } else {
        Ok(())
    }
}

fn verify_one(
    package: &OpcPackage,
    uri: PackURI,
    policy: &SignatureVerificationPolicy,
) -> Result<DigitalSignatureVerification> {
    let part = package.get_part(&uri).map_err(graph)?;
    if part.blob().len() > policy.max_signature_part_bytes {
        return Err(DigitalSignatureError::LimitExceeded(format!(
            "signature part {} is too large",
            uri.as_str()
        )));
    }
    let doc = Document::parse(part.blob(), policy)?;
    if !doc.is(doc.root, DS, "Signature") {
        return Err(DigitalSignatureError::InvalidXml(
            "root element must be ds:Signature".into(),
        ));
    }
    let signed = doc.required_child(doc.root, DS, "SignedInfo")?;
    let canon = doc.required_child(signed, DS, "CanonicalizationMethod")?;
    let canonicalization = c14n_mode(doc.attr(canon, "Algorithm")?)?;
    let signed_bytes = doc.canonicalize(signed, canonicalization);
    let method = doc.required_child(signed, DS, "SignatureMethod")?;
    let sig_alg = doc.attr(method, "Algorithm")?;
    sha1_policy(sig_alg, policy)?;
    if !matches!(
        sig_alg,
        RSA_SHA1 | RSA_SHA256 | RSA_SHA384 | RSA_SHA512 | ECDSA_SHA256
    ) {
        return Err(DigitalSignatureError::UnsupportedAlgorithm(sig_alg.into()));
    }
    let manifests = doc.descendants(doc.root, DS, "Manifest");
    if manifests.len() != 1 {
        return Err(DigitalSignatureError::InvalidXml(
            "an OPC signature must contain exactly one package Manifest".into(),
        ));
    }
    let manifest_references = doc.children(manifests[0], DS, "Reference");
    validate_manifest_coverage(package, &doc, &manifest_references)?;
    let mut refs = doc.children(signed, DS, "Reference");
    refs.extend(manifest_references);
    if refs.is_empty() {
        return Err(DigitalSignatureError::InvalidXml(
            "no signature references".into(),
        ));
    }
    if refs.len() > policy.max_references {
        return Err(DigitalSignatureError::LimitExceeded(format!(
            "{} references exceed policy",
            refs.len()
        )));
    }
    let mut reports = Vec::with_capacity(refs.len());
    let mut uses_sha1 = sig_alg == RSA_SHA1;
    let mut reference_uris = HashSet::new();
    for r in refs {
        let uri = doc.attr(r, "URI")?;
        if !reference_uris.insert(uri.to_string()) {
            return Err(DigitalSignatureError::InvalidXml(format!(
                "duplicate signature reference {uri}"
            )));
        }
        let (report, weak) = verify_reference(package, &doc, r, policy)?;
        uses_sha1 |= weak;
        reports.push(report);
    }
    let certs = extract_certificates(package, part, &doc, policy)?;
    let value = decode64(
        &doc.text(doc.required_child(doc.root, DS, "SignatureValue")?)?,
        policy.max_signature_part_bytes,
        "SignatureValue",
    )?;
    let sig_ok = match sig_alg {
        RSA_SHA1 | RSA_SHA256 | RSA_SHA384 | RSA_SHA512 => {
            let key = extract_key(&doc, &certs, policy)?;
            let signature = RsaSignature::try_from(value.as_slice())
                .map_err(|error| DigitalSignatureError::InvalidKey(error.to_string()))?;
            match sig_alg {
                RSA_SHA1 => VerifyingKey::<Sha1>::new(key).verify(&signed_bytes, &signature),
                RSA_SHA256 => VerifyingKey::<Sha256>::new(key).verify(&signed_bytes, &signature),
                RSA_SHA384 => VerifyingKey::<Sha384>::new(key).verify(&signed_bytes, &signature),
                RSA_SHA512 => VerifyingKey::<Sha512>::new(key).verify(&signed_bytes, &signature),
                _ => unreachable!(),
            }
        },
        ECDSA_SHA256 => {
            let key = extract_ec_key(&doc, &certs)?;
            let signature = EcdsaSignature::from_slice(&value)
                .map_err(|error| DigitalSignatureError::InvalidKey(error.to_string()))?;
            key.verify(&signed_bytes, &signature)
        },
        _ => unreachable!(),
    }
    .is_ok();
    let integrity = reports
        .iter()
        .all(|r| r.status == VerificationStatus::Valid);
    Ok(DigitalSignatureVerification {
        signature_part: uri,
        package_integrity: status(integrity),
        signature_value: status(sig_ok),
        certificate_trust: CertificateTrust::NotEvaluated,
        references: reports,
        certificates: certs,
        uses_sha1,
        signing_time: extract_signing_time(&doc)?,
    })
}

fn validate_manifest_coverage(
    package: &OpcPackage,
    document: &Document,
    references: &[usize],
) -> Result<()> {
    let mut expected: HashMap<String, Option<Vec<String>>> = HashMap::new();
    let mut parts: Vec<&dyn Part> = package
        .iter_parts()
        .filter(|part| !is_signature_infrastructure(*part))
        .collect();
    for part in &parts {
        expected.insert(
            format!(
                "{}?ContentType={}",
                part.partname().as_str(),
                part.content_type()
            ),
            None,
        );
    }
    insert_expected_relationships(&mut expected, "/_rels/.rels", package.rels());
    for part in parts.drain(..) {
        insert_expected_relationships(
            &mut expected,
            &relationship_part_path(part.partname().as_str())?,
            part.rels(),
        );
    }

    let mut seen = HashSet::new();
    for reference in references {
        let uri = document.attr(*reference, "URI")?;
        if !seen.insert(uri.to_string()) {
            return Err(DigitalSignatureError::InvalidXml(format!(
                "duplicate Manifest reference {uri}"
            )));
        }
        let relationship_ids = expected.remove(uri).ok_or_else(|| {
            DigitalSignatureError::InvalidGraph(format!(
                "Manifest contains an unexpected or incorrectly typed reference {uri}"
            ))
        })?;
        let transforms = transforms(document, *reference)?;
        if let Some(expected_ids) = relationship_ids {
            match transforms.as_slice() {
                [Transform::Relationships(actual_ids)]
                    if actual_ids.iter().all(|id| expected_ids.contains(id)) => {},
                [Transform::Relationships(actual_ids), Transform::Canonical(_)]
                    if actual_ids.iter().all(|id| expected_ids.contains(id)) => {},
                _ => {
                    return Err(DigitalSignatureError::InvalidGraph(format!(
                        "RelationshipTransform for {uri} selects an unknown relationship or has an invalid transform chain"
                    )));
                },
            }
        } else if !matches!(transforms.as_slice(), [] | [Transform::Canonical(_)]) {
            return Err(DigitalSignatureError::InvalidGraph(format!(
                "invalid transform chain for package part {uri}"
            )));
        }
    }
    Ok(())
}

fn insert_expected_relationships(
    expected: &mut HashMap<String, Option<Vec<String>>>,
    path: &str,
    relationships: &Relationships,
) {
    let mut ids: Vec<String> = relationships
        .iter()
        .filter(|relationship| {
            !matches!(
                relationship.reltype(),
                ORIGIN_REL | SIGNATURE_REL | CERTIFICATE_REL
            )
        })
        .map(|relationship| relationship.r_id().to_string())
        .collect();
    if ids.is_empty() {
        return;
    }
    ids.sort();
    expected.insert(
        format!("{path}?ContentType=application/vnd.openxmlformats-package.relationships+xml"),
        Some(ids),
    );
}
fn status(v: bool) -> VerificationStatus {
    if v {
        VerificationStatus::Valid
    } else {
        VerificationStatus::Invalid
    }
}
fn sha1_policy(a: &str, p: &SignatureVerificationPolicy) -> Result<()> {
    if matches!(a, SHA1 | RSA_SHA1) && p.sha1 == Sha1Policy::Reject {
        Err(DigitalSignatureError::Sha1Disallowed)
    } else {
        Ok(())
    }
}
#[derive(Debug, Clone, Copy)]
struct CanonMode {
    comments: bool,
    exclusive: bool,
}

fn c14n_mode(a: &str) -> Result<CanonMode> {
    match a {
        C14N => Ok(CanonMode {
            comments: false,
            exclusive: false,
        }),
        C14N_COMMENTS => Ok(CanonMode {
            comments: true,
            exclusive: false,
        }),
        EXCLUSIVE_C14N => Ok(CanonMode {
            comments: false,
            exclusive: true,
        }),
        EXCLUSIVE_C14N_COMMENTS => Ok(CanonMode {
            comments: true,
            exclusive: true,
        }),
        _ => Err(DigitalSignatureError::UnsupportedAlgorithm(a.into())),
    }
}

#[derive(Debug)]
enum Transform {
    Relationships(Vec<String>),
    Canonical(CanonMode),
}
fn transforms(doc: &Document, r: usize) -> Result<Vec<Transform>> {
    let Some(ts) = doc.child(r, DS, "Transforms") else {
        return Ok(vec![]);
    };
    let mut out = Vec::new();
    for t in doc.children(ts, DS, "Transform") {
        let alg = doc.attr(t, "Algorithm")?;
        match alg {
            REL_TRANSFORM => {
                if !out.is_empty() {
                    return Err(DigitalSignatureError::UnsupportedAlgorithm(
                        "RelationshipTransform must be first".into(),
                    ));
                }
                let mut ids = Vec::new();
                for n in doc.children(t, MDSSI, "RelationshipReference") {
                    ids.push(doc.attr(n, "SourceId")?.into())
                }
                if ids.is_empty() {
                    return Err(DigitalSignatureError::InvalidXml(
                        "empty RelationshipTransform".into(),
                    ));
                }
                let count = ids.len();
                ids.sort();
                ids.dedup();
                if ids.len() != count {
                    return Err(DigitalSignatureError::InvalidXml(
                        "duplicate RelationshipReference".into(),
                    ));
                }
                out.push(Transform::Relationships(ids));
            },
            C14N | C14N_COMMENTS | EXCLUSIVE_C14N | EXCLUSIVE_C14N_COMMENTS => {
                out.push(Transform::Canonical(c14n_mode(alg)?))
            },
            _ => return Err(DigitalSignatureError::UnsupportedAlgorithm(alg.into())),
        }
    }
    Ok(out)
}
fn verify_reference(
    package: &OpcPackage,
    doc: &Document,
    r: usize,
    p: &SignatureVerificationPolicy,
) -> Result<(ReferenceVerification, bool)> {
    let uri = doc.attr(r, "URI")?.to_string();
    let dm = doc.required_child(r, DS, "DigestMethod")?;
    let alg = doc.attr(dm, "Algorithm")?;
    sha1_policy(alg, p)?;
    let expected = decode64(
        &doc.text(doc.required_child(r, DS, "DigestValue")?)?,
        128,
        "DigestValue",
    )?;
    let data = dereference(package, doc, &uri, &transforms(doc, r)?, p)?;
    let actual = match alg {
        SHA1 => Sha1::digest(&data).to_vec(),
        SHA256 => Sha256::digest(&data).to_vec(),
        SHA384 => Sha384::digest(&data).to_vec(),
        SHA512 => Sha512::digest(&data).to_vec(),
        _ => return Err(DigitalSignatureError::UnsupportedAlgorithm(alg.into())),
    };
    let valid = actual.len() == expected.len() && bool::from(actual.ct_eq(&expected));
    Ok((
        ReferenceVerification {
            uri,
            status: status(valid),
        },
        alg == SHA1,
    ))
}
fn dereference(
    package: &OpcPackage,
    doc: &Document,
    uri: &str,
    ts: &[Transform],
    p: &SignatureVerificationPolicy,
) -> Result<Vec<u8>> {
    if let Some(id) = uri.strip_prefix('#') {
        let n = *doc
            .ids
            .get(id)
            .ok_or_else(|| DigitalSignatureError::InvalidXml(format!("unknown Id {id}")))?;
        let mut canonicalization = CanonMode {
            comments: false,
            exclusive: false,
        };
        for t in ts {
            match t {
                Transform::Canonical(mode) => canonicalization = *mode,
                Transform::Relationships(_) => {
                    return Err(DigitalSignatureError::UnsupportedAlgorithm(
                        "relationship transform on fragment".into(),
                    ));
                },
            }
        }
        return Ok(doc.canonicalize(n, canonicalization));
    }
    if !uri.starts_with('/') || uri.contains('#') {
        return Err(DigitalSignatureError::InvalidXml(format!(
            "invalid package reference {uri}"
        )));
    }
    let (path, q) = uri.split_once('?').ok_or_else(|| {
        DigitalSignatureError::InvalidXml(format!("reference lacks ContentType query: {uri}"))
    })?;
    let lower_path = path.to_ascii_lowercase();
    if path.contains('\\')
        || path.contains("//")
        || path.split('/').any(|segment| matches!(segment, "." | ".."))
        || lower_path.contains("%2f")
        || lower_path.contains("%5c")
        || lower_path.contains("%2e")
    {
        return Err(DigitalSignatureError::InvalidXml(format!(
            "unsafe package reference path {path}"
        )));
    }
    let ct = q
        .strip_prefix("ContentType=")
        .filter(|s| !s.is_empty() && !s.contains('&'))
        .ok_or_else(|| {
            DigitalSignatureError::InvalidXml(format!("invalid ContentType query: {uri}"))
        })?;
    if let Some(ids) = ts.iter().find_map(|t| {
        if let Transform::Relationships(v) = t {
            Some(v)
        } else {
            None
        }
    }) {
        if ct != "application/vnd.openxmlformats-package.relationships+xml"
            || !path.ends_with(".rels")
        {
            return Err(DigitalSignatureError::InvalidXml(
                "invalid relationship transform target".into(),
            ));
        }
        return canonical_relationships(relationships(package, path)?, ids);
    }
    let pu = PackURI::new(path.to_string()).map_err(DigitalSignatureError::InvalidXml)?;
    let part = package.get_part(&pu).map_err(graph)?;
    if part.content_type() != ct {
        return Err(DigitalSignatureError::InvalidGraph(format!(
            "reference content type mismatch for {path}"
        )));
    }
    if ts.is_empty() {
        return Ok(part.blob().to_vec());
    }
    if ts.len() != 1 {
        return Err(DigitalSignatureError::UnsupportedAlgorithm(
            "invalid part transforms".into(),
        ));
    }
    let Transform::Canonical(canonicalization) = ts[0] else {
        unreachable!()
    };
    let parsed = Document::parse(part.blob(), p)?;
    Ok(parsed.canonicalize(parsed.root, canonicalization))
}
fn relationships<'a>(p: &'a OpcPackage, path: &str) -> Result<&'a Relationships> {
    if path == "/_rels/.rels" {
        return Ok(p.rels());
    }
    let (dir, file) = path.rsplit_once("/_rels/").ok_or_else(|| {
        DigitalSignatureError::InvalidXml(format!("invalid relationship URI {path}"))
    })?;
    let source = file
        .strip_suffix(".rels")
        .filter(|s| !s.is_empty())
        .ok_or_else(|| {
            DigitalSignatureError::InvalidXml(format!("invalid relationship URI {path}"))
        })?;
    let uri = PackURI::new(format!("{dir}/{source}")).map_err(DigitalSignatureError::InvalidXml)?;
    Ok(p.get_part(&uri).map_err(graph)?.rels())
}
fn canonical_relationships(rels: &Relationships, ids: &[String]) -> Result<Vec<u8>> {
    let mut chosen = Vec::new();
    for id in ids {
        chosen.push(rels.get(id).ok_or_else(|| {
            DigitalSignatureError::InvalidGraph(format!("relationship {id} not found"))
        })?)
    }
    chosen.sort_by(|a, b| a.r_id().cmp(b.r_id()));
    let mut o = Vec::new();
    o.extend_from_slice(b"<Relationships xmlns=\"");
    attr_escape(&mut o, REL_NS);
    o.extend_from_slice(b"\">");
    for r in chosen {
        o.extend_from_slice(b"<Relationship Id=\"");
        attr_escape(&mut o, r.r_id());
        o.extend_from_slice(b"\" Target=\"");
        attr_escape(&mut o, r.target_ref());
        o.extend_from_slice(b"\" TargetMode=\"");
        o.extend_from_slice(if r.target_mode() == TargetMode::Internal {
            b"Internal"
        } else {
            b"External"
        });
        o.extend_from_slice(b"\" Type=\"");
        attr_escape(&mut o, r.reltype());
        o.extend_from_slice(b"\"></Relationship>");
    }
    o.extend_from_slice(b"</Relationships>");
    Ok(o)
}

fn extract_certificates(
    package: &OpcPackage,
    signature_part: &dyn Part,
    d: &Document,
    p: &SignatureVerificationPolicy,
) -> Result<Vec<EmbeddedCertificate>> {
    let mut certificates = Vec::new();
    for node in d.descendants(d.root, DS, "X509Certificate") {
        let der = decode64(
            &d.text(node)?,
            p.max_embedded_certificate_bytes,
            "X509Certificate",
        )?;
        push_certificate(&mut certificates, &der, p)?;
    }
    for relationship in signature_part
        .rels()
        .iter()
        .filter(|relationship| relationship.reltype() == CERTIFICATE_REL)
    {
        internal(relationship.target_mode(), "certificate")?;
        let uri = relationship.target_partname().map_err(graph)?;
        let certificate_part = package.get_part(&uri).map_err(graph)?;
        content_type(certificate_part, CERTIFICATE_CONTENT_TYPE, "certificate")?;
        push_certificate(&mut certificates, certificate_part.blob(), p)?;
    }
    Ok(certificates)
}

fn push_certificate(
    certificates: &mut Vec<EmbeddedCertificate>,
    der: &[u8],
    policy: &SignatureVerificationPolicy,
) -> Result<()> {
    if der.len() > policy.max_embedded_certificate_bytes {
        return Err(DigitalSignatureError::LimitExceeded(format!(
            "embedded certificate is {} bytes; limit is {}",
            der.len(),
            policy.max_embedded_certificate_bytes
        )));
    }
    if certificates
        .iter()
        .any(|certificate| certificate.der == der)
    {
        return Ok(());
    }
    if certificates.len() >= policy.max_certificates {
        return Err(DigitalSignatureError::LimitExceeded(format!(
            "certificate count exceeds {}",
            policy.max_certificates
        )));
    }
    let total_bytes = certificates
        .iter()
        .try_fold(der.len(), |total, certificate| {
            total.checked_add(certificate.der.len())
        })
        .ok_or_else(|| {
            DigitalSignatureError::LimitExceeded("certificate byte count overflow".into())
        })?;
    if total_bytes > policy.max_total_certificate_bytes {
        return Err(DigitalSignatureError::LimitExceeded(format!(
            "certificate data is {total_bytes} bytes; limit is {}",
            policy.max_total_certificate_bytes
        )));
    }
    certificates.push(EmbeddedCertificate { der: der.to_vec() });
    Ok(())
}
fn extract_key(
    d: &Document,
    certs: &[EmbeddedCertificate],
    p: &SignatureVerificationPolicy,
) -> Result<RsaPublicKey> {
    if let Some(k) = d.descendants(d.root, DS, "RSAKeyValue").first().copied() {
        let m = decode64(
            &d.text(d.required_child(k, DS, "Modulus")?)?,
            p.max_rsa_modulus_bits.div_ceil(8),
            "modulus",
        )?;
        let e = decode64(
            &d.text(d.required_child(k, DS, "Exponent")?)?,
            16,
            "exponent",
        )?;
        let bits = m.len() * 8 - m.first().map_or(0, |b| b.leading_zeros() as usize);
        if !(512..=p.max_rsa_modulus_bits).contains(&bits) {
            return Err(DigitalSignatureError::InvalidKey(format!(
                "RSA modulus is {bits} bits"
            )));
        }
        return RsaPublicKey::new(BigUint::from_bytes_be(&m), BigUint::from_bytes_be(&e))
            .map_err(|e| DigitalSignatureError::InvalidKey(e.to_string()));
    }
    let cert = certs
        .first()
        .ok_or_else(|| DigitalSignatureError::InvalidKey("no RSA key or certificate".into()))?;
    let key = RsaPublicKey::from_public_key_der(spki(&cert.der)?)
        .map_err(|e| DigitalSignatureError::InvalidKey(e.to_string()))?;
    if key.size() * 8 > p.max_rsa_modulus_bits {
        return Err(DigitalSignatureError::InvalidKey(
            "RSA key too large".into(),
        ));
    }
    Ok(key)
}

fn extract_ec_key(d: &Document, certs: &[EmbeddedCertificate]) -> Result<EcdsaVerifyingKey> {
    if let Some(key_value) = d.descendants(d.root, DSIG11, "ECKeyValue").first().copied() {
        let curve = d.required_child(key_value, DSIG11, "NamedCurve")?;
        if d.attr(curve, "URI")? != P256_CURVE {
            return Err(DigitalSignatureError::UnsupportedAlgorithm(
                "only the P-256 XMLDSIG curve is supported".into(),
            ));
        }
        let public_key = decode64(
            &d.text(d.required_child(key_value, DSIG11, "PublicKey")?)?,
            65,
            "P-256 public key",
        )?;
        return EcdsaVerifyingKey::from_sec1_bytes(&public_key)
            .map_err(|error| DigitalSignatureError::InvalidKey(error.to_string()));
    }
    let certificate = certs
        .first()
        .ok_or_else(|| DigitalSignatureError::InvalidKey("no P-256 key or certificate".into()))?;
    EcdsaVerifyingKey::from_public_key_der(spki(&certificate.der)?)
        .map_err(|error| DigitalSignatureError::InvalidKey(error.to_string()))
}

fn extract_signing_time(d: &Document) -> Result<Option<String>> {
    let values = d.descendants(d.root, MDSSI, "Value");
    if values.is_empty() {
        return Ok(None);
    }
    if values.len() != 1 {
        return Err(DigitalSignatureError::InvalidXml(
            "signature contains duplicate signing times".into(),
        ));
    }
    let value = d.text(values[0])?;
    validate_signing_time(&value)?;
    Ok(Some(value))
}

fn validate_signing_time(value: &str) -> Result<()> {
    let bytes = value.as_bytes();
    let structural = bytes.len() >= 20
        && bytes.get(4) == Some(&b'-')
        && bytes.get(7) == Some(&b'-')
        && bytes.get(10) == Some(&b'T')
        && bytes.get(13) == Some(&b':')
        && bytes.get(16) == Some(&b':')
        && (value.ends_with('Z')
            || (bytes.len() >= 25
                && matches!(bytes.get(bytes.len() - 6), Some(b'+') | Some(b'-'))
                && bytes.get(bytes.len() - 3) == Some(&b':')));
    if !structural || !bytes.iter().enumerate().all(|(index, byte)| {
        byte.is_ascii_digit()
            || matches!(index, 4 | 7 | 10 | 13 | 16)
            || matches!(*byte, b'Z' | b'+' | b'-' | b':' | b'.')
    }) {
        return Err(DigitalSignatureError::InvalidXml(
            "signing time must be an ISO 8601 date-time with a timezone".into(),
        ));
    }
    Ok(())
}

/// Add a signature while preserving existing signatures.
pub fn add_package_signature(
    package: &mut OpcPackage,
    signer: &PackageSigner,
) -> Result<PackURI> {
    let existing = verify_package(package, &SignatureVerificationPolicy::compatibility())?;
    if existing.iter().any(|verification| {
        verification.package_integrity != VerificationStatus::Valid
            || verification.signature_value != VerificationStatus::Valid
    }) {
        return Err(DigitalSignatureError::Signing(
            "cannot add a signature while an existing signature is invalid".into(),
        ));
    }
    let signature_uri = next_signature_uri(package)?;
    let signature_xml = build_signature(package, signer)?;
    install_signature(package, signature_uri, signature_xml)
}

/// Replace all signatures with one new signature.
pub fn resign_package(package: &mut OpcPackage, signer: &PackageSigner) -> Result<PackURI> {
    let signature_uri = PackURI::new("/_xmlsignatures/sig1.xml".to_string())
        .map_err(DigitalSignatureError::Signing)?;
    let signature_xml = build_signature(package, signer)?;
    clear_package_signatures(package)?;
    install_signature(package, signature_uri, signature_xml)
}

fn next_signature_uri(package: &OpcPackage) -> Result<PackURI> {
    for index in 1..=u32::MAX {
        let uri = PackURI::new(format!("/_xmlsignatures/sig{index}.xml"))
            .map_err(DigitalSignatureError::Signing)?;
        if !package.contains_part(&uri) {
            return Ok(uri);
        }
    }
    Err(DigitalSignatureError::Signing(
        "no signature part name is available".into(),
    ))
}

fn install_signature(
    package: &mut OpcPackage,
    signature_uri: PackURI,
    signature_xml: Vec<u8>,
) -> Result<PackURI> {
    let (origin_uri, has_origin_relationship) = {
        let origins: Vec<_> = package
            .rels()
            .iter()
            .filter(|relationship| relationship.reltype() == ORIGIN_REL)
            .collect();
        match origins.as_slice() {
            [] => (
                PackURI::new("/_xmlsignatures/origin.sigs".to_string())
                    .map_err(DigitalSignatureError::Signing)?,
                false,
            ),
            [origin] if origin.target_mode() == TargetMode::Internal => {
                (origin.target_partname().map_err(graph)?, true)
            },
            _ => {
                return Err(DigitalSignatureError::InvalidGraph(
                    "signature origin relationship is ambiguous or external".into(),
                ));
            },
        }
    };
    if package.contains_part(&signature_uri) {
        return Err(DigitalSignatureError::Signing(format!(
            "signature part {} already exists",
            signature_uri.as_str()
        )));
    }
    if !has_origin_relationship && package.contains_part(&origin_uri) {
        return Err(DigitalSignatureError::InvalidGraph(
            "orphan signature-origin part occupies the conventional origin URI".into(),
        ));
    }
    if has_origin_relationship {
        let origin = package.get_part(&origin_uri).map_err(graph)?;
        content_type(origin, ORIGIN_CONTENT_TYPE, "signature origin")?;
        if origin
            .rels()
            .iter()
            .any(|relationship| relationship.reltype() != SIGNATURE_REL)
        {
            return Err(DigitalSignatureError::InvalidGraph(
                "signature origin has a non-signature relationship".into(),
            ));
        }
    }
    let needs_origin = !has_origin_relationship;
    package
        .try_add_part(Box::new(crate::BlobPart::new(
            signature_uri.clone(),
            SIGNATURE_CONTENT_TYPE.to_string(),
            signature_xml,
        )))
        .map_err(graph)?;
    if needs_origin {
        package
            .try_add_part(Box::new(crate::BlobPart::new(
                origin_uri.clone(),
                ORIGIN_CONTENT_TYPE.to_string(),
                Vec::new(),
            )))
            .map_err(graph)?;
        package.relate_to(origin_uri.as_str(), ORIGIN_REL);
    }
    package
        .get_part_mut(&origin_uri)
        .map_err(graph)?
        .relate_to(signature_uri.as_str(), SIGNATURE_REL);
    Ok(signature_uri)
}

fn build_signature(package: &OpcPackage, signer: &PackageSigner) -> Result<Vec<u8>> {
    let canonicalization_uri = match signer.canonicalization {
        CanonicalizationMethod::Inclusive => C14N,
        CanonicalizationMethod::Exclusive => EXCLUSIVE_C14N,
    };
    let manifest = build_manifest(package)?;
    let signing_time = signer.signing_time.as_deref().unwrap_or("1970-01-01T00:00:00Z");
    let object = format!(
        "<Object xmlns=\"{DS}\" xmlns:mdssi=\"{MDSSI}\" Id=\"idPackageObject\"><Manifest>{manifest}</Manifest><SignatureProperties><SignatureProperty Id=\"idSignatureTime\" Target=\"#idPackageSignature\"><mdssi:SignatureTime><mdssi:Format>YYYY-MM-DDThh:mm:ssTZD</mdssi:Format><mdssi:Value>{}</mdssi:Value></mdssi:SignatureTime></SignatureProperty></SignatureProperties></Object>",
        escape_text_string(signing_time)
    );
    let mode = c14n_mode(canonicalization_uri)?;
    let object_digest = Sha256::digest(canonicalize_authored(&object, mode)?);
    let signed_info = format!(
        "<SignedInfo xmlns=\"{DS}\"><CanonicalizationMethod Algorithm=\"{canonicalization_uri}\"></CanonicalizationMethod><SignatureMethod Algorithm=\"{}\"></SignatureMethod><Reference URI=\"#idPackageObject\"><Transforms><Transform Algorithm=\"{canonicalization_uri}\"></Transform></Transforms><DigestMethod Algorithm=\"{SHA256}\"></DigestMethod><DigestValue>{}</DigestValue></Reference></SignedInfo>",
        match signer.algorithm() {
            SignatureAlgorithm::RsaSha256 => RSA_SHA256,
            SignatureAlgorithm::EcdsaP256Sha256 => ECDSA_SHA256,
        },
        encode64(&object_digest)
    );
    let signed_bytes = canonicalize_authored(&signed_info, mode)?;
    let signature_value = match &signer.material {
        SigningMaterial::Rsa(key) => {
            let signature: RsaSignature = RsaSigningKey::<Sha256>::new(key.as_ref().clone()).sign(&signed_bytes);
            signature.to_vec()
        },
        SigningMaterial::Ecdsa(key) => {
            let signature: EcdsaSignature = key.sign(&signed_bytes);
            signature.to_bytes().to_vec()
        },
    };
    let key_info = build_key_info(signer);
    Ok(format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\"?><Signature xmlns=\"{DS}\" Id=\"idPackageSignature\">{signed_info}<SignatureValue>{}</SignatureValue>{key_info}{object}</Signature>",
        encode64(&signature_value)
    )
    .into_bytes())
}

fn build_key_info(signer: &PackageSigner) -> String {
    let key_value = match &signer.material {
        SigningMaterial::Rsa(key) => {
            let public = key.to_public_key();
            format!(
                "<KeyValue><RSAKeyValue><Modulus>{}</Modulus><Exponent>{}</Exponent></RSAKeyValue></KeyValue>",
                encode64(&public.n().to_bytes_be()),
                encode64(&public.e().to_bytes_be())
            )
        },
        SigningMaterial::Ecdsa(key) => format!(
            "<KeyValue><dsig11:ECKeyValue xmlns:dsig11=\"{DSIG11}\"><dsig11:NamedCurve URI=\"{P256_CURVE}\"></dsig11:NamedCurve><dsig11:PublicKey>{}</dsig11:PublicKey></dsig11:ECKeyValue></KeyValue>",
            encode64(key.verifying_key().to_encoded_point(false).as_bytes())
        ),
    };
    if signer.certificates.is_empty() {
        return format!("<KeyInfo>{key_value}</KeyInfo>");
    }
    let certificates = signer
        .certificates
        .iter()
        .map(|certificate| format!("<X509Certificate>{}</X509Certificate>", encode64(certificate)))
        .collect::<String>();
    format!("<KeyInfo>{key_value}<X509Data>{certificates}</X509Data></KeyInfo>")
}

fn build_manifest(package: &OpcPackage) -> Result<String> {
    let mut references = Vec::new();
    let mut parts: Vec<&dyn Part> = package
        .iter_parts()
        .filter(|part| !is_signature_infrastructure(*part))
        .collect();
    parts.sort_by(|left, right| left.partname().as_str().cmp(right.partname().as_str()));
    for part in &parts {
        let uri = format!(
            "{}?ContentType={}",
            part.partname().as_str(),
            part.content_type()
        );
        references.push((
            uri,
            format!(
                "<DigestMethod Algorithm=\"{SHA256}\"></DigestMethod><DigestValue>{}</DigestValue>",
                encode64(&Sha256::digest(part.blob()))
            ),
        ));
    }
    push_relationship_reference(&mut references, "/_rels/.rels", package.rels())?;
    for part in parts {
        let path = relationship_part_path(part.partname().as_str())?;
        push_relationship_reference(&mut references, &path, part.rels())?;
    }
    references.sort_by(|left, right| left.0.cmp(&right.0));
    Ok(references
        .into_iter()
        .map(|(uri, body)| {
            format!(
                "<Reference URI=\"{}\">{body}</Reference>",
                escape_attribute_string(&uri)
            )
        })
        .collect())
}

fn push_relationship_reference(
    references: &mut Vec<(String, String)>,
    path: &str,
    relationships: &Relationships,
) -> Result<()> {
    let mut ids: Vec<String> = relationships
        .iter()
        .filter(|relationship| {
            !matches!(
                relationship.reltype(),
                ORIGIN_REL | SIGNATURE_REL | CERTIFICATE_REL
            )
        })
        .map(|relationship| relationship.r_id().to_string())
        .collect();
    if ids.is_empty() {
        return Ok(());
    }
    ids.sort();
    let digest = Sha256::digest(canonical_relationships(relationships, &ids)?);
    let relationship_parameters = ids
        .iter()
        .map(|id| {
            format!(
                "<mdssi:RelationshipReference SourceId=\"{}\"></mdssi:RelationshipReference>",
                escape_attribute_string(id)
            )
        })
        .collect::<String>();
    references.push((
        format!("{path}?ContentType=application/vnd.openxmlformats-package.relationships+xml"),
        format!(
            "<Transforms><Transform Algorithm=\"{REL_TRANSFORM}\" xmlns:mdssi=\"{MDSSI}\">{relationship_parameters}</Transform><Transform Algorithm=\"{C14N}\"></Transform></Transforms><DigestMethod Algorithm=\"{SHA256}\"></DigestMethod><DigestValue>{}</DigestValue>",
            encode64(&digest)
        ),
    ));
    Ok(())
}

fn relationship_part_path(source: &str) -> Result<String> {
    let (directory, file) = source.rsplit_once('/').ok_or_else(|| {
        DigitalSignatureError::Signing(format!("invalid relationship source {source}"))
    })?;
    if file.is_empty() {
        return Err(DigitalSignatureError::Signing(format!(
            "invalid relationship source {source}"
        )));
    }
    Ok(format!("{directory}/_rels/{file}.rels"))
}

fn is_signature_infrastructure(part: &dyn Part) -> bool {
    part.partname().as_str().starts_with("/_xmlsignatures/")
        || matches!(
            part.content_type(),
            ORIGIN_CONTENT_TYPE | SIGNATURE_CONTENT_TYPE | CERTIFICATE_CONTENT_TYPE
        )
}

fn canonicalize_authored(xml: &str, mode: CanonMode) -> Result<Vec<u8>> {
    let mut policy = SignatureVerificationPolicy::compatibility();
    policy.max_signature_part_bytes = xml.len().saturating_add(1);
    let document = Document::parse(xml.as_bytes(), &policy)?;
    Ok(document.canonicalize(document.root, mode))
}

fn encode64(bytes: &[u8]) -> String {
    base64::engine::general_purpose::STANDARD.encode(bytes)
}

fn escape_attribute_string(value: &str) -> String {
    let mut output = Vec::new();
    attr_escape(&mut output, value);
    String::from_utf8(output).expect("XML escaping preserves UTF-8")
}

fn escape_text_string(value: &str) -> String {
    let mut output = Vec::new();
    text_escape(&mut output, value);
    String::from_utf8(output).expect("XML escaping preserves UTF-8")
}

/// Remove signature relationships and all signature infrastructure parts.
pub fn clear_package_signatures(package: &mut OpcPackage) -> Result<()> {
    let relationship_ids: Vec<String> = package
        .rels()
        .iter()
        .filter(|relationship| relationship.reltype() == ORIGIN_REL)
        .map(|relationship| relationship.r_id().to_string())
        .collect();
    for relationship_id in relationship_ids {
        package.rels_mut().remove(&relationship_id);
    }
    let part_names: Vec<PackURI> = package
        .iter_parts()
        .filter(|part| {
            part.partname().as_str().starts_with("/_xmlsignatures/")
                || matches!(
                    part.content_type(),
                    ORIGIN_CONTENT_TYPE | SIGNATURE_CONTENT_TYPE | CERTIFICATE_CONTENT_TYPE
                )
        })
        .map(|part| part.partname().clone())
        .collect();
    for part_name in part_names {
        package.remove_part(&part_name);
    }
    Ok(())
}
fn spki(cert: &[u8]) -> Result<&[u8]> {
    let (tag, outer, end) = tlv(cert, 0)?;
    if tag != 0x30 || end != cert.len() {
        return Err(DigitalSignatureError::InvalidKey(
            "invalid DER certificate".into(),
        ));
    }
    let (tag, tbs, _) = tlv(outer, 0)?;
    if tag != 0x30 {
        return Err(DigitalSignatureError::InvalidKey(
            "invalid TBSCertificate".into(),
        ));
    }
    let mut pos = 0;
    if tbs.first() == Some(&0xa0) {
        pos = tlv(tbs, pos)?.2
    }
    for _ in 0..5 {
        pos = tlv(tbs, pos)?.2
    }
    let start = pos;
    let (tag, _, end) = tlv(tbs, pos)?;
    if tag != 0x30 {
        return Err(DigitalSignatureError::InvalidKey(
            "invalid SubjectPublicKeyInfo".into(),
        ));
    }
    Ok(&tbs[start..end])
}
fn tlv(d: &[u8], p: usize) -> Result<(u8, &[u8], usize)> {
    let tag = *d
        .get(p)
        .ok_or_else(|| DigitalSignatureError::InvalidKey("truncated DER".into()))?;
    let b = *d
        .get(p + 1)
        .ok_or_else(|| DigitalSignatureError::InvalidKey("truncated DER".into()))?;
    let (len, h) = if b & 128 == 0 {
        (b as usize, 2)
    } else {
        let n = (b & 127) as usize;
        if n == 0 || n > std::mem::size_of::<usize>() {
            return Err(DigitalSignatureError::InvalidKey(
                "invalid DER length".into(),
            ));
        }
        let mut l = 0usize;
        for x in d
            .get(p + 2..p + 2 + n)
            .ok_or_else(|| DigitalSignatureError::InvalidKey("truncated DER".into()))?
        {
            l = l
                .checked_mul(256)
                .and_then(|v| v.checked_add(*x as usize))
                .ok_or_else(|| DigitalSignatureError::InvalidKey("DER overflow".into()))?
        }
        if l < 128 {
            return Err(DigitalSignatureError::InvalidKey("non-minimal DER".into()));
        }
        (l, 2 + n)
    };
    let s = p
        .checked_add(h)
        .ok_or_else(|| DigitalSignatureError::InvalidKey("DER overflow".into()))?;
    let e = s
        .checked_add(len)
        .ok_or_else(|| DigitalSignatureError::InvalidKey("DER overflow".into()))?;
    Ok((
        tag,
        d.get(s..e)
            .ok_or_else(|| DigitalSignatureError::InvalidKey("truncated DER".into()))?,
        e,
    ))
}
fn decode64(s: &str, max: usize, what: &str) -> Result<Vec<u8>> {
    let compact: String = s.chars().filter(|c| !c.is_whitespace()).collect();
    if compact.len() > max.saturating_mul(4) / 3 + 8 {
        return Err(DigitalSignatureError::LimitExceeded(format!(
            "{what} too large"
        )));
    }
    let v = base64::engine::general_purpose::STANDARD
        .decode(compact)
        .map_err(|e| DigitalSignatureError::InvalidXml(e.to_string()))?;
    if v.len() > max {
        Err(DigitalSignatureError::LimitExceeded(format!(
            "{what} too large"
        )))
    } else {
        Ok(v)
    }
}

#[derive(Clone)]
struct Name {
    q: String,
    local: String,
    ns: String,
}
#[derive(Clone)]
struct Attribute {
    name: Name,
    value: String,
}
#[derive(Clone)]
enum Child {
    Element(usize),
    Text(String),
    Comment(String),
}
#[derive(Clone)]
struct Element {
    name: Name,
    attrs: Vec<Attribute>,
    namespaces: BTreeMap<String, String>,
    children: Vec<Child>,
}
struct Document {
    elements: Vec<Element>,
    root: usize,
    ids: HashMap<String, usize>,
}
impl Document {
    fn parse(bytes: &[u8], p: &SignatureVerificationPolicy) -> Result<Self> {
        if bytes.len() > p.max_signature_part_bytes {
            return Err(DigitalSignatureError::LimitExceeded("XML too large".into()));
        }
        let mut r = Reader::from_reader(bytes);
        r.config_mut().trim_text(false);
        let (mut elements, mut stack, mut ids, mut root, mut buf) =
            (Vec::new(), Vec::new(), HashMap::new(), None, Vec::new());
        loop {
            match r.read_event_into(&mut buf) {
                Ok(Event::Start(s)) => Self::start(
                    &s,
                    r.decoder(),
                    p,
                    &mut elements,
                    &mut stack,
                    &mut ids,
                    &mut root,
                )?,
                Ok(Event::Empty(s)) => {
                    Self::start(
                        &s,
                        r.decoder(),
                        p,
                        &mut elements,
                        &mut stack,
                        &mut ids,
                        &mut root,
                    )?;
                    stack.pop();
                },
                Ok(Event::End(_)) => {
                    if stack.pop().is_none() {
                        return Err(DigitalSignatureError::InvalidXml(
                            "unexpected end tag".into(),
                        ));
                    }
                },
                Ok(Event::Text(t)) => {
                    let raw = t.xml10_content().map_err(xml)?;
                    let text = quick_xml::escape::unescape(&raw).map_err(xml)?.into_owned();
                    Self::push_text(&mut elements, &stack, text)?
                },
                Ok(Event::CData(t)) => Self::push_text(
                    &mut elements,
                    &stack,
                    t.xml10_content().map_err(xml)?.into_owned(),
                )?,
                Ok(Event::Comment(c)) => {
                    if let Some(&n) = stack.last() {
                        elements[n]
                            .children
                            .push(Child::Comment(c.xml10_content().map_err(xml)?.into_owned()))
                    }
                },
                Ok(Event::Decl(_)) => {
                    if root.is_some() {
                        return Err(DigitalSignatureError::InvalidXml(
                            "late XML declaration".into(),
                        ));
                    }
                },
                Ok(Event::DocType(_) | Event::PI(_) | Event::GeneralRef(_)) => {
                    return Err(DigitalSignatureError::InvalidXml(
                        "DTD, PI, and entity references are rejected".into(),
                    ));
                },
                Ok(Event::Eof) => break,
                Err(e) => return Err(DigitalSignatureError::InvalidXml(e.to_string())),
            };
            buf.clear()
        }
        if !stack.is_empty() {
            return Err(DigitalSignatureError::InvalidXml("unclosed element".into()));
        }
        Ok(Self {
            elements,
            root: root.ok_or_else(|| DigitalSignatureError::InvalidXml("no root".into()))?,
            ids,
        })
    }
    #[allow(clippy::too_many_arguments)]
    fn start(
        s: &BytesStart<'_>,
        decoder: quick_xml::encoding::Decoder,
        p: &SignatureVerificationPolicy,
        es: &mut Vec<Element>,
        stack: &mut Vec<usize>,
        ids: &mut HashMap<String, usize>,
        root: &mut Option<usize>,
    ) -> Result<()> {
        if stack.len() >= p.max_xml_depth || es.len() >= p.max_xml_elements {
            return Err(DigitalSignatureError::LimitExceeded(
                "XML structure limit".into(),
            ));
        }
        let q = str::from_utf8(s.name().as_ref()).map_err(xml)?.to_string();
        let mut raw = Vec::new();
        for a in s.attributes().with_checks(true) {
            let a = a.map_err(xml)?;
            if raw.len() >= p.max_attributes_per_element {
                return Err(DigitalSignatureError::LimitExceeded(
                    "attribute limit".into(),
                ));
            }
            raw.push((
                str::from_utf8(a.key.as_ref()).map_err(xml)?.to_string(),
                a.decoded_and_normalized_value(quick_xml::XmlVersion::Explicit1_0, decoder)
                    .map_err(xml)?
                    .into_owned(),
            ));
        }
        let mut ns = stack.last().map_or_else(
            || {
                let mut m = BTreeMap::new();
                m.insert("xml".into(), XML_NS.into());
                m
            },
            |n| es[*n].namespaces.clone(),
        );
        for (a, v) in &raw {
            if a == "xmlns" {
                ns.insert("".into(), v.clone());
            } else if let Some(pre) = a.strip_prefix("xmlns:") {
                if pre.is_empty()
                    || pre == "xmlns"
                    || (pre == "xml" && v != XML_NS)
                    || (pre != "xml" && v == XML_NS)
                    || v.is_empty()
                {
                    return Err(DigitalSignatureError::InvalidXml(format!(
                        "invalid namespace {a}"
                    )));
                }
                ns.insert(pre.into(), v.clone());
            }
        }
        let name = expanded(&q, &ns, true)?;
        let mut attrs = Vec::new();
        let mut unique = HashSet::new();
        for (q, value) in raw {
            if q == "xmlns" || q.starts_with("xmlns:") {
                continue;
            }
            let name = expanded(&q, &ns, false)?;
            if !unique.insert((name.ns.clone(), name.local.clone())) {
                return Err(DigitalSignatureError::InvalidXml(format!(
                    "duplicate attribute {q}"
                )));
            }
            attrs.push(Attribute { name, value });
        }
        let n = es.len();
        es.push(Element {
            name,
            attrs,
            namespaces: ns,
            children: Vec::new(),
        });
        if let Some(&parent) = stack.last() {
            es[parent].children.push(Child::Element(n))
        } else if root.replace(n).is_some() {
            return Err(DigitalSignatureError::InvalidXml("multiple roots".into()));
        }
        for a in &es[n].attrs {
            if a.name.ns.is_empty()
                && a.name.local == "Id"
                && (a.value.is_empty() || ids.insert(a.value.clone(), n).is_some())
            {
                return Err(DigitalSignatureError::InvalidXml(
                    "empty or duplicate Id".into(),
                ));
            }
        }
        stack.push(n);
        Ok(())
    }
    fn push_text(es: &mut [Element], stack: &[usize], text: String) -> Result<()> {
        if let Some(&n) = stack.last() {
            es[n].children.push(Child::Text(text));
            Ok(())
        } else if text.trim().is_empty() {
            Ok(())
        } else {
            Err(DigitalSignatureError::InvalidXml(
                "text outside root".into(),
            ))
        }
    }
    fn is(&self, n: usize, ns: &str, local: &str) -> bool {
        self.elements[n].name.ns == ns && self.elements[n].name.local == local
    }
    fn children(&self, n: usize, ns: &str, local: &str) -> Vec<usize> {
        self.elements[n]
            .children
            .iter()
            .filter_map(|c| match c {
                Child::Element(i) if self.is(*i, ns, local) => Some(*i),
                _ => None,
            })
            .collect()
    }
    fn child(&self, n: usize, ns: &str, local: &str) -> Option<usize> {
        self.children(n, ns, local).first().copied()
    }
    fn required_child(&self, n: usize, ns: &str, local: &str) -> Result<usize> {
        let v = self.children(n, ns, local);
        if v.len() == 1 {
            Ok(v[0])
        } else {
            Err(DigitalSignatureError::InvalidXml(format!(
                "expected one {{{ns}}}{local}"
            )))
        }
    }
    fn descendants(&self, n: usize, ns: &str, local: &str) -> Vec<usize> {
        let mut out = Vec::new();
        let mut todo = vec![n];
        while let Some(x) = todo.pop() {
            for c in &self.elements[x].children {
                if let Child::Element(i) = c {
                    if self.is(*i, ns, local) {
                        out.push(*i)
                    }
                    todo.push(*i)
                }
            }
        }
        out
    }
    fn attr(&self, n: usize, local: &str) -> Result<&str> {
        let mut v = self.elements[n]
            .attrs
            .iter()
            .filter(|a| a.name.ns.is_empty() && a.name.local == local);
        let a = v.next().ok_or_else(|| {
            DigitalSignatureError::InvalidXml(format!("missing attribute {local}"))
        })?;
        if v.next().is_some() {
            return Err(DigitalSignatureError::InvalidXml(format!(
                "duplicate attribute {local}"
            )));
        }
        Ok(&a.value)
    }
    fn text(&self, n: usize) -> Result<String> {
        let mut out = String::new();
        for c in &self.elements[n].children {
            match c {
                Child::Text(s) => out.push_str(s),
                Child::Comment(_) => {},
                Child::Element(_) => {
                    return Err(DigitalSignatureError::InvalidXml(
                        "expected text-only element".into(),
                    ));
                },
            }
        }
        Ok(out)
    }
    fn canonicalize(&self, n: usize, mode: CanonMode) -> Vec<u8> {
        let mut out = Vec::new();
        let mut inherited = BTreeMap::new();
        inherited.insert("xml".into(), XML_NS.into());
        self.canon(n, &inherited, mode, &mut out);
        out
    }
    fn canon(
        &self,
        n: usize,
        inherited: &BTreeMap<String, String>,
        mode: CanonMode,
        out: &mut Vec<u8>,
    ) {
        let e = &self.elements[n];
        out.push(b'<');
        out.extend_from_slice(e.name.q.as_bytes());
        let visibly_used: HashSet<&str> = if mode.exclusive {
            std::iter::once(e.name.q.split_once(':').map_or("", |(prefix, _)| prefix))
                .chain(
                    e.attrs
                        .iter()
                        .filter_map(|attribute| attribute.name.q.split_once(':').map(|v| v.0)),
                )
                .collect()
        } else {
            e.namespaces.keys().map(String::as_str).collect()
        };
        let mut rendered = inherited.clone();
        for (pre, uri) in &e.namespaces {
            if pre == "xml" || inherited.get(pre) == Some(uri) {
                continue;
            }
            if mode.exclusive && !visibly_used.contains(pre.as_str()) {
                continue;
            }
            if pre.is_empty() {
                out.extend_from_slice(b" xmlns=\"")
            } else {
                out.extend_from_slice(b" xmlns:");
                out.extend_from_slice(pre.as_bytes());
                out.extend_from_slice(b"=\"")
            }
            attr_escape(out, uri);
            out.push(b'\"');
            rendered.insert(pre.clone(), uri.clone());
        }
        let mut attrs: Vec<_> = e.attrs.iter().collect();
        attrs.sort_by(|a, b| (&a.name.ns, &a.name.local).cmp(&(&b.name.ns, &b.name.local)));
        for a in attrs {
            out.push(b' ');
            out.extend_from_slice(a.name.q.as_bytes());
            out.extend_from_slice(b"=\"");
            attr_escape(out, &a.value);
            out.push(b'\"')
        }
        out.push(b'>');
        for c in &e.children {
            match c {
                Child::Element(i) => self.canon(*i, &rendered, mode, out),
                Child::Text(s) => text_escape(out, s),
                Child::Comment(s) if mode.comments => {
                    out.extend_from_slice(b"<!--");
                    out.extend_from_slice(s.as_bytes());
                    out.extend_from_slice(b"-->")
                },
                _ => {},
            }
        }
        out.extend_from_slice(b"</");
        out.extend_from_slice(e.name.q.as_bytes());
        out.push(b'>')
    }
}
fn xml<E: std::fmt::Display>(e: E) -> DigitalSignatureError {
    DigitalSignatureError::InvalidXml(e.to_string())
}
fn expanded(q: &str, ns: &BTreeMap<String, String>, element: bool) -> Result<Name> {
    if q.is_empty() || q.matches(':').count() > 1 {
        return Err(xml(format!("invalid name {q}")));
    }
    let (pre, local) = q.split_once(':').unwrap_or(("", q));
    if local.is_empty() || q.contains(':') && pre.is_empty() {
        return Err(xml(format!("invalid name {q}")));
    }
    let uri = if pre.is_empty() {
        if element {
            ns.get("").cloned().unwrap_or_default()
        } else {
            String::new()
        }
    } else {
        ns.get(pre)
            .cloned()
            .ok_or_else(|| xml(format!("unbound prefix {pre}")))?
    };
    Ok(Name {
        q: q.into(),
        local: local.into(),
        ns: uri,
    })
}
fn attr_escape(o: &mut Vec<u8>, s: &str) {
    for c in s.chars() {
        match c {
            '&' => o.extend_from_slice(b"&amp;"),
            '<' => o.extend_from_slice(b"&lt;"),
            '"' => o.extend_from_slice(b"&quot;"),
            '\t' => o.extend_from_slice(b"&#x9;"),
            '\n' => o.extend_from_slice(b"&#xA;"),
            '\r' => o.extend_from_slice(b"&#xD;"),
            _ => {
                let mut b = [0; 4];
                o.extend_from_slice(c.encode_utf8(&mut b).as_bytes())
            },
        }
    }
}
fn text_escape(o: &mut Vec<u8>, s: &str) {
    let mut prev = ['\0'; 2];
    for c in s.chars() {
        match c {
            '&' => o.extend_from_slice(b"&amp;"),
            '<' => o.extend_from_slice(b"&lt;"),
            c if c == '>' && prev == [']', ']'] => o.extend_from_slice(b"&gt;"),
            '\r' => o.extend_from_slice(b"&#xD;"),
            _ => {
                let mut b = [0; 4];
                o.extend_from_slice(c.encode_utf8(&mut b).as_bytes())
            },
        }
        prev = [prev[1], c]
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::part::BlobPart;
    const DOCX: &[u8] =
        include_bytes!("../../../test-data/poi/test-data/xmldsign/hello-world-signed.docx");
    const XLSX: &[u8] =
        include_bytes!("../../../test-data/poi/test-data/xmldsign/hello-world-signed.xlsx");
    const PPTX: &[u8] =
        include_bytes!("../../../test-data/poi/test-data/xmldsign/hello-world-signed.pptx");
    const TWICE: &[u8] =
        include_bytes!("../../../test-data/poi/test-data/xmldsign/hello-world-signed-twice.docx");
    const MICROSOFT_DOCX: &[u8] =
        include_bytes!("../../../test-data/poi/test-data/xmldsign/ms-office-2010-signed.docx");
    const MICROSOFT_XLSX: &[u8] =
        include_bytes!("../../../test-data/poi/test-data/xmldsign/ms-office-2010-signed.xlsx");
    const MICROSOFT_PPTX: &[u8] =
        include_bytes!("../../../test-data/poi/test-data/xmldsign/ms-office-2010-signed.pptx");
    fn valid(bytes: &[u8], count: usize) {
        let p = OpcPackage::from_bytes(bytes).unwrap();
        let reports = p
            .verify_digital_signatures(&SignatureVerificationPolicy::compatibility())
            .unwrap();
        assert_eq!(reports.len(), count);
        for r in reports {
            assert_eq!(r.package_integrity, VerificationStatus::Valid);
            assert_eq!(r.signature_value, VerificationStatus::Valid);
            assert_eq!(r.certificate_trust, CertificateTrust::NotEvaluated);
            assert!(r.uses_sha1);
            assert!(!r.certificates.is_empty())
        }
    }
    #[test]
    fn verifies_real_poi_office_fixtures() {
        valid(DOCX, 1);
        valid(XLSX, 1);
        valid(PPTX, 1)
    }
    #[test]
    fn verifies_real_poi_twice_signed_fixture() {
        valid(TWICE, 2)
    }
    #[test]
    fn verifies_real_microsoft_office_2010_fixtures() {
        valid(MICROSOFT_DOCX, 1);
        valid(MICROSOFT_XLSX, 1);
        valid(MICROSOFT_PPTX, 1);
    }
    #[test]
    fn strict_rejects_sha1() {
        let p = OpcPackage::from_bytes(DOCX).unwrap();
        assert!(matches!(
            p.verify_digital_signatures(&SignatureVerificationPolicy::strict()),
            Err(DigitalSignatureError::Sha1Disallowed)
        ))
    }
    #[test]
    fn tamper_is_reported_not_trusted() {
        let mut p = OpcPackage::from_bytes(DOCX).unwrap();
        let u = PackURI::new("/word/document.xml").unwrap();
        let part = p.get_part_mut(&u).unwrap();
        let mut b = part.blob().to_vec();
        b.push(b' ');
        part.set_blob(b);
        let r = p
            .verify_digital_signatures(&SignatureVerificationPolicy::compatibility())
            .unwrap();
        assert_eq!(r[0].package_integrity, VerificationStatus::Invalid);
        assert_eq!(r[0].signature_value, VerificationStatus::Valid)
    }
    #[test]
    fn verifies_certificate_stored_in_related_part() {
        let mut p = OpcPackage::from_bytes(DOCX).unwrap();
        let signature_uri = p
            .iter_parts()
            .find(|part| part.content_type() == SIGNATURE_CONTENT_TYPE)
            .unwrap()
            .partname()
            .clone();
        let policy = SignatureVerificationPolicy::compatibility();
        let (signature_xml, certificate) = {
            let part = p.get_part(&signature_uri).unwrap();
            let xml = str::from_utf8(part.blob()).unwrap();
            let certificate_start =
                xml.find("<X509Certificate>").unwrap() + "<X509Certificate>".len();
            let certificate_end =
                xml[certificate_start..].find("</X509Certificate>").unwrap() + certificate_start;
            let certificate = decode64(
                &xml[certificate_start..certificate_end],
                policy.max_embedded_certificate_bytes,
                "X509Certificate",
            )
            .unwrap();
            let data_start = xml.find("<X509Data>").unwrap();
            let data_end =
                xml[data_start..].find("</X509Data>").unwrap() + data_start + "</X509Data>".len();
            let mut without = Vec::with_capacity(xml.len() - (data_end - data_start));
            without.extend_from_slice(&xml.as_bytes()[..data_start]);
            without.extend_from_slice(&xml.as_bytes()[data_end..]);
            (without, certificate)
        };
        {
            let signature = p.get_part_mut(&signature_uri).unwrap();
            signature.set_blob(signature_xml);
            signature
                .rels_mut()
                .try_add_relationship(
                    CERTIFICATE_REL.to_string(),
                    "cert1.cer".to_string(),
                    "rIdCertificate".to_string(),
                    TargetMode::Internal,
                )
                .unwrap();
        }
        let certificate_uri = PackURI::new("/_xmlsignatures/cert1.cer").unwrap();
        p.try_add_part(Box::new(BlobPart::new(
            certificate_uri,
            CERTIFICATE_CONTENT_TYPE.to_string(),
            certificate.clone(),
        )))
        .unwrap();
        let reports = p.verify_digital_signatures(&policy).unwrap();
        assert_eq!(reports.len(), 1);
        assert_eq!(reports[0].package_integrity, VerificationStatus::Valid);
        assert_eq!(reports[0].signature_value, VerificationStatus::Valid);
        assert_eq!(
            reports[0].certificates.as_slice(),
            &[EmbeddedCertificate { der: certificate }]
        );
    }
    #[test]
    fn certificate_resource_limits_are_deduplication_aware() {
        let mut policy = SignatureVerificationPolicy::compatibility();
        policy.max_certificates = 1;
        policy.max_total_certificate_bytes = 3;
        let mut certificates = Vec::new();
        push_certificate(&mut certificates, b"abc", &policy).unwrap();
        push_certificate(&mut certificates, b"abc", &policy).unwrap();
        assert_eq!(certificates.len(), 1);
        assert!(matches!(
            push_certificate(&mut certificates, b"def", &policy),
            Err(DigitalSignatureError::LimitExceeded(_))
        ));
        policy.max_certificates = 2;
        assert!(matches!(
            push_certificate(&mut certificates, b"d", &policy),
            Err(DigitalSignatureError::LimitExceeded(_))
        ));
    }

    fn authored_package() -> OpcPackage {
        let mut package = OpcPackage::new();
        let document_uri = PackURI::new("/document.xml").unwrap();
        package
            .try_add_part(Box::new(BlobPart::new(
                document_uri,
                "application/xml".to_string(),
                b"<document><value>signed</value></document>".to_vec(),
            )))
            .unwrap();
        package.relate_to(
            "/document.xml",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
        );
        package
    }

    fn p256_signer() -> PackageSigner {
        let key = EcdsaSigningKey::from_bytes((&[7u8; 32]).into()).unwrap();
        let mut signer = PackageSigner::ecdsa_p256_sha256(key);
        signer
            .set_signing_time(Some("2026-07-19T12:34:56+08:00"))
            .unwrap();
        signer
    }

    #[test]
    fn authors_and_verifies_rsa_sha256_round_trip() {
        let key = RsaPrivateKey::new(&mut rsa::rand_core::OsRng, 2048).unwrap();
        let mut signer = PackageSigner::rsa_sha256(key).unwrap();
        signer
            .set_signing_time(Some("2026-07-19T12:34:56Z"))
            .unwrap();
        let mut package = authored_package();
        let uri = package.add_digital_signature(&signer).unwrap();
        assert_eq!(uri.as_str(), "/_xmlsignatures/sig1.xml");
        let verification = package
            .verify_digital_signatures(&SignatureVerificationPolicy::strict())
            .unwrap();
        assert_eq!(verification.len(), 1);
        assert_eq!(verification[0].package_integrity, VerificationStatus::Valid);
        assert_eq!(verification[0].signature_value, VerificationStatus::Valid);
        assert_eq!(
            verification[0].signing_time.as_deref(),
            Some("2026-07-19T12:34:56Z")
        );
    }

    #[test]
    fn authors_multiple_p256_signatures_with_exclusive_canonicalization() {
        let mut package = authored_package();
        let mut first = p256_signer();
        first.set_canonicalization(CanonicalizationMethod::Exclusive);
        package.add_digital_signature(&first).unwrap();
        package.add_digital_signature(&p256_signer()).unwrap();
        let verification = package
            .verify_digital_signatures(&SignatureVerificationPolicy::strict())
            .unwrap();
        assert_eq!(verification.len(), 2);
        assert!(verification.iter().all(|report| {
            report.package_integrity == VerificationStatus::Valid
                && report.signature_value == VerificationStatus::Valid
                && report.certificate_trust == CertificateTrust::NotEvaluated
        }));
        package.clear_digital_signatures().unwrap();
        assert!(package
            .verify_digital_signatures(&SignatureVerificationPolicy::strict())
            .unwrap()
            .is_empty());
    }

    #[test]
    fn authored_tamper_and_atomic_failures_are_detected() {
        let mut package = authored_package();
        package.add_digital_signature(&p256_signer()).unwrap();
        let document_uri = PackURI::new("/document.xml").unwrap();
        package
            .get_part_mut(&document_uri)
            .unwrap()
            .set_blob(b"<document>tampered</document>".to_vec());
        let verification = package
            .verify_digital_signatures(&SignatureVerificationPolicy::strict())
            .unwrap();
        assert_eq!(verification[0].package_integrity, VerificationStatus::Invalid);
        assert_eq!(verification[0].signature_value, VerificationStatus::Valid);

        let mut signer = p256_signer();
        assert!(signer.set_certificates(vec![vec![1, 2, 3]]).is_err());
        assert!(signer.certificates().is_empty());

        let mut spoofed = authored_package();
        spoofed
            .try_add_part(Box::new(BlobPart::new(
                PackURI::new("/_xmlsignatures/origin.sigs").unwrap(),
                "application/octet-stream".into(),
                b"preserve".to_vec(),
            )))
            .unwrap();
        let before = spoofed.part_count();
        assert!(spoofed.add_digital_signature(&p256_signer()).is_err());
        assert_eq!(spoofed.part_count(), before);
        assert!(!spoofed.contains_part(
            &PackURI::new("/_xmlsignatures/sig1.xml").unwrap()
        ));
    }

    #[test]
    fn authored_signatures_round_trip_for_all_ooxml_package_families() {
        for (part_name, content_type) in [
            (
                "/word/document.xml",
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
            ),
            (
                "/xl/workbook.xml",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
            ),
            (
                "/ppt/presentation.xml",
                "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml",
            ),
            (
                "/xl/workbook.bin",
                "application/vnd.ms-excel.sheet.binary.macroEnabled.main",
            ),
        ] {
            let mut package = OpcPackage::new();
            package
                .try_add_part(Box::new(BlobPart::new(
                    PackURI::new(part_name).unwrap(),
                    content_type.to_string(),
                    if part_name.ends_with(".bin") {
                        vec![0, 1, 2, 3, 0xff]
                    } else {
                        b"<root xmlns=\"urn:test\"><value>roundtrip</value></root>".to_vec()
                    },
                )))
                .unwrap();
            package.relate_to(
                part_name,
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
            );
            package.add_digital_signature(&p256_signer()).unwrap();
            let mut serialized = std::io::Cursor::new(Vec::new());
            package.to_stream(&mut serialized).unwrap();
            let reopened = OpcPackage::from_bytes(serialized.get_ref()).unwrap();
            let verification = reopened
                .verify_digital_signatures(&SignatureVerificationPolicy::strict())
                .unwrap();
            assert_eq!(verification.len(), 1, "failed package family {part_name}");
            assert_eq!(verification[0].package_integrity, VerificationStatus::Valid);
            assert_eq!(verification[0].signature_value, VerificationStatus::Valid);
        }
    }

    #[test]
    fn rejects_wrong_key_relationship_mismatch_and_unsafe_part_reference() {
        let mut wrong_key_package = authored_package();
        let signer = p256_signer();
        wrong_key_package.add_digital_signature(&signer).unwrap();
        let signature_uri = PackURI::new("/_xmlsignatures/sig1.xml").unwrap();
        let correct_key = match &signer.material {
            SigningMaterial::Ecdsa(key) => encode64(key.verifying_key().to_encoded_point(false).as_bytes()),
            SigningMaterial::Rsa(_) => unreachable!(),
        };
        let wrong_signer = PackageSigner::ecdsa_p256_sha256(
            EcdsaSigningKey::from_bytes((&[9u8; 32]).into()).unwrap(),
        );
        let wrong_key = match &wrong_signer.material {
            SigningMaterial::Ecdsa(key) => encode64(key.verifying_key().to_encoded_point(false).as_bytes()),
            SigningMaterial::Rsa(_) => unreachable!(),
        };
        let xml = String::from_utf8(
            wrong_key_package
                .get_part(&signature_uri)
                .unwrap()
                .blob()
                .to_vec(),
        )
        .unwrap()
        .replace(&correct_key, &wrong_key);
        wrong_key_package
            .get_part_mut(&signature_uri)
            .unwrap()
            .set_blob(xml.into_bytes());
        let verification = wrong_key_package
            .verify_digital_signatures(&SignatureVerificationPolicy::strict())
            .unwrap();
        assert_eq!(verification[0].signature_value, VerificationStatus::Invalid);

        let mut relationship_spoof = authored_package();
        relationship_spoof.add_digital_signature(&p256_signer()).unwrap();
        let xml = String::from_utf8(
            relationship_spoof
                .get_part(&signature_uri)
                .unwrap()
                .blob()
                .to_vec(),
        )
        .unwrap()
        .replace("SourceId=\"rId1\"", "SourceId=\"rIdMissing\"");
        relationship_spoof
            .get_part_mut(&signature_uri)
            .unwrap()
            .set_blob(xml.into_bytes());
        assert!(matches!(
            relationship_spoof
                .verify_digital_signatures(&SignatureVerificationPolicy::strict()),
            Err(DigitalSignatureError::InvalidGraph(_))
        ));

        let mut path_spoof = authored_package();
        path_spoof.add_digital_signature(&p256_signer()).unwrap();
        let xml = String::from_utf8(
            path_spoof
                .get_part(&signature_uri)
                .unwrap()
                .blob()
                .to_vec(),
        )
        .unwrap()
        .replace("/document.xml?ContentType=", "/../document.xml?ContentType=");
        path_spoof
            .get_part_mut(&signature_uri)
            .unwrap()
            .set_blob(xml.into_bytes());
        assert!(path_spoof
            .verify_digital_signatures(&SignatureVerificationPolicy::strict())
            .is_err());
    }
}
