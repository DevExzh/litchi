//! Trust-neutral digital signatures for Office-family containers.
//!
//! The crate deliberately separates mathematical integrity from PKI trust.
//! Parsing and verification never fetch certificates, execute document content,
//! or consult an async runtime. Container crates resolve signed bytes; this
//! crate owns bounded XMLDSig processing and cryptographic key handling.

#![forbid(unsafe_code)]

pub mod cfb;
pub mod xml;

use chrono::DateTime;
use p256::ecdsa::{SigningKey as EcSigningKey, VerifyingKey as EcVerifyingKey};
use p256::pkcs8::{DecodePrivateKey as _, DecodePublicKey as _};
use rsa::pkcs1v15::{Signature as RsaSignature, SigningKey as RsaSigningKey};
use rsa::sha2::Sha256;
use rsa::traits::PublicKeyParts as _;
use rsa::{RsaPrivateKey, RsaPublicKey};
use signature::{SignatureEncoding as _, Signer as _};
use std::fmt;
use thiserror::Error as ThisError;
use x509_cert::Certificate as X509Certificate;
use x509_cert::der::{Decode as _, Encode as _};
use zeroize::Zeroizing;

/// Result type used by every signing layer.
pub type Result<T> = std::result::Result<T, Error>;

/// Errors are typed by failure domain and contain no container-specific IDs.
#[derive(Debug, ThisError)]
#[non_exhaustive]
pub enum Error {
    #[error("invalid or unsafe signature XML: {0}")]
    Xml(String),
    #[error("signature resource limit exceeded: {0}")]
    Limit(String),
    #[error("unsupported signature algorithm or transform: {0}")]
    Unsupported(String),
    #[error("SHA-1 is disallowed by the verification policy")]
    Sha1,
    #[error("invalid signing or verification key: {0}")]
    Key(String),
    #[error("signature authoring failed: {0}")]
    Sign(String),
    #[error("invalid signature container: {0}")]
    Container(String),
    #[error("CFB error: {0}")]
    Cfb(#[from] litchi_cfb::OleError),
    #[error("encrypted binary Office documents require a decrypted storage before signing")]
    Encrypted,
    #[error("legacy `_signatures` CryptoAPI signatures are recognized but unsupported")]
    Legacy,
}

/// Resource ceilings shared by XML and container adapters.
///
/// Fields are private so callers cannot manufacture a zero-sized or internally
/// inconsistent policy with a struct literal.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Limits {
    signature_bytes: usize,
    xml_depth: usize,
    xml_elements: usize,
    attributes: usize,
    references: usize,
    certificate_bytes: usize,
    certificates: usize,
    total_certificate_bytes: usize,
    rsa_bits: usize,
    cfb_bytes: usize,
    cfb_entries: usize,
    cfb_depth: usize,
    cfb_streams: usize,
    signatures: usize,
}

impl Limits {
    /// Conservative defaults suitable for untrusted Office documents.
    pub fn standard() -> Self {
        Self {
            signature_bytes: 8 * 1024 * 1024,
            xml_depth: 128,
            xml_elements: 100_000,
            attributes: 256,
            references: 20_000,
            certificate_bytes: 1024 * 1024,
            certificates: 256,
            total_certificate_bytes: 16 * 1024 * 1024,
            rsa_bits: 16_384,
            cfb_bytes: 1024 * 1024 * 1024,
            cfb_entries: 100_000,
            cfb_depth: 128,
            cfb_streams: 20_000,
            signatures: 64,
        }
    }

    /// Set the maximum signature XML size.
    pub fn signature_bytes(mut self, value: usize) -> Result<Self> {
        self.signature_bytes = positive(value, "signature byte limit")?;
        Ok(self)
    }

    /// Set the maximum parsed XML depth.
    pub fn xml_depth(mut self, value: usize) -> Result<Self> {
        self.xml_depth = positive(value, "XML depth limit")?;
        Ok(self)
    }

    /// Set the maximum parsed XML element count.
    pub fn xml_elements(mut self, value: usize) -> Result<Self> {
        self.xml_elements = positive(value, "XML element limit")?;
        Ok(self)
    }

    /// Set the maximum attributes on one element.
    pub fn attributes(mut self, value: usize) -> Result<Self> {
        self.attributes = positive(value, "XML attribute limit")?;
        Ok(self)
    }

    /// Set the maximum number of signed references.
    pub fn references(mut self, value: usize) -> Result<Self> {
        self.references = positive(value, "reference limit")?;
        Ok(self)
    }

    /// Set the per-certificate DER ceiling.
    pub fn certificate_bytes(mut self, value: usize) -> Result<Self> {
        self.certificate_bytes = positive(value, "certificate byte limit")?;
        Ok(self)
    }

    /// Set the certificate-count ceiling.
    pub fn certificates(mut self, value: usize) -> Result<Self> {
        self.certificates = positive(value, "certificate count limit")?;
        Ok(self)
    }

    /// Set the aggregate certificate DER ceiling.
    pub fn total_certificate_bytes(mut self, value: usize) -> Result<Self> {
        self.total_certificate_bytes = positive(value, "total certificate byte limit")?;
        Ok(self)
    }

    /// Set the largest accepted RSA modulus.
    pub fn rsa_bits(mut self, value: usize) -> Result<Self> {
        if value < 2_048 {
            return Err(Error::Limit(
                "RSA modulus ceiling must be at least 2048 bits".into(),
            ));
        }
        self.rsa_bits = value;
        Ok(self)
    }

    /// Set the maximum CFB byte size accepted by the editor.
    pub fn cfb_bytes(mut self, value: usize) -> Result<Self> {
        self.cfb_bytes = positive(value, "CFB byte limit")?;
        Ok(self)
    }

    /// Set the maximum total directory entries traversed in a CFB container.
    pub fn cfb_entries(mut self, value: usize) -> Result<Self> {
        self.cfb_entries = positive(value, "CFB directory-entry limit")?;
        Ok(self)
    }

    /// Set the maximum storage/stream path depth in a CFB container.
    pub fn cfb_depth(mut self, value: usize) -> Result<Self> {
        self.cfb_depth = positive(value, "CFB path-depth limit")?;
        Ok(self)
    }

    /// Set the maximum number of CFB streams.
    pub fn cfb_streams(mut self, value: usize) -> Result<Self> {
        self.cfb_streams = positive(value, "CFB stream limit")?;
        Ok(self)
    }

    /// Set the maximum number of signatures in one container.
    pub fn signatures(mut self, value: usize) -> Result<Self> {
        self.signatures = positive(value, "signature count limit")?;
        Ok(self)
    }

    pub fn max_signature_bytes(&self) -> usize {
        self.signature_bytes
    }

    pub fn max_references(&self) -> usize {
        self.references
    }

    pub fn max_cfb_bytes(&self) -> usize {
        self.cfb_bytes
    }

    pub fn max_cfb_entries(&self) -> usize {
        self.cfb_entries
    }

    pub fn max_cfb_depth(&self) -> usize {
        self.cfb_depth
    }

    pub fn max_cfb_streams(&self) -> usize {
        self.cfb_streams
    }

    pub fn max_signatures(&self) -> usize {
        self.signatures
    }

    pub(crate) fn max_xml_depth(&self) -> usize {
        self.xml_depth
    }

    pub(crate) fn max_xml_elements(&self) -> usize {
        self.xml_elements
    }

    pub(crate) fn max_attributes(&self) -> usize {
        self.attributes
    }

    pub fn max_certificate_bytes(&self) -> usize {
        self.certificate_bytes
    }

    pub fn max_certificates(&self) -> usize {
        self.certificates
    }

    pub fn max_total_certificate_bytes(&self) -> usize {
        self.total_certificate_bytes
    }

    pub(crate) fn max_rsa_bits(&self) -> usize {
        self.rsa_bits
    }
}

impl Default for Limits {
    fn default() -> Self {
        Self::standard()
    }
}

fn positive(value: usize, description: &str) -> Result<usize> {
    if value == 0 {
        Err(Error::Limit(format!("{description} must be non-zero")))
    } else {
        Ok(value)
    }
}

/// Treatment of legacy SHA-1 while verifying existing documents.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Weak {
    Allow,
    Reject,
}

/// Immutable verification policy.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Policy {
    weak: Weak,
    partial: bool,
    limits: Limits,
}

impl Policy {
    pub fn compatible() -> Self {
        Self {
            weak: Weak::Allow,
            partial: true,
            limits: Limits::standard(),
        }
    }

    pub fn strict() -> Self {
        Self {
            weak: Weak::Reject,
            partial: false,
            limits: Limits::standard(),
        }
    }

    pub fn with_limits(mut self, limits: Limits) -> Self {
        self.limits = limits;
        self
    }

    pub fn limits(&self) -> &Limits {
        &self.limits
    }

    pub fn weak(&self) -> Weak {
        self.weak
    }

    /// Whether a cryptographically valid signed subset is reportable rather
    /// than rejected. Compatibility mode permits this because real Office
    /// producers omit some package parts and relationship IDs.
    pub fn allows_partial_coverage(&self) -> bool {
        self.partial
    }
}

impl Default for Policy {
    fn default() -> Self {
        Self::strict()
    }
}

/// Mathematical verification state.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Status {
    Valid,
    Invalid,
}

/// PKI trust state. This crate intentionally performs no ambient trust lookup.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Trust {
    NotChecked,
}

/// Whether every container-owned part and relationship was covered.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Coverage {
    Complete,
    Partial,
}

impl Coverage {
    pub(crate) fn combine(self, other: Self) -> Self {
        if self == Self::Complete && other == Self::Complete {
            Self::Complete
        } else {
            Self::Partial
        }
    }
}

/// One embedded X.509 certificate.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Cert {
    der: Vec<u8>,
}

impl Cert {
    pub fn der(&self) -> &[u8] {
        &self.der
    }
}

/// Result for one signed reference.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Reference {
    uri: String,
    status: Status,
    coverage: Coverage,
}

impl Reference {
    pub fn uri(&self) -> &str {
        &self.uri
    }

    pub fn status(&self) -> Status {
        self.status
    }

    pub fn coverage(&self) -> Coverage {
        self.coverage
    }
}

/// Trust-neutral verification result.
#[derive(Debug, Clone)]
pub struct Report {
    integrity: Status,
    signature: Status,
    trust: Trust,
    coverage: Coverage,
    references: Vec<Reference>,
    certificates: Vec<Cert>,
    uses_sha1: bool,
    time: Option<String>,
}

impl Report {
    pub fn integrity(&self) -> Status {
        self.integrity
    }

    pub fn signature(&self) -> Status {
        self.signature
    }

    pub fn trust(&self) -> Trust {
        self.trust
    }

    pub fn coverage(&self) -> Coverage {
        self.coverage
    }

    pub fn references(&self) -> &[Reference] {
        &self.references
    }

    pub fn certificates(&self) -> &[Cert] {
        &self.certificates
    }

    pub fn uses_sha1(&self) -> bool {
        self.uses_sha1
    }

    pub fn time(&self) -> Option<&str> {
        self.time.as_deref()
    }

    pub(crate) fn new(
        integrity: Status,
        signature: Status,
        coverage: Coverage,
        references: Vec<Reference>,
        certificates: Vec<Cert>,
        uses_sha1: bool,
        time: Option<String>,
    ) -> Self {
        Self {
            integrity,
            signature,
            trust: Trust::NotChecked,
            coverage,
            references,
            certificates,
            uses_sha1,
            time,
        }
    }
}

impl Reference {
    pub(crate) fn new(uri: String, status: Status, coverage: Coverage) -> Self {
        Self {
            uri,
            status,
            coverage,
        }
    }
}

impl Cert {
    pub(crate) fn new(der: Vec<u8>) -> Self {
        Self { der }
    }
}

enum Key {
    Rsa {
        signing: Box<RsaSigningKey<Sha256>>,
        public: RsaPublicKey,
    },
    P256(EcSigningKey),
}

/// Validated, move-only signing capability.
///
/// `Signer` intentionally does not implement `Clone`; duplicating private key
/// material is never an implicit operation.
pub struct Signer {
    key: Key,
    certificates: Vec<Vec<u8>>,
    time: Option<String>,
    canon: xml::Canon,
}

impl fmt::Debug for Signer {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Signer")
            .field("method", &self.method())
            .field("certificates", &self.certificates.len())
            .field("time", &self.time)
            .field("canon", &self.canon)
            .finish_non_exhaustive()
    }
}

impl Signer {
    /// Consume a validated RSA private key.
    pub fn rsa(key: RsaPrivateKey) -> Result<Self> {
        if key.n().bits() < 2_048 {
            return Err(Error::Key(
                "RSA signing keys must be at least 2048 bits".into(),
            ));
        }
        key.validate()
            .map_err(|error| Error::Key(error.to_string()))?;
        let public = key.to_public_key();
        Ok(Self {
            key: Key::Rsa {
                signing: Box::new(RsaSigningKey::<Sha256>::new(key)),
                public,
            },
            certificates: Vec::new(),
            time: None,
            canon: xml::Canon::Inclusive,
        })
    }

    /// Consume a P-256 private key.
    pub fn p256(key: EcSigningKey) -> Self {
        Self {
            key: Key::P256(key),
            certificates: Vec::new(),
            time: None,
            canon: xml::Canon::Inclusive,
        }
    }

    /// Parse a zeroizing PKCS#8 RSA buffer once and retain the parsed key.
    pub fn rsa_pkcs8(der: Zeroizing<Vec<u8>>) -> Result<Self> {
        let key =
            RsaPrivateKey::from_pkcs8_der(&der).map_err(|error| Error::Key(error.to_string()))?;
        Self::rsa(key)
    }

    /// Parse a zeroizing PKCS#8 P-256 buffer once and retain the parsed key.
    pub fn p256_pkcs8(der: Zeroizing<Vec<u8>>) -> Result<Self> {
        let key =
            EcSigningKey::from_pkcs8_der(&der).map_err(|error| Error::Key(error.to_string()))?;
        Ok(Self::p256(key))
    }

    /// Attach a certificate chain after checking its leaf against the key.
    pub fn certs(self, certificates: Vec<Vec<u8>>) -> Result<Self> {
        self.certs_with(certificates, &Limits::standard())
    }

    /// Attach a certificate chain with explicit pre-parse resource bounds.
    pub fn certs_with(mut self, certificates: Vec<Vec<u8>>, limits: &Limits) -> Result<Self> {
        if certificates.len() > limits.max_certificates() {
            return Err(Error::Limit("too many signing certificates".into()));
        }
        let mut total = 0_usize;
        for certificate in &certificates {
            if certificate.len() > limits.max_certificate_bytes() {
                return Err(Error::Limit("signing certificate is too large".into()));
            }
            total = total
                .checked_add(certificate.len())
                .ok_or_else(|| Error::Limit("certificate byte count overflow".into()))?;
            if total > limits.max_total_certificate_bytes() {
                return Err(Error::Limit("signing certificates are too large".into()));
            }
        }
        if certificates.iter().any(Vec::is_empty) {
            return Err(Error::Key("certificate DER must not be empty".into()));
        }
        for certificate in &certificates {
            X509Certificate::from_der(certificate)
                .map_err(|error| Error::Key(format!("invalid X.509 certificate: {error}")))?;
        }
        if let Some(leaf) = certificates.first() {
            self.check_leaf(leaf)?;
        }
        self.certificates = certificates;
        Ok(self)
    }

    /// Attach an RFC 3339 signing time.
    pub fn time(mut self, value: impl Into<String>) -> Result<Self> {
        let value = value.into();
        if value.len() > 64 {
            return Err(Error::Sign("signing time exceeds 64 bytes".into()));
        }
        DateTime::parse_from_rfc3339(&value)
            .map_err(|error| Error::Sign(format!("invalid signing time: {error}")))?;
        self.time = Some(value);
        Ok(self)
    }

    pub fn canon(mut self, value: xml::Canon) -> Self {
        self.canon = value;
        self
    }

    pub fn method(&self) -> xml::Method {
        match self.key {
            Key::Rsa { .. } => xml::Method::RsaSha256,
            Key::P256(_) => xml::Method::EcdsaP256Sha256,
        }
    }

    pub fn certificates(&self) -> &[Vec<u8>] {
        &self.certificates
    }

    pub fn signing_time(&self) -> Option<&str> {
        self.time.as_deref()
    }

    pub fn canonicalization(&self) -> xml::Canon {
        self.canon
    }

    fn check_leaf(&self, certificate: &[u8]) -> Result<()> {
        let certificate = X509Certificate::from_der(certificate)
            .map_err(|error| Error::Key(format!("invalid X.509 certificate: {error}")))?;
        let spki = certificate
            .tbs_certificate
            .subject_public_key_info
            .to_der()
            .map_err(|error| Error::Key(error.to_string()))?;
        match &self.key {
            Key::Rsa { public, .. } => {
                let leaf = RsaPublicKey::from_public_key_der(&spki)
                    .map_err(|error| Error::Key(error.to_string()))?;
                if &leaf != public {
                    return Err(Error::Key("leaf certificate does not match RSA key".into()));
                }
            },
            Key::P256(key) => {
                let leaf = EcVerifyingKey::from_public_key_der(&spki)
                    .map_err(|error| Error::Key(error.to_string()))?;
                if leaf != *key.verifying_key() {
                    return Err(Error::Key(
                        "leaf certificate does not match P-256 key".into(),
                    ));
                }
            },
        }
        Ok(())
    }

    pub(crate) fn sign(&self, message: &[u8]) -> Vec<u8> {
        match &self.key {
            Key::Rsa { signing, .. } => {
                let signature: RsaSignature = signing.sign(message);
                signature.to_vec()
            },
            Key::P256(key) => {
                let signature: p256::ecdsa::Signature = key.sign(message);
                signature.to_bytes().to_vec()
            },
        }
    }

    pub(crate) fn rsa_public(&self) -> Option<&RsaPublicKey> {
        match &self.key {
            Key::Rsa { public, .. } => Some(public),
            Key::P256(_) => None,
        }
    }

    pub(crate) fn p256_public(&self) -> Option<&EcVerifyingKey> {
        match &self.key {
            Key::Rsa { .. } => None,
            Key::P256(key) => Some(key.verifying_key()),
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn limits_reject_invalid_values() {
        assert!(Limits::standard().signature_bytes(0).is_err());
        assert!(Limits::standard().rsa_bits(1_024).is_err());
    }

    #[test]
    fn signing_time_is_typed_at_construction() {
        let key = EcSigningKey::from_bytes((&[7_u8; 32]).into()).unwrap();
        assert!(Signer::p256(key).time("not-a-time").is_err());
    }
}
