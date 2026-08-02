//! Presentation protection support for PowerPoint presentations.
//!
//! This module provides PresentationML modification protection, read-only
//! recommendations, and slide restrictions. File-open encryption belongs to
//! the outer package and is configured when the package is saved.

use crate::error::{OoxmlError, Result};
use base64::Engine;
use base64::engine::general_purpose::STANDARD as BASE64_ENGINE;
use quick_xml::Reader;
use quick_xml::events::Event;
use rand::TryRng;
use rand::rngs::SysRng;
use sha2::digest::Output;
use sha2::{Digest, Sha512};
use zeroize::{Zeroize, Zeroizing};

/// SHA-512 output whose storage is erased when it leaves scope.
///
/// `digest::Output` deliberately does not implement [`Zeroize`], so this
/// local wrapper provides the operation without allocating an intermediate
/// password-hash `Vec`.
struct Sha512Output(Output<Sha512>);

impl Zeroize for Sha512Output {
    fn zeroize(&mut self) {
        self.0.as_mut_slice().fill(0);
    }
}

fn modify_password_hash(password: &str, salt: &[u8], spin_count: u32) -> Zeroizing<Sha512Output> {
    // [MS-OE376] requires H(password || salt), followed by
    // H(previous || little_endian_iteration) for iterations starting at zero.
    let mut hasher = Sha512::new();
    for unit in password.encode_utf16() {
        hasher.update(unit.to_le_bytes());
    }
    hasher.update(salt);
    let mut hash = Zeroizing::new(Sha512Output(Output::<Sha512>::default()));
    hasher.finalize_into(&mut hash.0);

    for iteration in 0..spin_count {
        let mut hasher = Sha512::new();
        hasher.update(hash.0.as_slice());
        hasher.update(iteration.to_le_bytes());
        hasher.finalize_into(&mut hash.0);
    }
    hash
}

/// Type of protection applied to a presentation.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ProtectionType {
    /// No protection
    None,
    /// Read-only recommended (shows a dialog but can be bypassed)
    ReadOnlyRecommended,
    /// Password required to modify
    ModifyPassword,
}

/// Cryptographic algorithm for password hashing.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum CryptoAlgorithm {
    /// SHA-1 (legacy)
    Sha1,
    /// SHA-256
    #[default]
    Sha256,
    /// SHA-384
    Sha384,
    /// SHA-512
    Sha512,
}

impl CryptoAlgorithm {
    /// Get the algorithm URI for XML.
    pub fn uri(&self) -> &'static str {
        match self {
            CryptoAlgorithm::Sha1 => "http://www.w3.org/2000/09/xmldsig#sha1",
            CryptoAlgorithm::Sha256 => "http://www.w3.org/2001/04/xmlenc#sha256",
            CryptoAlgorithm::Sha384 => "http://www.w3.org/2001/04/xmldsig-more#sha384",
            CryptoAlgorithm::Sha512 => "http://www.w3.org/2001/04/xmlenc#sha512",
        }
    }

    /// Parse an exact OOXML algorithm URI or registered algorithm name.
    pub fn from_uri(value: &str) -> Result<Self> {
        match value {
            "http://www.w3.org/2000/09/xmldsig#sha1" | "SHA-1" => Ok(Self::Sha1),
            "http://www.w3.org/2001/04/xmlenc#sha256" | "SHA-256" => Ok(Self::Sha256),
            "http://www.w3.org/2001/04/xmldsig-more#sha384" | "SHA-384" => Ok(Self::Sha384),
            "http://www.w3.org/2001/04/xmlenc#sha512" | "SHA-512" => Ok(Self::Sha512),
            _ => Err(OoxmlError::InvalidFormat(format!(
                "unsupported presentation protection hash algorithm '{value}'"
            ))),
        }
    }

    fn from_sid(sid: u32) -> Result<Self> {
        match sid {
            4 => Ok(Self::Sha1),
            12 => Ok(Self::Sha256),
            13 => Ok(Self::Sha384),
            14 => Ok(Self::Sha512),
            _ => Err(OoxmlError::InvalidFormat(format!(
                "unsupported presentation protection hash SID {sid}"
            ))),
        }
    }

    const fn output_bytes(self) -> usize {
        match self {
            Self::Sha1 => 20,
            Self::Sha256 => 32,
            Self::Sha384 => 48,
            Self::Sha512 => 64,
        }
    }
}

fn set_once<T>(slot: &mut Option<T>, value: T, field: &'static str) -> Result<()> {
    if slot.is_some() {
        return Err(OoxmlError::InvalidFormat(format!(
            "duplicate presentation modify-verifier {field}"
        )));
    }
    *slot = Some(value);
    Ok(())
}

fn verifier_value<'a>(bytes: &'a [u8], field: &'static str) -> Result<&'a str> {
    std::str::from_utf8(bytes).map_err(|_| {
        OoxmlError::InvalidFormat(format!("presentation modify-verifier {field} is not UTF-8"))
    })
}

/// Validated, read-only presentation modification verifier.
#[derive(Clone)]
pub struct ModifyVerifier {
    algorithm: CryptoAlgorithm,
    spin_count: u32,
    hash: String,
    salt: String,
}

impl ModifyVerifier {
    /// Hash algorithm recorded by the verifier.
    pub const fn algorithm(&self) -> CryptoAlgorithm {
        self.algorithm
    }

    /// Number of iterative hash rounds after the initial password/salt hash.
    pub const fn spins(&self) -> u32 {
        self.spin_count
    }

    /// Borrow the Base64-encoded verifier hash.
    pub fn hash(&self) -> &str {
        &self.hash
    }

    /// Borrow the Base64-encoded verifier salt.
    pub fn salt(&self) -> &str {
        &self.salt
    }
}

impl std::fmt::Debug for ModifyVerifier {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("ModifyVerifier")
            .field("algorithm", &self.algorithm)
            .field("spin_count", &self.spin_count)
            .field("hash_bytes", &self.hash.len())
            .field("salt_bytes", &self.salt.len())
            .finish_non_exhaustive()
    }
}

/// Protection settings for a presentation.
#[derive(Debug, Clone, Default)]
pub struct PresentationProtection {
    /// Whether the presentation is marked as read-only recommended
    pub read_only_recommended: bool,
    /// Modification protection is one validated aggregate, so callers cannot
    /// independently desynchronize its algorithm, rounds, salt, and hash.
    modify: Option<ModifyVerifier>,
    /// Prevent editing of individual slides
    pub protect_structure: bool,
    /// Prevent changing windows/views
    pub protect_windows: bool,
}

impl PresentationProtection {
    /// Create new protection settings with no protection.
    pub fn new() -> Self {
        Self::default()
    }

    /// Set read-only recommended flag.
    pub fn with_read_only_recommended(mut self, value: bool) -> Self {
        self.read_only_recommended = value;
        self
    }

    /// Set structure protection.
    pub fn with_structure_protection(mut self, value: bool) -> Self {
        self.protect_structure = value;
        self
    }

    /// Set window protection.
    pub fn with_window_protection(mut self, value: bool) -> Self {
        self.protect_windows = value;
        self
    }

    /// Check if any protection is enabled.
    pub fn is_protected(&self) -> bool {
        self.read_only_recommended
            || self.modify.is_some()
            || self.protect_structure
            || self.protect_windows
    }

    /// Get the protection type.
    pub fn protection_type(&self) -> ProtectionType {
        if self.modify.is_some() {
            ProtectionType::ModifyPassword
        } else if self.read_only_recommended {
            ProtectionType::ReadOnlyRecommended
        } else {
            ProtectionType::None
        }
    }

    /// Borrow the validated modification verifier, when present.
    pub fn modify(&self) -> Option<&ModifyVerifier> {
        self.modify.as_ref()
    }

    /// Set the Office-compatible modify-password verifier.
    ///
    /// New verifier state is staged and committed together, so a random-source
    /// failure leaves the previous protection settings unchanged.
    pub fn set_modify_password(&mut self, password: &str) -> Result<()> {
        const SPIN_COUNT: u32 = 100_000;

        // Generate random salt (16 bytes, as commonly used by Office)
        let mut salt = [0u8; 16];
        let mut rng = SysRng;
        rng.try_fill_bytes(&mut salt).map_err(|e| {
            OoxmlError::Other(format!(
                "failed to generate random salt for modify password: {e}"
            ))
        })?;

        let hash = modify_password_hash(password, &salt, SPIN_COUNT);
        let encoded_hash = BASE64_ENGINE.encode(hash.0.as_slice());
        let encoded_salt = BASE64_ENGINE.encode(salt);

        self.modify = Some(ModifyVerifier {
            algorithm: CryptoAlgorithm::Sha512,
            spin_count: SPIN_COUNT,
            hash: encoded_hash,
            salt: encoded_salt,
        });
        Ok(())
    }

    /// Clear modify password protection.
    pub fn clear_modify_password(&mut self) {
        self.modify = None;
    }

    /// Parse protection settings from presentation properties XML.
    pub fn parse_xml(xml: &str) -> Result<Self> {
        let mut protection = Self::new();
        let xml = litchi_ooxml_common::mce::process_str(xml)?;
        let mut reader = Reader::from_str(xml.as_ref());
        reader.config_mut().trim_text(true);
        let mut verifier_seen = false;

        loop {
            match reader.read_event() {
                Ok(Event::Empty(e)) | Ok(Event::Start(e))
                    if e.local_name().as_ref() == b"modifyVerifier" =>
                {
                    if verifier_seen {
                        return Err(OoxmlError::InvalidFormat(
                            "duplicate presentation modifyVerifier".to_string(),
                        ));
                    }
                    verifier_seen = true;

                    let mut hash = None;
                    let mut salt = None;
                    let mut spin_count = None;
                    let mut algorithm = None;
                    for attr in e.attributes() {
                        let attr = attr.map_err(|error| OoxmlError::Xml(error.to_string()))?;
                        match attr.key.as_ref() {
                            // ISO-style attributes
                            b"hashValue" | b"hashData" => {
                                let value = verifier_value(attr.value.as_ref(), "hash")?;
                                set_once(&mut hash, value.to_owned(), "hash")?;
                            },
                            b"saltValue" | b"saltData" => {
                                let value = verifier_value(attr.value.as_ref(), "salt")?;
                                set_once(&mut salt, value.to_owned(), "salt")?;
                            },
                            b"spinCount" | b"spinValue" => {
                                let value = verifier_value(attr.value.as_ref(), "spin count")?;
                                let value = value.parse::<u32>().map_err(|_| {
                                    OoxmlError::InvalidFormat(
                                        "presentation modify-verifier spin count is not a u32"
                                            .to_string(),
                                    )
                                })?;
                                if !(1..=10_000_000).contains(&value) {
                                    return Err(OoxmlError::InvalidFormat(
                                        "presentation modify-verifier spin count must be between 1 and 10000000"
                                            .to_string(),
                                    ));
                                }
                                set_once(&mut spin_count, value, "spin count")?;
                            },
                            b"algorithmName" => {
                                let value = verifier_value(attr.value.as_ref(), "algorithm")?;
                                set_once(
                                    &mut algorithm,
                                    CryptoAlgorithm::from_uri(value)?,
                                    "algorithm",
                                )?;
                            },
                            b"algIdExt" => {
                                return Err(OoxmlError::InvalidFormat(
                                    "extended CryptoAPI presentation protection algorithms are unsupported"
                                        .to_string(),
                                ));
                            },
                            // Legacy SID-based form
                            b"cryptAlgorithmSid" => {
                                let value = verifier_value(attr.value.as_ref(), "algorithm SID")?;
                                let sid = value.parse::<u32>().map_err(|_| {
                                    OoxmlError::InvalidFormat(
                                        "presentation modify-verifier algorithm SID is not a u32"
                                            .to_string(),
                                    )
                                })?;
                                set_once(
                                    &mut algorithm,
                                    CryptoAlgorithm::from_sid(sid)?,
                                    "algorithm",
                                )?;
                            },
                            _ => {},
                        }
                    }

                    let hash = hash.ok_or_else(|| {
                        OoxmlError::InvalidFormat(
                            "presentation modifyVerifier is missing its hash".to_string(),
                        )
                    })?;
                    let salt = salt.ok_or_else(|| {
                        OoxmlError::InvalidFormat(
                            "presentation modifyVerifier is missing its salt".to_string(),
                        )
                    })?;
                    let spin_count = spin_count.ok_or_else(|| {
                        OoxmlError::InvalidFormat(
                            "presentation modifyVerifier is missing its spin count".to_string(),
                        )
                    })?;
                    let algorithm = algorithm.ok_or_else(|| {
                        OoxmlError::InvalidFormat(
                            "presentation modifyVerifier is missing its hash algorithm".to_string(),
                        )
                    })?;

                    if hash.len() > 128 || salt.len() > 1_368 {
                        return Err(OoxmlError::InvalidFormat(
                            "presentation modify-verifier Base64 field exceeds its bound"
                                .to_string(),
                        ));
                    }
                    let decoded_hash = BASE64_ENGINE.decode(&hash).map_err(|_| {
                        OoxmlError::InvalidFormat(
                            "presentation modify-verifier hash is not valid Base64".to_string(),
                        )
                    })?;
                    if decoded_hash.len() != algorithm.output_bytes() {
                        return Err(OoxmlError::InvalidFormat(format!(
                            "presentation modify-verifier hash has {} bytes, expected {}",
                            decoded_hash.len(),
                            algorithm.output_bytes()
                        )));
                    }
                    let decoded_salt = BASE64_ENGINE.decode(&salt).map_err(|_| {
                        OoxmlError::InvalidFormat(
                            "presentation modify-verifier salt is not valid Base64".to_string(),
                        )
                    })?;
                    if decoded_salt.is_empty() || decoded_salt.len() > 1_024 {
                        return Err(OoxmlError::InvalidFormat(
                            "presentation modify-verifier salt must contain 1 to 1024 bytes"
                                .to_string(),
                        ));
                    }

                    protection.modify = Some(ModifyVerifier {
                        algorithm,
                        spin_count,
                        hash,
                        salt,
                    });
                },
                Ok(Event::Eof) => break,
                Err(e) => return Err(OoxmlError::Xml(e.to_string())),
                _ => {},
            }
        }

        Ok(protection)
    }

    /// Generate XML for presentation.xml modification protection.
    pub fn to_xml(&self) -> String {
        let mut xml = String::new();

        if let Some(verifier) = &self.modify {
            let sid = match verifier.algorithm {
                CryptoAlgorithm::Sha1 => 4u32,
                CryptoAlgorithm::Sha256 => 12u32,
                CryptoAlgorithm::Sha384 => 13u32,
                CryptoAlgorithm::Sha512 => 14u32,
            };

            // Emit only the legacy SID-based attributes, matching
            // PowerPoint's own output for modify password protection.
            xml.push_str(&format!(
                r#"<p:modifyVerifier cryptProviderType="rsaAES" cryptAlgorithmClass="hash" cryptAlgorithmType="typeAny" cryptAlgorithmSid="{}" spinCount="{}" saltData="{}" hashData="{}"/>"#,
                sid,
                verifier.spin_count,
                verifier.salt,
                verifier.hash,
            ));
        }

        xml
    }

    /// Generate XML for presProps.xml (read-only recommended flag).
    pub fn to_pres_props_xml(&self) -> String {
        let mut xml = String::new();

        if self.read_only_recommended {
            xml.push_str(r#"<p:extLst><p:ext uri="{E76CE94A-603C-4142-B9EB-6D1370010A27}"><p14:discardImageEditData xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" val="0"/></p:ext></p:extLst>"#);
        }

        xml
    }
}

/// Slide-level protection settings.
#[derive(Debug, Clone, Default)]
pub struct SlideProtection {
    /// Prevent selection of shapes
    pub no_select: bool,
    /// Prevent moving shapes
    pub no_move: bool,
    /// Prevent resizing shapes
    pub no_resize: bool,
    /// Prevent editing shape text
    pub no_edit_text: bool,
    /// Prevent ungrouping
    pub no_ungroup: bool,
    /// Prevent changing z-order
    pub no_change_z_order: bool,
}

impl SlideProtection {
    /// Create new slide protection with no restrictions.
    pub fn new() -> Self {
        Self::default()
    }

    /// Set all protections on.
    pub fn protect_all(mut self) -> Self {
        self.no_select = true;
        self.no_move = true;
        self.no_resize = true;
        self.no_edit_text = true;
        self.no_ungroup = true;
        self.no_change_z_order = true;
        self
    }

    /// Check if any protection is enabled.
    pub fn is_protected(&self) -> bool {
        self.no_select
            || self.no_move
            || self.no_resize
            || self.no_edit_text
            || self.no_ungroup
            || self.no_change_z_order
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_protection_type() {
        let mut prot = PresentationProtection::new();
        assert_eq!(prot.protection_type(), ProtectionType::None);

        prot.read_only_recommended = true;
        assert_eq!(prot.protection_type(), ProtectionType::ReadOnlyRecommended);

        prot.modify = Some(ModifyVerifier {
            algorithm: CryptoAlgorithm::Sha512,
            spin_count: 1,
            hash: BASE64_ENGINE.encode([0; 64]),
            salt: BASE64_ENGINE.encode([0; 16]),
        });
        assert_eq!(prot.protection_type(), ProtectionType::ModifyPassword);
    }

    #[test]
    fn test_crypto_algorithm() {
        assert!(matches!(
            CryptoAlgorithm::from_uri("http://www.w3.org/2001/04/xmlenc#sha256"),
            Ok(CryptoAlgorithm::Sha256)
        ));
        assert!(CryptoAlgorithm::from_uri("vendor-sha512-ish").is_err());
    }

    #[test]
    fn modify_password_hashing_does_not_retain_or_debug_the_password() {
        let password = "unique plaintext password 9bQ!";
        let mut protection = PresentationProtection::new();

        protection.set_modify_password(password).unwrap();

        let verifier = protection.modify().expect("modify verifier");
        assert_eq!(verifier.algorithm(), CryptoAlgorithm::Sha512);
        assert_eq!(verifier.spins(), 100_000);
        assert_eq!(BASE64_ENGINE.decode(verifier.hash()).unwrap().len(), 64);
        assert_eq!(BASE64_ENGINE.decode(verifier.salt()).unwrap().len(), 16);
        assert!(!format!("{protection:?}").contains(password));
        assert!(!format!("{:?}", protection.clone()).contains(password));
    }

    #[test]
    fn modify_password_hash_matches_office_password_then_salt_order() {
        let salt: Vec<u8> = (0..16).collect();
        let hash = modify_password_hash("Päss😀", &salt, 2);

        assert_eq!(
            BASE64_ENGINE.encode(hash.0.as_slice()),
            "3ACFcYR0/M+PsEwOXR4/mcgYsTN1VMXMunIrbpt1lY+1Kal3nCkZOJjIEw+LWRlQzI3rL5HZnVIoL87I6R8tNw=="
        );
    }

    #[test]
    fn modify_verifier_parser_requires_exact_typed_metadata() {
        let hash = BASE64_ENGINE.encode([0x5a; 64]);
        let salt = BASE64_ENGINE.encode([0xa5; 16]);
        let xml = format!(
            r#"<p:modifyVerifier xmlns:p="urn:p" cryptAlgorithmSid="14" spinCount="100000" saltData="{salt}" hashData="{hash}"/>"#
        );
        let parsed = PresentationProtection::parse_xml(&xml).expect("valid verifier");
        let verifier = parsed.modify().expect("parsed verifier");
        assert_eq!(verifier.algorithm(), CryptoAlgorithm::Sha512);
        assert_eq!(verifier.spins(), 100_000);

        for malformed in [
            format!(
                r#"<p:modifyVerifier xmlns:p="urn:p" cryptAlgorithmSid="99" spinCount="100000" saltData="{salt}" hashData="{hash}"/>"#
            ),
            format!(
                r#"<p:modifyVerifier xmlns:p="urn:p" cryptAlgorithmSid="14" spinCount="many" saltData="{salt}" hashData="{hash}"/>"#
            ),
            format!(
                r#"<p:modifyVerifier xmlns:p="urn:p" cryptAlgorithmSid="14" spinCount="100000" saltData="{salt}" hashData="not-base64"/>"#
            ),
            format!(
                r#"<p:modifyVerifier xmlns:p="urn:p" cryptAlgorithmSid="14" spinCount="100000" saltData="{salt}"/>"#
            ),
        ] {
            assert!(PresentationProtection::parse_xml(&malformed).is_err());
        }
    }

    #[test]
    fn test_slide_protection() {
        let prot = SlideProtection::new().protect_all();
        assert!(prot.is_protected());
        assert!(prot.no_select);
        assert!(prot.no_move);
    }
}
