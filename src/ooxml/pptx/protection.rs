//! Presentation protection support for PowerPoint presentations.
//!
//! This module provides support for password protection, read-only mode,
//! and other security settings for presentations.

use crate::ooxml::error::{OoxmlError, Result};
use base64::Engine;
use base64::engine::general_purpose::STANDARD as BASE64_ENGINE;
use quick_xml::Reader;
use quick_xml::events::Event;
use rand::TryRngCore;
use rand::rngs::OsRng;
use sha2::{Digest, Sha512};

/// Type of protection applied to a presentation.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ProtectionType {
    /// No protection
    None,
    /// Read-only recommended (shows a dialog but can be bypassed)
    ReadOnlyRecommended,
    /// Password required to modify
    ModifyPassword,
    /// Password required to open
    OpenPassword,
    /// Both open and modify passwords required
    FullProtection,
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

    /// Parse from URI string.
    pub fn from_uri(uri: &str) -> Self {
        if uri.contains("sha512") {
            CryptoAlgorithm::Sha512
        } else if uri.contains("sha384") {
            CryptoAlgorithm::Sha384
        } else if uri.contains("sha256") {
            CryptoAlgorithm::Sha256
        } else {
            CryptoAlgorithm::Sha1
        }
    }
}

/// Protection settings for a presentation.
#[derive(Debug, Clone, Default)]
pub struct PresentationProtection {
    /// Whether the presentation is marked as read-only recommended
    pub read_only_recommended: bool,
    /// Whether modification requires a password
    pub modify_password_protected: bool,
    /// Hashed modify password (base64 encoded)
    pub modify_password_hash: Option<String>,
    /// Salt for modify password (base64 encoded)
    pub modify_password_salt: Option<String>,
    /// Spin count for modify password hashing
    pub modify_spin_count: u32,
    /// Algorithm used for modify password
    pub modify_algorithm: CryptoAlgorithm,
    /// Whether opening requires a password (handled by encryption, not here)
    pub open_password_protected: bool,
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
            || self.modify_password_protected
            || self.open_password_protected
            || self.protect_structure
            || self.protect_windows
    }

    /// Get the protection type.
    pub fn protection_type(&self) -> ProtectionType {
        if self.open_password_protected && self.modify_password_protected {
            ProtectionType::FullProtection
        } else if self.open_password_protected {
            ProtectionType::OpenPassword
        } else if self.modify_password_protected {
            ProtectionType::ModifyPassword
        } else if self.read_only_recommended {
            ProtectionType::ReadOnlyRecommended
        } else {
            ProtectionType::None
        }
    }

    /// Set modify password (hashes the password).
    /// Note: This is a simplified implementation. Real implementation would use
    /// proper OOXML password hashing algorithm.
    pub fn set_modify_password(&mut self, password: &str) -> Result<()> {
        // Enable modify protection and configure algorithm parameters
        self.modify_password_protected = true;
        self.modify_spin_count = 100000;
        self.modify_algorithm = CryptoAlgorithm::Sha512;

        // Generate random salt (16 bytes, as commonly used by Office)
        let mut salt = [0u8; 16];
        let mut rng = OsRng;
        rng.try_fill_bytes(&mut salt).map_err(|e| {
            OoxmlError::Other(format!(
                "failed to generate random salt for modify password: {e}"
            ))
        })?;

        // Encode password as UTF-16LE bytes
        let mut pw_bytes = Vec::with_capacity(password.len() * 2);
        for ch in password.encode_utf16() {
            pw_bytes.extend_from_slice(&ch.to_le_bytes());
        }

        // Initial hash: H[init] = H(salt || password)
        let mut hasher = Sha512::new();
        hasher.update(salt);
        hasher.update(&pw_bytes);
        let mut hash = hasher.finalize().to_vec();

        // Iterative hashing: H[n] = H(H[n-1] || count_le_u32), for spinCount cycles
        for i in 0..self.modify_spin_count {
            let mut hasher = Sha512::new();
            hasher.update(&hash);
            hasher.update(i.to_le_bytes());
            hash = hasher.finalize().to_vec();
        }

        self.modify_password_hash = Some(BASE64_ENGINE.encode(&hash));
        self.modify_password_salt = Some(BASE64_ENGINE.encode(salt));
        Ok(())
    }

    /// Clear modify password protection.
    pub fn clear_modify_password(&mut self) {
        self.modify_password_protected = false;
        self.modify_password_hash = None;
        self.modify_password_salt = None;
    }

    pub fn set_open_password(&mut self, _password: &str) -> Result<()> {
        Err(OoxmlError::Other(
            "PPTX open-password protection (file encryption) is not implemented yet; only modify-password protection is currently supported.".to_string(),
        ))
    }

    pub fn clear_open_password(&mut self) {
        self.open_password_protected = false;
    }

    /// Parse protection settings from presentation properties XML.
    pub fn parse_xml(xml: &str) -> Result<Self> {
        let mut protection = Self::new();
        let mut reader = Reader::from_str(xml);
        reader.config_mut().trim_text(true);

        loop {
            match reader.read_event() {
                Ok(Event::Empty(e)) | Ok(Event::Start(e)) => {
                    if e.local_name().as_ref() == b"modifyVerifier" {
                        protection.modify_password_protected = true;
                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                // ISO-style attributes
                                b"hashValue" | b"hashData" => {
                                    protection.modify_password_hash = Some(
                                        std::str::from_utf8(&attr.value).unwrap_or("").to_string(),
                                    );
                                },
                                b"saltValue" | b"saltData" => {
                                    protection.modify_password_salt = Some(
                                        std::str::from_utf8(&attr.value).unwrap_or("").to_string(),
                                    );
                                },
                                b"spinCount" | b"spinValue" => {
                                    protection.modify_spin_count = std::str::from_utf8(&attr.value)
                                        .ok()
                                        .and_then(|s| s.parse().ok())
                                        .unwrap_or(100000);
                                },
                                b"algorithmName" | b"algIdExt" => {
                                    if let Ok(uri) = std::str::from_utf8(&attr.value) {
                                        protection.modify_algorithm =
                                            CryptoAlgorithm::from_uri(uri);
                                    }
                                },
                                // Legacy SID-based form
                                b"cryptAlgorithmSid" => {
                                    if let Ok(text) = std::str::from_utf8(&attr.value)
                                        && let Ok(sid) = text.parse::<u32>()
                                    {
                                        protection.modify_algorithm = match sid {
                                            4 => CryptoAlgorithm::Sha1,
                                            12 => CryptoAlgorithm::Sha256,
                                            13 => CryptoAlgorithm::Sha384,
                                            14 => CryptoAlgorithm::Sha512,
                                            _ => protection.modify_algorithm,
                                        };
                                    }
                                },
                                _ => {},
                            }
                        }
                    }
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

        if self.modify_password_protected
            && let (Some(hash), Some(salt)) =
                (&self.modify_password_hash, &self.modify_password_salt)
        {
            let sid = match self.modify_algorithm {
                CryptoAlgorithm::Sha1 => Some(4u32),
                CryptoAlgorithm::Sha256 => Some(12u32),
                CryptoAlgorithm::Sha384 => Some(13u32),
                CryptoAlgorithm::Sha512 => Some(14u32),
            };

            if let Some(sid) = sid {
                // Emit only the legacy SID-based attributes, matching
                // PowerPoint's own output for modify password protection.
                // Example (from a PowerPoint-generated file):
                // <p:modifyVerifier cryptProviderType="rsaAES" cryptAlgorithmClass="hash"
                //                  cryptAlgorithmType="typeAny" cryptAlgorithmSid="14"
                //                  spinCount="100000" saltData="..." hashData="..."/>
                xml.push_str(&format!(
                    r#"<p:modifyVerifier cryptProviderType="rsaAES" cryptAlgorithmClass="hash" cryptAlgorithmType="typeAny" cryptAlgorithmSid="{}" spinCount="{}" saltData="{}" hashData="{}"/>"#,
                    sid,
                    self.modify_spin_count,
                    salt,
                    hash,
                ));
            } else {
                // Fallback: ISO-style only, in case we ever have an
                // algorithm without a corresponding SID.
                xml.push_str(&format!(
                    r#"<p:modifyVerifier algorithmName="SHA-512" hashValue="{}" saltValue="{}" spinCount="{}"/>"#,
                    hash,
                    salt,
                    self.modify_spin_count,
                ));
            }
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

        prot.modify_password_protected = true;
        assert_eq!(prot.protection_type(), ProtectionType::ModifyPassword);
    }

    #[test]
    fn test_crypto_algorithm() {
        assert_eq!(
            CryptoAlgorithm::from_uri("http://www.w3.org/2001/04/xmlenc#sha256"),
            CryptoAlgorithm::Sha256
        );
    }

    #[test]
    fn test_slide_protection() {
        let prot = SlideProtection::new().protect_all();
        assert!(prot.is_protected());
        assert!(prot.no_select);
        assert!(prot.no_move);
    }
}
