//! Typed, namespace-aware parsing of `META-INF/manifest.xml`.

use crate::package as common_package;
use base64::Engine as _;
use base64::engine::general_purpose::STANDARD as BASE64_STANDARD;
use litchi_core::{Error, Result};
use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use std::borrow::Cow;
use std::collections::HashMap;
use std::num::NonZeroU32;

const MANIFEST_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:manifest:1.0";
const LOEXT_NAMESPACE: &[u8] =
    b"urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0";
pub(crate) const MAX_PBKDF2_ITERATIONS: u32 = 10_000_000;
const MAX_SALT_BYTES: usize = 1_024;
pub(crate) const MAX_ARGON2_ITERATIONS: u32 = 10;
pub(crate) const MAX_ARGON2_MEMORY_KIB: u32 = 262_144;
pub(crate) const MAX_ARGON2_LANES: u32 = 16;
pub(crate) const MAX_ARGON2_WORK_KIB_PASSES: u64 = 1_048_576;

fn try_copy_bytes(value: &[u8], resource: &'static str) -> Result<Vec<u8>> {
    let mut output = Vec::new();
    output
        .try_reserve_exact(value.len())
        .map_err(|source| Error::Allocation { resource, source })?;
    output.extend_from_slice(value);
    Ok(output)
}

fn try_copy_string(value: &str, resource: &'static str) -> Result<String> {
    let mut output = String::new();
    output
        .try_reserve_exact(value.len())
        .map_err(|source| Error::Allocation { resource, source })?;
    output.push_str(value);
    Ok(output)
}

/// ODF manifest (`META-INF/manifest.xml`).
#[derive(Debug, Clone)]
pub struct Manifest {
    pub mimetype: String,
    pub entries: HashMap<String, ManifestEntry>,
}

/// One file entry in the ODF manifest.
#[derive(Debug, Clone)]
pub struct ManifestEntry {
    pub media_type: String,
    pub size: Option<u64>,
    pub encryption: Option<ManifestEncryption>,
}

/// Complete password-encryption descriptor for one package entry.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ManifestEncryption {
    pub checksum: Option<ManifestChecksum>,
    pub algorithm: ManifestEncryptionAlgorithm,
    pub start_key: ManifestStartKeyGeneration,
    pub key_derivation: ManifestKeyDerivation,
}

/// Integrity checksum over the first 1024 decrypted compressed bytes.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ManifestChecksum {
    pub algorithm: ManifestChecksumAlgorithm,
    pub value: Vec<u8>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ManifestChecksumAlgorithm {
    Sha1First1024,
    Sha256First1024,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum ManifestEncryptionAlgorithm {
    Aes128Cbc { iv: [u8; 16] },
    Aes192Cbc { iv: [u8; 16] },
    Aes256Cbc { iv: [u8; 16] },
    Aes128Gcm { iv: [u8; 12] },
    Aes192Gcm { iv: [u8; 12] },
    Aes256Gcm { iv: [u8; 12] },
    BlowfishCfb8 { iv: [u8; 8] },
}

impl ManifestEncryptionAlgorithm {
    pub(crate) const fn accepts_key_size(&self, key_size: u16) -> bool {
        match self {
            Self::Aes128Cbc { .. } | Self::Aes128Gcm { .. } => key_size == 16,
            Self::Aes192Cbc { .. } | Self::Aes192Gcm { .. } => key_size == 24,
            Self::Aes256Cbc { .. } | Self::Aes256Gcm { .. } => key_size == 32,
            Self::BlowfishCfb8 { .. } => key_size >= 4 && key_size <= 56,
        }
    }

    const fn key_size_description(&self) -> &'static str {
        match self {
            Self::Aes128Cbc { .. } | Self::Aes128Gcm { .. } => "16",
            Self::Aes192Cbc { .. } | Self::Aes192Gcm { .. } => "24",
            Self::Aes256Cbc { .. } | Self::Aes256Gcm { .. } => "32",
            Self::BlowfishCfb8 { .. } => "4 through 56",
        }
    }

    pub(crate) const fn is_aead(&self) -> bool {
        matches!(
            self,
            Self::Aes128Gcm { .. } | Self::Aes192Gcm { .. } | Self::Aes256Gcm { .. }
        )
    }

    pub(crate) const fn fixed_key_size(&self) -> Option<u16> {
        match self {
            Self::Aes128Cbc { .. } | Self::Aes128Gcm { .. } => Some(16),
            Self::Aes192Cbc { .. } | Self::Aes192Gcm { .. } => Some(24),
            Self::Aes256Cbc { .. } | Self::Aes256Gcm { .. } => Some(32),
            Self::BlowfishCfb8 { .. } => None,
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ManifestStartKeyGeneration {
    Sha1,
    Sha256,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum ManifestKeyDerivation {
    Pbkdf2 {
        salt: Vec<u8>,
        iterations: NonZeroU32,
        key_size: u16,
    },
    Argon2id {
        salt: Vec<u8>,
        iterations: NonZeroU32,
        memory_kib: NonZeroU32,
        lanes: NonZeroU32,
        key_size: Option<u16>,
    },
}

#[derive(Default)]
struct PartialEncryption {
    checksum: Option<ManifestChecksum>,
    algorithm: Option<ManifestEncryptionAlgorithm>,
    start_key: Option<ManifestStartKeyGeneration>,
    key_derivation: Option<ManifestKeyDerivation>,
}

#[derive(Default)]
struct TypedManifestObserver {
    current_path: Option<String>,
    current_size: Option<u64>,
    encryption: Option<PartialEncryption>,
    completed_encryption: Option<ManifestEncryption>,
    encrypted_entries: HashMap<String, ManifestEncryption>,
}

impl common_package::ManifestScanObserver for TypedManifestObserver {
    fn event<'namespace, 'event>(
        &mut self,
        reader: &NsReader<&[u8]>,
        namespace: &ResolveResult<'namespace>,
        event: &Event<'event>,
        file_entry_path: Option<&str>,
        file_entry_size: Option<u64>,
    ) -> Result<()> {
        match event {
            Event::Start(element) if is_manifest_element(namespace, element, b"file-entry") => {
                if self.current_path.is_some() {
                    return Err(Error::InvalidFormat(
                        "Nested manifest file entries are invalid".to_string(),
                    ));
                }
                let path = file_entry_path.ok_or_else(|| {
                    Error::InvalidFormat("Manifest file entry has no full path".to_string())
                })?;
                self.current_path = Some(try_copy_string(path, "ODF encrypted manifest path")?);
                self.current_size = file_entry_size;
            },
            Event::Empty(element) if is_manifest_element(namespace, element, b"file-entry") => {
                // Neutral file-entry validation and indexing are owned by
                // litchi-odf-common. There is no encryption subtree to
                // inspect on an empty element.
            },
            Event::Start(element)
                if is_manifest_element(namespace, element, b"encryption-data") =>
            {
                if self.current_path.is_none() || self.encryption.is_some() {
                    return Err(Error::InvalidFormat(
                        "Misplaced or duplicate manifest encryption data".to_string(),
                    ));
                }
                self.encryption = Some(PartialEncryption {
                    checksum: parse_encryption_checksum(reader, element)?,
                    ..PartialEncryption::default()
                });
            },
            Event::Empty(element)
                if is_manifest_element(namespace, element, b"encryption-data") =>
            {
                if self.current_path.is_none() || self.encryption.is_some() {
                    return Err(Error::InvalidFormat(
                        "Misplaced or duplicate manifest encryption data".to_string(),
                    ));
                }
                parse_encryption_checksum(reader, element)?;
                return Err(Error::InvalidFormat(
                    "ODF package has encrypted entries with incomplete encryption metadata"
                        .to_string(),
                ));
            },
            Event::Start(element) | Event::Empty(element)
                if is_manifest_element(namespace, element, b"algorithm") =>
            {
                let algorithm = parse_algorithm(reader, element)?;
                set_once(
                    &mut encryption_mut(&mut self.encryption)?.algorithm,
                    algorithm,
                    "algorithm",
                )?;
            },
            Event::Start(element) | Event::Empty(element)
                if is_manifest_element(namespace, element, b"start-key-generation") =>
            {
                let start_key = parse_start_key(reader, element)?;
                set_once(
                    &mut encryption_mut(&mut self.encryption)?.start_key,
                    start_key,
                    "start-key-generation",
                )?;
            },
            Event::Start(element) | Event::Empty(element)
                if is_manifest_element(namespace, element, b"key-derivation") =>
            {
                let key_derivation = parse_key_derivation(reader, element)?;
                set_once(
                    &mut encryption_mut(&mut self.encryption)?.key_derivation,
                    key_derivation,
                    "key-derivation",
                )?;
            },
            Event::End(element)
                if namespace_is_manifest(namespace)
                    && element.local_name().as_ref() == b"encryption-data" =>
            {
                let path = self.current_path.as_ref().ok_or_else(|| {
                    Error::InvalidFormat("Encryption data has no file entry".to_string())
                })?;
                let descriptor = finish_encryption(self.encryption.take().ok_or_else(|| {
                    Error::InvalidFormat("Unexpected encryption-data end".to_string())
                })?)?;
                if self.current_size.is_none() {
                    return Err(Error::InvalidFormat(format!(
                        "Encrypted manifest entry '{path}' has no plaintext size"
                    )));
                }
                self.completed_encryption = Some(descriptor);
            },
            Event::End(element)
                if namespace_is_manifest(namespace)
                    && element.local_name().as_ref() == b"file-entry" =>
            {
                if self.encryption.is_some() {
                    return Err(Error::InvalidFormat(
                        "Unterminated manifest encryption data".to_string(),
                    ));
                }
                if let Some(descriptor) = self.completed_encryption.take() {
                    let path = self.current_path.take().ok_or_else(|| {
                        Error::InvalidFormat("Manifest entry has no file path".to_string())
                    })?;
                    self.encrypted_entries.insert(path, descriptor);
                } else {
                    self.current_path = None;
                }
                self.current_size = None;
            },
            Event::Eof => {
                if self.current_path.is_some() || self.encryption.is_some() {
                    return Err(Error::InvalidFormat(
                        "Incomplete manifest file entry".to_string(),
                    ));
                }
            },
            Event::Start(_)
            | Event::End(_)
            | Event::Empty(_)
            | Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::DocType(_)
            | Event::GeneralRef(_) => {},
        }
        Ok(())
    }
}

impl Manifest {
    /// Parse a `META-INF/manifest.xml` document.
    ///
    /// # Errors
    ///
    /// Returns an error when the XML is malformed or its ODF manifest and
    /// encryption descriptors are invalid.
    pub fn parse(xml: &str) -> Result<Self> {
        let mut observer = TypedManifestObserver::default();
        let (
            common_package::Manifest {
                mimetype,
                entries: neutral_entries,
            },
            observer_error,
        ) = common_package::parse_manifest_with_observer(xml, &mut observer)?;
        if let Some(error) = observer_error {
            return Err(error);
        }
        let mut entries: HashMap<String, ManifestEntry> = HashMap::new();
        entries
            .try_reserve(neutral_entries.len())
            .map_err(|source| Error::Allocation {
                resource: "ODF typed manifest entries",
                source,
            })?;
        for (path, entry) in neutral_entries {
            entries.insert(
                path,
                ManifestEntry {
                    media_type: entry.media_type,
                    size: entry.size,
                    encryption: observer.encrypted_entries.remove(&path),
                },
            );
        }
        Ok(Self { mimetype, entries })
    }

    #[must_use]
    pub fn get_media_type(&self, path: &str) -> Option<&str> {
        self.entries
            .get(path)
            .map(|entry| entry.media_type.as_str())
    }

    #[must_use]
    pub fn has_path(&self, path: &str) -> bool {
        self.entries.contains_key(path)
    }

    pub fn paths(&self) -> impl Iterator<Item = &str> {
        self.entries.keys().map(String::as_str)
    }

    #[must_use]
    pub fn get_entry(&self, path: &str) -> Option<&ManifestEntry> {
        self.entries.get(path)
    }

    #[must_use]
    pub fn has_encrypted_entries(&self) -> bool {
        self.entries
            .values()
            .any(|entry| entry.encryption.is_some())
    }
}

fn is_manifest_element(
    namespace: &ResolveResult<'_>,
    element: &BytesStart<'_>,
    local: &[u8],
) -> bool {
    namespace_is_manifest(namespace) && element.local_name().as_ref() == local
}

fn namespace_is_manifest(namespace: &ResolveResult<'_>) -> bool {
    matches!(namespace, ResolveResult::Bound(Namespace(uri)) if *uri == MANIFEST_NAMESPACE)
}

fn manifest_attributes(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
) -> Result<HashMap<Vec<u8>, String>> {
    let mut values = HashMap::new();
    for raw_attribute in element.attributes() {
        let attribute = raw_attribute.map_err(|error| {
            Error::InvalidFormat(format!("Invalid manifest attribute: {error}"))
        })?;
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        if !matches!(namespace, ResolveResult::Bound(Namespace(uri)) if uri == MANIFEST_NAMESPACE) {
            continue;
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
            .map_err(|error| {
                Error::InvalidFormat(format!("Invalid manifest attribute value: {error}"))
            })?;
        if values.contains_key(local.as_ref()) {
            return Err(Error::InvalidFormat(
                "Duplicate manifest attribute".to_string(),
            ));
        }
        values.try_reserve(1).map_err(|source| Error::Allocation {
            resource: "ODF typed manifest attributes",
            source,
        })?;
        let key = try_copy_bytes(local.as_ref(), "ODF typed manifest attribute name")?;
        let value = match value {
            Cow::Owned(value) => value,
            Cow::Borrowed(value) => try_copy_string(value, "ODF typed manifest attribute value")?,
        };
        values.insert(key, value);
    }
    Ok(values)
}

#[allow(
    clippy::type_complexity,
    reason = "Manifest and LibreOffice extension attributes must be returned separately."
)]
#[allow(
    clippy::wildcard_enum_match_arm,
    reason = "quick-xml's non-exhaustive ResolveResult needs forward-compatible namespace handling."
)]
fn manifest_and_loext_attributes(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
) -> Result<(HashMap<Vec<u8>, String>, HashMap<Vec<u8>, String>)> {
    let mut manifest = HashMap::new();
    let mut loext = HashMap::new();
    for raw_attribute in element.attributes() {
        let attribute = raw_attribute.map_err(|error| {
            Error::InvalidFormat(format!("Invalid manifest attribute: {error}"))
        })?;
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        let target = match namespace {
            ResolveResult::Bound(Namespace(uri)) if uri == MANIFEST_NAMESPACE => &mut manifest,
            ResolveResult::Bound(Namespace(uri)) if uri == LOEXT_NAMESPACE => &mut loext,
            _ => continue,
        };
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
            .map_err(|error| {
                Error::InvalidFormat(format!("Invalid manifest attribute value: {error}"))
            })?;
        if target.contains_key(local.as_ref()) {
            return Err(Error::InvalidFormat(
                "Duplicate manifest key-derivation attribute".to_string(),
            ));
        }
        target.try_reserve(1).map_err(|source| Error::Allocation {
            resource: "ODF key-derivation attributes",
            source,
        })?;
        let key = try_copy_bytes(local.as_ref(), "ODF key-derivation attribute name")?;
        let value = match value {
            Cow::Owned(value) => value,
            Cow::Borrowed(value) => try_copy_string(value, "ODF key-derivation attribute value")?,
        };
        target.insert(key, value);
    }
    Ok((manifest, loext))
}

fn required<'a>(attributes: &'a HashMap<Vec<u8>, String>, name: &[u8]) -> Result<&'a str> {
    attributes.get(name).map(String::as_str).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "Missing manifest attribute '{}'",
            String::from_utf8_lossy(name)
        ))
    })
}

fn decode_base64(value: &str, field: &str) -> Result<Vec<u8>> {
    let mut output = Vec::new();
    output
        .try_reserve_exact(base64::decoded_len_estimate(value.len()))
        .map_err(|source| Error::Allocation {
            resource: "ODF manifest Base64 value",
            source,
        })?;
    BASE64_STANDARD
        .decode_vec(value, &mut output)
        .map_err(|error| {
            Error::InvalidFormat(format!("Invalid Base64 in manifest {field}: {error}"))
        })?;
    Ok(output)
}

fn parse_checksum(reader: &NsReader<&[u8]>, element: &BytesStart<'_>) -> Result<ManifestChecksum> {
    let attributes = manifest_attributes(reader, element)?;
    let algorithm = match required(&attributes, b"checksum-type")? {
        "SHA1/1K" | "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0#sha1-1k" => {
            ManifestChecksumAlgorithm::Sha1First1024
        },
        "SHA256/1K" | "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0#sha256-1k" => {
            ManifestChecksumAlgorithm::Sha256First1024
        },
        value => {
            return Err(Error::InvalidFormat(format!(
                "Unsupported ODF checksum algorithm '{value}'"
            )));
        },
    };
    let value = decode_base64(required(&attributes, b"checksum")?, "checksum")?;
    let expected = match algorithm {
        ManifestChecksumAlgorithm::Sha1First1024 => 20,
        ManifestChecksumAlgorithm::Sha256First1024 => 32,
    };
    if value.len() != expected {
        return Err(Error::InvalidFormat(format!(
            "ODF checksum has length {}, expected {expected}",
            value.len()
        )));
    }
    Ok(ManifestChecksum { algorithm, value })
}

fn parse_encryption_checksum(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
) -> Result<Option<ManifestChecksum>> {
    let attributes = manifest_attributes(reader, element)?;
    let has_type = attributes.contains_key(b"checksum-type".as_slice());
    let has_value = attributes.contains_key(b"checksum".as_slice());
    if has_type != has_value {
        return Err(Error::InvalidFormat(
            "ODF encryption metadata must contain both checksum attributes or neither".to_string(),
        ));
    }
    if !has_type {
        return Ok(None);
    }
    parse_checksum(reader, element).map(Some).map_err(|error| {
        Error::InvalidFormat(format!(
            "ODF package has encrypted entries with invalid metadata: {error}"
        ))
    })
}

fn parse_algorithm(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
) -> Result<ManifestEncryptionAlgorithm> {
    let attributes = manifest_attributes(reader, element)?;
    let name = required(&attributes, b"algorithm-name")?;
    let iv = decode_base64(
        required(&attributes, b"initialisation-vector")?,
        "initialisation-vector",
    )?;
    match name {
        "http://www.w3.org/2001/04/xmlenc#aes128-cbc" => {
            Ok(ManifestEncryptionAlgorithm::Aes128Cbc {
                iv: fixed_iv(iv, 16, "AES-CBC")?,
            })
        },
        "http://www.w3.org/2001/04/xmlenc#aes192-cbc" => {
            Ok(ManifestEncryptionAlgorithm::Aes192Cbc {
                iv: fixed_iv(iv, 16, "AES-CBC")?,
            })
        },
        "http://www.w3.org/2001/04/xmlenc#aes256-cbc" => {
            Ok(ManifestEncryptionAlgorithm::Aes256Cbc {
                iv: fixed_iv(iv, 16, "AES-CBC")?,
            })
        },
        "http://www.w3.org/2009/xmlenc11#aes128-gcm" => {
            Ok(ManifestEncryptionAlgorithm::Aes128Gcm {
                iv: fixed_iv(iv, 12, "AES-GCM")?,
            })
        },
        "http://www.w3.org/2009/xmlenc11#aes192-gcm" => {
            Ok(ManifestEncryptionAlgorithm::Aes192Gcm {
                iv: fixed_iv(iv, 12, "AES-GCM")?,
            })
        },
        "http://www.w3.org/2009/xmlenc11#aes256-gcm" => {
            Ok(ManifestEncryptionAlgorithm::Aes256Gcm {
                iv: fixed_iv(iv, 12, "AES-GCM")?,
            })
        },
        "Blowfish CFB" | "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0#blowfish" => {
            Ok(ManifestEncryptionAlgorithm::BlowfishCfb8 {
                iv: fixed_iv(iv, 8, "Blowfish CFB")?,
            })
        },
        _ => Err(Error::InvalidFormat(format!(
            "Unsupported ODF encryption algorithm '{name}'"
        ))),
    }
}

fn fixed_iv<const N: usize>(
    decoded_iv: Vec<u8>,
    expected: usize,
    algorithm: &str,
) -> Result<[u8; N]> {
    decoded_iv.try_into().map_err(|invalid_iv: Vec<u8>| {
        Error::InvalidFormat(format!(
            "{algorithm} IV has length {}, expected {expected}",
            invalid_iv.len()
        ))
    })
}

fn parse_start_key(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
) -> Result<ManifestStartKeyGeneration> {
    let attributes = manifest_attributes(reader, element)?;
    let (algorithm, expected_size) = match required(&attributes, b"start-key-generation-name")? {
        "SHA1" | "http://www.w3.org/2000/09/xmldsig#sha1" => (ManifestStartKeyGeneration::Sha1, 20),
        "http://www.w3.org/2001/04/xmlenc#sha256" | "http://www.w3.org/2000/09/xmldsig#sha256" => {
            (ManifestStartKeyGeneration::Sha256, 32)
        },
        value => {
            return Err(Error::InvalidFormat(format!(
                "Unsupported ODF start-key algorithm '{value}'"
            )));
        },
    };
    let key_size = required(&attributes, b"key-size")?
        .parse::<usize>()
        .map_err(|error| Error::InvalidFormat(format!("Invalid start-key size: {error}")))?;
    if key_size != expected_size {
        return Err(Error::InvalidFormat(format!(
            "Start-key size {key_size} does not match its digest"
        )));
    }
    Ok(algorithm)
}

fn parse_key_derivation(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
) -> Result<ManifestKeyDerivation> {
    let (attributes, loext_attributes) = manifest_and_loext_attributes(reader, element)?;
    let name = required(&attributes, b"key-derivation-name")?;
    let salt = decode_base64(required(&attributes, b"salt")?, "salt")?;
    match name {
        "PBKDF2" | "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0#pbkdf2" => {
            if salt.is_empty() || salt.len() > MAX_SALT_BYTES {
                return Err(Error::InvalidFormat(format!(
                    "ODF PBKDF2 salt length {} is outside supported limits",
                    salt.len()
                )));
            }
            let iterations = parse_nonzero_u32(
                required(&attributes, b"iteration-count")?,
                "PBKDF2 iteration count",
            )?;
            if iterations.get() > MAX_PBKDF2_ITERATIONS {
                return Err(Error::InvalidFormat(format!(
                    "PBKDF2 iteration count {} exceeds the supported limit",
                    iterations.get()
                )));
            }
            let key_size = parse_optional_key_size(&attributes)?.unwrap_or(16);
            if !(4..=56).contains(&key_size) {
                return Err(Error::InvalidFormat(format!(
                    "Unsupported PBKDF2 derived key size {key_size}"
                )));
            }
            Ok(ManifestKeyDerivation::Pbkdf2 {
                salt,
                iterations,
                key_size,
            })
        },
        "urn:oasis:names:tc:opendocument:xmlns:manifest:1.5#argon2id" => {
            if has_argon2_parameters(&loext_attributes) {
                return Err(Error::InvalidFormat(
                    "Standard Argon2id metadata mixes LibreOffice extension attributes".to_string(),
                ));
            }
            parse_argon2id(&attributes, salt)
        },
        "urn:org:documentfoundation:names:experimental:office:manifest:argon2id" => {
            if has_argon2_parameters(&attributes) {
                return Err(Error::InvalidFormat(
                    "Experimental Argon2id metadata mixes standard manifest attributes".to_string(),
                ));
            }
            let mut parameters = loext_attributes;
            if let Some(key_size) = attributes.get(b"key-size".as_slice()) {
                parameters
                    .try_reserve(1)
                    .map_err(|source| Error::Allocation {
                        resource: "ODF Argon2id parameters",
                        source,
                    })?;
                parameters.insert(
                    try_copy_bytes(b"key-size", "ODF Argon2id parameter name")?,
                    try_copy_string(key_size, "ODF Argon2id key size")?,
                );
            }
            parse_argon2id(&parameters, salt)
        },
        _ => Err(Error::InvalidFormat(format!(
            "Unsupported ODF key derivation '{name}'"
        ))),
    }
}

fn parse_optional_key_size(attributes: &HashMap<Vec<u8>, String>) -> Result<Option<u16>> {
    attributes
        .get(b"key-size".as_slice())
        .map(|value| {
            value
                .parse::<u16>()
                .map_err(|error| Error::InvalidFormat(format!("Invalid derived key size: {error}")))
        })
        .transpose()
}

fn parse_nonzero_u32(value: &str, field: &str) -> Result<NonZeroU32> {
    let parsed = value
        .parse::<u32>()
        .map_err(|error| Error::InvalidFormat(format!("Invalid {field}: {error}")))?;
    NonZeroU32::new(parsed).ok_or_else(|| Error::InvalidFormat(format!("Invalid {field}: zero")))
}

fn has_argon2_parameters(attributes: &HashMap<Vec<u8>, String>) -> bool {
    [
        b"argon2-iterations".as_slice(),
        b"argon2-memory".as_slice(),
        b"argon2-lanes".as_slice(),
    ]
    .iter()
    .any(|name| attributes.contains_key(*name))
}

fn parse_argon2id(
    parameters: &HashMap<Vec<u8>, String>,
    salt: Vec<u8>,
) -> Result<ManifestKeyDerivation> {
    if !(8..=64).contains(&salt.len()) {
        return Err(Error::InvalidFormat(format!(
            "ODF Argon2id salt length {} is outside supported limits",
            salt.len()
        )));
    }
    let iterations = parse_nonzero_u32(
        required(parameters, b"argon2-iterations")?,
        "Argon2id iteration count",
    )?;
    let memory_kib = parse_nonzero_u32(
        required(parameters, b"argon2-memory")?,
        "Argon2id memory cost",
    )?;
    let lanes = parse_nonzero_u32(
        required(parameters, b"argon2-lanes")?,
        "Argon2id lane count",
    )?;
    if iterations.get() > MAX_ARGON2_ITERATIONS
        || memory_kib.get() > MAX_ARGON2_MEMORY_KIB
        || lanes.get() > MAX_ARGON2_LANES
        || memory_kib.get() < 8 * lanes.get()
        || u64::from(memory_kib.get()) * u64::from(iterations.get()) > MAX_ARGON2_WORK_KIB_PASSES
    {
        return Err(Error::InvalidFormat(
            "ODF Argon2id parameters exceed supported resource limits".to_string(),
        ));
    }
    let key_size = parse_optional_key_size(parameters)?;
    if key_size.is_some_and(|size| !matches!(size, 16 | 24 | 32)) {
        return Err(Error::InvalidFormat(
            "ODF Argon2id derived key size must be 16, 24, or 32 bytes".to_string(),
        ));
    }
    Ok(ManifestKeyDerivation::Argon2id {
        salt,
        iterations,
        memory_kib,
        lanes,
        key_size,
    })
}

fn encryption_mut(encryption: &mut Option<PartialEncryption>) -> Result<&mut PartialEncryption> {
    encryption.as_mut().ok_or_else(|| {
        Error::InvalidFormat("Encryption child appears outside encryption-data".to_string())
    })
}

fn set_once<T>(slot: &mut Option<T>, value: T, name: &str) -> Result<()> {
    if slot.replace(value).is_some() {
        return Err(Error::InvalidFormat(format!(
            "Duplicate manifest {name} element"
        )));
    }
    Ok(())
}

fn finish_encryption(partial: PartialEncryption) -> Result<ManifestEncryption> {
    let algorithm = partial
        .algorithm
        .ok_or_else(|| Error::InvalidFormat("Missing encryption algorithm".to_string()))?;
    let key_derivation = partial
        .key_derivation
        .ok_or_else(|| Error::InvalidFormat("Missing key derivation".to_string()))?;
    let derived_key_size = match &key_derivation {
        ManifestKeyDerivation::Pbkdf2 { key_size, .. } => *key_size,
        ManifestKeyDerivation::Argon2id { key_size, .. } => {
            (*key_size).or(algorithm.fixed_key_size()).ok_or_else(|| {
                Error::InvalidFormat(
                    "Argon2id key size cannot be inferred for this cipher".to_string(),
                )
            })?
        },
    };
    if !algorithm.accepts_key_size(derived_key_size) {
        return Err(Error::InvalidFormat(format!(
            "Encryption algorithm requires a derived key size of {} bytes, found {derived_key_size}",
            algorithm.key_size_description()
        )));
    }
    let checksum = partial.checksum;
    if checksum.is_none() && !algorithm.is_aead() {
        return Err(Error::InvalidFormat(
            "CBC and Blowfish encryption require checksum metadata".to_string(),
        ));
    }
    Ok(ManifestEncryption {
        checksum,
        algorithm,
        start_key: partial
            .start_key
            .ok_or_else(|| Error::InvalidFormat("Missing start-key generation".to_string()))?,
        key_derivation,
    })
}

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    reason = "Test fixtures use infallible parsing setup so assertions can focus on manifest validation."
)]
mod tests {
    use super::*;

    #[test]
    fn parses_typed_encryption_with_an_arbitrary_prefix() {
        let xml = r#"<m:manifest xmlns:m="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0">
          <m:file-entry m:full-path="/" m:media-type="application/vnd.oasis.opendocument.text"/>
          <m:file-entry m:full-path="content.xml" m:media-type="text/xml" m:size="12">
            <m:encryption-data m:checksum-type="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0#sha256-1k" m:checksum="AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA=">
              <m:algorithm m:algorithm-name="http://www.w3.org/2001/04/xmlenc#aes256-cbc" m:initialisation-vector="AAAAAAAAAAAAAAAAAAAAAA=="/>
              <m:start-key-generation m:start-key-generation-name="http://www.w3.org/2000/09/xmldsig#sha256" m:key-size="32"/>
              <m:key-derivation m:key-derivation-name="PBKDF2" m:salt="AQIDBA==" m:iteration-count="100000" m:key-size="32"/>
            </m:encryption-data>
          </m:file-entry>
        </m:manifest>"#;
        let manifest = Manifest::parse(xml).unwrap();
        let encryption = manifest.entries["content.xml"].encryption.as_ref().unwrap();
        assert_eq!(encryption.start_key, ManifestStartKeyGeneration::Sha256);
        assert!(encryption.checksum.is_some());
        assert!(matches!(
            encryption.algorithm,
            ManifestEncryptionAlgorithm::Aes256Cbc { iv } if iv == [0; 16]
        ));
    }

    #[test]
    fn rejects_incomplete_and_excessive_encryption_metadata() {
        let incomplete = r#"<m:manifest xmlns:m="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><m:file-entry m:full-path="content.xml" m:size="1"><m:encryption-data m:checksum-type="SHA1/1K" m:checksum="AAAAAAAAAAAAAAAAAAAAAAAAAAA="/></m:file-entry></m:manifest>"#;
        assert!(Manifest::parse(incomplete).is_err());

        let excessive = r#"<m:manifest xmlns:m="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><m:file-entry m:full-path="content.xml" m:size="1"><m:encryption-data m:checksum-type="SHA1/1K" m:checksum="AAAAAAAAAAAAAAAAAAAAAAAAAAA="><m:algorithm m:algorithm-name="http://www.w3.org/2001/04/xmlenc#aes256-cbc" m:initialisation-vector="AAAAAAAAAAAAAAAAAAAAAA=="/><m:start-key-generation m:start-key-generation-name="SHA1" m:key-size="20"/><m:key-derivation m:key-derivation-name="PBKDF2" m:salt="AQ==" m:iteration-count="10000001" m:key-size="32"/></m:encryption-data></m:file-entry></m:manifest>"#;
        assert!(Manifest::parse(excessive).is_err());
    }

    #[test]
    fn rejects_cipher_and_derived_key_size_mismatch() {
        let xml = r#"<m:manifest xmlns:m="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><m:file-entry m:full-path="content.xml" m:size="1"><m:encryption-data m:checksum-type="SHA1/1K" m:checksum="AAAAAAAAAAAAAAAAAAAAAAAAAAA="><m:algorithm m:algorithm-name="http://www.w3.org/2001/04/xmlenc#aes128-cbc" m:initialisation-vector="AAAAAAAAAAAAAAAAAAAAAA=="/><m:start-key-generation m:start-key-generation-name="SHA1" m:key-size="20"/><m:key-derivation m:key-derivation-name="PBKDF2" m:salt="AQ==" m:iteration-count="1000" m:key-size="32"/></m:encryption-data></m:file-entry></m:manifest>"#;
        assert!(Manifest::parse(xml).is_err());
    }

    #[test]
    fn accepts_checksumless_gcm_but_not_checksumless_cbc() {
        let gcm = r#"<m:manifest xmlns:m="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><m:file-entry m:full-path="content.xml" m:size="1"><m:encryption-data><m:algorithm m:algorithm-name="http://www.w3.org/2009/xmlenc11#aes256-gcm" m:initialisation-vector="AAAAAAAAAAAAAAAA"/><m:start-key-generation m:start-key-generation-name="SHA1" m:key-size="20"/><m:key-derivation m:key-derivation-name="PBKDF2" m:salt="AQ==" m:iteration-count="1000" m:key-size="32"/></m:encryption-data></m:file-entry></m:manifest>"#;
        let parsed = Manifest::parse(gcm).unwrap();
        assert!(
            parsed.entries["content.xml"]
                .encryption
                .as_ref()
                .unwrap()
                .checksum
                .is_none()
        );

        let cbc = gcm
            .replace(
                "http://www.w3.org/2009/xmlenc11#aes256-gcm",
                "http://www.w3.org/2001/04/xmlenc#aes256-cbc",
            )
            .replace("AAAAAAAAAAAAAAAA", "AAAAAAAAAAAAAAAAAAAAAA==");
        assert!(Manifest::parse(&cbc).is_err());
    }

    #[test]
    fn validates_standard_argon2id_parameters_and_resource_limits() {
        let xml = r#"<m:manifest xmlns:m="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" xmlns:loext="urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0"><m:file-entry m:full-path="content.xml" m:size="1"><m:encryption-data><m:algorithm m:algorithm-name="http://www.w3.org/2009/xmlenc11#aes256-gcm" m:initialisation-vector="AAAAAAAAAAAAAAAA"/><m:start-key-generation m:start-key-generation-name="http://www.w3.org/2001/04/xmlenc#sha256" m:key-size="32"/><m:key-derivation m:key-derivation-name="urn:oasis:names:tc:opendocument:xmlns:manifest:1.5#argon2id" m:salt="AAAAAAAAAAAAAAAAAAAAAA==" m:argon2-iterations="1" m:argon2-memory="1024" m:argon2-lanes="1"/></m:encryption-data></m:file-entry></m:manifest>"#;
        let parsed = Manifest::parse(xml).unwrap();
        assert!(matches!(
            parsed.entries["content.xml"]
                .encryption
                .as_ref()
                .unwrap()
                .key_derivation,
            ManifestKeyDerivation::Argon2id {
                iterations,
                memory_kib,
                lanes,
                key_size: None,
                ..
            } if iterations.get() == 1 && memory_kib.get() == 1024 && lanes.get() == 1
        ));

        let excessive = xml.replace("m:argon2-memory=\"1024\"", "m:argon2-memory=\"262145\"");
        assert!(Manifest::parse(&excessive).is_err());
        let incomplete = xml.replace(" m:argon2-lanes=\"1\"", "");
        assert!(Manifest::parse(&incomplete).is_err());
        let mixed = xml.replace(
            "m:argon2-lanes=\"1\"",
            "m:argon2-lanes=\"1\" loext:argon2-lanes=\"1\"",
        );
        assert!(Manifest::parse(&mixed).is_err());
    }

    #[test]
    fn fused_scan_matches_neutral_index_and_typed_projection() {
        let xml = r#"<p:manifest xmlns:p="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0">
          <p:file-entry p:full-path="/" p:media-type="application/vnd.oasis.opendocument.text"/>
          <p:file-entry p:full-path="content.xml" p:media-type="text/xml" p:size="12">
            <p:encryption-data p:checksum-type="SHA256/1K" p:checksum="AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA=">
              <p:algorithm p:algorithm-name="http://www.w3.org/2009/xmlenc11#aes256-gcm" p:initialisation-vector="AAAAAAAAAAAAAAAA"/>
              <p:start-key-generation p:start-key-generation-name="SHA1" p:key-size="20"/>
              <p:key-derivation p:key-derivation-name="PBKDF2" p:salt="AQIDBA==" p:iteration-count="1000" p:key-size="32"/>
            </p:encryption-data>
          </p:file-entry>
        </p:manifest>"#;
        let neutral = crate::package::parse_manifest(xml).unwrap();
        let typed = Manifest::parse(xml).unwrap();

        assert_eq!(typed.mimetype, neutral.mimetype);
        for (path, neutral_entry) in &neutral.entries {
            let typed_entry = typed.entries.get(path).unwrap();
            assert_eq!(typed_entry.media_type, neutral_entry.media_type);
            assert_eq!(typed_entry.size, neutral_entry.size);
        }
        assert!(typed.entries["content.xml"].encryption.is_some());
    }

    #[test]
    fn neutral_errors_keep_precedence_over_earlier_encryption_errors() {
        let duplicate_after_bad_encryption = r#"<m:manifest xmlns:m="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0">
          <m:file-entry m:full-path="content.xml" m:media-type="text/xml" m:size="1">
            <m:encryption-data>
              <m:algorithm m:algorithm-name="vendor-cipher" m:initialisation-vector="AQ=="/>
            </m:encryption-data>
          </m:file-entry>
          <m:file-entry m:full-path="content.xml" m:media-type="text/xml"/>
        </m:manifest>"#;
        let error = Manifest::parse(duplicate_after_bad_encryption)
            .unwrap_err()
            .to_string();
        assert!(error.contains("Duplicate manifest file entry 'content.xml'"));

        let malformed_size_before_encryption = r#"<m:manifest xmlns:m="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0">
          <m:file-entry m:full-path="content.xml" m:media-type="text/xml" m:size="not-a-number">
            <m:encryption-data><m:algorithm m:algorithm-name="vendor-cipher" m:initialisation-vector="AQ=="/></m:encryption-data>
          </m:file-entry>
        </m:manifest>"#;
        let error = Manifest::parse(malformed_size_before_encryption)
            .unwrap_err()
            .to_string();
        assert!(error.contains("Invalid manifest size for 'content.xml'"));
    }

    #[test]
    fn duplicate_entries_with_alternate_prefix_are_rejected_once() {
        let xml = r#"<alt:manifest xmlns:alt="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0">
          <alt:file-entry alt:full-path="/" alt:media-type="application/vnd.oasis.opendocument.text"/>
          <alt:file-entry alt:full-path="content.xml" alt:media-type="text/xml"/>
          <alt:file-entry alt:full-path="content.xml" alt:media-type="text/xml"/>
        </alt:manifest>"#;
        let error = Manifest::parse(xml).unwrap_err().to_string();
        assert!(error.contains("Duplicate manifest file entry 'content.xml'"));
    }
}
