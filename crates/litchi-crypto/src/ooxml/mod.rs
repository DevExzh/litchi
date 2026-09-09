//! Bounded ECMA-376 package encryption.
//!
//! This module owns the application-neutral `[MS-OFFCRYPTO]` envelope used by
//! DOCX, PPTX, XLSX, and the other OPC package families. It never opens the
//! decrypted package or selects a concrete document model.

mod agile;
mod container;
mod standard;

use std::fmt;
use std::io::{ErrorKind, Read};

use litchi_cfb::is_ole_file;
use thiserror::Error;
use zeroize::Zeroizing;

/// Maximum spin count admitted by the published Agile schema.
pub const SPEC_MAX_SPIN_COUNT: u32 = 10_000_000;

/// Result returned by this module.
pub type Result<T> = std::result::Result<T, Error>;

/// Errors from bounded OOXML package encryption.
#[derive(Debug, Error)]
#[non_exhaustive]
pub enum Error {
    /// A caller-supplied resource policy is internally invalid.
    #[error("invalid OOXML encryption limit: {0}")]
    InvalidLimit(&'static str),
    /// Input or prospective output exceeds an explicit resource ceiling.
    #[error("OOXML encryption {resource} uses {actual} bytes/items, maximum {maximum}")]
    Limit {
        resource: &'static str,
        actual: u64,
        maximum: u64,
    },
    /// The encrypted package uses an otherwise valid profile this build does not implement.
    #[error("unsupported OOXML encryption profile: {0}")]
    Unsupported(String),
    /// The encrypted package is structurally malformed.
    #[error("malformed OOXML encryption data: {0}")]
    Malformed(String),
    /// The supplied password does not match the encrypted verifier.
    #[error("incorrect OOXML package password")]
    Password,
    /// Safe-by-default authoring requires a non-empty new password.
    #[error("a new OOXML encryption password must not be empty")]
    PasswordRequired,
    /// An operation that requires encrypted input received an ordinary package.
    #[error("OOXML package is not encrypted")]
    NotEncrypted,
    /// Cryptographically authenticated package data has been modified.
    #[error("OOXML encrypted package integrity check failed")]
    Integrity,
    /// Parsing the Agile XML descriptor failed.
    #[error("invalid Agile encryption XML: {0}")]
    Xml(String),
    /// The operating system could not provide cryptographic randomness.
    #[error("OOXML encryption random source failed: {0}")]
    Random(String),
    /// A bounded buffer could not reserve the memory required for an operation.
    #[error("OOXML encryption could not reserve memory for {0}")]
    Allocation(&'static str),
    /// The storage container or `StrongEncryptionDataSpace` graph is malformed.
    #[error("OOXML encrypted container error: {0}")]
    Container(String),
    /// A bounded output sink failed.
    #[error("OOXML encrypted output failed: {0}")]
    Io(#[from] std::io::Error),
}

/// Explicit resource policy for one encryption or decryption operation.
///
/// Public fields intentionally support concise struct-update customization:
///
/// ```
/// use litchi_crypto::ooxml::Limits;
///
/// let limits = Limits {
///     max_plaintext_bytes: 16 * 1024 * 1024,
///     ..Limits::default()
/// };
/// assert_eq!(limits.max_plaintext_bytes, 16 * 1024 * 1024);
/// ```
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Limits {
    /// Maximum bytes accepted at the outer API boundary.
    pub max_input_bytes: usize,
    /// Maximum bytes accepted from the `EncryptionInfo` stream.
    pub max_info_bytes: usize,
    /// Maximum bytes accepted in the Agile XML descriptor.
    pub max_xml_bytes: usize,
    /// Maximum nested XML element depth.
    pub max_xml_depth: usize,
    /// Maximum XML element, text, comment, and processing-instruction events.
    pub max_xml_nodes: usize,
    /// Maximum aggregate attributes in the Agile descriptor.
    pub max_xml_attributes: usize,
    /// Maximum password-hash iterations accepted from Agile input.
    pub max_spin_count: u32,
    /// Maximum Unicode scalar values in a password.
    ///
    /// This crate caps the configurable ceiling at 255 as an interoperability
    /// and security policy.
    pub max_password_chars: usize,
    /// Maximum declared or supplied clear OPC package bytes.
    pub max_plaintext_bytes: usize,
    /// Maximum bytes in the encrypted package stream.
    pub max_encrypted_bytes: usize,
    /// Maximum bytes emitted for the complete compound container.
    pub max_output_bytes: usize,
    /// Maximum padded directory stream accepted from the CFB container.
    pub max_cfb_directory_bytes: usize,
    /// Maximum decoded FAT/DIFAT/MiniFAT table bytes accepted from the CFB
    /// container.
    pub max_cfb_allocation_table_bytes: usize,
    /// Accept `LibreOffice`'s nonconforming encrypted container without `DataSpaces`.
    pub allow_missing_data_spaces: bool,
}

impl Default for Limits {
    fn default() -> Self {
        Self {
            max_input_bytes: 512 * 1024 * 1024,
            max_info_bytes: 1024 * 1024,
            max_xml_bytes: 1024 * 1024,
            max_xml_depth: 64,
            max_xml_nodes: 4_096,
            max_xml_attributes: 4_096,
            // Office normally writes 100,000. A caller must explicitly opt in
            // to the schema maximum when handling unusually expensive input.
            max_spin_count: 1_000_000,
            max_password_chars: 255,
            max_plaintext_bytes: 512 * 1024 * 1024,
            max_encrypted_bytes: 520 * 1024 * 1024,
            max_output_bytes: 528 * 1024 * 1024,
            max_cfb_directory_bytes: litchi_cfb::OleFileLimits::DEFAULT_MAX_DIRECTORY_BYTES
                as usize,
            max_cfb_allocation_table_bytes:
                litchi_cfb::OleFileLimits::DEFAULT_MAX_ALLOCATION_TABLE_BYTES as usize,
            allow_missing_data_spaces: false,
        }
    }
}

impl Limits {
    fn validate(self) -> Result<Self> {
        if self.max_input_bytes == 0 {
            return Err(Error::InvalidLimit("max_input_bytes must be nonzero"));
        }
        if self.max_info_bytes < 12 {
            return Err(Error::InvalidLimit("max_info_bytes must be at least 12"));
        }
        if self.max_xml_bytes == 0 {
            return Err(Error::InvalidLimit("max_xml_bytes must be nonzero"));
        }
        if self.max_xml_depth == 0 {
            return Err(Error::InvalidLimit("max_xml_depth must be nonzero"));
        }
        if self.max_xml_nodes == 0 {
            return Err(Error::InvalidLimit("max_xml_nodes must be nonzero"));
        }
        if self.max_xml_attributes == 0 {
            return Err(Error::InvalidLimit("max_xml_attributes must be nonzero"));
        }
        if self.max_spin_count > SPEC_MAX_SPIN_COUNT {
            return Err(Error::InvalidLimit(
                "max_spin_count exceeds the MS-OFFCRYPTO schema maximum",
            ));
        }
        if self.max_password_chars == 0 || self.max_password_chars > 255 {
            return Err(Error::InvalidLimit(
                "max_password_chars must be between 1 and 255",
            ));
        }
        if self.max_plaintext_bytes == 0 {
            return Err(Error::InvalidLimit("max_plaintext_bytes must be nonzero"));
        }
        if self.max_encrypted_bytes < 24 {
            return Err(Error::InvalidLimit(
                "max_encrypted_bytes must be at least 24",
            ));
        }
        if self.max_output_bytes < 512 {
            return Err(Error::InvalidLimit("max_output_bytes must be at least 512"));
        }
        let max_cfb_input = u64::try_from(self.max_input_bytes).unwrap_or(u64::MAX);
        if max_cfb_input > litchi_cfb::OleFileLimits::MAX_INPUT_BYTES {
            return Err(Error::InvalidLimit(
                "max_input_bytes exceeds the supported CFB input ceiling",
            ));
        }
        let max_cfb_directory = u64::try_from(self.max_cfb_directory_bytes).unwrap_or(u64::MAX);
        if max_cfb_directory == 0
            || max_cfb_directory > litchi_cfb::OleFileLimits::MAX_DIRECTORY_BYTES
        {
            return Err(Error::InvalidLimit(
                "max_cfb_directory_bytes exceeds the supported CFB directory ceiling",
            ));
        }
        let max_cfb_allocation =
            u64::try_from(self.max_cfb_allocation_table_bytes).unwrap_or(u64::MAX);
        if max_cfb_allocation == 0
            || max_cfb_allocation > litchi_cfb::OleFileLimits::MAX_ALLOCATION_TABLE_BYTES
        {
            return Err(Error::InvalidLimit(
                "max_cfb_allocation_table_bytes exceeds the supported CFB allocation-table ceiling",
            ));
        }
        Ok(self)
    }

    fn bytes(resource: &'static str, actual: usize, maximum: usize) -> Result<()> {
        if actual > maximum {
            return Err(Error::Limit {
                resource,
                actual: u64::try_from(actual).unwrap_or(u64::MAX),
                maximum: u64::try_from(maximum).unwrap_or(u64::MAX),
            });
        }
        Ok(())
    }

    fn count(resource: &'static str, actual: usize, maximum: usize) -> Result<()> {
        Self::bytes(resource, actual, maximum)
    }
}

/// An owned password that is cleared on drop and never exposed by `Debug`.
///
/// This type is intentionally non-`Clone`; move it into long-lived document
/// state and borrow it only for a crypto operation.
pub struct Password(Zeroizing<String>);

impl Password {
    /// Take ownership of a password without copying its allocation.
    #[must_use]
    pub fn new(value: String) -> Self {
        Self(Zeroizing::new(value))
    }

    /// Borrow the password for one operation.
    #[must_use]
    pub fn as_str(&self) -> &str {
        self.0.as_str()
    }
}

impl From<String> for Password {
    fn from(value: String) -> Self {
        Self::new(value)
    }
}

impl fmt::Debug for Password {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str("Password([REDACTED])")
    }
}

/// Supported encryption profile.
///
/// `Standard` and `Agile` retain their historical meaning: AES-128/SHA-1.
/// The additional variants identify the non-default profiles on input and on
/// explicit authoring calls. Keeping the profile in this value is important
/// for re-encryption: a caller that opens an AES-256 or SHA-512 package must
/// not accidentally write it back as the publishing default.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Mode {
    /// Binary Standard Encryption using AES-128 ECB and SHA-1.
    Standard,
    /// Binary Standard Encryption using AES-192 ECB and SHA-1.
    StandardAes192,
    /// Binary Standard Encryption using AES-256 ECB and SHA-1.
    StandardAes256,
    /// XML-described Agile Encryption using AES-128 CBC, SHA-1, and integrity data.
    Agile,
    /// XML-described Agile Encryption using AES-128 CBC, SHA-256, and integrity data.
    AgileSha256,
    /// XML-described Agile Encryption using AES-128 CBC, SHA-512, and integrity data.
    AgileSha512,
    /// XML-described Agile Encryption using AES-192 CBC, SHA-1, and integrity data.
    AgileAes192,
    /// XML-described Agile Encryption using AES-256 CBC, SHA-1, and integrity data.
    AgileAes256,
    /// XML-described Agile Encryption using AES-192 CBC, SHA-256, and integrity data.
    AgileAes192Sha256,
    /// XML-described Agile Encryption using AES-192 CBC, SHA-512, and integrity data.
    AgileAes192Sha512,
    /// XML-described Agile Encryption using AES-256 CBC, SHA-256, and integrity data.
    AgileAes256Sha256,
    /// XML-described Agile Encryption using AES-256 CBC, SHA-512, and integrity data.
    AgileAes256Sha512,
}

impl Mode {
    /// Return whether this is a Standard (binary) profile.
    #[must_use]
    pub const fn is_standard(self) -> bool {
        matches!(
            self,
            Self::Standard | Self::StandardAes192 | Self::StandardAes256
        )
    }

    /// Return whether this is an Agile (XML-described) profile.
    #[must_use]
    pub const fn is_agile(self) -> bool {
        !self.is_standard()
    }

    /// Construct an Agile profile from its cipher and hash names.
    ///
    /// The compact variants cover every AES-128/192/256 and SHA-1/256/512
    /// combination currently admitted by the owner.  Unknown combinations
    /// are not representable and therefore cannot accidentally reach an
    /// authoring or decryption path.
    #[must_use]
    pub const fn agile(cipher: AgileCipher, hash: AgileHash) -> Self {
        match (cipher, hash) {
            (AgileCipher::Aes128, AgileHash::Sha1) => Self::Agile,
            (AgileCipher::Aes128, AgileHash::Sha256) => Self::AgileSha256,
            (AgileCipher::Aes128, AgileHash::Sha512) => Self::AgileSha512,
            (AgileCipher::Aes192, AgileHash::Sha1) => Self::AgileAes192,
            (AgileCipher::Aes192, AgileHash::Sha256) => Self::AgileAes192Sha256,
            (AgileCipher::Aes192, AgileHash::Sha512) => Self::AgileAes192Sha512,
            (AgileCipher::Aes256, AgileHash::Sha1) => Self::AgileAes256,
            (AgileCipher::Aes256, AgileHash::Sha256) => Self::AgileAes256Sha256,
            (AgileCipher::Aes256, AgileHash::Sha512) => Self::AgileAes256Sha512,
        }
    }

    pub(crate) const fn agile_cipher(self) -> AgileCipher {
        match self {
            Self::Agile | Self::AgileSha256 | Self::AgileSha512 => AgileCipher::Aes128,
            Self::AgileAes192 | Self::AgileAes192Sha256 | Self::AgileAes192Sha512 => {
                AgileCipher::Aes192
            },
            Self::AgileAes256 | Self::AgileAes256Sha256 | Self::AgileAes256Sha512 => {
                AgileCipher::Aes256
            },
            _ => AgileCipher::Aes128,
        }
    }

    pub(crate) const fn agile_hash(self) -> AgileHash {
        match self {
            Self::Agile | Self::AgileAes192 | Self::AgileAes256 => AgileHash::Sha1,
            Self::AgileSha256 | Self::AgileAes192Sha256 | Self::AgileAes256Sha256 => {
                AgileHash::Sha256
            },
            Self::AgileSha512 | Self::AgileAes192Sha512 | Self::AgileAes256Sha512 => {
                AgileHash::Sha512
            },
            _ => AgileHash::Sha1,
        }
    }
}

/// AES key sizes admitted by Agile Encryption.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum AgileCipher {
    /// AES with a 128-bit key.
    Aes128,
    /// AES with a 192-bit key.
    Aes192,
    /// AES with a 256-bit key.
    Aes256,
}

/// Hash functions admitted by Agile Encryption.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum AgileHash {
    /// SHA-1, the historical Office publishing default.
    Sha1,
    /// SHA-256.
    Sha256,
    /// SHA-512.
    Sha512,
}

/// Password-free outer-package classification.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Kind {
    /// Ordinary non-CFB input. The bytes are not parsed as an OPC package.
    Plain,
    /// A strictly parsed supported encryption profile.
    Encrypted(Mode),
}

/// Policy controlling whether an Agile package may omit its optional
/// `dataIntegrity` element while being opened.
///
/// The default reader policy is [`Self::RequireAuthenticated`].
/// [`Self::AllowUnauthenticated`] is an explicit compatibility escape hatch
/// for the legal schema profile used by some non-Office producers. It affects
/// reading only; every Agile authoring API in this module always emits
/// authenticated `dataIntegrity`. Standard Encryption has no Agile
/// `dataIntegrity` element and reports [`IntegrityStatus::Unauthenticated`].
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum IntegrityPolicy {
    /// Require authenticated Agile `dataIntegrity` before returning clear
    /// package bytes. Standard Encryption remains unauthenticated because it
    /// has no Agile `dataIntegrity` element.
    #[default]
    RequireAuthenticated,
    /// Permit an Agile package with no `dataIntegrity` element to be read.
    AllowUnauthenticated,
}

/// Integrity result recorded for an encrypted package that was opened.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum IntegrityStatus {
    /// Agile `dataIntegrity` was present and verified against the encrypted
    /// package stream.
    Authenticated,
    /// No integrity check was available. This includes Standard Encryption,
    /// which has no package-integrity field, and an Agile package opened under
    /// [`IntegrityPolicy::AllowUnauthenticated`].
    Unauthenticated,
}

impl fmt::Display for Mode {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::Standard => "Standard",
            Self::StandardAes192 => "Standard AES-192",
            Self::StandardAes256 => "Standard AES-256",
            Self::Agile => "Agile",
            Self::AgileSha256 => "Agile SHA-256",
            Self::AgileSha512 => "Agile SHA-512",
            Self::AgileAes192 => "Agile AES-192",
            Self::AgileAes256 => "Agile AES-256",
            Self::AgileAes192Sha256 => "Agile AES-192/SHA-256",
            Self::AgileAes192Sha512 => "Agile AES-192/SHA-512",
            Self::AgileAes256Sha256 => "Agile AES-256/SHA-256",
            Self::AgileAes256Sha512 => "Agile AES-256/SHA-512",
        })
    }
}

/// A moved clear package and its optional source encryption profile.
pub struct Opened {
    mode: Option<Mode>,
    integrity: Option<IntegrityStatus>,
    bytes: Vec<u8>,
}

impl fmt::Debug for Opened {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Opened")
            .field("mode", &self.mode)
            .field("integrity", &self.integrity)
            .field("byte_len", &self.bytes.len())
            .finish_non_exhaustive()
    }
}

impl Opened {
    /// Encryption profile used by the input, or `None` for an ordinary OPC package.
    #[must_use]
    pub const fn mode(&self) -> Option<Mode> {
        self.mode
    }

    /// Integrity result for encrypted input, or `None` for an ordinary package.
    #[must_use]
    pub const fn integrity(&self) -> Option<IntegrityStatus> {
        self.integrity
    }

    /// Borrow the clear package bytes.
    #[must_use]
    pub fn bytes(&self) -> &[u8] {
        &self.bytes
    }

    /// Move the clear package allocation to the caller.
    #[must_use]
    pub fn into_bytes(self) -> Vec<u8> {
        self.bytes
    }
}

/// Open an ordinary or encrypted package with safe default limits.
///
/// Ordinary OPC input is returned in the exact caller-provided allocation.
///
/// # Errors
///
/// Returns an [`enum@Error`] when the input violates the default resource policy
/// or the encrypted package cannot be validated or decrypted.
pub fn open(bytes: Vec<u8>, password: &str) -> Result<Opened> {
    open_with(bytes, password, &Limits::default())
}

/// Classify an ordinary or encrypted package without reading encrypted content.
///
/// # Errors
///
/// Returns an [`enum@Error`] when the input violates the default resource policy
/// or the encryption metadata is malformed or unsupported.
pub fn inspect(bytes: &[u8]) -> Result<Kind> {
    inspect_with(bytes, &Limits::default())
}

/// Classify a package under an explicit resource and Agile-integrity policy.
///
/// The ordinary [`inspect`] and [`inspect_with`] entry points require
/// authenticated Agile metadata. Use this method when an application has
/// deliberately opted into the schema's unauthenticated Agile profile.
pub fn inspect_with_policy(
    bytes: &[u8],
    limits: &Limits,
    integrity_policy: IntegrityPolicy,
) -> Result<Kind> {
    let validated = limits.validate()?;
    Limits::bytes("input", bytes.len(), validated.max_input_bytes)?;
    if !is_ole_file(bytes) {
        return Ok(Kind::Plain);
    }
    let info = container::read_info(bytes, &validated)?;
    let mode = profile_with_policy(&info, &validated, integrity_policy)?;
    Ok(Kind::Encrypted(mode))
}

/// Classify a package under an explicit resource policy.
///
/// This parses the CFB header and `EncryptionInfo` only; it does not allocate
/// or read the potentially large `EncryptedPackage` stream.
///
/// # Errors
///
/// Returns an [`enum@Error`] when the limits are invalid, the input exceeds them,
/// or the encryption metadata is malformed or unsupported.
pub fn inspect_with(bytes: &[u8], limits: &Limits) -> Result<Kind> {
    inspect_with_policy(bytes, limits, IntegrityPolicy::RequireAuthenticated)
}

/// Read and open an ordinary or encrypted package with safe default limits.
///
/// The reader is runtime-neutral and never consumes more than the configured
/// input ceiling plus one detection byte.
///
/// # Errors
///
/// Returns an [`enum@Error`] when reading fails, the input violates the default
/// resource policy, or the encrypted package cannot be opened.
pub fn load<R: Read>(reader: R, password: &str) -> Result<Opened> {
    load_with(reader, password, &Limits::default())
}

/// Read and open a package under an explicit resource and Agile-integrity
/// policy.
pub fn load_with_policy<R: Read>(
    reader: R,
    password: &str,
    limits: &Limits,
    integrity_policy: IntegrityPolicy,
) -> Result<Opened> {
    load_with_policy_inner(reader, password, limits, integrity_policy)
}

/// Read and open an ordinary or encrypted package under an explicit policy.
///
/// # Errors
///
/// Returns an [`enum@Error`] when reading fails, the limits are invalid or
/// exceeded, or the encrypted package cannot be opened.
pub fn load_with<R: Read>(reader: R, password: &str, limits: &Limits) -> Result<Opened> {
    load_with_policy_inner(
        reader,
        password,
        limits,
        IntegrityPolicy::RequireAuthenticated,
    )
}

fn load_with_policy_inner<R: Read>(
    mut reader: R,
    password: &str,
    limits: &Limits,
    integrity_policy: IntegrityPolicy,
) -> Result<Opened> {
    let validated = limits.validate()?;
    let mut bytes = Vec::new();
    let mut buffer = [0u8; 8 * 1024];
    while bytes.len() < validated.max_input_bytes {
        let remaining = validated.max_input_bytes - bytes.len();
        let read_len = remaining.min(buffer.len());
        let count = match reader.read(&mut buffer[..read_len]) {
            Ok(0) => return open_with_policy(bytes, password, &validated, integrity_policy),
            Ok(count) => count,
            Err(error) if error.kind() == ErrorKind::Interrupted => continue,
            Err(error) => return Err(Error::Io(error)),
        };
        bytes
            .try_reserve(count)
            .map_err(|_err| Error::Allocation("OOXML package input"))?;
        bytes.extend_from_slice(&buffer[..count]);
    }

    let mut probe = [0u8; 1];
    loop {
        match reader.read(&mut probe) {
            Ok(0) => return open_with_policy(bytes, password, &validated, integrity_policy),
            Ok(_) => {
                return Err(Error::Limit {
                    resource: "input",
                    actual: u64::try_from(validated.max_input_bytes)
                        .unwrap_or(u64::MAX)
                        .saturating_add(1),
                    maximum: u64::try_from(validated.max_input_bytes).unwrap_or(u64::MAX),
                });
            },
            Err(error) if error.kind() == ErrorKind::Interrupted => {},
            Err(error) => return Err(Error::Io(error)),
        }
    }
}

/// Open an ordinary or encrypted package under an explicit resource policy.
///
/// # Errors
///
/// Returns an [`enum@Error`] when the limits are invalid or exceeded, the password
/// violates policy or is incorrect, or the package is malformed.
pub fn open_with(bytes: Vec<u8>, password: &str, limits: &Limits) -> Result<Opened> {
    open_with_policy(
        bytes,
        password,
        limits,
        IntegrityPolicy::RequireAuthenticated,
    )
}

/// Open an ordinary or encrypted package under an explicit resource and
/// Agile-integrity policy.
///
/// `AllowUnauthenticated` applies only to an Agile descriptor that omits the
/// optional `dataIntegrity` element. It never disables verification when the
/// element is present, and it never changes authoring behavior.
pub fn open_with_policy(
    bytes: Vec<u8>,
    password: &str,
    limits: &Limits,
    integrity_policy: IntegrityPolicy,
) -> Result<Opened> {
    let validated = limits.validate()?;
    Limits::bytes("input", bytes.len(), validated.max_input_bytes)?;

    if !is_ole_file(&bytes) {
        Limits::bytes("plaintext", bytes.len(), validated.max_plaintext_bytes)?;
        return Ok(Opened {
            mode: None,
            integrity: None,
            bytes,
        });
    }

    validate_password(password, &validated)?;
    let (info, encrypted) = container::read(bytes, &validated)?;
    let mode = profile_with_policy(&info, &validated, integrity_policy)?;
    let (package, integrity) = match mode.is_standard() {
        true => (
            standard::decrypt(&info, encrypted, password, &validated)?,
            IntegrityStatus::Unauthenticated,
        ),
        false => {
            agile::decrypt_with_policy(&info, encrypted, password, &validated, integrity_policy)?
        },
    };
    Ok(Opened {
        mode: Some(mode),
        integrity: Some(integrity),
        bytes: package,
    })
}

/// Encrypt a moved OPC package with safe default limits.
///
/// # Errors
///
/// Returns an [`enum@Error`] when the package is empty, the password is empty or
/// violates policy, the default limits are exceeded, or encryption fails.
pub fn encrypt(bytes: Vec<u8>, password: &str, mode: Mode) -> Result<Vec<u8>> {
    encrypt_with(bytes, password, mode, &Limits::default())
}

/// Encrypt a moved OPC package under an explicit resource policy.
///
/// # Errors
///
/// Returns an [`enum@Error`] when the package is empty, the password is empty or
/// violates policy, the limits are invalid or exceeded, or encryption fails.
pub fn encrypt_with(
    bytes: Vec<u8>,
    password: &str,
    mode: Mode,
    limits: &Limits,
) -> Result<Vec<u8>> {
    let validated = limits.validate()?;
    if bytes.is_empty() {
        return Err(malformed("cannot encrypt an empty OPC package"));
    }
    if password.is_empty() {
        return Err(Error::PasswordRequired);
    }
    Limits::bytes("plaintext", bytes.len(), validated.max_plaintext_bytes)?;
    validate_password(password, &validated)?;

    match mode {
        Mode::Standard | Mode::StandardAes192 | Mode::StandardAes256 => {
            standard::encrypt(bytes, password, mode, &validated)
        },
        _ => agile::encrypt(bytes, password, mode, &validated),
    }
}

/// Change an encrypted package password while preserving its encryption mode.
///
/// # Errors
///
/// Returns an [`enum@Error`] when the input is not encrypted, the old password is
/// incorrect, the new password violates policy, or the default limits are
/// exceeded.
pub fn rekey(bytes: Vec<u8>, old_password: &str, new_password: &str) -> Result<Vec<u8>> {
    rekey_with(bytes, old_password, new_password, &Limits::default())
}

/// Change a password under an explicit policy without copying clear package bytes.
///
/// # Errors
///
/// Returns an [`enum@Error`] when the input is not encrypted, the old password is
/// incorrect, the new password violates policy, or the limits are invalid or
/// exceeded.
pub fn rekey_with(
    bytes: Vec<u8>,
    old_password: &str,
    new_password: &str,
    limits: &Limits,
) -> Result<Vec<u8>> {
    rekey_with_policy(
        bytes,
        old_password,
        new_password,
        limits,
        IntegrityPolicy::RequireAuthenticated,
    )
}

/// Change a password under an explicit resource and Agile-integrity policy.
///
/// Regardless of the input policy, the returned package is authored with the
/// authenticated Agile `dataIntegrity` element; Standard Encryption output has
/// no Agile integrity element.
pub fn rekey_with_policy(
    bytes: Vec<u8>,
    old_password: &str,
    new_password: &str,
    limits: &Limits,
    integrity_policy: IntegrityPolicy,
) -> Result<Vec<u8>> {
    let opened = open_with_policy(bytes, old_password, limits, integrity_policy)?;
    let mode = opened.mode().ok_or(Error::NotEncrypted)?;
    encrypt_with(opened.into_bytes(), new_password, mode, limits)
}

fn family_mode(info: &[u8]) -> Result<Mode> {
    if info.len() < 8 {
        return Err(malformed(
            "EncryptionInfo is shorter than its version header",
        ));
    }
    let major = u16::from_le_bytes([info[0], info[1]]);
    let minor = u16::from_le_bytes([info[2], info[3]]);
    match (major, minor) {
        (2..=4, 2) => Ok(Mode::Standard),
        (4, 4) => Ok(Mode::Agile),
        _ => Err(Error::Unsupported(format!(
            "EncryptionInfo version {major}.{minor}"
        ))),
    }
}

fn profile_with_policy(
    info: &[u8],
    limits: &Limits,
    integrity_policy: IntegrityPolicy,
) -> Result<Mode> {
    match family_mode(info)? {
        Mode::Standard => standard::profile(info, limits),
        Mode::Agile => agile::profile_with_policy(info, limits, integrity_policy),
        _ => Err(malformed("encryption profile family is not representable")),
    }
}

fn validate_password(password: &str, limits: &Limits) -> Result<usize> {
    let characters = password.chars().count();
    Limits::count("password characters", characters, limits.max_password_chars)?;
    let units = password.encode_utf16().count();
    units
        .checked_mul(2)
        .ok_or_else(|| malformed("password UTF-16 byte length overflows usize"))
}

fn password_bytes(password: &str, limits: &Limits) -> Result<Zeroizing<Vec<u8>>> {
    let byte_len = validate_password(password, limits)?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(byte_len)
        .map_err(|_err| Error::Allocation("UTF-16 password"))?;
    for unit in password.encode_utf16() {
        output.extend_from_slice(&unit.to_le_bytes());
    }
    Ok(Zeroizing::new(output))
}

fn declared_size(value: u64, limits: &Limits) -> Result<usize> {
    let size = usize::try_from(value).map_err(|_err| Error::Limit {
        resource: "declared plaintext",
        actual: value,
        maximum: u64::try_from(limits.max_plaintext_bytes).unwrap_or(u64::MAX),
    })?;
    Limits::bytes("declared plaintext", size, limits.max_plaintext_bytes)?;
    Ok(size)
}

fn malformed(message: impl Into<String>) -> Error {
    Error::Malformed(message.into())
}

#[cfg(test)]
mod tests {
    #![allow(
        clippy::expect_used,
        reason = "test code panics on failure; expect keeps assertions concise"
    )]
    use std::io::Cursor;

    use super::*;

    #[test]
    fn ordinary_input_moves_without_reallocation() {
        let bytes = Vec::from(&b"PK\x03\x04ordinary package"[..]);
        let pointer = bytes.as_ptr();
        let opened = open(bytes, "unused").expect("ordinary input");
        assert_eq!(opened.mode(), None);
        assert_eq!(opened.integrity(), None);
        let moved = opened.into_bytes();
        assert_eq!(moved.as_ptr(), pointer);
    }

    #[test]
    fn opened_debug_never_discloses_package_bytes() {
        let marker = b"confidential document marker 7qZ";
        let opened = open(marker.to_vec(), "unused").expect("ordinary input");
        let debug = format!("{opened:?}");

        assert!(debug.contains("byte_len"));
        assert!(!debug.contains("confidential document marker"));
    }

    #[test]
    fn ordinary_input_ignores_irrelevant_password_policy() {
        let bytes = Vec::from(&b"PK\x03\x04ordinary package"[..]);
        let pointer = bytes.as_ptr();
        let limits = Limits {
            max_password_chars: 1,
            ..Limits::default()
        };
        let opened = open_with(bytes, "deliberately over limit", &limits)
            .expect("password is irrelevant to ordinary OPC input");
        let moved = opened.into_bytes();
        assert_eq!(moved.as_ptr(), pointer);
    }

    #[test]
    fn policy_rejects_schema_exceeding_spin_limit() {
        let limits = Limits {
            max_spin_count: SPEC_MAX_SPIN_COUNT + 1,
            ..Limits::default()
        };
        assert!(matches!(
            limits.validate(),
            Err(Error::InvalidLimit(
                "max_spin_count exceeds the MS-OFFCRYPTO schema maximum"
            ))
        ));
    }

    #[test]
    fn password_limit_counts_unicode_scalar_values() {
        let limits = Limits {
            max_password_chars: 1,
            ..Limits::default()
        };
        assert!(matches!(
            validate_password("ab", &limits),
            Err(Error::Limit {
                resource: "password characters",
                actual: 2,
                maximum: 1,
            })
        ));
        assert!(matches!(validate_password("😀", &limits), Ok(4)));
    }

    #[test]
    fn bounded_reader_stops_after_one_detection_byte() {
        let limits = Limits {
            max_input_bytes: 4,
            ..Limits::default()
        };
        assert!(matches!(
            load_with(Cursor::new(b"12345"), "unused", &limits),
            Err(Error::Limit {
                resource: "input",
                actual: 5,
                maximum: 4,
            })
        ));
    }

    #[test]
    fn inspect_distinguishes_plain_from_malformed_cfb() {
        assert!(matches!(inspect(b"PK\x03\x04ordinary"), Ok(Kind::Plain)));
        let mut malformed_cfb = vec![0u8; 1_536];
        malformed_cfb[..8].copy_from_slice(&[0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]);
        assert!(inspect(&malformed_cfb).is_err());
    }

    #[test]
    fn rekey_rejects_plain_input_and_password_debug_is_redacted() {
        assert!(matches!(
            rekey(Vec::from(&b"PK\x03\x04ordinary"[..]), "old", "new"),
            Err(Error::NotEncrypted)
        ));
        let password = Password::new("never print me".to_string());
        assert_eq!(format!("{password:?}"), "Password([REDACTED])");
    }

    #[test]
    fn authoring_requires_a_nonempty_password() {
        assert!(matches!(
            encrypt(Vec::from(&b"PK\x03\x04ordinary"[..]), "", Mode::Standard),
            Err(Error::PasswordRequired)
        ));
    }
}
