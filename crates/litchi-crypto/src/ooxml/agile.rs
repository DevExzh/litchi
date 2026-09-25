//! `[MS-OFFCRYPTO]` Agile Encryption profiles.
//!
//! This owner supports the AES-128/192/256 CBC profiles with SHA-1, SHA-256,
//! or SHA-512 password derivation and authenticated `dataIntegrity`. The
//! schema's optional `dataIntegrity` profile is readable only through the
//! explicit unauthenticated-read policy; authoring always emits the element.

use std::fmt::Write as _;

use aes::cipher::{BlockModeDecrypt, BlockModeEncrypt, KeyIvInit, block_padding::NoPadding};
use aes::{Aes128, Aes192, Aes256};
use base64::Engine as _;
use base64::engine::general_purpose::STANDARD as BASE64;
use cbc::{Decryptor, Encryptor};
use hmac::{Hmac, Mac, digest::KeyInit};
use quick_xml::XmlVersion;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesDecl, BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use rand::TryRng;
use rand::rngs::SysRng;
use sha1::Sha1;
use sha2::{Digest, Sha256, Sha512};
use subtle::ConstantTimeEq;
use zeroize::{Zeroize, Zeroizing};

use super::{
    AgileCipher, AgileHash, Error, IntegrityPolicy, IntegrityStatus, Limits, Mode, Result,
    SPEC_MAX_SPIN_COUNT, container, declared_size, malformed, password_bytes,
};
use litchi_ole_common::xml_attributes::BytesStartExt as _;

const BLOCK: usize = 16;
const SALT_BYTES: usize = 16;
const SPIN_COUNT: u32 = 100_000;
const SEGMENT: usize = 4_096;
const ENC_NS: &[u8] = b"http://schemas.microsoft.com/office/2006/encryption";
const PASSWORD_NS: &[u8] = b"http://schemas.microsoft.com/office/2006/keyEncryptor/password";

const VERIFIER_INPUT_BLOCK: [u8; 8] = [0xfe, 0xa7, 0xd2, 0x76, 0x3b, 0x4b, 0x9e, 0x79];
const HASHED_VERIFIER_BLOCK: [u8; 8] = [0xd7, 0xaa, 0x0f, 0x6d, 0x30, 0x61, 0x34, 0x4e];
const CRYPTO_KEY_BLOCK: [u8; 8] = [0x14, 0x6e, 0x0b, 0xe7, 0xab, 0xac, 0xd0, 0xd6];
const INTEGRITY_KEY_BLOCK: [u8; 8] = [0x5f, 0xb2, 0xad, 0x01, 0x0c, 0xb9, 0xe1, 0xf6];
const INTEGRITY_VALUE_BLOCK: [u8; 8] = [0xa0, 0x67, 0x7f, 0x02, 0xb2, 0x2c, 0x84, 0x33];

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct Profile {
    cipher: AgileCipher,
    hash: AgileHash,
}

impl Profile {
    fn from_mode(mode: Mode) -> Result<Self> {
        if !mode.is_agile() {
            return Err(Error::Unsupported("mode is not an Agile profile".into()));
        }
        Ok(Self {
            cipher: mode.agile_cipher(),
            hash: mode.agile_hash(),
        })
    }

    const fn key_bytes(self) -> usize {
        match self.cipher {
            AgileCipher::Aes128 => 16,
            AgileCipher::Aes192 => 24,
            AgileCipher::Aes256 => 32,
        }
    }

    const fn key_bits(self) -> u32 {
        match self.cipher {
            AgileCipher::Aes128 => 128,
            AgileCipher::Aes192 => 192,
            AgileCipher::Aes256 => 256,
        }
    }

    const fn hash_bytes(self) -> usize {
        match self.hash {
            AgileHash::Sha1 => 20,
            AgileHash::Sha256 => 32,
            AgileHash::Sha512 => 64,
        }
    }

    const fn hash_name(self) -> &'static str {
        match self.hash {
            AgileHash::Sha1 => "SHA-1",
            AgileHash::Sha256 => "SHA256",
            AgileHash::Sha512 => "SHA512",
        }
    }

    const fn mode(self) -> Mode {
        Mode::agile(self.cipher, self.hash)
    }

    const fn encrypted_hash_bytes(self) -> usize {
        round_up_const(self.hash_bytes())
    }

    // Office's published SHA-1 profile stores a 20-byte integrity key in a
    // 32-byte encrypted block even though keyData.saltSize is 16. Preserve
    // that interoperable shape and generalize it to the selected digest.
    const fn integrity_bytes(self) -> usize {
        if self.hash_bytes() > SALT_BYTES {
            self.hash_bytes()
        } else {
            SALT_BYTES
        }
    }

    const fn encrypted_integrity_bytes(self) -> usize {
        round_up_const(self.integrity_bytes())
    }
}

const fn round_up_const(value: usize) -> usize {
    value.div_ceil(BLOCK) * BLOCK
}

struct Material {
    verifier_salt: [u8; SALT_BYTES],
    verifier: [u8; BLOCK],
    key_salt: [u8; SALT_BYTES],
    content_key: [u8; 32],
    integrity_salt: [u8; 64],
    integrity_len: usize,
}

impl Zeroize for Material {
    fn zeroize(&mut self) {
        self.verifier_salt.zeroize();
        self.verifier.zeroize();
        self.key_salt.zeroize();
        self.content_key.zeroize();
        self.integrity_salt.zeroize();
        self.integrity_len = 0;
    }
}

struct Info {
    profile: Profile,
    wrap_profile: Profile,
    spin_count: u32,
    key_salt: [u8; SALT_BYTES],
    verifier_salt: [u8; SALT_BYTES],
    encrypted_verifier: Zeroizing<Vec<u8>>,
    encrypted_verifier_hash: Zeroizing<Vec<u8>>,
    encrypted_key: Zeroizing<Vec<u8>>,
    encrypted_hmac_key: Option<Zeroizing<Vec<u8>>>,
    encrypted_hmac_value: Option<Zeroizing<Vec<u8>>>,
}

#[derive(Default)]
struct Parsed {
    profile: Option<Profile>,
    wrap_profile: Option<Profile>,
    key_salt: Option<[u8; SALT_BYTES]>,
    verifier_salt: Option<[u8; SALT_BYTES]>,
    encrypted_verifier: Option<Zeroizing<Vec<u8>>>,
    encrypted_verifier_hash: Option<Zeroizing<Vec<u8>>>,
    encrypted_key: Option<Zeroizing<Vec<u8>>>,
    encrypted_hmac_key: Option<Zeroizing<Vec<u8>>>,
    encrypted_hmac_value: Option<Zeroizing<Vec<u8>>>,
    spin_count: Option<u32>,
}

#[derive(Clone, Copy)]
enum Direction {
    Encrypt,
    Decrypt,
}

/// A digest kept inline so a high-spin password derivation does not allocate
/// once per iteration. The final password hash is copied into one bounded
/// zeroizing vector at the API boundary.
struct DigestValue {
    bytes: [u8; 64],
    len: usize,
}

impl DigestValue {
    fn as_slice(&self) -> &[u8] {
        &self.bytes[..self.len]
    }
}

impl Zeroize for DigestValue {
    fn zeroize(&mut self) {
        self.bytes.zeroize();
        self.len = 0;
    }
}

impl Drop for DigestValue {
    fn drop(&mut self) {
        self.zeroize();
    }
}

pub(super) fn profile_with_policy(
    info: &[u8],
    limits: &Limits,
    integrity_policy: IntegrityPolicy,
) -> Result<Mode> {
    Limits::bytes("EncryptionInfo", info.len(), limits.max_info_bytes)?;
    let parsed = parse(info, limits)?;
    require_integrity_policy(&parsed, integrity_policy)?;
    Ok(parsed.profile.mode())
}

pub(super) fn encrypt(
    package: Vec<u8>,
    password: &str,
    mode: Mode,
    limits: &Limits,
) -> Result<Vec<u8>> {
    let profile = Profile::from_mode(mode)?;
    let mut material = Zeroizing::new(Material {
        verifier_salt: [0; SALT_BYTES],
        verifier: [0; BLOCK],
        key_salt: [0; SALT_BYTES],
        content_key: [0; 32],
        integrity_salt: [0; 64],
        integrity_len: profile.integrity_bytes(),
    });
    let mut rng = SysRng;
    fill_random(&mut rng, &mut material.verifier_salt, "Agile verifier salt")?;
    fill_random(&mut rng, &mut material.verifier, "Agile verifier")?;
    fill_random(&mut rng, &mut material.key_salt, "Agile key salt")?;
    fill_random(
        &mut rng,
        &mut material.content_key[..profile.key_bytes()],
        "Agile content key",
    )?;
    let integrity_len = material.integrity_len;
    fill_random(
        &mut rng,
        &mut material.integrity_salt[..integrity_len],
        "Agile integrity salt",
    )?;
    let (info, encrypted) = encrypt_parts(package, password, profile, &material, limits)?;
    container::write(&info, encrypted, limits)
}

pub(super) fn decrypt_with_policy(
    info: &[u8],
    encrypted: Vec<u8>,
    password: &str,
    limits: &Limits,
    integrity_policy: IntegrityPolicy,
) -> Result<(Vec<u8>, IntegrityStatus)> {
    Limits::bytes("EncryptionInfo", info.len(), limits.max_info_bytes)?;
    Limits::bytes(
        "EncryptedPackage",
        encrypted.len(),
        limits.max_encrypted_bytes,
    )?;
    let parsed = parse(info, limits)?;
    let profile = parsed.profile;
    let wrap_profile = parsed.wrap_profile;
    let password_hash = password_hash(
        password,
        &parsed.verifier_salt,
        parsed.spin_count,
        wrap_profile.hash,
        limits,
    )?;
    let verifier = decrypt_value(
        &parsed.verifier_salt,
        &password_hash,
        wrap_profile,
        VERIFIER_INPUT_BLOCK,
        &parsed.encrypted_verifier,
        BLOCK,
    )?;
    let verifier_hash = decrypt_value(
        &parsed.verifier_salt,
        &password_hash,
        wrap_profile,
        HASHED_VERIFIER_BLOCK,
        &parsed.encrypted_verifier_hash,
        wrap_profile.hash_bytes(),
    )?;
    let expected_hash = digest(wrap_profile.hash, verifier.as_slice());
    if !bool::from(
        verifier_hash
            .as_slice()
            .get(..wrap_profile.hash_bytes())
            .ok_or_else(|| malformed("Agile verifier hash is shorter than its profile"))?
            .ct_eq(expected_hash.as_slice()),
    ) {
        return Err(Error::Password);
    }
    let content_key = decrypt_value(
        &parsed.verifier_salt,
        &password_hash,
        wrap_profile,
        CRYPTO_KEY_BLOCK,
        &parsed.encrypted_key,
        profile.key_bytes(),
    )?;
    let integrity = verify_integrity_or_policy(
        &parsed,
        profile,
        content_key.as_slice(),
        &encrypted,
        integrity_policy,
    )?;
    let clear_len = package_size(&encrypted, limits)?;
    let package = decrypt_package(
        encrypted,
        clear_len,
        profile,
        content_key.as_slice(),
        &parsed.key_salt,
    )?;
    Ok((package, integrity))
}

fn fill_random(rng: &mut SysRng, bytes: &mut [u8], name: &'static str) -> Result<()> {
    rng.try_fill_bytes(bytes)
        .map_err(|error| Error::Random(format!("{name}: {error}")))
}

fn encrypt_parts(
    package: Vec<u8>,
    password: &str,
    profile: Profile,
    material: &Material,
    limits: &Limits,
) -> Result<(Vec<u8>, Vec<u8>)> {
    encrypt_parts_with_wrapper(package, password, profile, profile, material, limits)
}

fn encrypt_parts_with_wrapper(
    package: Vec<u8>,
    password: &str,
    profile: Profile,
    wrap_profile: Profile,
    material: &Material,
    limits: &Limits,
) -> Result<(Vec<u8>, Vec<u8>)> {
    check_spin(SPIN_COUNT, limits)?;
    let password_hash = password_hash(
        password,
        &material.verifier_salt,
        SPIN_COUNT,
        wrap_profile.hash,
        limits,
    )?;
    let encrypted_verifier = encrypt_value(
        &material.verifier_salt,
        &password_hash,
        wrap_profile,
        VERIFIER_INPUT_BLOCK,
        &material.verifier,
    )?;
    let verifier_hash = digest(wrap_profile.hash, &material.verifier);
    let encrypted_verifier_hash = encrypt_value(
        &material.verifier_salt,
        &password_hash,
        wrap_profile,
        HASHED_VERIFIER_BLOCK,
        verifier_hash.as_slice(),
    )?;
    let encrypted_key = encrypt_value(
        &material.verifier_salt,
        &password_hash,
        wrap_profile,
        CRYPTO_KEY_BLOCK,
        &material.content_key[..profile.key_bytes()],
    )?;

    let encrypted = encrypt_package(
        package,
        profile,
        &material.content_key[..profile.key_bytes()],
        &material.key_salt,
        limits,
    )?;
    let integrity_value = hmac(
        profile.hash,
        &material.integrity_salt[..material.integrity_len],
        &encrypted,
    )?;
    let encrypted_hmac_key = encrypt_content(
        profile,
        &material.content_key[..profile.key_bytes()],
        &material.key_salt,
        INTEGRITY_KEY_BLOCK,
        &material.integrity_salt[..material.integrity_len],
    )?;
    let encrypted_hmac_value = encrypt_content(
        profile,
        &material.content_key[..profile.key_bytes()],
        &material.key_salt,
        INTEGRITY_VALUE_BLOCK,
        integrity_value.as_slice(),
    )?;
    let info = Info {
        profile,
        wrap_profile,
        spin_count: SPIN_COUNT,
        key_salt: material.key_salt,
        verifier_salt: material.verifier_salt,
        encrypted_verifier,
        encrypted_verifier_hash,
        encrypted_key,
        encrypted_hmac_key: Some(encrypted_hmac_key),
        encrypted_hmac_value: Some(encrypted_hmac_value),
    };
    Ok((build_info(&info, limits)?, encrypted))
}

fn password_hash(
    password: &str,
    salt: &[u8; SALT_BYTES],
    spin_count: u32,
    hash: AgileHash,
    limits: &Limits,
) -> Result<Zeroizing<Vec<u8>>> {
    check_spin(spin_count, limits)?;
    let encoded = password_bytes(password, limits)?;
    let mut hash_value = digest_parts_inline(hash, salt, encoded.as_slice());
    for iterator in 0..spin_count {
        hash_value = digest_parts_inline(hash, &iterator.to_le_bytes(), hash_value.as_slice());
    }
    let mut output = Zeroizing::new(Vec::new());
    output
        .try_reserve_exact(hash_value.len)
        .map_err(|_err| Error::Allocation("Agile password hash"))?;
    output.extend_from_slice(hash_value.as_slice());
    Ok(output)
}

fn check_spin(spin_count: u32, limits: &Limits) -> Result<()> {
    if spin_count > SPEC_MAX_SPIN_COUNT {
        return Err(malformed("Agile spinCount exceeds the schema maximum"));
    }
    if spin_count > limits.max_spin_count {
        return Err(Error::Limit {
            resource: "Agile spin count",
            actual: u64::from(spin_count),
            maximum: u64::from(limits.max_spin_count),
        });
    }
    Ok(())
}

fn digest(hash: AgileHash, bytes: &[u8]) -> Zeroizing<Vec<u8>> {
    match hash {
        AgileHash::Sha1 => Zeroizing::new(Sha1::digest(bytes).to_vec()),
        AgileHash::Sha256 => Zeroizing::new(Sha256::digest(bytes).to_vec()),
        AgileHash::Sha512 => Zeroizing::new(Sha512::digest(bytes).to_vec()),
    }
}

fn digest_parts(hash: AgileHash, first: &[u8], second: &[u8]) -> Zeroizing<Vec<u8>> {
    match hash {
        AgileHash::Sha1 => {
            let mut hasher = Sha1::new();
            hasher.update(first);
            hasher.update(second);
            Zeroizing::new(hasher.finalize().to_vec())
        },
        AgileHash::Sha256 => {
            let mut hasher = Sha256::new();
            hasher.update(first);
            hasher.update(second);
            Zeroizing::new(hasher.finalize().to_vec())
        },
        AgileHash::Sha512 => {
            let mut hasher = Sha512::new();
            hasher.update(first);
            hasher.update(second);
            Zeroizing::new(hasher.finalize().to_vec())
        },
    }
}

fn digest_parts_inline(hash: AgileHash, first: &[u8], second: &[u8]) -> DigestValue {
    let mut output = DigestValue {
        bytes: [0; 64],
        len: match hash {
            AgileHash::Sha1 => 20,
            AgileHash::Sha256 => 32,
            AgileHash::Sha512 => 64,
        },
    };
    match hash {
        AgileHash::Sha1 => {
            let mut hasher = Sha1::new();
            hasher.update(first);
            hasher.update(second);
            output.bytes[..20].copy_from_slice(&hasher.finalize());
        },
        AgileHash::Sha256 => {
            let mut hasher = Sha256::new();
            hasher.update(first);
            hasher.update(second);
            output.bytes[..32].copy_from_slice(&hasher.finalize());
        },
        AgileHash::Sha512 => {
            let mut hasher = Sha512::new();
            hasher.update(first);
            hasher.update(second);
            output.bytes[..64].copy_from_slice(&hasher.finalize());
        },
    }
    output
}

fn derive_key(password_hash: &[u8], profile: Profile, block_key: [u8; 8]) -> Zeroizing<Vec<u8>> {
    let digest = digest_parts(profile.hash, password_hash, &block_key);
    let mut key = Zeroizing::new(vec![0x36; profile.key_bytes()]);
    let copy_len = digest.len().min(key.len());
    key[..copy_len].copy_from_slice(&digest[..copy_len]);
    key
}

fn iv(
    salt: &[u8; SALT_BYTES],
    profile: Profile,
    block_key: Option<&[u8]>,
) -> Zeroizing<[u8; BLOCK]> {
    let mut output = Zeroizing::new([0x36; BLOCK]);
    match block_key {
        Some(block) => {
            let digest = digest_parts_inline(profile.hash, salt, block);
            let copy_len = digest.len.min(BLOCK);
            output.as_mut()[..copy_len].copy_from_slice(&digest.bytes[..copy_len]);
        },
        None => output.as_mut().copy_from_slice(salt),
    }
    output
}

fn encrypt_value(
    salt: &[u8; SALT_BYTES],
    password_hash: &[u8],
    profile: Profile,
    block_key: [u8; 8],
    input: &[u8],
) -> Result<Zeroizing<Vec<u8>>> {
    let output_len = round_up(input.len())?;
    let key = derive_key(password_hash, profile, block_key);
    let mut output = Zeroizing::new(vec![0u8; output_len]);
    output[..input.len()].copy_from_slice(input);
    cbc_crypt(
        profile,
        key.as_slice(),
        &iv(salt, profile, None),
        output.as_mut_slice(),
        Direction::Encrypt,
    )?;
    Ok(output)
}

fn decrypt_value(
    salt: &[u8; SALT_BYTES],
    password_hash: &[u8],
    profile: Profile,
    block_key: [u8; 8],
    encrypted: &[u8],
    output_len: usize,
) -> Result<Zeroizing<Vec<u8>>> {
    if encrypted.is_empty()
        || !encrypted.len().is_multiple_of(BLOCK)
        || encrypted.len() < output_len
    {
        return Err(malformed(
            "Agile encrypted password value has an invalid size",
        ));
    }
    let key = derive_key(password_hash, profile, block_key);
    let mut buffer = Zeroizing::new(encrypted.to_vec());
    cbc_crypt(
        profile,
        key.as_slice(),
        &iv(salt, profile, None),
        buffer.as_mut_slice(),
        Direction::Decrypt,
    )?;
    buffer.truncate(output_len);
    Ok(buffer)
}

fn encrypt_content(
    profile: Profile,
    key: &[u8],
    salt: &[u8; SALT_BYTES],
    block_key: [u8; 8],
    input: &[u8],
) -> Result<Zeroizing<Vec<u8>>> {
    let output_len = round_up(input.len())?;
    let mut output = Zeroizing::new(vec![0u8; output_len]);
    output[..input.len()].copy_from_slice(input);
    cbc_crypt(
        profile,
        key,
        &iv(salt, profile, Some(&block_key)),
        output.as_mut_slice(),
        Direction::Encrypt,
    )?;
    Ok(output)
}

fn decrypt_content(
    profile: Profile,
    key: &[u8],
    salt: &[u8; SALT_BYTES],
    block_key: [u8; 8],
    encrypted: &[u8],
    output_len: usize,
) -> Result<Zeroizing<Vec<u8>>> {
    if encrypted.is_empty()
        || !encrypted.len().is_multiple_of(BLOCK)
        || encrypted.len() < output_len
    {
        return Err(malformed("Agile integrity value has an invalid size"));
    }
    let mut buffer = Zeroizing::new(encrypted.to_vec());
    cbc_crypt(
        profile,
        key,
        &iv(salt, profile, Some(&block_key)),
        buffer.as_mut_slice(),
        Direction::Decrypt,
    )?;
    buffer.truncate(output_len);
    Ok(buffer)
}

fn hmac(hash: AgileHash, key: &[u8], bytes: &[u8]) -> Result<Zeroizing<Vec<u8>>> {
    macro_rules! run {
        ($ty:ty) => {{
            let mut mac = <Hmac<$ty> as KeyInit>::new_from_slice(key)
                .map_err(|_err| malformed("Agile HMAC key length invariant was violated"))?;
            mac.update(bytes);
            Zeroizing::new(mac.finalize().into_bytes().to_vec())
        }};
    }
    Ok(match hash {
        AgileHash::Sha1 => run!(Sha1),
        AgileHash::Sha256 => run!(Sha256),
        AgileHash::Sha512 => run!(Sha512),
    })
}

fn verify_integrity_or_policy(
    info: &Info,
    profile: Profile,
    content_key: &[u8],
    encrypted: &[u8],
    integrity_policy: IntegrityPolicy,
) -> Result<IntegrityStatus> {
    let (Some(encrypted_hmac_key), Some(encrypted_hmac_value)) = (
        info.encrypted_hmac_key.as_ref(),
        info.encrypted_hmac_value.as_ref(),
    ) else {
        return match integrity_policy {
            IntegrityPolicy::RequireAuthenticated => Err(Error::Unsupported(
                "Agile profile without authenticated dataIntegrity".into(),
            )),
            IntegrityPolicy::AllowUnauthenticated => Ok(IntegrityStatus::Unauthenticated),
        };
    };
    let integrity_len = integrity_key_length(profile, encrypted_hmac_key.len())?;
    let integrity_salt = decrypt_content(
        profile,
        content_key,
        &info.key_salt,
        INTEGRITY_KEY_BLOCK,
        encrypted_hmac_key,
        integrity_len,
    )?;
    let stored = decrypt_content(
        profile,
        content_key,
        &info.key_salt,
        INTEGRITY_VALUE_BLOCK,
        encrypted_hmac_value,
        profile.hash_bytes(),
    )?;
    let expected = hmac(profile.hash, integrity_salt.as_slice(), encrypted)?;
    if !bool::from(stored.as_slice().ct_eq(expected.as_slice())) {
        return Err(Error::Integrity);
    }
    Ok(IntegrityStatus::Authenticated)
}

fn integrity_key_length(profile: Profile, encrypted_length: usize) -> Result<usize> {
    if encrypted_length == BLOCK {
        return Ok(SALT_BYTES);
    }
    if encrypted_length == profile.encrypted_integrity_bytes() {
        return Ok(profile.integrity_bytes());
    }
    Err(malformed(format!(
        "Agile encryptedHmacKey has {encrypted_length} bytes, expected {BLOCK} or {}",
        profile.encrypted_integrity_bytes()
    )))
}

fn encrypt_package(
    mut package: Vec<u8>,
    profile: Profile,
    content_key: &[u8],
    key_salt: &[u8; SALT_BYTES],
    limits: &Limits,
) -> Result<Vec<u8>> {
    let clear_len = package.len();
    let cipher_len = round_up(clear_len)?;
    let total = cipher_len
        .checked_add(8)
        .ok_or_else(|| malformed("Agile EncryptedPackage size overflows usize"))?;
    Limits::bytes("EncryptedPackage", total, limits.max_encrypted_bytes)?;
    package
        .try_reserve_exact(total.saturating_sub(package.len()))
        .map_err(|_err| Error::Allocation("Agile EncryptedPackage"))?;
    package.resize(total, 0);
    package.copy_within(0..clear_len, 8);
    package
        .get_mut(..8)
        .ok_or_else(|| malformed("Agile EncryptedPackage prefix is unavailable"))?
        .copy_from_slice(
            &u64::try_from(clear_len)
                .map_err(|_err| malformed("plaintext size does not fit u64"))?
                .to_le_bytes(),
        );
    crypt_segments(
        &mut package,
        clear_len,
        profile,
        content_key,
        key_salt,
        Direction::Encrypt,
    )?;
    Ok(package)
}

fn package_size(encrypted: &[u8], limits: &Limits) -> Result<usize> {
    if encrypted.len() < 8 + BLOCK {
        return Err(malformed("Agile EncryptedPackage is too short"));
    }
    let prefix = encrypted
        .get(..8)
        .ok_or_else(|| malformed("Agile EncryptedPackage has no StreamSize"))?;
    let clear_len = declared_size(
        u64::from_le_bytes(
            prefix
                .try_into()
                .map_err(|_| malformed("Agile StreamSize has the wrong length"))?,
        ),
        limits,
    )?;
    if clear_len == 0 {
        return Err(malformed(
            "Agile EncryptedPackage declares an empty package",
        ));
    }
    let expected = round_up(clear_len)?
        .checked_add(8)
        .ok_or_else(|| malformed("Agile EncryptedPackage length overflows usize"))?;
    if encrypted.len() != expected {
        return Err(malformed(
            "Agile EncryptedPackage length disagrees with StreamSize",
        ));
    }
    Ok(clear_len)
}

fn decrypt_package(
    mut encrypted: Vec<u8>,
    clear_len: usize,
    profile: Profile,
    content_key: &[u8],
    key_salt: &[u8; SALT_BYTES],
) -> Result<Vec<u8>> {
    crypt_segments(
        &mut encrypted,
        clear_len,
        profile,
        content_key,
        key_salt,
        Direction::Decrypt,
    )?;
    let source_end = clear_len
        .checked_add(8)
        .ok_or_else(|| malformed("Agile plaintext range overflows usize"))?;
    if encrypted.get(8..source_end).is_none() {
        return Err(malformed("Agile decrypted package is truncated"));
    }
    encrypted.copy_within(8..source_end, 0);
    encrypted.truncate(clear_len);
    Ok(encrypted)
}

fn crypt_segments(
    bytes: &mut [u8],
    clear_len: usize,
    profile: Profile,
    content_key: &[u8],
    key_salt: &[u8; SALT_BYTES],
    direction: Direction,
) -> Result<()> {
    let segments = clear_len
        .checked_add(SEGMENT - 1)
        .ok_or_else(|| malformed("Agile segment count overflows usize"))?
        / SEGMENT;
    for index in 0..segments {
        let clear_start = index
            .checked_mul(SEGMENT)
            .ok_or_else(|| malformed("Agile segment offset overflows usize"))?;
        let clear_segment = (clear_len - clear_start).min(SEGMENT);
        let cipher_segment = round_up(clear_segment)?;
        let start = clear_start
            .checked_add(8)
            .ok_or_else(|| malformed("Agile segment start overflows usize"))?;
        let end = start
            .checked_add(cipher_segment)
            .ok_or_else(|| malformed("Agile segment end overflows usize"))?;
        let segment = bytes
            .get_mut(start..end)
            .ok_or_else(|| malformed("Agile encrypted segment is truncated"))?;
        let block = u32::try_from(index)
            .map_err(|_err| malformed("Agile segment index exceeds u32"))?
            .to_le_bytes();
        cbc_crypt(
            profile,
            content_key,
            &iv(key_salt, profile, Some(&block)),
            segment,
            direction,
        )?;
    }
    Ok(())
}

fn cbc_crypt(
    profile: Profile,
    key: &[u8],
    iv: &[u8; BLOCK],
    bytes: &mut [u8],
    direction: Direction,
) -> Result<()> {
    if !bytes.len().is_multiple_of(BLOCK) {
        return Err(malformed("Agile AES data is not block aligned"));
    }
    let message_len = bytes.len();
    macro_rules! run {
        ($ty:ty) => {{
            match direction {
                Direction::Encrypt => Encryptor::<$ty>::new_from_slices(key, iv)
                    .map_err(|_err| malformed("Agile AES key or IV length invariant was violated"))?
                    .encrypt_padded::<NoPadding>(bytes, message_len)
                    .map(|_| ())
                    .map_err(|_err| malformed("Agile CBC encryption failed")),
                Direction::Decrypt => Decryptor::<$ty>::new_from_slices(key, iv)
                    .map_err(|_err| malformed("Agile AES key or IV length invariant was violated"))?
                    .decrypt_padded::<NoPadding>(bytes)
                    .map(|_| ())
                    .map_err(|_err| malformed("Agile CBC decryption failed")),
            }
        }};
    }
    match profile.cipher {
        AgileCipher::Aes128 => run!(Aes128),
        AgileCipher::Aes192 => run!(Aes192),
        AgileCipher::Aes256 => run!(Aes256),
    }
}

fn round_up(value: usize) -> Result<usize> {
    value
        .checked_add(BLOCK - 1)
        .map(|padded| padded / BLOCK * BLOCK)
        .ok_or_else(|| malformed("Agile block length overflows usize"))
}

fn build_info(info: &Info, limits: &Limits) -> Result<Vec<u8>> {
    let key_salt = BASE64.encode(info.key_salt);
    let verifier_salt = BASE64.encode(info.verifier_salt);
    let encrypted_verifier = BASE64.encode(&info.encrypted_verifier);
    let encrypted_verifier_hash = BASE64.encode(&info.encrypted_verifier_hash);
    let encrypted_key = BASE64.encode(&info.encrypted_key);
    let encrypted_hmac_key = BASE64.encode(
        info.encrypted_hmac_key
            .as_ref()
            .ok_or_else(|| malformed("Agile authoring requires encrypted HMAC key"))?,
    );
    let encrypted_hmac_value = BASE64.encode(
        info.encrypted_hmac_value
            .as_ref()
            .ok_or_else(|| malformed("Agile authoring requires encrypted HMAC value"))?,
    );
    let mut xml = String::new();
    xml.try_reserve(1_024)
        .map_err(|_err| Error::Allocation("Agile EncryptionInfo XML"))?;
    write!(
        xml,
        concat!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
            r#"<encryption xmlns="http://schemas.microsoft.com/office/2006/encryption" xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password">"#,
            r#"<keyData saltSize="16" blockSize="16" keyBits="{}" hashSize="{}" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="{}" saltValue="{}"/>"#,
            r#"<dataIntegrity encryptedHmacKey="{}" encryptedHmacValue="{}"/>"#,
            r#"<keyEncryptors><keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password">"#,
            r#"<p:encryptedKey spinCount="{}" saltSize="16" blockSize="16" keyBits="{}" hashSize="{}" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="{}" saltValue="{}" encryptedVerifierHashInput="{}" encryptedVerifierHashValue="{}" encryptedKeyValue="{}"/>"#,
            r#"</keyEncryptor></keyEncryptors></encryption>"#,
        ),
        info.profile.key_bits(),
        info.profile.hash_bytes(),
        info.profile.hash_name(),
        key_salt,
        encrypted_hmac_key,
        encrypted_hmac_value,
        info.spin_count,
        info.wrap_profile.key_bits(),
        info.wrap_profile.hash_bytes(),
        info.wrap_profile.hash_name(),
        verifier_salt,
        encrypted_verifier,
        encrypted_verifier_hash,
        encrypted_key,
    )
    .map_err(|_err| Error::Allocation("Agile EncryptionInfo XML"))?;
    Limits::bytes("Agile XML", xml.len(), limits.max_xml_bytes)?;
    let total = xml
        .len()
        .checked_add(8)
        .ok_or_else(|| malformed("Agile EncryptionInfo length overflows usize"))?;
    Limits::bytes("EncryptionInfo", total, limits.max_info_bytes)?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(total)
        .map_err(|_err| Error::Allocation("Agile EncryptionInfo"))?;
    output.extend_from_slice(&4u16.to_le_bytes());
    output.extend_from_slice(&4u16.to_le_bytes());
    output.extend_from_slice(&0x40u32.to_le_bytes());
    output.extend_from_slice(xml.as_bytes());
    Ok(output)
}

fn parse(bytes: &[u8], limits: &Limits) -> Result<Info> {
    Limits::bytes("EncryptionInfo", bytes.len(), limits.max_info_bytes)?;
    let prefix = bytes
        .get(..8)
        .ok_or_else(|| malformed("Agile EncryptionInfo is shorter than its header"))?;
    let major = u16::from_le_bytes([prefix[0], prefix[1]]);
    let minor = u16::from_le_bytes([prefix[2], prefix[3]]);
    if (major, minor) != (4, 4) {
        return Err(Error::Unsupported(format!(
            "Agile EncryptionInfo version {major}.{minor}"
        )));
    }
    let reserved = u32::from_le_bytes([prefix[4], prefix[5], prefix[6], prefix[7]]);
    if reserved != 0x40 {
        return Err(malformed(format!(
            "Agile EncryptionInfo reserved field is {reserved:#x}, expected 0x40"
        )));
    }
    let xml = bytes
        .get(8..)
        .ok_or_else(|| malformed("Agile EncryptionInfo has no XML descriptor"))?;
    if xml.is_empty() {
        return Err(malformed("Agile EncryptionInfo XML is empty"));
    }
    Limits::bytes("Agile XML", xml.len(), limits.max_xml_bytes)?;

    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().check_comments = true;
    reader.config_mut().expand_empty_elements = true;
    let mut phase = 0u8;
    let mut depth = 0usize;
    let mut nodes = 0usize;
    let mut attributes = 0usize;
    let mut declaration_seen = false;
    let mut first_event = true;
    let mut parsed = Parsed::default();
    let mut key_encryptors = 0u8;

    loop {
        let decoder = reader.decoder();
        let (namespace, event) = reader
            .read_resolved_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        let was_first = first_event;
        first_event = false;
        match event {
            Event::Decl(declaration) => {
                count(&mut nodes, limits.max_xml_nodes, "Agile XML nodes")?;
                if !was_first || declaration_seen || phase != 0 {
                    return Err(malformed(
                        "Agile XML declaration must be the first event and occur once",
                    ));
                }
                validate_declaration(&declaration)?;
                declaration_seen = true;
            },
            Event::Start(element) => {
                count(&mut nodes, limits.max_xml_nodes, "Agile XML nodes")?;
                let child_depth = depth
                    .checked_add(1)
                    .ok_or_else(|| limit("Agile XML depth", usize::MAX, limits.max_xml_depth))?;
                if child_depth > limits.max_xml_depth {
                    return Err(limit("Agile XML depth", child_depth, limits.max_xml_depth));
                }
                match phase {
                    0 => {
                        element_is(&namespace, &element, ENC_NS, b"encryption")?;
                        exact_attributes(
                            &reader,
                            &element,
                            decoder,
                            &[],
                            &mut attributes,
                            limits,
                            |_, _| Ok(()),
                        )?;
                        phase = 1;
                    },
                    1 => {
                        element_is(&namespace, &element, ENC_NS, b"keyData")?;
                        let (salt, profile) =
                            parse_key_data(&reader, &element, decoder, &mut attributes, limits)?;
                        parsed.key_salt = Some(salt);
                        parsed.profile = Some(profile);
                        phase = 2;
                    },
                    3 => {
                        if element.local_name().as_ref() == b"keyEncryptors" {
                            element_is(&namespace, &element, ENC_NS, b"keyEncryptors")?;
                            exact_attributes(
                                &reader,
                                &element,
                                decoder,
                                &[],
                                &mut attributes,
                                limits,
                                |_, _| Ok(()),
                            )?;
                            phase = 6;
                        } else {
                            element_is(&namespace, &element, ENC_NS, b"dataIntegrity")?;
                            let profile = parsed
                                .profile
                                .ok_or_else(|| malformed("Agile dataIntegrity precedes keyData"))?;
                            let (key, value) = parse_data_integrity(
                                &reader,
                                &element,
                                decoder,
                                &mut attributes,
                                limits,
                                profile,
                            )?;
                            parsed.encrypted_hmac_key = Some(key);
                            parsed.encrypted_hmac_value = Some(value);
                            phase = 4;
                        }
                    },
                    5 => {
                        element_is(&namespace, &element, ENC_NS, b"keyEncryptors")?;
                        exact_attributes(
                            &reader,
                            &element,
                            decoder,
                            &[],
                            &mut attributes,
                            limits,
                            |_, _| Ok(()),
                        )?;
                        phase = 6;
                    },
                    6 => {
                        element_is(&namespace, &element, ENC_NS, b"keyEncryptor")?;
                        if key_encryptors != 0 {
                            return Err(malformed(
                                "Agile XML contains multiple keyEncryptor elements",
                            ));
                        }
                        key_encryptors = 1;
                        parse_key_encryptor(&reader, &element, decoder, &mut attributes, limits)?;
                        phase = 7;
                    },
                    7 => {
                        element_is(&namespace, &element, PASSWORD_NS, b"encryptedKey")?;
                        let profile = parsed
                            .profile
                            .ok_or_else(|| malformed("Agile encryptedKey precedes keyData"))?;
                        parse_password_key(
                            &reader,
                            &element,
                            decoder,
                            &mut attributes,
                            limits,
                            profile,
                            &mut parsed,
                        )?;
                        phase = 8;
                    },
                    _ => return Err(malformed("Agile XML element order is invalid")),
                }
                depth = child_depth;
            },
            Event::End(element) => {
                count(&mut nodes, limits.max_xml_nodes, "Agile XML nodes")?;
                match phase {
                    2 => {
                        end_is(&namespace, &element, ENC_NS, b"keyData")?;
                        phase = 3;
                    },
                    4 => {
                        end_is(&namespace, &element, ENC_NS, b"dataIntegrity")?;
                        phase = 5;
                    },
                    8 => {
                        end_is(&namespace, &element, PASSWORD_NS, b"encryptedKey")?;
                        phase = 9;
                    },
                    9 => {
                        end_is(&namespace, &element, ENC_NS, b"keyEncryptor")?;
                        phase = 10;
                    },
                    10 => {
                        end_is(&namespace, &element, ENC_NS, b"keyEncryptors")?;
                        phase = 11;
                    },
                    11 => {
                        end_is(&namespace, &element, ENC_NS, b"encryption")?;
                        phase = 12;
                    },
                    _ => return Err(malformed("Agile XML end-element order is invalid")),
                }
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| malformed("Agile XML has an unexpected end element"))?;
            },
            Event::Text(text) => {
                count(&mut nodes, limits.max_xml_nodes, "Agile XML nodes")?;
                let value = text
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|error| Error::Xml(error.to_string()))?;
                if !value.trim().is_empty() {
                    return Err(malformed("Agile XML cannot contain character data"));
                }
            },
            Event::Comment(comment) => {
                count(&mut nodes, limits.max_xml_nodes, "Agile XML nodes")?;
                comment
                    .decode()
                    .map_err(|error| Error::Xml(error.to_string()))?;
            },
            Event::PI(instruction) => {
                count(&mut nodes, limits.max_xml_nodes, "Agile XML nodes")?;
                decoder
                    .decode(instruction.as_ref())
                    .map_err(|error| Error::Xml(error.to_string()))?;
            },
            Event::DocType(_) => return Err(malformed("DTD is forbidden in Agile XML")),
            Event::CData(_) | Event::GeneralRef(_) => {
                return Err(malformed("Agile XML cannot contain CDATA or entity nodes"));
            },
            Event::Empty(_) => return Err(malformed("Agile XML empty-element expansion failed")),
            Event::Eof => break,
        }
    }
    if phase != 12 || depth != 0 || key_encryptors != 1 {
        return Err(malformed("Agile XML descriptor is incomplete"));
    }
    if parsed.encrypted_hmac_key.is_some() != parsed.encrypted_hmac_value.is_some() {
        return Err(malformed("Agile dataIntegrity is incomplete"));
    }
    let profile = parsed
        .profile
        .ok_or_else(|| malformed("Agile XML has no keyData"))?;
    let wrap_profile = parsed
        .wrap_profile
        .ok_or_else(|| malformed("Agile XML has no encryptedKey profile"))?;
    Ok(Info {
        profile,
        wrap_profile,
        spin_count: parsed
            .spin_count
            .ok_or_else(|| malformed("Agile encryptedKey has no spinCount"))?,
        key_salt: parsed
            .key_salt
            .ok_or_else(|| malformed("Agile XML has no keyData salt"))?,
        verifier_salt: parsed
            .verifier_salt
            .ok_or_else(|| malformed("Agile XML has no password salt"))?,
        encrypted_verifier: parsed
            .encrypted_verifier
            .ok_or_else(|| malformed("Agile XML has no encrypted verifier"))?,
        encrypted_verifier_hash: parsed
            .encrypted_verifier_hash
            .ok_or_else(|| malformed("Agile XML has no encrypted verifier hash"))?,
        encrypted_key: parsed
            .encrypted_key
            .ok_or_else(|| malformed("Agile XML has no encrypted content key"))?,
        encrypted_hmac_key: parsed.encrypted_hmac_key,
        encrypted_hmac_value: parsed.encrypted_hmac_value,
    })
}

fn require_integrity_policy(info: &Info, policy: IntegrityPolicy) -> Result<()> {
    if info.encrypted_hmac_key.is_some() {
        return Ok(());
    }
    match policy {
        IntegrityPolicy::RequireAuthenticated => Err(Error::Unsupported(
            "Agile profile without authenticated dataIntegrity".into(),
        )),
        IntegrityPolicy::AllowUnauthenticated => Ok(()),
    }
}

fn element_is(
    namespace: &ResolveResult<'_>,
    element: &BytesStart<'_>,
    expected_namespace: &[u8],
    expected_local: &[u8],
) -> Result<()> {
    namespace_is(namespace, expected_namespace)?;
    if element.local_name().as_ref() != expected_local {
        return Err(malformed(format!(
            "unexpected Agile XML element '{}'",
            String::from_utf8_lossy(element.name().as_ref())
        )));
    }
    Ok(())
}

fn end_is(
    namespace: &ResolveResult<'_>,
    element: &quick_xml::events::BytesEnd<'_>,
    expected_namespace: &[u8],
    expected_local: &[u8],
) -> Result<()> {
    namespace_is(namespace, expected_namespace)?;
    if element.local_name().as_ref() != expected_local {
        return Err(malformed(format!(
            "unexpected Agile XML end element '{}'",
            String::from_utf8_lossy(element.name().as_ref())
        )));
    }
    Ok(())
}

fn namespace_is(namespace: &ResolveResult<'_>, expected: &[u8]) -> Result<()> {
    match namespace {
        ResolveResult::Bound(Namespace(actual)) if *actual == expected => Ok(()),
        ResolveResult::Bound(Namespace(actual)) => Err(malformed(format!(
            "unexpected Agile XML namespace '{}'",
            String::from_utf8_lossy(actual)
        ))),
        ResolveResult::Unbound => Err(malformed("Agile XML element has no namespace")),
        ResolveResult::Unknown(prefix) => Err(malformed(format!(
            "Agile XML namespace prefix '{}' is unbound",
            String::from_utf8_lossy(prefix)
        ))),
    }
}

fn parse_key_data(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    decoder: Decoder,
    total: &mut usize,
    limits: &Limits,
) -> Result<([u8; SALT_BYTES], Profile)> {
    const NAMES: &[&[u8]] = &[
        b"saltSize",
        b"blockSize",
        b"keyBits",
        b"hashSize",
        b"cipherAlgorithm",
        b"cipherChaining",
        b"hashAlgorithm",
        b"saltValue",
    ];
    let mut salt = None;
    let mut key_bits = None;
    let mut hash_size = None;
    let mut hash = None;
    exact_attributes(
        reader,
        element,
        decoder,
        NAMES,
        total,
        limits,
        |name, value| match name {
            b"saltSize" => exact_number(value, SALT_BYTES as u32, "keyData.saltSize"),
            b"blockSize" => exact_number(value, BLOCK as u32, "keyData.blockSize"),
            b"keyBits" => {
                key_bits = Some(number(value, "keyData.keyBits")?);
                Ok(())
            },
            b"hashSize" => {
                hash_size = Some(number(value, "keyData.hashSize")?);
                Ok(())
            },
            b"cipherAlgorithm" => exact_text(value, "AES", "keyData.cipherAlgorithm"),
            b"cipherChaining" => exact_text(value, "ChainingModeCBC", "keyData.cipherChaining"),
            b"hashAlgorithm" => {
                hash = Some(parse_hash(value, "keyData.hashAlgorithm")?);
                Ok(())
            },
            b"saltValue" => {
                salt = Some(decode_fixed::<SALT_BYTES>(value, "keyData.saltValue")?);
                Ok(())
            },
            _ => Err(malformed("unknown keyData attribute")),
        },
    )?;
    let key_bits = key_bits.ok_or_else(|| malformed("keyData.keyBits is missing"))?;
    let hash = hash.ok_or_else(|| malformed("keyData.hashAlgorithm is missing"))?;
    let profile = profile_from_values(
        key_bits,
        hash_size.ok_or_else(|| malformed("keyData.hashSize is missing"))?,
        hash,
        "keyData",
    )?;
    Ok((
        salt.ok_or_else(|| malformed("keyData.saltValue is missing"))?,
        profile,
    ))
}

fn parse_data_integrity(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    decoder: Decoder,
    total: &mut usize,
    limits: &Limits,
    profile: Profile,
) -> Result<(Zeroizing<Vec<u8>>, Zeroizing<Vec<u8>>)> {
    const NAMES: &[&[u8]] = &[b"encryptedHmacKey", b"encryptedHmacValue"];
    let mut key = None;
    let mut value = None;
    exact_attributes(
        reader,
        element,
        decoder,
        NAMES,
        total,
        limits,
        |name, raw| match name {
            b"encryptedHmacKey" => {
                key = Some(decode_integrity_key(raw, profile)?);
                Ok(())
            },
            b"encryptedHmacValue" => {
                value = Some(decode_array(
                    raw,
                    "dataIntegrity.encryptedHmacValue",
                    profile.encrypted_integrity_bytes(),
                )?);
                Ok(())
            },
            _ => Err(malformed("unknown dataIntegrity attribute")),
        },
    )?;
    Ok((
        key.ok_or_else(|| malformed("dataIntegrity.encryptedHmacKey is missing"))?,
        value.ok_or_else(|| malformed("dataIntegrity.encryptedHmacValue is missing"))?,
    ))
}

fn decode_integrity_key(raw: &str, profile: Profile) -> Result<Zeroizing<Vec<u8>>> {
    let office_size = profile.encrypted_integrity_bytes();
    let max_size = office_size.max(BLOCK);
    let max_encoded = max_size
        .checked_add(2)
        .and_then(|length| length.checked_mul(4).map(|value| value / 3 + 4))
        .ok_or_else(|| malformed("dataIntegrity.encryptedHmacKey length overflows usize"))?;
    if raw.len() > max_encoded {
        return Err(malformed(
            "dataIntegrity.encryptedHmacKey exceeds its bounded encoded size",
        ));
    }
    let decoded = BASE64.decode(raw).map_err(|error| {
        malformed(format!(
            "dataIntegrity.encryptedHmacKey is not valid base64: {error}"
        ))
    })?;
    if decoded.len() != BLOCK && decoded.len() != office_size {
        return Err(malformed(format!(
            "dataIntegrity.encryptedHmacKey has {} bytes, expected {BLOCK} or {office_size}",
            decoded.len()
        )));
    }
    Ok(Zeroizing::new(decoded))
}

fn parse_key_encryptor(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    decoder: Decoder,
    total: &mut usize,
    limits: &Limits,
) -> Result<()> {
    exact_attributes(
        reader,
        element,
        decoder,
        &[b"uri"],
        total,
        limits,
        |_, value| {
            exact_text(
                value,
                "http://schemas.microsoft.com/office/2006/keyEncryptor/password",
                "keyEncryptor.uri",
            )
        },
    )
}

fn parse_password_key(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    decoder: Decoder,
    total: &mut usize,
    limits: &Limits,
    profile: Profile,
    parsed: &mut Parsed,
) -> Result<()> {
    const NAMES: &[&[u8]] = &[
        b"spinCount",
        b"saltSize",
        b"blockSize",
        b"keyBits",
        b"hashSize",
        b"cipherAlgorithm",
        b"cipherChaining",
        b"hashAlgorithm",
        b"saltValue",
        b"encryptedVerifierHashInput",
        b"encryptedVerifierHashValue",
        b"encryptedKeyValue",
    ];
    let mut wrap_key_bits = None;
    let mut wrap_hash_size = None;
    let mut wrap_hash = None;
    let mut encoded_verifier_hash = None;
    exact_attributes(
        reader,
        element,
        decoder,
        NAMES,
        total,
        limits,
        |name, value| match name {
            b"spinCount" => {
                let count = number(value, "encryptedKey.spinCount")?;
                check_spin(count, limits)?;
                parsed.spin_count = Some(count);
                Ok(())
            },
            b"saltSize" => exact_number(value, SALT_BYTES as u32, "encryptedKey.saltSize"),
            b"blockSize" => exact_number(value, BLOCK as u32, "encryptedKey.blockSize"),
            b"keyBits" => {
                wrap_key_bits = Some(number(value, "encryptedKey.keyBits")?);
                Ok(())
            },
            b"hashSize" => {
                wrap_hash_size = Some(number(value, "encryptedKey.hashSize")?);
                Ok(())
            },
            b"cipherAlgorithm" => exact_text(value, "AES", "encryptedKey.cipherAlgorithm"),
            b"cipherChaining" => {
                exact_text(value, "ChainingModeCBC", "encryptedKey.cipherChaining")
            },
            b"hashAlgorithm" => {
                wrap_hash = Some(parse_hash(value, "encryptedKey.hashAlgorithm")?);
                Ok(())
            },
            b"saltValue" => {
                parsed.verifier_salt =
                    Some(decode_fixed::<SALT_BYTES>(value, "encryptedKey.saltValue")?);
                Ok(())
            },
            b"encryptedVerifierHashInput" => {
                parsed.encrypted_verifier =
                    Some(decode_array(value, "encryptedVerifierHashInput", BLOCK)?);
                Ok(())
            },
            b"encryptedVerifierHashValue" => {
                encoded_verifier_hash = Some(Zeroizing::new(value.to_owned()));
                Ok(())
            },
            b"encryptedKeyValue" => {
                parsed.encrypted_key = Some(decode_array(
                    value,
                    "encryptedKeyValue",
                    round_up_const(profile.key_bytes()),
                )?);
                Ok(())
            },
            _ => Err(malformed("unknown encryptedKey attribute")),
        },
    )?;
    let wrap_key_bits =
        wrap_key_bits.ok_or_else(|| malformed("encryptedKey.keyBits is missing"))?;
    let wrap_hash_size =
        wrap_hash_size.ok_or_else(|| malformed("encryptedKey.hashSize is missing"))?;
    let wrap_hash = wrap_hash.ok_or_else(|| malformed("encryptedKey.hashAlgorithm is missing"))?;
    let wrap_profile =
        profile_from_values(wrap_key_bits, wrap_hash_size, wrap_hash, "encryptedKey")?;
    if !is_supported_wrapper_profile(profile, wrap_profile) {
        return Err(malformed(
            "encryptedKey profile is not supported with keyData",
        ));
    }
    let encoded_verifier_hash = encoded_verifier_hash
        .ok_or_else(|| malformed("encryptedKey.encryptedVerifierHashValue is missing"))?;
    parsed.encrypted_verifier_hash = Some(decode_array(
        encoded_verifier_hash.as_str(),
        "encryptedVerifierHashValue",
        wrap_profile.encrypted_hash_bytes(),
    )?);
    parsed.wrap_profile = Some(wrap_profile);
    Ok(())
}

fn is_supported_wrapper_profile(data: Profile, wrapper: Profile) -> bool {
    wrapper.hash == data.hash
        // A POI 60320 workbook uses an AES-256/SHA-512 password wrapper for
        // an AES-128/SHA-1 keyData profile. Keep that producer quirk narrow;
        // all other hash mismatches remain rejected against the published
        // PasswordKeyEncryptor requirements.
        || (data.cipher == AgileCipher::Aes128
            && data.hash == AgileHash::Sha1
            && wrapper.cipher == AgileCipher::Aes256
            && wrapper.hash == AgileHash::Sha512)
}

fn profile_from_values(
    key_bits: u32,
    hash_size: u32,
    hash: AgileHash,
    field: &'static str,
) -> Result<Profile> {
    let cipher = match key_bits {
        128 => AgileCipher::Aes128,
        192 => AgileCipher::Aes192,
        256 => AgileCipher::Aes256,
        _ => return Err(Error::Unsupported(format!("{field}.keyBits is {key_bits}"))),
    };
    let profile = Profile { cipher, hash };
    if hash_size != profile.hash_bytes() as u32 {
        return Err(Error::Unsupported(format!(
            "{field}.hashSize is {hash_size}, expected {}",
            profile.hash_bytes()
        )));
    }
    Ok(profile)
}

fn parse_hash(value: &str, field: &'static str) -> Result<AgileHash> {
    match value {
        // Office and several compatible producers use both spellings. Keep
        // the canonical writer spelling in `Profile::hash_name`.
        "SHA-1" | "SHA1" => Ok(AgileHash::Sha1),
        "SHA256" => Ok(AgileHash::Sha256),
        "SHA512" => Ok(AgileHash::Sha512),
        _ => Err(Error::Unsupported(format!("{field} is '{value}'"))),
    }
}

fn exact_attributes(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    decoder: Decoder,
    allowed: &[&[u8]],
    total: &mut usize,
    limits: &Limits,
    mut visitor: impl FnMut(&[u8], &str) -> Result<()>,
) -> Result<()> {
    if allowed.len() > u16::BITS as usize {
        return Err(Error::InvalidLimit(
            "Agile attribute schema exceeds its bitset",
        ));
    }
    let mut seen = 0u16;
    for raw_attribute in element.checked_attributes() {
        count(total, limits.max_xml_attributes, "Agile XML attributes")?;
        let attribute = raw_attribute.map_err(|error| Error::Xml(error.to_string()))?;
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
            .map_err(|error| Error::Xml(error.to_string()))?;
        let name = attribute.key.as_ref();
        if is_namespace_declaration(name) {
            validate_namespace_declaration(name, &value)?;
            continue;
        }
        match reader.resolver().resolve_attribute(attribute.key).0 {
            ResolveResult::Unbound => {},
            ResolveResult::Bound(_) => {
                return Err(malformed(format!(
                    "Agile attribute '{}' must be unqualified",
                    String::from_utf8_lossy(name)
                )));
            },
            ResolveResult::Unknown(prefix) => {
                return Err(malformed(format!(
                    "Agile attribute prefix '{}' is unbound",
                    String::from_utf8_lossy(&prefix)
                )));
            },
        }
        let index = allowed
            .iter()
            .position(|expected| *expected == name)
            .ok_or_else(|| {
                malformed(format!(
                    "unexpected Agile attribute '{}'",
                    String::from_utf8_lossy(name)
                ))
            })?;
        let bit = 1u16
            .checked_shl(
                u32::try_from(index).map_err(|_| malformed("Agile attribute index exceeds u32"))?,
            )
            .ok_or_else(|| malformed("Agile attribute index exceeds its bitset"))?;
        if seen & bit != 0 {
            return Err(malformed(format!(
                "duplicate Agile attribute '{}'",
                String::from_utf8_lossy(name)
            )));
        }
        seen |= bit;
        visitor(name, &value)?;
    }
    for (index, name) in allowed.iter().enumerate() {
        let bit = 1u16
            .checked_shl(
                u32::try_from(index).map_err(|_| malformed("Agile attribute index exceeds u32"))?,
            )
            .ok_or_else(|| malformed("Agile attribute index exceeds its bitset"))?;
        if seen & bit == 0 {
            return Err(malformed(format!(
                "missing Agile attribute '{}'",
                String::from_utf8_lossy(name)
            )));
        }
    }
    Ok(())
}

fn is_namespace_declaration(name: &[u8]) -> bool {
    name == b"xmlns" || name.starts_with(b"xmlns:")
}

fn validate_namespace_declaration(name: &[u8], value: &str) -> Result<()> {
    if value.is_empty()
        || value
            .bytes()
            .any(|byte| matches!(byte, b' ' | b'\t' | b'\r' | b'\n'))
    {
        return Err(malformed(
            "Agile namespace URI is empty or contains whitespace",
        ));
    }
    if value == "http://www.w3.org/2000/xmlns/" {
        return Err(malformed("the xmlns namespace cannot be rebound"));
    }
    if let Some(prefix) = name.strip_prefix(b"xmlns:") {
        if prefix.is_empty() || prefix.contains(&b':') || prefix == b"xmlns" {
            return Err(malformed("Agile namespace prefix is invalid"));
        }
        if (prefix == b"xml") != (value == "http://www.w3.org/XML/1998/namespace") {
            return Err(malformed("the XML namespace may be bound only to 'xml'"));
        }
    } else if value == "http://www.w3.org/XML/1998/namespace" {
        return Err(malformed("the XML namespace may be bound only to 'xml'"));
    }
    Ok(())
}

fn validate_declaration(declaration: &BytesDecl<'_>) -> Result<()> {
    let version = declaration
        .xml_version()
        .map_err(|error| Error::Xml(error.to_string()))?;
    if version != XmlVersion::Explicit1_0 {
        return Err(malformed("Agile XML declaration must use version 1.0"));
    }
    if let Some(encoding) = declaration.encoding() {
        let encoding_name = encoding.map_err(|error| Error::Xml(error.to_string()))?;
        if !encoding_name.eq_ignore_ascii_case(b"UTF-8") {
            return Err(Error::Unsupported(format!(
                "Agile XML encoding '{}'",
                String::from_utf8_lossy(&encoding_name)
            )));
        }
    }
    if let Some(standalone) = declaration.standalone() {
        let standalone_value = standalone.map_err(|error| Error::Xml(error.to_string()))?;
        if !matches!(standalone_value.as_ref(), b"yes" | b"no") {
            return Err(malformed("Agile XML standalone must be 'yes' or 'no'"));
        }
    }
    Ok(())
}

fn exact_number(value: &str, expected: u32, field: &'static str) -> Result<()> {
    let actual = number(value, field)?;
    if actual != expected {
        return Err(Error::Unsupported(format!(
            "{field} is {actual}, expected {expected}"
        )));
    }
    Ok(())
}

fn number(value: &str, field: &'static str) -> Result<u32> {
    value
        .parse()
        .map_err(|_err| malformed(format!("{field} is not a u32")))
}

fn exact_text(value: &str, expected: &str, field: &'static str) -> Result<()> {
    if value != expected {
        return Err(Error::Unsupported(format!(
            "{field} is '{value}', expected '{expected}'"
        )));
    }
    Ok(())
}

fn decode_array(value: &str, field: &'static str, expected: usize) -> Result<Zeroizing<Vec<u8>>> {
    let max_encoded = expected
        .checked_add(2)
        .and_then(|length| length.checked_mul(4).map(|value| value / 3 + 4))
        .ok_or_else(|| malformed(format!("{field} length overflows usize")))?;
    if value.len() > max_encoded {
        return Err(malformed(format!(
            "{field} exceeds its bounded encoded size"
        )));
    }
    let decoded = BASE64
        .decode(value)
        .map_err(|error| malformed(format!("{field} is not valid base64: {error}")))?;
    if decoded.len() != expected {
        return Err(malformed(format!(
            "{field} has {} bytes, expected {expected}",
            decoded.len()
        )));
    }
    Ok(Zeroizing::new(decoded))
}

fn decode_fixed<const N: usize>(value: &str, field: &'static str) -> Result<[u8; N]> {
    let decoded = decode_array(value, field, N)?;
    decoded
        .as_slice()
        .try_into()
        .map_err(|_| malformed(format!("{field} has an invalid fixed length")))
}

fn count(value: &mut usize, maximum: usize, resource: &'static str) -> Result<()> {
    let next = value
        .checked_add(1)
        .ok_or_else(|| limit(resource, usize::MAX, maximum))?;
    if next > maximum {
        return Err(limit(resource, next, maximum));
    }
    *value = next;
    Ok(())
}

fn limit(resource: &'static str, actual: usize, maximum: usize) -> Error {
    Error::Limit {
        resource,
        actual: u64::try_from(actual).unwrap_or(u64::MAX),
        maximum: u64::try_from(maximum).unwrap_or(u64::MAX),
    }
}

#[cfg(test)]
mod tests {
    #![allow(
        clippy::expect_used,
        clippy::unwrap_used,
        reason = "test code panics on failure; assertions stay concise"
    )]
    use super::*;
    use crate::ooxml::{Mode, open_with};

    const MATERIAL: Material = Material {
        verifier_salt: [
            0x9a, 0x4c, 0x79, 0x4b, 0x45, 0x20, 0x8c, 0xf6, 0x2c, 0x8a, 0xf5, 0xcd, 0x3a, 0xb6,
            0x9c, 0xe4,
        ],
        verifier: *b"fixed verifier!!",
        key_salt: [
            0xfd, 0xae, 0x22, 0x5a, 0xa3, 0xf2, 0x22, 0xf1, 0x36, 0x71, 0x4a, 0x25, 0x24, 0xc2,
            0xab, 0x23,
        ],
        content_key: [0x43; 32],
        integrity_salt: [0x49; 64],
        integrity_len: 20,
    };

    #[test]
    fn agile_profiles_round_trip_with_authenticated_integrity() {
        let limits = Limits::default();
        let clear = Vec::from(&b"PK\x03\x04deterministic Agile package"[..]);
        for mode in [
            Mode::Agile,
            Mode::AgileSha256,
            Mode::AgileSha512,
            Mode::AgileAes192,
            Mode::AgileAes192Sha256,
            Mode::AgileAes192Sha512,
            Mode::AgileAes256,
            Mode::AgileAes256Sha256,
            Mode::AgileAes256Sha512,
        ] {
            let profile = Profile::from_mode(mode).expect("profile");
            let mut material = Zeroizing::new(Material {
                verifier_salt: MATERIAL.verifier_salt,
                verifier: MATERIAL.verifier,
                key_salt: MATERIAL.key_salt,
                content_key: MATERIAL.content_key,
                integrity_salt: MATERIAL.integrity_salt,
                integrity_len: profile.integrity_bytes(),
            });
            let (info, encrypted) =
                encrypt_parts(clear.clone(), "correct horse", profile, &material, &limits)
                    .expect("encrypt Agile parts");
            let compound = container::write(&info, encrypted, &limits).expect("wrap Agile parts");
            let opened = open_with(compound, "correct horse", &limits).expect("open Agile package");
            assert_eq!(opened.mode(), Some(mode));
            assert_eq!(opened.integrity(), Some(IntegrityStatus::Authenticated));
            assert_eq!(opened.bytes(), clear);
            material.zeroize();
        }
    }

    #[test]
    fn agile_sha1_accepts_literal_salt_size_integrity_key() {
        let limits = Limits::default();
        let profile = Profile::from_mode(Mode::Agile).expect("profile");
        let material = Zeroizing::new(Material {
            verifier_salt: MATERIAL.verifier_salt,
            verifier: MATERIAL.verifier,
            key_salt: MATERIAL.key_salt,
            content_key: MATERIAL.content_key,
            integrity_salt: MATERIAL.integrity_salt,
            integrity_len: SALT_BYTES,
        });
        let (info, encrypted) = encrypt_parts(
            b"PK\x03\x04literal salt-size integrity package".to_vec(),
            "correct horse",
            profile,
            &material,
            &limits,
        )
        .expect("encrypt literal salt-size form");
        let parsed = parse(&info, &limits).expect("parse literal salt-size form");
        assert_eq!(
            parsed
                .encrypted_hmac_key
                .as_ref()
                .expect("integrity key")
                .len(),
            BLOCK
        );
        let (opened, status) = decrypt_with_policy(
            &info,
            encrypted,
            "correct horse",
            &limits,
            IntegrityPolicy::RequireAuthenticated,
        )
        .expect("decrypt literal salt-size form");
        assert_eq!(status, IntegrityStatus::Authenticated);
        assert_eq!(opened, b"PK\x03\x04literal salt-size integrity package");
    }

    #[test]
    fn agile_reader_accepts_sha1_alias_but_authoring_stays_canonical() {
        let limits = Limits::default();
        let profile = Profile::from_mode(Mode::Agile).expect("profile");
        assert_eq!(profile.hash_name(), "SHA-1");
        let material = Zeroizing::new(Material {
            verifier_salt: MATERIAL.verifier_salt,
            verifier: MATERIAL.verifier,
            key_salt: MATERIAL.key_salt,
            content_key: MATERIAL.content_key,
            integrity_salt: MATERIAL.integrity_salt,
            integrity_len: profile.integrity_bytes(),
        });
        let clear = b"PK\x03\x04SHA1 XML alias";
        let (info, encrypted) =
            encrypt_parts(clear.to_vec(), "correct horse", profile, &material, &limits)
                .expect("encrypt canonical profile");
        let xml = std::str::from_utf8(&info[8..])
            .expect("UTF-8 Agile XML")
            .to_owned();
        assert!(xml.contains("hashAlgorithm=\"SHA-1\""));
        let xml = xml.replace("SHA-1", "SHA1");
        let mut alias_info = info[..8].to_vec();
        alias_info.extend_from_slice(xml.as_bytes());
        let compound =
            container::write(&alias_info, encrypted, &limits).expect("wrap SHA1 alias profile");
        let opened =
            open_with(compound, "correct horse", &limits).expect("read SHA1 alias profile");
        assert_eq!(opened.mode(), Some(Mode::Agile));
        assert_eq!(opened.integrity(), Some(IntegrityStatus::Authenticated));
        assert_eq!(opened.bytes(), clear);
    }

    #[test]
    fn missing_data_integrity_requires_explicit_read_opt_in() {
        let limits = Limits::default();
        let profile = Profile::from_mode(Mode::Agile).expect("profile");
        let material = Zeroizing::new(Material {
            verifier_salt: MATERIAL.verifier_salt,
            verifier: MATERIAL.verifier,
            key_salt: MATERIAL.key_salt,
            content_key: MATERIAL.content_key,
            integrity_salt: MATERIAL.integrity_salt,
            integrity_len: profile.integrity_bytes(),
        });
        let clear = b"PK\x03\x04unauthenticated agile profile";
        let (info, encrypted) =
            encrypt_parts(clear.to_vec(), "correct horse", profile, &material, &limits)
                .expect("encrypt profile");
        let xml = std::str::from_utf8(&info[8..]).expect("UTF-8 Agile XML");
        let start = xml.find("<dataIntegrity ").expect("dataIntegrity start");
        let end = xml[start..]
            .find("/>")
            .map(|offset| start + offset + 2)
            .expect("dataIntegrity end");
        let xml_bytes = xml.as_bytes();
        let mut without_integrity = info[..8].to_vec();
        without_integrity.extend_from_slice(&xml_bytes[..start]);
        without_integrity.extend_from_slice(&xml_bytes[end..]);
        let compound = container::write(&without_integrity, encrypted, &limits)
            .expect("wrap schema-legal descriptor");

        assert!(matches!(
            open_with(compound.clone(), "correct horse", &limits),
            Err(Error::Unsupported(message)) if message.contains("dataIntegrity")
        ));
        assert!(matches!(
            crate::ooxml::inspect_with(&compound, &limits),
            Err(Error::Unsupported(message)) if message.contains("dataIntegrity")
        ));
        assert_eq!(
            crate::ooxml::inspect_with_policy(
                &compound,
                &limits,
                IntegrityPolicy::AllowUnauthenticated,
            )
            .expect("explicit unauthenticated inspection policy"),
            crate::ooxml::Kind::Encrypted(Mode::Agile)
        );
        let opened = crate::ooxml::open_with_policy(
            compound,
            "correct horse",
            &limits,
            IntegrityPolicy::AllowUnauthenticated,
        )
        .expect("explicit unauthenticated read policy");
        assert_eq!(opened.mode(), Some(Mode::Agile));
        assert_eq!(opened.integrity(), Some(IntegrityStatus::Unauthenticated));
        assert_eq!(opened.bytes(), clear);
    }

    #[test]
    fn agile_xml_limits_reject_malformed_depth_nodes_and_spin() {
        let limits = Limits::default();
        let profile = Profile::from_mode(Mode::AgileSha256).expect("profile");
        let material = Zeroizing::new(Material {
            verifier_salt: MATERIAL.verifier_salt,
            verifier: MATERIAL.verifier,
            key_salt: MATERIAL.key_salt,
            content_key: MATERIAL.content_key,
            integrity_salt: MATERIAL.integrity_salt,
            integrity_len: profile.integrity_bytes(),
        });
        let (info, _) = encrypt_parts(
            b"PK\x03\x04XML limits".to_vec(),
            "correct horse",
            profile,
            &material,
            &limits,
        )
        .expect("encrypt profile");

        let shallow = Limits {
            max_xml_depth: 2,
            ..limits
        };
        assert!(matches!(
            parse(&info, &shallow),
            Err(Error::Limit {
                resource: "Agile XML depth",
                ..
            })
        ));

        let few_nodes = Limits {
            max_xml_nodes: 3,
            ..limits
        };
        assert!(matches!(
            parse(&info, &few_nodes),
            Err(Error::Limit {
                resource: "Agile XML nodes",
                ..
            })
        ));

        let few_spin = Limits {
            max_spin_count: 99_999,
            ..limits
        };
        assert!(matches!(
            parse(&info, &few_spin),
            Err(Error::Limit {
                resource: "Agile spin count",
                ..
            })
        ));

        let mut malformed = info;
        malformed.pop();
        assert!(matches!(
            parse(&malformed, &limits),
            Err(Error::Xml(_)) | Err(Error::Malformed(_))
        ));
    }

    #[test]
    fn parses_independent_agile_wrapper_and_data_key_sizes() {
        let limits = Limits::default();
        let encrypted = include_bytes!(
            "../../tests/data/ooxml/component-agile-aes128-wrap-aes256-data-sha512.docx"
        );
        let info = container::read_info(encrypted, &limits).expect("read fixture EncryptionInfo");
        let parsed = parse(&info, &limits).expect("parse fixture Agile XML");
        assert_eq!(parsed.profile.cipher, AgileCipher::Aes256);
        assert_eq!(parsed.profile.hash, AgileHash::Sha512);
        assert_eq!(parsed.wrap_profile.cipher, AgileCipher::Aes128);
        assert_eq!(parsed.wrap_profile.hash, AgileHash::Sha512);
    }

    #[test]
    fn agile_reader_supports_the_known_mixed_wrapper_hash_profile() {
        let limits = Limits::default();
        let data_profile = Profile::from_mode(Mode::Agile).expect("data profile");
        let wrapper_profile = Profile {
            cipher: AgileCipher::Aes256,
            hash: AgileHash::Sha512,
        };
        let material = Zeroizing::new(Material {
            verifier_salt: MATERIAL.verifier_salt,
            verifier: MATERIAL.verifier,
            key_salt: MATERIAL.key_salt,
            content_key: MATERIAL.content_key,
            integrity_salt: MATERIAL.integrity_salt,
            integrity_len: data_profile.integrity_bytes(),
        });
        let clear = b"PK\x03\x04mixed wrapper profile";
        let (info, encrypted) = encrypt_parts_with_wrapper(
            clear.to_vec(),
            "correct horse",
            data_profile,
            wrapper_profile,
            &material,
            &limits,
        )
        .expect("encrypt mixed wrapper profile");
        let parsed = parse(&info, &limits).expect("parse mixed wrapper profile");
        assert_eq!(parsed.profile, data_profile);
        assert_eq!(parsed.wrap_profile, wrapper_profile);
        let verifier_hash = parsed.encrypted_verifier_hash.as_slice();
        assert_eq!(verifier_hash.len(), wrapper_profile.encrypted_hash_bytes());
        let compound = container::write(&info, encrypted, &limits).expect("wrap mixed profile");
        let opened =
            open_with(compound, "correct horse", &limits).expect("open mixed wrapper profile");
        assert_eq!(opened.mode(), Some(Mode::Agile));
        assert_eq!(opened.integrity(), Some(IntegrityStatus::Authenticated));
        assert_eq!(opened.bytes(), clear);
    }

    #[test]
    fn mixed_wrapper_hash_compatibility_stays_narrow() {
        let data_profile = Profile::from_mode(Mode::Agile).expect("data profile");
        assert!(is_supported_wrapper_profile(
            data_profile,
            Profile {
                cipher: AgileCipher::Aes256,
                hash: AgileHash::Sha512,
            }
        ));
        assert!(!is_supported_wrapper_profile(
            data_profile,
            Profile {
                cipher: AgileCipher::Aes256,
                hash: AgileHash::Sha256,
            }
        ));
    }

    #[test]
    fn agile_wrong_password_and_tampering_are_typed() {
        let limits = Limits::default();
        let profile = Profile::from_mode(Mode::AgileSha256).expect("profile");
        let material = Zeroizing::new(Material {
            verifier_salt: MATERIAL.verifier_salt,
            verifier: MATERIAL.verifier,
            key_salt: MATERIAL.key_salt,
            content_key: MATERIAL.content_key,
            integrity_salt: MATERIAL.integrity_salt,
            integrity_len: profile.integrity_bytes(),
        });
        let (info, mut encrypted) = encrypt_parts(
            b"PK\x03\x04package".to_vec(),
            "correct horse",
            profile,
            &material,
            &limits,
        )
        .expect("encrypt");
        assert!(matches!(
            decrypt_with_policy(
                &info,
                encrypted.clone(),
                "wrong",
                &limits,
                IntegrityPolicy::RequireAuthenticated,
            ),
            Err(Error::Password)
        ));
        let last = encrypted.len() - 1;
        encrypted[last] ^= 1;
        assert!(matches!(
            decrypt_with_policy(
                &info,
                encrypted,
                "correct horse",
                &limits,
                IntegrityPolicy::RequireAuthenticated,
            ),
            Err(Error::Integrity)
        ));
    }
}
