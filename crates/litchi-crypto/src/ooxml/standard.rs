//! `[MS-OFFCRYPTO]` Standard Encryption profiles.
//!
//! Standard Encryption always uses SHA-1 for password derivation and AES-128,
//! AES-192, or AES-256 for the package cipher. Legacy RC4 containers are
//! intentionally outside this OOXML profile owner; callers receive an
//! unsupported-profile error rather than silently accepting a weak downgrade.

use aes::cipher::{Block, BlockCipherDecrypt, BlockCipherEncrypt, KeyInit};
use aes::{Aes128, Aes192, Aes256};
use rand::TryRng;
use rand::rngs::SysRng;
use sha1::{Digest, Sha1};
use subtle::ConstantTimeEq;
use zeroize::Zeroizing;

use super::{Error, Limits, Mode, Result, container, declared_size, malformed, password_bytes};

const BLOCK: usize = 16;
const BLOCK_U32: u32 = 16;
const SPIN_COUNT: u32 = 50_000;
const FLAGS_AES: u32 = 0x24;
const CRYPTO_API: u32 = 0x04;
const DOC_PROPERTIES: u32 = 0x08;
const EXTERNAL: u32 = 0x10;
const AES: u32 = 0x20;
const ALG_AES_128: u32 = 0x660e;
const ALG_AES_192: u32 = 0x660f;
const ALG_AES_256: u32 = 0x6610;
const ALG_SHA1: u32 = 0x8004;
const KEY_BITS_AES_128: u32 = 128;
const KEY_BITS_AES_192: u32 = 192;
const KEY_BITS_AES_256: u32 = 256;
const PROVIDER_AES: u32 = 0x18;
const PROVIDER_AES_NAME: &str = "Microsoft Enhanced RSA and AES Cryptographic Provider";

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Cipher {
    Aes128,
    Aes192,
    Aes256,
}

impl Cipher {
    const fn key_bits(self) -> u32 {
        match self {
            Self::Aes128 => KEY_BITS_AES_128,
            Self::Aes192 => KEY_BITS_AES_192,
            Self::Aes256 => KEY_BITS_AES_256,
        }
    }

    const fn alg_id(self) -> u32 {
        match self {
            Self::Aes128 => ALG_AES_128,
            Self::Aes192 => ALG_AES_192,
            Self::Aes256 => ALG_AES_256,
        }
    }

    const fn mode(self) -> Mode {
        match self {
            Self::Aes128 => Mode::Standard,
            Self::Aes192 => Mode::StandardAes192,
            Self::Aes256 => Mode::StandardAes256,
        }
    }
}

#[derive(Debug, Clone, Copy)]
struct Verifier {
    cipher: Cipher,
    salt: [u8; BLOCK],
    encrypted: [u8; BLOCK],
    hash: [u8; BLOCK * 2],
    hash_len: usize,
}

#[derive(Clone, Copy)]
enum Direction {
    Encrypt,
    Decrypt,
}

pub(super) fn profile(info: &[u8], limits: &Limits) -> Result<Mode> {
    Limits::bytes("EncryptionInfo", info.len(), limits.max_info_bytes)?;
    Ok(parse(info)?.cipher.mode())
}

pub(super) fn encrypt(
    package: Vec<u8>,
    password: &str,
    mode: Mode,
    limits: &Limits,
) -> Result<Vec<u8>> {
    let cipher = cipher_for_mode(mode)?;
    let mut rng = SysRng;
    let mut salt = Zeroizing::new([0u8; BLOCK]);
    let mut verifier = Zeroizing::new([0u8; BLOCK]);
    rng.try_fill_bytes(salt.as_mut())
        .map_err(|error| Error::Random(error.to_string()))?;
    rng.try_fill_bytes(verifier.as_mut())
        .map_err(|error| Error::Random(error.to_string()))?;
    encrypt_with(package, password, cipher, &salt, &verifier, limits)
}

fn cipher_for_mode(mode: Mode) -> Result<Cipher> {
    match mode {
        Mode::Standard => Ok(Cipher::Aes128),
        Mode::StandardAes192 => Ok(Cipher::Aes192),
        Mode::StandardAes256 => Ok(Cipher::Aes256),
        _ => Err(Error::Unsupported(format!(
            "{mode} is not a Standard Encryption profile"
        ))),
    }
}

fn encrypt_with(
    package: Vec<u8>,
    password: &str,
    cipher: Cipher,
    salt: &[u8; BLOCK],
    verifier: &[u8; BLOCK],
    limits: &Limits,
) -> Result<Vec<u8>> {
    let key = key(password, salt, cipher.key_bits(), limits)?;
    let (encrypted_verifier, encrypted_hash, hash_len) = encrypt_verifier(&key, cipher, verifier)?;
    let info = build_info(cipher, salt, &encrypted_verifier, &encrypted_hash, hash_len)?;
    Limits::bytes("EncryptionInfo", info.len(), limits.max_info_bytes)?;
    let encrypted = encrypt_package(package, &key, cipher, limits)?;
    container::write(&info, encrypted, limits)
}

pub(super) fn decrypt(
    info: &[u8],
    encrypted: Vec<u8>,
    password: &str,
    limits: &Limits,
) -> Result<Vec<u8>> {
    Limits::bytes("EncryptionInfo", info.len(), limits.max_info_bytes)?;
    Limits::bytes(
        "EncryptedPackage",
        encrypted.len(),
        limits.max_encrypted_bytes,
    )?;
    let verifier = parse(info)?;
    let key = key(password, &verifier.salt, verifier.cipher.key_bits(), limits)?;
    verify(&key, &verifier)?;
    decrypt_package(encrypted, &key, verifier.cipher, limits)
}

fn key(
    password: &str,
    salt: &[u8; BLOCK],
    key_bits: u32,
    limits: &Limits,
) -> Result<Zeroizing<Vec<u8>>> {
    if !matches!(
        key_bits,
        KEY_BITS_AES_128 | KEY_BITS_AES_192 | KEY_BITS_AES_256
    ) {
        return Err(Error::Unsupported(format!(
            "Standard key size {key_bits} bits"
        )));
    }
    let encoded = password_bytes(password, limits)?;
    let mut hasher = Sha1::new();
    hasher.update(salt);
    hasher.update(encoded.as_slice());
    let mut hash = Zeroizing::new(<[u8; 20]>::from(hasher.finalize()));
    for iterator in 0..SPIN_COUNT {
        let mut spin = Sha1::new();
        spin.update(iterator.to_le_bytes());
        spin.update(hash.as_slice());
        let digest = spin.finalize();
        hash.copy_from_slice(&digest);
    }
    let mut finalizer = Sha1::new();
    finalizer.update(hash.as_slice());
    finalizer.update([0u8; 4]);
    let final_hash = Zeroizing::new(<[u8; 20]>::from(finalizer.finalize()));
    let x1 = digest_xor(final_hash.as_slice(), 0x36);
    let x2 = digest_xor(final_hash.as_slice(), 0x5c);
    let required = usize::try_from(key_bits / 8)
        .map_err(|_err| malformed("Standard key size does not fit usize"))?;
    let mut output = Zeroizing::new(Vec::new());
    output
        .try_reserve_exact(required)
        .map_err(|_err| Error::Allocation("Standard derived key"))?;
    output.extend(x1.as_slice().iter().chain(x2.as_slice()).take(required));
    Ok(output)
}

fn digest_xor(input: &[u8], fill: u8) -> Zeroizing<[u8; 20]> {
    let mut buffer = Zeroizing::new([fill; 64]);
    for (destination, source) in buffer.iter_mut().zip(input) {
        *destination ^= source;
    }
    let mut sha = Sha1::new();
    sha.update(buffer.as_slice());
    Zeroizing::new(<[u8; 20]>::from(sha.finalize()))
}

fn encrypt_verifier(
    key: &[u8],
    cipher: Cipher,
    verifier: &[u8; BLOCK],
) -> Result<([u8; BLOCK], [u8; BLOCK * 2], usize)> {
    let mut encrypted = *verifier;
    let mut hash = [0u8; BLOCK * 2];
    let digest = Zeroizing::new(<[u8; 20]>::from(Sha1::digest(verifier)));
    aes_crypt(key, cipher, &mut encrypted, Direction::Encrypt)?;
    hash[..digest.len()].copy_from_slice(digest.as_slice());
    let hash_len = BLOCK * 2;
    aes_crypt(key, cipher, &mut hash[..hash_len], Direction::Encrypt)?;
    Ok((encrypted, hash, hash_len))
}

fn build_info(
    cipher: Cipher,
    salt: &[u8; BLOCK],
    verifier: &[u8; BLOCK],
    hash: &[u8; BLOCK * 2],
    hash_len: usize,
) -> Result<Vec<u8>> {
    let flags = FLAGS_AES;
    let provider_type = PROVIDER_AES;
    let provider = PROVIDER_AES_NAME;
    let mut output = Vec::new();
    output
        .try_reserve_exact(256)
        .map_err(|_err| Error::Allocation("Standard EncryptionInfo"))?;
    output.extend_from_slice(&3u16.to_le_bytes());
    output.extend_from_slice(&2u16.to_le_bytes());
    output.extend_from_slice(&flags.to_le_bytes());
    let size_offset = output.len();
    output.extend_from_slice(&0u32.to_le_bytes());
    output.extend_from_slice(&flags.to_le_bytes());
    output.extend_from_slice(&0u32.to_le_bytes());
    output.extend_from_slice(&cipher.alg_id().to_le_bytes());
    output.extend_from_slice(&ALG_SHA1.to_le_bytes());
    output.extend_from_slice(&cipher.key_bits().to_le_bytes());
    output.extend_from_slice(&provider_type.to_le_bytes());
    output.extend_from_slice(&0u32.to_le_bytes());
    output.extend_from_slice(&0u32.to_le_bytes());
    for unit in provider.encode_utf16() {
        output.extend_from_slice(&unit.to_le_bytes());
    }
    output.extend_from_slice(&0u16.to_le_bytes());
    let header_size = output
        .len()
        .checked_sub(size_offset + 4)
        .and_then(|size| u32::try_from(size).ok())
        .ok_or_else(|| malformed("Standard EncryptionHeader size overflows u32"))?;
    output
        .get_mut(size_offset..size_offset + 4)
        .ok_or_else(|| malformed("Standard EncryptionHeader size field is unavailable"))?
        .copy_from_slice(&header_size.to_le_bytes());
    output.extend_from_slice(&BLOCK_U32.to_le_bytes());
    output.extend_from_slice(salt);
    output.extend_from_slice(verifier);
    output.extend_from_slice(&20u32.to_le_bytes());
    output.extend_from_slice(&hash[..hash_len]);
    Ok(output)
}

fn parse(info: &[u8]) -> Result<Verifier> {
    if info.len() < 12 {
        return Err(malformed(
            "Standard EncryptionInfo is shorter than its header",
        ));
    }
    let major = u16::from_le_bytes([info[0], info[1]]);
    let minor = u16::from_le_bytes([info[2], info[3]]);
    if !(2..=4).contains(&major) || minor != 2 {
        return Err(Error::Unsupported(format!(
            "Standard EncryptionInfo version {major}.{minor}"
        )));
    }
    let outer_flags = read_u32(info, 4, "Standard outer flags")?;
    validate_flags(outer_flags)?;
    let header_size = usize::try_from(read_u32(info, 8, "EncryptionHeaderSize")?)
        .map_err(|_err| malformed("EncryptionHeaderSize does not fit usize"))?;
    let header_end = 12usize
        .checked_add(header_size)
        .ok_or_else(|| malformed("EncryptionHeader size overflows usize"))?;
    if header_size < 34 || header_end > info.len() {
        return Err(malformed("Standard EncryptionHeader has an invalid size"));
    }
    let header = &info[12..header_end];
    let inner_flags = read_u32(header, 0, "EncryptionHeader.Flags")?;
    validate_flags(inner_flags)?;
    if inner_flags != outer_flags {
        return Err(malformed(
            "Standard outer flags are not a copy of EncryptionHeader.Flags",
        ));
    }
    require_u32(header, 4, 0, "EncryptionHeader.SizeExtra")?;
    let algorithm = read_u32(header, 8, "EncryptionHeader.AlgID")?;
    let hash_algorithm = read_u32(header, 12, "EncryptionHeader.AlgIDHash")?;
    let key_bits = read_u32(header, 16, "EncryptionHeader.KeySize")?;
    let cipher = validate_profile(outer_flags, algorithm, hash_algorithm, key_bits)?;
    let _provider = read_u32(header, 20, "EncryptionHeader.ProviderType")?;
    let _reserved1 = read_u32(header, 24, "EncryptionHeader.Reserved1")?;
    require_u32(header, 28, 0, "EncryptionHeader.Reserved2")?;
    validate_provider_name(&header[32..])?;

    let hash_len = BLOCK * 2;
    let verifier_len = 4 + BLOCK + BLOCK + 4 + hash_len;
    let verifier_end = header_end
        .checked_add(verifier_len)
        .ok_or_else(|| malformed("Standard verifier size overflows usize"))?;
    if verifier_end != info.len() {
        return Err(malformed(
            "Standard EncryptionInfo verifier is truncated or has trailing bytes",
        ));
    }
    let verifier = &info[header_end..verifier_end];
    require_u32(verifier, 0, BLOCK_U32, "EncryptionVerifier.SaltSize")?;
    let salt = array::<BLOCK>(verifier, 4, "EncryptionVerifier.Salt")?;
    let encrypted = array::<BLOCK>(verifier, 4 + BLOCK, "EncryptedVerifier")?;
    require_u32(
        verifier,
        4 + BLOCK + BLOCK,
        20,
        "EncryptionVerifier.VerifierHashSize",
    )?;
    let mut hash = [0u8; BLOCK * 2];
    hash[..hash_len].copy_from_slice(array_slice(verifier, 4 + BLOCK + BLOCK + 4, hash_len)?);
    Ok(Verifier {
        cipher,
        salt,
        encrypted,
        hash,
        hash_len,
    })
}

fn validate_provider_name(bytes: &[u8]) -> Result<()> {
    if bytes.len() < 2 || !bytes.len().is_multiple_of(2) {
        return Err(malformed(
            "Standard CSPName is not a terminated UTF-16LE string",
        ));
    }
    let body_end = bytes.len() - 2;
    if bytes.get(body_end..) != Some(&[0, 0][..])
        || bytes[..body_end].as_chunks::<2>().0.contains(&[0, 0])
    {
        return Err(malformed(
            "Standard CSPName terminator is missing or not final",
        ));
    }
    let units = bytes[..body_end]
        .as_chunks::<2>()
        .0
        .iter()
        .map(|pair| u16::from_le_bytes([pair[0], pair[1]]));
    if char::decode_utf16(units).any(|character| character.is_err()) {
        return Err(malformed("Standard CSPName contains invalid UTF-16"));
    }
    Ok(())
}

fn verify(key: &[u8], verifier: &Verifier) -> Result<()> {
    let mut clear = Zeroizing::new(verifier.encrypted);
    let mut stored = Zeroizing::new(verifier.hash);
    aes_crypt(key, verifier.cipher, clear.as_mut(), Direction::Decrypt)?;
    aes_crypt(
        key,
        verifier.cipher,
        &mut stored[..verifier.hash_len],
        Direction::Decrypt,
    )?;
    let expected = Zeroizing::new(<[u8; 20]>::from(Sha1::digest(clear.as_slice())));
    if !bool::from(stored[..20].ct_eq(expected.as_slice())) {
        return Err(Error::Password);
    }
    Ok(())
}

fn encrypt_package(
    mut package: Vec<u8>,
    key: &[u8],
    cipher: Cipher,
    limits: &Limits,
) -> Result<Vec<u8>> {
    let clear_len = package.len();
    let cipher_len = round_up(clear_len, BLOCK)?;
    let total = cipher_len
        .checked_add(8)
        .ok_or_else(|| malformed("Standard EncryptedPackage size overflows usize"))?;
    Limits::bytes("EncryptedPackage", total, limits.max_encrypted_bytes)?;
    package
        .try_reserve_exact(total.saturating_sub(package.len()))
        .map_err(|_err| Error::Allocation("Standard EncryptedPackage"))?;
    package.resize(total, 0);
    package.copy_within(0..clear_len, 8);
    package
        .get_mut(..8)
        .ok_or_else(|| malformed("Standard EncryptedPackage prefix is unavailable"))?
        .copy_from_slice(
            &u64::try_from(clear_len)
                .map_err(|_err| malformed("plaintext size does not fit u64"))?
                .to_le_bytes(),
        );
    let ciphertext = package
        .get_mut(8..)
        .ok_or_else(|| malformed("Standard EncryptedPackage ciphertext is unavailable"))?;
    aes_crypt(key, cipher, ciphertext, Direction::Encrypt)?;
    Ok(package)
}

fn decrypt_package(
    mut encrypted: Vec<u8>,
    key: &[u8],
    cipher: Cipher,
    limits: &Limits,
) -> Result<Vec<u8>> {
    let minimum = 8 + BLOCK;
    if encrypted.len() < minimum {
        return Err(malformed("Standard EncryptedPackage is too short"));
    }
    let declared = u64::from_le_bytes(array::<8>(&encrypted, 0, "StreamSize")?);
    let clear_len = declared_size(declared, limits)?;
    if clear_len == 0 {
        return Err(malformed(
            "Standard EncryptedPackage declares an empty package",
        ));
    }
    let cipher_len = round_up(clear_len, BLOCK)?;
    if encrypted.len() != cipher_len + 8 {
        return Err(malformed(
            "Standard EncryptedPackage length disagrees with StreamSize",
        ));
    }
    let ciphertext = encrypted
        .get_mut(8..)
        .ok_or_else(|| malformed("Standard EncryptedPackage ciphertext is unavailable"))?;
    aes_crypt(key, cipher, ciphertext, Direction::Decrypt)?;
    let source_end = clear_len
        .checked_add(8)
        .ok_or_else(|| malformed("Standard decrypted package range overflows usize"))?;
    if encrypted.get(8..source_end).is_none() {
        return Err(malformed("Standard decrypted package is truncated"));
    }
    encrypted.copy_within(8..source_end, 0);
    encrypted.truncate(clear_len);
    Ok(encrypted)
}

fn aes_crypt(key: &[u8], cipher: Cipher, bytes: &mut [u8], direction: Direction) -> Result<()> {
    if !bytes.len().is_multiple_of(BLOCK) {
        return Err(malformed("AES data is not aligned to a 16-byte block"));
    }
    macro_rules! run {
        ($ty:ty) => {{
            let cipher = <$ty as KeyInit>::new_from_slice(key)
                .map_err(|_err| malformed("AES key length invariant was violated"))?;
            for chunk in bytes.as_chunks_mut::<BLOCK>().0 {
                let block: &mut Block<$ty> = (&mut chunk[..])
                    .try_into()
                    .map_err(|_err| malformed("AES block conversion failed"))?;
                match direction {
                    Direction::Encrypt => cipher.encrypt_block(block),
                    Direction::Decrypt => cipher.decrypt_block(block),
                }
            }
            Ok(())
        }};
    }
    match cipher {
        Cipher::Aes128 => run!(Aes128),
        Cipher::Aes192 => run!(Aes192),
        Cipher::Aes256 => run!(Aes256),
    }
}

fn round_up(value: usize, multiple: usize) -> Result<usize> {
    value
        .checked_add(multiple - 1)
        .map(|padded| padded / multiple * multiple)
        .ok_or_else(|| malformed("encrypted block length overflows usize"))
}

fn read_u32(bytes: &[u8], offset: usize, field: &'static str) -> Result<u32> {
    Ok(u32::from_le_bytes(array::<4>(bytes, offset, field)?))
}

fn validate_flags(flags: u32) -> Result<()> {
    if flags & CRYPTO_API == 0 || flags & (DOC_PROPERTIES | EXTERNAL) != 0 {
        return Err(malformed(format!(
            "Standard EncryptionHeader flags {flags:#010x} violate the Standard profile"
        )));
    }
    Ok(())
}

fn validate_profile(
    flags: u32,
    algorithm: u32,
    hash_algorithm: u32,
    key_bits: u32,
) -> Result<Cipher> {
    if hash_algorithm != ALG_SHA1 {
        return Err(malformed(format!(
            "EncryptionHeader.AlgIDHash is {hash_algorithm:#010x}, expected SHA-1"
        )));
    }
    if flags & AES == 0 {
        return Err(Error::Unsupported(
            "Standard RC4/legacy non-AES encryption is not an OOXML AES profile".into(),
        ));
    }
    let cipher = match algorithm {
        ALG_AES_128 => Cipher::Aes128,
        ALG_AES_192 => Cipher::Aes192,
        ALG_AES_256 => Cipher::Aes256,
        _ => {
            return Err(malformed(format!(
                "EncryptionHeader.AlgID {algorithm:#010x} is not a Standard AES algorithm"
            )));
        },
    };
    if key_bits != cipher.key_bits() {
        return Err(malformed(format!(
            "EncryptionHeader.KeySize {key_bits} contradicts AlgID {algorithm:#x}"
        )));
    }
    Ok(cipher)
}

fn require_u32(bytes: &[u8], offset: usize, expected: u32, field: &'static str) -> Result<()> {
    let actual = read_u32(bytes, offset, field)?;
    if actual != expected {
        return Err(malformed(format!(
            "{field} is {actual:#010x}, expected {expected:#010x}"
        )));
    }
    Ok(())
}

fn array<const N: usize>(bytes: &[u8], offset: usize, field: &'static str) -> Result<[u8; N]> {
    let end = offset
        .checked_add(N)
        .ok_or_else(|| malformed(format!("{field} offset overflows usize")))?;
    bytes
        .get(offset..end)
        .ok_or_else(|| malformed(format!("{field} is truncated")))?
        .try_into()
        .map_err(|_err| malformed(format!("{field} has the wrong length")))
}

fn array_slice(bytes: &[u8], offset: usize, length: usize) -> Result<&[u8]> {
    let end = offset
        .checked_add(length)
        .ok_or_else(|| malformed("array slice offset overflows usize"))?;
    bytes
        .get(offset..end)
        .ok_or_else(|| malformed("array slice is truncated"))
}

#[cfg(test)]
mod tests {
    #![allow(
        clippy::expect_used,
        reason = "test code panics on failure; expect keeps assertions concise"
    )]
    use super::*;
    use crate::ooxml::{Kind, inspect, open_with};

    const SALT: [u8; BLOCK] = [
        0x92, 0x25, 0x50, 0xf6, 0xb6, 0x4f, 0xfe, 0x5b, 0xd3, 0x96, 0xdf, 0x5e, 0xe9, 0x17, 0xda,
        0x3a,
    ];
    const VERIFIER: [u8; BLOCK] = *b"fixed verifier!!";

    #[test]
    fn standard_aes_key_sizes_round_trip() {
        let limits = Limits::default();
        let clear = Vec::from(&b"PK\x03\x04deterministic Standard package"[..]);
        for mode in [Mode::Standard, Mode::StandardAes192, Mode::StandardAes256] {
            let encrypted = encrypt(clear.clone(), "correct horse", mode, &limits)
                .expect("encrypt Standard package");
            assert_eq!(
                inspect(&encrypted).expect("classify package"),
                Kind::Encrypted(mode)
            );
            let opened = open_with(encrypted, "correct horse", &limits).expect("decrypt package");
            assert_eq!(opened.mode(), Some(mode));
            assert_eq!(
                opened.integrity(),
                Some(crate::ooxml::IntegrityStatus::Unauthenticated)
            );
            assert_eq!(opened.bytes(), clear);
        }
    }

    #[test]
    fn standard_wrong_password_is_typed() {
        let limits = Limits::default();
        let encrypted = encrypt(
            Vec::from(&b"PK\x03\x04deterministic package"[..]),
            "correct horse",
            Mode::StandardAes256,
            &limits,
        )
        .expect("encrypt Standard package");
        assert!(matches!(
            open_with(encrypted, "wrong", &limits),
            Err(Error::Password)
        ));
    }

    #[test]
    fn parses_the_published_ms_offcrypto_3_8_header_vector() {
        let info = hex("03 00 02 00 24 00 00 00 A4 00 00 00 24 00 00 00 \
             00 00 00 00 0E 66 00 00 04 80 00 00 80 00 00 00 \
             18 00 00 00 E0 BC 3B 07 00 00 00 00 4D 00 69 00 \
             63 00 72 00 6F 00 73 00 6F 00 66 00 74 00 20 00 \
             45 00 6E 00 68 00 61 00 6E 00 63 00 65 00 64 00 \
             20 00 52 00 53 00 41 00 20 00 61 00 6E 00 64 00 \
             20 00 41 00 45 00 53 00 20 00 43 00 72 00 79 00 \
             70 00 74 00 6F 00 67 00 72 00 61 00 70 00 68 00 \
             69 00 63 00 20 00 50 00 72 00 6F 00 76 00 69 00 \
             64 00 65 00 72 00 20 00 28 00 50 00 72 00 6F 00 \
             74 00 6F 00 74 00 79 00 70 00 65 00 29 00 00 00 \
             10 00 00 00 92 25 50 F6 B6 4F FE 5B D3 96 DF 5E \
             E9 17 DA 3A BF 86 E1 8F 64 9D 17 D0 A5 41 D9 45 \
             CE FD 96 0C 14 00 00 00 12 FF DC 88 A1 BD 26 23 \
             59 32 27 1F 73 0B 8F 79 4E 45 DA B3 AB 08 04 F4 \
             0B B9 50 46 D3 91 41 84");
        let parsed = parse(&info).expect("published Standard vector");
        assert_eq!(parsed.salt, SALT);
        assert_eq!(parsed.cipher, Cipher::Aes128);
        assert_eq!(
            parsed.encrypted,
            [
                0xbf, 0x86, 0xe1, 0x8f, 0x64, 0x9d, 0x17, 0xd0, 0xa5, 0x41, 0xd9, 0x45, 0xce, 0xfd,
                0x96, 0x0c,
            ]
        );
    }

    #[test]
    fn valid_aes_profiles_are_admitted_to_the_parser() {
        let key = Zeroizing::new(vec![0u8; 16]);
        let (encrypted, hash, hash_len) =
            encrypt_verifier(&key, Cipher::Aes128, &VERIFIER).expect("verifier");
        for (cipher, mode, algorithm, key_bits) in [
            (Cipher::Aes192, Mode::StandardAes192, ALG_AES_192, 192u32),
            (Cipher::Aes256, Mode::StandardAes256, ALG_AES_256, 256u32),
        ] {
            let mut info = build_info(cipher, &SALT, &encrypted, &hash, hash_len).expect("info");
            info[20..24].copy_from_slice(&algorithm.to_le_bytes());
            info[28..32].copy_from_slice(&key_bits.to_le_bytes());
            assert_eq!(profile(&info, &Limits::default()).expect("profile"), mode);
        }
    }

    fn hex(value: &str) -> Vec<u8> {
        value
            .split_ascii_whitespace()
            .map(|byte| u8::from_str_radix(byte, 16).expect("test hex"))
            .collect()
    }
}
