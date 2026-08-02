//! `[MS-OFFCRYPTO]` Standard Encryption (AES-128/SHA-1 profile).

use aes::Aes128;
use aes::cipher::{Block, BlockCipherDecrypt, BlockCipherEncrypt, KeyInit};
use rand::TryRng;
use rand::rngs::SysRng;
use sha1::{Digest, Sha1};
use subtle::ConstantTimeEq;
use zeroize::Zeroizing;

use super::{Error, Limits, Result, container, declared_size, malformed, password_bytes};

const BLOCK: usize = 16;
const BLOCK_U32: u32 = 16;
const SPIN_COUNT: u32 = 50_000;
const FLAGS: u32 = 0x24;
const REQUIRED_FLAGS: u32 = 0x24;
const FORBIDDEN_FLAGS: u32 = 0x18;
const ALG_AES_128: u32 = 0x660e;
const ALG_SHA1: u32 = 0x8004;
const KEY_BITS: u32 = 128;
const PROVIDER_AES: u32 = 0x18;
const PROVIDER: &str = "Microsoft Enhanced RSA and AES Cryptographic Provider";

#[derive(Debug, Clone, Copy)]
struct Verifier {
    salt: [u8; BLOCK],
    encrypted: [u8; BLOCK],
    hash: [u8; 32],
}

pub(super) fn encrypt(package: Vec<u8>, password: &str, limits: &Limits) -> Result<Vec<u8>> {
    let mut rng = SysRng;
    let mut salt = Zeroizing::new([0u8; BLOCK]);
    let mut verifier = Zeroizing::new([0u8; BLOCK]);
    rng.try_fill_bytes(salt.as_mut())
        .map_err(|error| Error::Random(error.to_string()))?;
    rng.try_fill_bytes(verifier.as_mut())
        .map_err(|error| Error::Random(error.to_string()))?;
    encrypt_with(package, password, &salt, &verifier, limits)
}

fn encrypt_with(
    package: Vec<u8>,
    password: &str,
    salt: &[u8; BLOCK],
    verifier: &[u8; BLOCK],
    limits: &Limits,
) -> Result<Vec<u8>> {
    let key = key(password, salt, limits)?;
    let (encrypted, hash) = encrypt_verifier(&key, verifier)?;
    let info = build_info(salt, &encrypted, &hash)?;
    limits.bytes("EncryptionInfo", info.len(), limits.max_info_bytes)?;
    let encrypted = encrypt_package(package, &key, limits)?;
    container::write(&info, encrypted, limits)
}

pub(super) fn decrypt(
    info: &[u8],
    encrypted: Vec<u8>,
    password: &str,
    limits: &Limits,
) -> Result<Vec<u8>> {
    limits.bytes("EncryptionInfo", info.len(), limits.max_info_bytes)?;
    limits.bytes(
        "EncryptedPackage",
        encrypted.len(),
        limits.max_encrypted_bytes,
    )?;
    let verifier = parse(info)?;
    let key = key(password, &verifier.salt, limits)?;
    verify(&key, &verifier)?;
    decrypt_package(encrypted, &key, limits)
}

pub(super) fn validate_info(info: &[u8], limits: &Limits) -> Result<()> {
    limits.bytes("EncryptionInfo", info.len(), limits.max_info_bytes)?;
    parse(info).map(|_| ())
}

fn key(password: &str, salt: &[u8; BLOCK], limits: &Limits) -> Result<Zeroizing<[u8; BLOCK]>> {
    let password = password_bytes(password, limits)?;
    let mut sha = Sha1::new();
    sha.update(salt);
    sha.update(password.as_slice());
    let mut hash = Zeroizing::new(<[u8; 20]>::from(sha.finalize()));

    for iterator in 0..SPIN_COUNT {
        let mut sha = Sha1::new();
        sha.update(iterator.to_le_bytes());
        sha.update(hash.as_slice());
        hash = Zeroizing::new(<[u8; 20]>::from(sha.finalize()));
    }

    let mut sha = Sha1::new();
    sha.update(hash.as_slice());
    sha.update([0u8; 4]);
    let intermediate = Zeroizing::new(<[u8; 20]>::from(sha.finalize()));
    let x1 = digest_xor(intermediate.as_slice(), 0x36);
    let x2 = digest_xor(intermediate.as_slice(), 0x5c);
    let mut output = Zeroizing::new([0u8; BLOCK]);
    output.copy_from_slice(&x1[..BLOCK.min(x1.len())]);
    // AES-128 requires only the first 16 bytes of X1; retaining this explicit
    // branch makes the X1 || X2 rule visible without allocating X3.
    let _ = x2;
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

fn encrypt_verifier(key: &[u8; BLOCK], verifier: &[u8; BLOCK]) -> Result<([u8; BLOCK], [u8; 32])> {
    let cipher = cipher(key)?;
    let mut encrypted = *verifier;
    cipher.encrypt_block((&mut encrypted).into());

    let mut sha = Sha1::new();
    sha.update(verifier);
    let hash = Zeroizing::new(<[u8; 20]>::from(sha.finalize()));
    let mut padded = Zeroizing::new([0u8; 32]);
    padded[..hash.len()].copy_from_slice(hash.as_slice());
    crypt_blocks(&cipher, padded.as_mut(), Direction::Encrypt)?;
    Ok((encrypted, *padded))
}

fn build_info(salt: &[u8; BLOCK], verifier: &[u8; BLOCK], hash: &[u8; 32]) -> Result<Vec<u8>> {
    let mut output = Vec::new();
    output
        .try_reserve_exact(256)
        .map_err(|_| Error::Allocation("Standard EncryptionInfo"))?;
    output.extend_from_slice(&3u16.to_le_bytes());
    output.extend_from_slice(&2u16.to_le_bytes());
    output.extend_from_slice(&FLAGS.to_le_bytes());

    let size_offset = output.len();
    output.extend_from_slice(&0u32.to_le_bytes());
    output.extend_from_slice(&FLAGS.to_le_bytes());
    output.extend_from_slice(&0u32.to_le_bytes());
    output.extend_from_slice(&ALG_AES_128.to_le_bytes());
    output.extend_from_slice(&ALG_SHA1.to_le_bytes());
    output.extend_from_slice(&KEY_BITS.to_le_bytes());
    output.extend_from_slice(&PROVIDER_AES.to_le_bytes());
    output.extend_from_slice(&0u32.to_le_bytes());
    output.extend_from_slice(&0u32.to_le_bytes());
    for unit in PROVIDER.encode_utf16() {
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
    output.extend_from_slice(hash);
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
        .map_err(|_| malformed("EncryptionHeaderSize does not fit usize"))?;
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
    require_u32(header, 8, ALG_AES_128, "EncryptionHeader.AlgID")?;
    require_u32(header, 12, ALG_SHA1, "EncryptionHeader.AlgIDHash")?;
    require_u32(header, 16, KEY_BITS, "EncryptionHeader.KeySize")?;
    // ProviderType is a SHOULD in the specification and is not security
    // authoritative, so any value remains readable.
    let _provider = read_u32(header, 20, "EncryptionHeader.ProviderType")?;
    let _reserved1 = read_u32(header, 24, "EncryptionHeader.Reserved1")?;
    require_u32(header, 28, 0, "EncryptionHeader.Reserved2")?;
    validate_provider_name(&header[32..])?;

    const VERIFIER_LEN: usize = 4 + BLOCK + BLOCK + 4 + 32;
    let verifier_end = header_end
        .checked_add(VERIFIER_LEN)
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
    let hash = array::<32>(verifier, 4 + BLOCK + BLOCK + 4, "EncryptedVerifierHash")?;
    Ok(Verifier {
        salt,
        encrypted,
        hash,
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
        || bytes[..body_end].chunks_exact(2).any(|pair| pair == [0, 0])
    {
        return Err(malformed(
            "Standard CSPName terminator is missing or not final",
        ));
    }
    let units = bytes[..body_end]
        .chunks_exact(2)
        .map(|pair| u16::from_le_bytes([pair[0], pair[1]]));
    if char::decode_utf16(units).any(|character| character.is_err()) {
        return Err(malformed("Standard CSPName contains invalid UTF-16"));
    }
    Ok(())
}

fn verify(key: &[u8; BLOCK], verifier: &Verifier) -> Result<()> {
    let cipher = cipher(key)?;
    let mut clear = Zeroizing::new(verifier.encrypted);
    cipher.decrypt_block((&mut *clear).into());
    let mut sha = Sha1::new();
    sha.update(clear.as_slice());
    let expected = Zeroizing::new(<[u8; 20]>::from(sha.finalize()));
    let mut stored = Zeroizing::new(verifier.hash);
    crypt_blocks(&cipher, stored.as_mut(), Direction::Decrypt)?;
    if !bool::from(stored[..20].ct_eq(expected.as_slice())) {
        return Err(Error::Password);
    }
    Ok(())
}

fn encrypt_package(mut package: Vec<u8>, key: &[u8; BLOCK], limits: &Limits) -> Result<Vec<u8>> {
    let clear_len = package.len();
    let cipher_len = round_up(clear_len, BLOCK)?;
    let total = cipher_len
        .checked_add(8)
        .ok_or_else(|| malformed("Standard EncryptedPackage size overflows usize"))?;
    limits.bytes("EncryptedPackage", total, limits.max_encrypted_bytes)?;

    package
        .try_reserve_exact(total.saturating_sub(package.len()))
        .map_err(|_| Error::Allocation("Standard EncryptedPackage"))?;
    package.resize(total, 0);
    package.copy_within(0..clear_len, 8);
    package
        .get_mut(..8)
        .ok_or_else(|| malformed("Standard EncryptedPackage prefix is unavailable"))?
        .copy_from_slice(
            &u64::try_from(clear_len)
                .map_err(|_| malformed("plaintext size does not fit u64"))?
                .to_le_bytes(),
        );
    let cipher = cipher(key)?;
    let ciphertext = package
        .get_mut(8..)
        .ok_or_else(|| malformed("Standard EncryptedPackage ciphertext is unavailable"))?;
    crypt_blocks(&cipher, ciphertext, Direction::Encrypt)?;
    Ok(package)
}

fn decrypt_package(mut encrypted: Vec<u8>, key: &[u8; BLOCK], limits: &Limits) -> Result<Vec<u8>> {
    if encrypted.len() < 8 + BLOCK {
        return Err(malformed("Standard EncryptedPackage is too short"));
    }
    let declared = u64::from_le_bytes(array::<8>(&encrypted, 0, "StreamSize")?);
    let clear_len = declared_size(declared, limits)?;
    if clear_len == 0 {
        return Err(malformed(
            "Standard EncryptedPackage declares an empty package",
        ));
    }
    let expected_cipher = round_up(clear_len, BLOCK)?;
    if encrypted.len() != expected_cipher + 8 {
        return Err(malformed(
            "Standard EncryptedPackage length disagrees with StreamSize",
        ));
    }
    let cipher = cipher(key)?;
    let ciphertext = encrypted
        .get_mut(8..)
        .ok_or_else(|| malformed("Standard EncryptedPackage ciphertext is unavailable"))?;
    crypt_blocks(&cipher, ciphertext, Direction::Decrypt)?;
    let source_end = clear_len
        .checked_add(8)
        .ok_or_else(|| malformed("Standard plaintext range overflows usize"))?;
    if encrypted.get(8..source_end).is_none() {
        return Err(malformed("Standard decrypted package is truncated"));
    }
    encrypted.copy_within(8..source_end, 0);
    encrypted.truncate(clear_len);
    Ok(encrypted)
}

fn cipher(key: &[u8; BLOCK]) -> Result<Aes128> {
    Aes128::new_from_slice(key).map_err(|_| malformed("AES-128 key length invariant was violated"))
}

#[derive(Clone, Copy)]
enum Direction {
    Encrypt,
    Decrypt,
}

fn crypt_blocks(cipher: &Aes128, bytes: &mut [u8], direction: Direction) -> Result<()> {
    if !bytes.len().is_multiple_of(BLOCK) {
        return Err(malformed("AES data is not aligned to a 16-byte block"));
    }
    for chunk in bytes.chunks_exact_mut(BLOCK) {
        let block: &mut Block<Aes128> = chunk
            .try_into()
            .map_err(|_| malformed("AES block conversion failed"))?;
        match direction {
            Direction::Encrypt => cipher.encrypt_block(block),
            Direction::Decrypt => cipher.decrypt_block(block),
        }
    }
    Ok(())
}

fn round_up(value: usize, multiple: usize) -> Result<usize> {
    value
        .checked_add(multiple - 1)
        .map(|value| value / multiple * multiple)
        .ok_or_else(|| malformed("encrypted block length overflows usize"))
}

fn read_u32(bytes: &[u8], offset: usize, field: &'static str) -> Result<u32> {
    Ok(u32::from_le_bytes(array::<4>(bytes, offset, field)?))
}

fn validate_flags(flags: u32) -> Result<()> {
    if flags & REQUIRED_FLAGS != REQUIRED_FLAGS || flags & FORBIDDEN_FLAGS != 0 {
        return Err(Error::Unsupported(format!(
            "Standard EncryptionHeader flags {flags:#010x} do not select ECMA-376 AES"
        )));
    }
    Ok(())
}

fn require_u32(bytes: &[u8], offset: usize, expected: u32, field: &'static str) -> Result<()> {
    let actual = read_u32(bytes, offset, field)?;
    if actual != expected {
        return Err(Error::Unsupported(format!(
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
        .map_err(|_| malformed(format!("{field} has the wrong length")))
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::ooxml::{Kind, Mode, inspect, open_with, rekey};

    const SALT: [u8; BLOCK] = [
        0x92, 0x25, 0x50, 0xf6, 0xb6, 0x4f, 0xfe, 0x5b, 0xd3, 0x96, 0xdf, 0x5e, 0xe9, 0x17, 0xda,
        0x3a,
    ];
    const VERIFIER: [u8; BLOCK] = *b"fixed verifier!!";

    #[test]
    fn standard_round_trip_is_move_first_and_wrong_password_is_typed() {
        let limits = Limits::default();
        let clear = Vec::from(&b"PK\x03\x04deterministic standard package"[..]);
        let encrypted = encrypt_with(clear.clone(), "correct horse", &SALT, &VERIFIER, &limits)
            .expect("encrypt Standard package");
        assert_eq!(
            inspect(&encrypted).expect("classify Standard package"),
            Kind::Encrypted(Mode::Standard)
        );
        let opened = open_with(encrypted.clone(), "correct horse", &limits)
            .expect("decrypt Standard package");
        assert_eq!(opened.mode(), Some(Mode::Standard));
        assert_eq!(opened.bytes(), clear);
        assert!(matches!(
            open_with(encrypted, "wrong", &limits),
            Err(Error::Password)
        ));

        let encrypted = encrypt_with(clear.clone(), "old", &SALT, &VERIFIER, &limits)
            .expect("encrypt package for rekey");
        let rekeyed = rekey(encrypted, "old", "new").expect("rekey Standard package");
        assert_eq!(
            open_with(rekeyed.clone(), "new", &limits)
                .expect("open rekeyed package")
                .bytes(),
            clear
        );
        assert!(matches!(
            open_with(rekeyed, "old", &limits),
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
        assert_eq!(
            parsed.encrypted,
            [
                0xbf, 0x86, 0xe1, 0x8f, 0x64, 0x9d, 0x17, 0xd0, 0xa5, 0x41, 0xd9, 0x45, 0xce, 0xfd,
                0x96, 0x0c,
            ]
        );
    }

    #[test]
    fn declared_size_is_bounded_before_output_allocation() {
        let limits = Limits {
            max_plaintext_bytes: 32,
            ..Limits::default()
        };
        let key = Zeroizing::new([0u8; BLOCK]);
        let mut encrypted = vec![0u8; 8 + 48];
        encrypted[..8].copy_from_slice(&33u64.to_le_bytes());
        assert!(matches!(
            decrypt_package(encrypted, &key, &limits),
            Err(Error::Limit {
                resource: "declared plaintext",
                actual: 33,
                maximum: 32,
            })
        ));
    }

    #[test]
    fn header_profile_is_validated_instead_of_skipped() {
        let (encrypted, hash) = encrypt_verifier(&[0u8; BLOCK], &VERIFIER).expect("verifier");
        let mut info = build_info(&SALT, &encrypted, &hash).expect("info");
        info[20] ^= 1;
        assert!(matches!(parse(&info), Err(Error::Unsupported(_))));
    }

    #[test]
    fn flags_ignore_undefined_bits_but_require_matching_safe_bits() {
        let (encrypted, hash) = encrypt_verifier(&[0u8; BLOCK], &VERIFIER).expect("verifier");
        let mut info = build_info(&SALT, &encrypted, &hash).expect("info");
        let extended = FLAGS | 0x8000_0000;
        info[4..8].copy_from_slice(&extended.to_le_bytes());
        info[12..16].copy_from_slice(&extended.to_le_bytes());
        parse(&info).expect("undefined flag bits are ignored");

        info[4..8].copy_from_slice(&FLAGS.to_le_bytes());
        assert!(matches!(parse(&info), Err(Error::Malformed(_))));

        let forbidden = FLAGS | 0x08;
        info[4..8].copy_from_slice(&forbidden.to_le_bytes());
        info[12..16].copy_from_slice(&forbidden.to_le_bytes());
        assert!(matches!(parse(&info), Err(Error::Unsupported(_))));
    }

    fn hex(value: &str) -> Vec<u8> {
        value
            .split_ascii_whitespace()
            .map(|byte| u8::from_str_radix(byte, 16).expect("test hex"))
            .collect()
    }
}
