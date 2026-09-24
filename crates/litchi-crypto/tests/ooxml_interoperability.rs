//! Independent OOXML encryption vectors generated outside `litchi-crypto`.

#![cfg(feature = "ooxml")]

use litchi_crypto::ooxml::{Kind, Limits, Mode, inspect_with, open_with, rekey_with};
use litchi_crypto::spaces;

const PASSWORD: &str = "Litchi synthetic crypto fixture 2026";
const CLEAR: &[u8] = include_bytes!("data/ooxml/clear.docx");

#[test]
fn opens_independent_standard_aes192_and_aes256_vectors() {
    let limits = Limits::default();
    for (name, expected_mode, encrypted) in [
        (
            "Standard AES-192",
            Mode::StandardAes192,
            include_bytes!("data/ooxml/component-standard-aes192-sha1.docx").as_slice(),
        ),
        (
            "Standard AES-256",
            Mode::StandardAes256,
            include_bytes!("data/ooxml/component-standard-aes256-sha1.docx").as_slice(),
        ),
    ] {
        assert_eq!(
            inspect_with(encrypted, &limits).unwrap_or_else(|error| {
                panic!("{name} profile classification failed: {error}")
            }),
            Kind::Encrypted(expected_mode),
        );
        let opened = open_with(encrypted.to_vec(), PASSWORD, &limits)
            .unwrap_or_else(|error| panic!("{name} fixture must open: {error}"));
        assert_eq!(opened.mode(), Some(expected_mode));
        assert_eq!(opened.bytes(), CLEAR, "{name} clear package bytes");
    }
}

#[test]
fn opens_independent_agile_sha512_and_mixed_key_size_vectors() {
    let limits = Limits::default();
    for (name, encrypted) in [
        (
            "msoffcrypto Agile AES-256/SHA-512",
            include_bytes!("data/ooxml/msoffcrypto-agile-aes256-sha512-valid-graph.docx")
                .as_slice(),
        ),
        (
            "mixed Agile AES-128 wrapper/AES-256 data",
            include_bytes!("data/ooxml/component-agile-aes128-wrap-aes256-data-sha512.docx")
                .as_slice(),
        ),
    ] {
        assert_eq!(
            inspect_with(encrypted, &limits).unwrap_or_else(|error| {
                panic!("{name} profile classification failed: {error}")
            }),
            Kind::Encrypted(Mode::AgileAes256Sha512),
        );
        let opened = open_with(encrypted.to_vec(), PASSWORD, &limits)
            .unwrap_or_else(|error| panic!("{name} fixture must open: {error}"));
        assert_eq!(opened.mode(), Some(Mode::AgileAes256Sha512));
        assert_eq!(opened.bytes(), CLEAR, "{name} clear package bytes");
    }
}

#[test]
fn rekey_preserves_an_opened_non_default_profile() {
    let limits = Limits::default();
    let input =
        include_bytes!("data/ooxml/msoffcrypto-agile-aes256-sha512-valid-graph.docx").to_vec();
    let rekeyed = rekey_with(input, PASSWORD, "new fixture password", &limits)
        .expect("rekey independent Agile fixture");
    assert_eq!(
        inspect_with(&rekeyed, &limits).expect("inspect rekeyed fixture"),
        Kind::Encrypted(Mode::AgileAes256Sha512)
    );
    let opened = open_with(rekeyed, "new fixture password", &limits).expect("open rekeyed fixture");
    assert_eq!(opened.mode(), Some(Mode::AgileAes256Sha512));
    assert_eq!(opened.bytes(), CLEAR);
}

#[test]
fn ooxml_read_compatibility_normalizes_msoffcrypto_zero_block_graph() {
    let encrypted = include_bytes!("data/ooxml/msoffcrypto-agile-aes256-sha512.docx");
    assert!(spaces::inspect_bytes(encrypted).is_err());
    let opened = open_with(encrypted.to_vec(), PASSWORD, &Limits::default())
        .expect("OOXML compatibility reader accepts advisory zero block size");
    assert_eq!(opened.mode(), Some(Mode::AgileAes256Sha512));
    assert_eq!(opened.bytes(), CLEAR);
}
