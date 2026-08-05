use super::codec::*;
use crate::package::DocError;
use crate::parts::fib::FileInformationBlock;
use litchi_crypto::rc4 as office_rc4;
use zeroize::Zeroizing;

fn xor_fib(verifier: u32, language_id: u16) -> FileInformationBlock {
    let mut data = vec![0u8; FIB_BASE_LEN];
    data[0..2].copy_from_slice(&0xa5ecu16.to_le_bytes());
    data[2..4].copy_from_slice(&0x00c1u16.to_le_bytes());
    data[6..8].copy_from_slice(&language_id.to_le_bytes());
    data[10..12].copy_from_slice(&0x8100u16.to_le_bytes());
    data[14..18].copy_from_slice(&verifier.to_le_bytes());
    FileInformationBlock::parse(&data).unwrap()
}

#[test]
fn xor_method_two_matches_poi_and_libreoffice_vectors() {
    let password = xor_password_bytes("abc");
    assert_eq!(xor_password_verifier(&password), 0x514a_cc1a);
    assert_eq!(
        create_word_xor_array(&password),
        [
            0x95, 0x99, 0x94, 0x75, 0xda, 0x57, 0x78, 0x57, 0xda, 0x74, 0x65, 0xa8, 0x7a, 0x2f,
            0x25, 0x77,
        ]
    );

    let context = XorContext {
        array: Zeroizing::new([0x10; 16]),
    };
    let mut special = [0x00, 0x10, 0x11];
    apply_xor_stream(&mut special, 0, &context).unwrap();
    assert_eq!(special, [0x00, 0x10, 0x01]);
}

#[test]
fn xor_decrypts_all_streams_at_absolute_offsets_after_verification() {
    let password = xor_password_bytes("abc");
    let fib = xor_fib(xor_password_verifier(&password), 0x0409);
    let context = XorContext {
        array: Zeroizing::new(create_word_xor_array(&password)),
    };
    let original_word = vec![0x5a; 101];
    let original_table = vec![0x6b; 37];
    let original_data = vec![0x7c; 41];
    let mut word = original_word.clone();
    let mut table = original_table.clone();
    let mut data = original_data.clone();
    apply_xor_stream(&mut word[FIB_BASE_LEN..], FIB_BASE_LEN, &context).unwrap();
    apply_xor_stream(&mut table, 0, &context).unwrap();
    apply_xor_stream(&mut data, 0, &context).unwrap();
    let encrypted_word = word.clone();
    let encrypted_table = table.clone();
    let encrypted_data = data.clone();

    assert!(matches!(
        decrypt_document_streams(&fib, &mut word, &mut table, Some(&mut data), None),
        Err(DocError::PasswordRequired)
    ));
    assert_eq!(word, encrypted_word);
    assert_eq!(table, encrypted_table);
    assert_eq!(data, encrypted_data);
    assert!(matches!(
        decrypt_document_streams(&fib, &mut word, &mut table, Some(&mut data), Some("wrong"),),
        Err(DocError::InvalidPassword)
    ));
    assert_eq!(word, encrypted_word);
    assert_eq!(table, encrypted_table);
    assert_eq!(data, encrypted_data);

    decrypt_document_streams(&fib, &mut word, &mut table, Some(&mut data), Some("abc")).unwrap();
    assert_eq!(word, original_word);
    assert_eq!(table, original_table);
    assert_eq!(data, original_data);
    assert_eq!(&word[..FIB_BASE_LEN], &encrypted_word[..FIB_BASE_LEN]);
}

#[test]
fn xor_accepts_lcid_ansi_password_conversion_and_truncates_to_fifteen_bytes() {
    assert_eq!(
        xor_password_bytes("abcdefghijklmnop").as_slice(),
        b"abcdefghijklmno"
    );
    assert_eq!(
        ansi_password_bytes("abcdefghijklmnop", 0x0409).as_slice(),
        b"abcdefghijklmno"
    );
    assert_eq!(xor_password_bytes("€").as_slice(), &[0xac]);
    let ansi = ansi_password_bytes("€", 0x0409);
    assert_eq!(ansi.as_slice(), &[0x80]);

    let fib = xor_fib(xor_password_verifier(&ansi), 0x0409);
    let context = XorContext {
        array: Zeroizing::new(create_word_xor_array(&ansi)),
    };
    let original_word = vec![0x42; 84];
    let original_table = vec![0x24; 19];
    let mut word = original_word.clone();
    let mut table = original_table.clone();
    apply_xor_stream(&mut word[FIB_BASE_LEN..], FIB_BASE_LEN, &context).unwrap();
    apply_xor_stream(&mut table, 0, &context).unwrap();
    decrypt_document_streams(&fib, &mut word, &mut table, None, Some("€")).unwrap();
    assert_eq!(word, original_word);
    assert_eq!(table, original_table);
}

#[test]
fn binary_rc4_secret_matches_apache_poi_vector() {
    let salt = [
        0x17, 0xf6, 0xd1, 0x6b, 0x09, 0xb1, 0x5f, 0x7b, 0x4c, 0x9d, 0x03, 0xb4, 0x81, 0xb5, 0xb4,
        0x4a,
    ];
    assert_eq!(
        derive_secret("MoneyForNothing", &salt).as_ref(),
        &[0xc2, 0xd9, 0x56, 0xb2, 0x6b]
    );
}

#[test]
fn stream_cipher_preserves_absolute_block_position() {
    let secret = [1, 2, 3, 4, 5];
    let mut data = vec![0x5a; 80];
    let expected = data.clone();
    apply_stream_cipher(&mut data, 500, &secret).unwrap();
    assert_ne!(data, expected);
    apply_stream_cipher(&mut data, 500, &secret).unwrap();
    assert_eq!(data, expected);
}

#[test]
fn cryptoapi_stream_rekeys_at_512_byte_boundaries() {
    let context = office_rc4::context("stream-position", &[0x42; 16], 120).unwrap();
    let original = vec![0x5a; 80];
    let mut data = original.clone();
    apply_cryptoapi_stream(&mut data, 500, &context).unwrap();
    assert_ne!(data, original);
    apply_cryptoapi_stream(&mut data, 500, &context).unwrap();
    assert_eq!(data, original);
}

#[test]
fn cryptoapi_clear_prefix_offsets_are_consumed() {
    let context = office_rc4::context("clear-prefix", &[0x24; 16], 56).unwrap();
    let mut word = vec![0x11; 620];
    let original = word.clone();
    apply_cryptoapi_stream(&mut word[FIB_BASE_LEN..], FIB_BASE_LEN, &context).unwrap();
    assert_eq!(&word[..FIB_BASE_LEN], &original[..FIB_BASE_LEN]);
    apply_cryptoapi_stream(&mut word[FIB_BASE_LEN..], FIB_BASE_LEN, &context).unwrap();
    assert_eq!(word, original);
}
