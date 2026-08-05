use litchi_cfb::OleFile;
use litchi_doc::writer::{EncryptionProfile, FootnoteEntry, Writer};
use litchi_doc::{Error, OpenOptions, Package};
use std::io::Cursor;

fn encrypted_document(profile: EncryptionProfile, password: &str) -> Vec<u8> {
    let mut writer = Writer::new();
    writer.add_paragraph("Main 文本").unwrap();
    writer.set_odd_header("Header");
    writer.add_footnote(FootnoteEntry::new(2, "Footnote text", 1));
    writer.set_password(password, profile).unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    output.into_inner()
}

fn streams(file: &[u8]) -> (Vec<u8>, Vec<u8>, Vec<u8>) {
    let mut ole = OleFile::open(Cursor::new(file)).unwrap();
    (
        ole.open_stream(&["WordDocument"]).unwrap(),
        ole.open_stream(&["1Table"]).unwrap(),
        ole.open_stream(&["Data"]).unwrap(),
    )
}

fn assert_round_trip(profile: EncryptionProfile, password: &str) -> Vec<u8> {
    let bytes = encrypted_document(profile, password);
    let mut package = Package::from_reader(Cursor::new(bytes.clone())).unwrap();
    assert!(matches!(package.document(), Err(Error::PasswordRequired)));
    assert!(matches!(
        package.document_with_options(OpenOptions {
            password: Some("wrong"),
            ..Default::default()
        }),
        Err(Error::InvalidPassword)
    ));
    let document = package
        .document_with_options(OpenOptions {
            password: Some(password),
            ..Default::default()
        })
        .unwrap();
    assert!(document.text().unwrap().contains("Main 文本"));
    assert_eq!(
        document.footnotes().unwrap()[0].text(),
        "\u{0002}Footnote text\r"
    );
    assert_eq!(document.headers_footers().unwrap()[0].text(), "Header\r\r");
    bytes
}

#[test]
fn all_profiles_round_trip_all_document_streams_and_exact_headers() {
    let xor = assert_round_trip(EncryptionProfile::WordXorObfuscation, "abc");
    let (word, table, data) = streams(&xor);
    assert_eq!(
        u16::from_le_bytes(word[10..12].try_into().unwrap()) & 0x8100,
        0x8100
    );
    assert_eq!(
        u32::from_le_bytes(word[14..18].try_into().unwrap()),
        0x514a_cc1a
    );
    assert!(table.iter().any(|byte| *byte != 0));
    assert_eq!(data, vec![0u8; data.len()]);

    let binary = assert_round_trip(EncryptionProfile::OfficeBinaryRc4, "密码🔐");
    let (word, table, data) = streams(&binary);
    assert_eq!(&table[..4], &[1, 0, 1, 0]);
    assert_eq!(u32::from_le_bytes(word[14..18].try_into().unwrap()), 52);
    assert!(data.iter().any(|byte| *byte != 0));

    for key_bits in [40, 56, 120, 128] {
        let bytes = assert_round_trip(EncryptionProfile::CryptoApiRc4 { key_bits }, "Unicode 密码");
        let (word, table, data) = streams(&bytes);
        let header_len = u32::from_le_bytes(word[14..18].try_into().unwrap()) as usize;
        assert_eq!(&table[..8], &[2, 0, 2, 0, 4, 0, 0, 0]);
        assert_eq!(
            u32::from_le_bytes(table[28..32].try_into().unwrap()),
            u32::from(key_bits)
        );
        assert_eq!(
            header_len,
            12 + u32::from_le_bytes(table[8..12].try_into().unwrap()) as usize + 60
        );
        assert!(data.iter().any(|byte| *byte != 0));
    }
}

#[test]
fn salts_are_nondeterministic_and_setter_validation_is_atomic() {
    let first = encrypted_document(EncryptionProfile::OfficeBinaryRc4, "secret");
    let second = encrypted_document(EncryptionProfile::OfficeBinaryRc4, "secret");
    let (_, first_table, _) = streams(&first);
    let (_, second_table, _) = streams(&second);
    assert_ne!(&first_table[4..20], &second_table[4..20]);

    let mut writer = Writer::new();
    writer
        .set_password("kept", EncryptionProfile::OfficeBinaryRc4)
        .unwrap();
    assert!(
        writer
            .set_password("", EncryptionProfile::WordXorObfuscation)
            .is_err()
    );
    assert!(
        writer
            .set_password("abcdefghijklmnop", EncryptionProfile::WordXorObfuscation,)
            .is_err()
    );
    assert!(
        writer
            .set_password("bad", EncryptionProfile::CryptoApiRc4 { key_bits: 41 })
            .is_err()
    );
    assert_eq!(
        writer.encryption_profile(),
        Some(EncryptionProfile::OfficeBinaryRc4)
    );
    writer.clear_password();
    assert_eq!(writer.encryption_profile(), None);
}
