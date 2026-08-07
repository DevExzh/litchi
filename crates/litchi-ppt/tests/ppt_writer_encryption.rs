#![cfg(feature = "encryption")]

use litchi_cfb::OleFile;
use litchi_ppt::writer::{EncryptionProfile, PictureKind, Writer};
use litchi_ppt::{Error, OpenOptions, Package};
use std::collections::BTreeMap;
use std::io::Cursor;

fn write(profile: EncryptionProfile, password: &str, picture: bool) -> Vec<u8> {
    let mut writer = Writer::new();
    let slide = writer.add_slide().unwrap();
    writer
        .add_textbox(slide, 10, 10, 300, 40, "Encrypted 文本")
        .unwrap();
    writer.set_slide_notes(slide, "Speaker notes").unwrap();
    if picture {
        writer
            .add_picture_data_as(
                vec![0x89, b'P', b'N', b'G', 0x0d, 0x0a, 0x1a, 0x0a, 1, 2, 3, 4],
                PictureKind::Png,
            )
            .unwrap();
    }
    writer.set_password(password, profile).unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    output.into_inner()
}

fn streams(file: &[u8]) -> (Vec<u8>, Vec<u8>, Option<Vec<u8>>) {
    let mut ole = OleFile::open(Cursor::new(file)).unwrap();
    let document = ole.open_stream(&["PowerPoint Document"]).unwrap();
    let current_user = ole.open_stream(&["Current User"]).unwrap();
    let pictures = ole.open_stream(&["Pictures"]).ok();
    (document, current_user, pictures)
}

fn persist_mappings(document: &[u8], offset: usize) -> BTreeMap<u32, u32> {
    let len = u32::from_le_bytes(document[offset + 4..offset + 8].try_into().unwrap()) as usize;
    let mut cursor = offset + 8;
    let end = cursor + len;
    let mut output = BTreeMap::new();
    while cursor < end {
        let info = u32::from_le_bytes(document[cursor..cursor + 4].try_into().unwrap());
        cursor += 4;
        let base = info & 0x000f_ffff;
        for index in 0..(info >> 20) {
            let value = u32::from_le_bytes(document[cursor..cursor + 4].try_into().unwrap());
            cursor += 4;
            output.insert(base + index, value);
        }
    }
    output
}

fn assert_round_trip(profile: EncryptionProfile, password: &str, picture: bool) -> Vec<u8> {
    let bytes = write(profile, password, picture);
    let mut package = Package::from_reader(Cursor::new(bytes.clone())).unwrap();
    assert!(matches!(
        package.presentation(),
        Err(Error::PasswordRequired)
    ));
    assert!(matches!(
        package.presentation_with_options(OpenOptions {
            password: Some("wrong"),
            ..OpenOptions::default()
        }),
        Err(Error::InvalidPassword)
    ));
    let presentation = package
        .presentation_with_options(OpenOptions {
            password: Some(password),
            ..OpenOptions::default()
        })
        .unwrap();
    assert!(presentation.text().unwrap().contains("Encrypted 文本"));
    if picture {
        assert!(presentation.has_pictures());
        assert_eq!(presentation.images().unwrap().len(), 1);
    }
    bytes
}

#[test]
fn cryptoapi_profiles_round_trip_and_emit_exact_bootstrap() {
    for key_bits in [40, 56, 120, 128] {
        let bytes = assert_round_trip(
            EncryptionProfile::CryptoApiRc4 { key_bits },
            "密码🔐",
            false,
        );
        let (document, current_user, _) = streams(&bytes);
        assert_eq!(
            u32::from_le_bytes(current_user[12..16].try_into().unwrap()),
            0xf3d1_c4df
        );
        let user_edit_offset =
            u32::from_le_bytes(current_user[16..20].try_into().unwrap()) as usize;
        assert_eq!(
            u16::from_le_bytes(
                document[user_edit_offset + 2..user_edit_offset + 4]
                    .try_into()
                    .unwrap()
            ),
            4085
        );
        assert_eq!(
            u32::from_le_bytes(
                document[user_edit_offset + 4..user_edit_offset + 8]
                    .try_into()
                    .unwrap()
            ),
            32
        );
        let directory_offset = u32::from_le_bytes(
            document[user_edit_offset + 20..user_edit_offset + 24]
                .try_into()
                .unwrap(),
        ) as usize;
        let session_id = u32::from_le_bytes(
            document[user_edit_offset + 36..user_edit_offset + 40]
                .try_into()
                .unwrap(),
        );
        let mappings = persist_mappings(&document, directory_offset);
        let session_offset = mappings[&session_id] as usize;
        assert_eq!(
            u16::from_le_bytes(
                document[session_offset + 2..session_offset + 4]
                    .try_into()
                    .unwrap()
            ),
            12052
        );
        assert_eq!(
            &document[session_offset + 8..session_offset + 16],
            &[2, 0, 2, 0, 12, 0, 0, 0]
        );
        assert_eq!(
            u32::from_le_bytes(
                document[session_offset + 36..session_offset + 40]
                    .try_into()
                    .unwrap()
            ),
            u32::from(key_bits)
        );
        assert_ne!(u16::from_le_bytes(document[2..4].try_into().unwrap()), 1000);
    }
}

#[test]
fn pictures_salts_and_atomic_profile_validation_are_covered() {
    let first = assert_round_trip(
        EncryptionProfile::CryptoApiRc4 { key_bits: 128 },
        "secret",
        true,
    );
    let second = write(
        EncryptionProfile::CryptoApiRc4 { key_bits: 128 },
        "secret",
        false,
    );
    let (first_document, first_current, pictures) = streams(&first);
    assert!(pictures.unwrap().iter().any(|byte| *byte != 0));
    let (second_document, second_current, _) = streams(&second);
    let salt = |document: &[u8], current: &[u8]| {
        let user = u32::from_le_bytes(current[16..20].try_into().unwrap()) as usize;
        let directory =
            u32::from_le_bytes(document[user + 20..user + 24].try_into().unwrap()) as usize;
        let session = u32::from_le_bytes(document[user + 36..user + 40].try_into().unwrap());
        let offset = persist_mappings(document, directory)[&session] as usize;
        let header_size =
            u32::from_le_bytes(document[offset + 16..offset + 20].try_into().unwrap()) as usize;
        let salt_offset = offset + 24 + header_size;
        document[salt_offset..salt_offset + 16].to_vec()
    };
    assert_ne!(
        salt(&first_document, &first_current),
        salt(&second_document, &second_current)
    );

    let mut writer = Writer::new();
    writer
        .set_password("kept", EncryptionProfile::CryptoApiRc4 { key_bits: 128 })
        .unwrap();
    assert!(
        writer
            .set_password("", EncryptionProfile::CryptoApiRc4 { key_bits: 40 })
            .is_err()
    );
    assert!(
        writer
            .set_password("bad", EncryptionProfile::CryptoApiRc4 { key_bits: 41 })
            .is_err()
    );
    assert_eq!(
        writer.encryption_profile(),
        Some(EncryptionProfile::CryptoApiRc4 { key_bits: 128 })
    );
    writer.clear_password();
    assert_eq!(writer.encryption_profile(), None);
}
