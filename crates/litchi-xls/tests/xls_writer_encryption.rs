use std::io::Cursor;

use litchi_cfb::OleFile;
use litchi_core::sheet::{Cell, CellValue, WorkbookTrait};
use litchi_xls::writer::{EncryptionProfile, Writer};
use litchi_xls::{Error, OpenOptions, WeakEncryptionPolicy, Workbook};

fn encrypted_workbook(profile: EncryptionProfile, password: &str) -> Vec<u8> {
    let mut writer = Writer::new();
    let first = writer.add_worksheet("Data").unwrap();
    let second = writer.add_worksheet("Second").unwrap();
    writer.write_string(first, 0, 0, "shared text").unwrap();
    writer.write_string(second, 0, 0, "shared text").unwrap();
    writer.write_number(first, 1, 0, 42.5).unwrap();
    writer.write_formula(first, 2, 0, "A2*2").unwrap();
    writer
        .add_comment(first, 0, 0, "author", "comment text")
        .unwrap();
    match profile {
        EncryptionProfile::XorObfuscation => writer
            .set_xor_obfuscation_password(password, WeakEncryptionPolicy::allow_xor_obfuscation())
            .unwrap(),
        _ => writer.set_password(password, profile).unwrap(),
    }
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    output.into_inner()
}

fn open<'a>(bytes: &'a [u8], password: Option<&str>) -> Result<Workbook<Cursor<&'a [u8]>>, Error> {
    Workbook::new_with_options(
        Cursor::new(bytes),
        OpenOptions::new().with_password(password),
    )
}

fn workbook_stream(bytes: &[u8]) -> Vec<u8> {
    let mut ole = OleFile::open(Cursor::new(bytes)).unwrap();
    ole.open_stream(&["Workbook"]).unwrap()
}

fn filepass(stream: &[u8]) -> (usize, &[u8]) {
    let mut offset = 0usize;
    let mut found = None;
    while offset < stream.len() {
        let sid = u16::from_le_bytes(stream[offset..offset + 2].try_into().unwrap());
        let len = usize::from(u16::from_le_bytes(
            stream[offset + 2..offset + 4].try_into().unwrap(),
        ));
        if sid == 0x002f {
            assert!(found.is_none());
            found = Some((offset, &stream[offset + 4..offset + 4 + len]));
        }
        offset += 4 + len;
    }
    found.unwrap()
}

fn assert_round_trip(profile: EncryptionProfile, password: &str) -> Vec<u8> {
    let bytes = encrypted_workbook(profile, password);
    assert!(matches!(open(&bytes, None), Err(Error::PasswordRequired)));
    assert!(matches!(
        open(&bytes, Some("wrong")),
        Err(Error::InvalidPassword)
    ));
    let workbook = open(&bytes, Some(password)).unwrap();
    assert_eq!(workbook.worksheet_count(), 2);
    let first = workbook.xls_worksheet(0).unwrap();
    assert!(
        matches!(first.get_cell(0, 0).unwrap().value(), CellValue::String(value) if value == "shared text")
    );
    assert!(
        matches!(first.get_cell(1, 0).unwrap().value(), CellValue::Float(value) if *value == 42.5)
    );
    assert_eq!(first.comments()[0].text(), "comment text");
    assert!(
        matches!(workbook.xls_worksheet(1).unwrap().get_cell(0, 0).unwrap().value(), CellValue::String(value) if value == "shared text")
    );
    bytes
}

#[test]
fn all_profiles_round_trip_and_emit_exact_filepass_families() {
    let xor = assert_round_trip(EncryptionProfile::XorObfuscation, "cafe");
    let xor_stream = workbook_stream(&xor);
    let (_, xor_pass) = filepass(&xor_stream);
    assert_eq!(xor_pass.len(), 6);
    assert_eq!(&xor_pass[..2], &[0, 0]);

    let binary = assert_round_trip(EncryptionProfile::OfficeBinaryRc4, "密码🔐");
    let binary_stream = workbook_stream(&binary);
    let (_, binary_pass) = filepass(&binary_stream);
    assert_eq!(binary_pass.len(), 54);
    assert_eq!(&binary_pass[..6], &[1, 0, 1, 0, 1, 0]);

    for key_bits in [40, 56, 120, 128] {
        let bytes = assert_round_trip(EncryptionProfile::CryptoApiRc4 { key_bits }, "密码🔐");
        let stream = workbook_stream(&bytes);
        let (offset, pass) = filepass(&stream);
        assert_eq!(&pass[..6], &[1, 0, 2, 0, 2, 0]);
        assert_eq!(
            u32::from_le_bytes(pass[30..34].try_into().unwrap()),
            u32::from(key_bits)
        );
        let next = offset + 4 + pass.len();
        assert_eq!(
            u16::from_le_bytes(stream[next..next + 2].try_into().unwrap()),
            0x0042
        );
    }
}

#[test]
fn salts_are_nondeterministic_and_configuration_changes_are_atomic() {
    let first = encrypted_workbook(EncryptionProfile::OfficeBinaryRc4, "secret");
    let second = encrypted_workbook(EncryptionProfile::OfficeBinaryRc4, "secret");
    let first_stream = workbook_stream(&first);
    let second_stream = workbook_stream(&second);
    assert_ne!(
        &filepass(&first_stream).1[6..22],
        &filepass(&second_stream).1[6..22]
    );

    let mut writer = Writer::new();
    writer.add_worksheet("Sheet").unwrap();
    writer
        .set_password("kept", EncryptionProfile::OfficeBinaryRc4)
        .unwrap();
    assert!(
        writer
            .set_password("", EncryptionProfile::XorObfuscation)
            .is_err()
    );
    assert!(
        writer
            .set_password("not ANSI 🔐", EncryptionProfile::XorObfuscation)
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
