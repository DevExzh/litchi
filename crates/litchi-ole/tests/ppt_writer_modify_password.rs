use std::io::Cursor;

use litchi_ole::ppt::writer::{PptEncryptionProfile, PptWriter};
use litchi_ole::ppt::{Package, PptError, PptOpenOptions};

fn encrypted_presentation(modify_password: &str, open_password: &str) -> Vec<u8> {
    let mut writer = PptWriter::new();
    writer.add_slide().unwrap();
    writer
        .set_password(
            open_password,
            PptEncryptionProfile::CryptoApiRc4 { key_bits: 128 },
        )
        .unwrap();
    writer.set_modify_password(modify_password).unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    output.into_inner()
}

fn open(bytes: &[u8], password: Option<&str>) -> Result<litchi_ole::ppt::Presentation, PptError> {
    let mut package = Package::from_reader(Cursor::new(bytes)).unwrap();
    package.presentation_with_options(PptOpenOptions { password })
}

#[test]
fn encrypted_modify_password_round_trips_and_open_password_behavior_is_unchanged() {
    let bytes = encrypted_presentation("修改🔐", "open secret");
    assert!(matches!(open(&bytes, None), Err(PptError::PasswordRequired)));
    assert!(matches!(open(&bytes, Some("wrong")), Err(PptError::InvalidPassword)));
    let presentation = open(&bytes, Some("open secret")).unwrap();
    let password = presentation.modify_password().unwrap().unwrap();
    assert_eq!(password.expose_secret(), "修改🔐");
    let debug = format!("{password:?}");
    assert!(debug.contains("[REDACTED]"));
    assert!(!debug.contains("修改"));
}

#[test]
fn setter_replacement_clear_and_failures_are_atomic() {
    let mut writer = PptWriter::new();
    writer.add_slide().unwrap();
    writer.set_modify_password("first").unwrap();
    writer.set_modify_password("replacement").unwrap();
    assert_eq!(writer.modify_password().unwrap().expose_secret(), "replacement");
    assert!(writer.set_modify_password("bad\nvalue").is_err());
    assert!(writer.set_modify_password("x".repeat(256)).is_err());
    assert_eq!(writer.modify_password().unwrap().expose_secret(), "replacement");

    let mut output = Cursor::new(Vec::new());
    assert!(writer.write_to(&mut output).is_err());
    assert!(output.get_ref().is_empty());

    writer
        .set_password(
            "open",
            PptEncryptionProfile::CryptoApiRc4 { key_bits: 40 },
        )
        .unwrap();
    writer.clear_modify_password();
    assert!(writer.modify_password().is_none());
    writer.write_to(&mut output).unwrap();
    let presentation = open(output.get_ref(), Some("open")).unwrap();
    assert!(presentation.modify_password().unwrap().is_none());
}
