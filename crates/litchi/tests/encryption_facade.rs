#![cfg(feature = "encryption")]

const PASSWORD: &str = "FacadeSecret123";
const NEW_PASSWORD: &str = "FacadeSecret456";

#[cfg(feature = "xlsx")]
#[test]
fn xlsx_facade_exposes_managed_package_and_workbook_encryption() {
    use litchi::xlsx::{Package, Workbook, encryption::Mode};
    use std::io::Cursor;

    let workbook = Workbook::create().expect("create workbook");
    let encrypted = workbook
        .to_encrypted(PASSWORD, Mode::Agile)
        .expect("encrypt workbook");
    let reopened = Workbook::from_reader_with_password(Cursor::new(encrypted), PASSWORD)
        .expect("open encrypted workbook");
    assert_eq!(reopened.encryption(), Some(Mode::Agile));

    let encrypted = reopened
        .to_reencrypted(NEW_PASSWORD)
        .expect("re-encrypt workbook");
    let package = Package::from_reader_with_password(Cursor::new(encrypted), NEW_PASSWORD)
        .expect("open encrypted package");
    assert_eq!(package.encryption(), Some(Mode::Agile));
}

#[cfg(feature = "pptx")]
#[test]
fn pptx_facade_exposes_managed_package_encryption() {
    use litchi::pptx::{Package, encryption::Mode};
    use std::io::Cursor;

    let mut package = Package::new().expect("create presentation package");
    let encrypted = package
        .to_encrypted(PASSWORD, Mode::Standard)
        .expect("encrypt presentation package");
    let mut reopened = Package::from_reader_with_password(Cursor::new(encrypted), PASSWORD)
        .expect("open encrypted presentation package");
    assert_eq!(reopened.encryption(), Some(Mode::Standard));

    let encrypted = reopened
        .to_reencrypted(NEW_PASSWORD)
        .expect("re-encrypt presentation package");
    let reopened = Package::from_reader_with_password(Cursor::new(encrypted), NEW_PASSWORD)
        .expect("open re-encrypted presentation package");
    assert_eq!(reopened.encryption(), Some(Mode::Standard));
}
