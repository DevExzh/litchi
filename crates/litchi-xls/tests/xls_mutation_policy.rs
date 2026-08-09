use std::io::Cursor;

use litchi_core::sheet::Worksheet as _;
use litchi_xls::writer::Writer;
use litchi_xls::{EncryptionProfile, Error, WeakEncryptionPolicy, Workbook, Worksheet};

fn decoded_worksheet() -> Worksheet {
    let mut writer = Writer::new();
    writer.add_worksheet("Data").unwrap();
    let mut bytes = Cursor::new(Vec::new());
    writer.write_to(&mut bytes).unwrap();
    let workbook = Workbook::new(Cursor::new(bytes.into_inner())).unwrap();
    workbook.xls_worksheet(0).unwrap().clone()
}

#[test]
fn decoded_worksheet_mutations_refuse_without_changing_state() {
    let mut worksheet = decoded_worksheet();
    assert!(worksheet.is_source_bound());
    let row_count_before = worksheet.row_count();
    assert!(matches!(
        worksheet.set_dimensions(0, 12, 0, 3),
        Err(Error::SourceBoundWorksheetMutation {
            operation: "set_dimensions"
        })
    ));
    assert_eq!(worksheet.row_count(), row_count_before);
    assert!(matches!(
        worksheet.protection_mut(),
        Err(Error::SourceBoundWorksheetMutation {
            operation: "protection_mut"
        })
    ));
}

#[test]
fn authored_worksheet_mutations_remain_available() {
    let mut worksheet = Worksheet::new("Authored".to_string());
    worksheet.set_dimensions(0, 12, 0, 3).unwrap();
    assert!(!worksheet.is_source_bound());
    assert_eq!(worksheet.row_count(), 12);
    assert!(worksheet.protection_mut().is_ok());
}

#[test]
fn xor_authoring_requires_an_explicit_policy() {
    let mut writer = Writer::new();
    assert!(matches!(
        writer.set_password("legacy", EncryptionProfile::XorObfuscation),
        Err(Error::WeakEncryptionRequiresExplicitPolicy)
    ));
    assert_eq!(writer.encryption_profile(), None);

    writer
        .set_xor_obfuscation_password("legacy", WeakEncryptionPolicy::allow_xor_obfuscation())
        .unwrap();
    assert_eq!(
        writer.encryption_profile(),
        Some(EncryptionProfile::XorObfuscation)
    );
}
