//! Fresh authoring publishes through a validated immutable owner.
#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "acceptance assertions fail fast"
)]

use litchi_core::sheet::CellValue;
use litchi_xls::cell_values::Reference;
use litchi_xls::{Builder, EncryptionProfile, Package};

#[test]
fn default_builder_and_minimal_package_agree() {
    let mut builder = Builder::new();
    let first = builder.build().unwrap();
    assert_eq!(first.workbook().worksheet_names().unwrap(), ["Sheet1"]);
    assert_eq!(
        first.to_bytes().unwrap(),
        Package::new().unwrap().to_bytes().unwrap()
    );
    assert_eq!(
        first.to_bytes().unwrap(),
        builder.build().unwrap().to_bytes().unwrap()
    );
}

#[test]
fn builds_are_detached_from_later_authoring() {
    let mut builder = Builder::new();
    let sheet = builder.add_worksheet("Data").unwrap();
    builder.write_number(sheet, 0, 0, 1.0).unwrap();
    let first = builder.build().unwrap();
    let bytes = first.to_bytes().unwrap();
    builder.write_number(sheet, 0, 0, 2.0).unwrap();
    let second = builder.build().unwrap();
    drop(builder);
    assert_eq!(first.to_bytes().unwrap(), bytes);
    let reference = Reference::new(0, 0).unwrap();
    assert_eq!(
        first
            .workbook()
            .worksheet(0_usize)
            .unwrap()
            .unwrap()
            .cell_value(reference)
            .unwrap(),
        Some(CellValue::Float(1.0))
    );
    assert_eq!(
        second
            .workbook()
            .worksheet(0_usize)
            .unwrap()
            .unwrap()
            .cell_value(reference)
            .unwrap(),
        Some(CellValue::Float(2.0))
    );
    let mut edit = second.edit().unwrap();
    edit.set_number(0_usize, reference, 3.0).unwrap();
    let commit = edit.commit().unwrap();
    assert_eq!(
        commit
            .package()
            .apply(&commit.patch().inverse())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap(),
        second.to_bytes().unwrap()
    );
}

#[test]
fn encrypted_build_retains_the_ciphertext_and_requires_explicit_output() {
    let mut builder = Builder::new();
    let sheet = builder.add_worksheet("Data").unwrap();
    builder.write_string(sheet, 0, 0, "protected").unwrap();
    builder
        .set_password("secret", EncryptionProfile::OfficeBinaryRc4)
        .unwrap();
    builder.set_encryption_entropy(litchi_xls::EncryptionEntropy::generate().unwrap());
    let package = builder.build().unwrap();
    assert!(package.is_encrypted());
    assert!(matches!(
        package.to_bytes(),
        Err(litchi_xls::package::Error::EncryptionPolicy)
    ));
    let encrypted = package.to_encrypted_bytes().unwrap();
    let reopened = Package::from_bytes_with_options(
        &encrypted,
        litchi_xls::OpenOptions::new().password(Some("secret")),
    )
    .unwrap();
    assert!(reopened.is_encrypted());
    assert_eq!(reopened.to_encrypted_bytes().unwrap(), encrypted);
    assert!(reopened.workbook().text().unwrap().contains("protected"));
    builder.clear_password();
    let plaintext = builder.build().unwrap();
    assert!(!plaintext.is_encrypted());
    assert!(plaintext.to_bytes().is_ok());
    assert_eq!(package.to_encrypted_bytes().unwrap(), encrypted);
}

#[test]
fn build_refuses_pivot_page_entries_without_declared_fields() {
    let mut builder = Builder::new();
    let sheet = builder.add_worksheet("Invalid").unwrap();
    builder
        .add_pivot_table(
            sheet,
            litchi_xls::writer::PivotTableConfig {
                name: "Pivot".into(),
                source_type: 1,
                source_sheet_name: "Invalid".into(),
                source_first_row: 0,
                source_last_row: 0,
                source_first_col: 0,
                source_last_col: 0,
                first_row: 4,
                last_row: 4,
                first_col: 0,
                last_col: 0,
                first_header_row: 4,
                first_data_row: 4,
                first_data_col: 0,
                data_field_name: "Values".into(),
                data_axis: 0,
                data_position: 0,
                fields: Vec::new(),
                data_items: Vec::new(),
                page_entries: vec![(0, 0, 1)],
                source_data: Vec::new(),
            },
        )
        .unwrap();
    builder
        .set_password("secret", EncryptionProfile::OfficeBinaryRc4)
        .unwrap();
    builder.set_encryption_entropy(litchi_xls::EncryptionEntropy::generate().unwrap());
    assert!(matches!(
        builder.build(),
        Err(litchi_xls::package::Error::Reader(
            litchi_xls::Error::InvalidRecord {
                record_type: 0x00b6,
                ..
            }
        ))
    ));
    // Strict admission failed after encryption; the attempted build still consumed entropy.
    assert!(matches!(
        builder.build(),
        Err(litchi_xls::package::Error::Reader(
            litchi_xls::Error::EncryptionEntropyRequired
        ))
    ));
}
