//! Direct reader checks for the unmodified native Excel corpus.

use litchi_core::sheet::WorkbookTrait;
use litchi_xls::raw::Reader;

#[test]
fn native_excel_cryptoapi_password_contract_retains_profile_ciphertext_and_cells() {
    use litchi_core::sheet::CellValue;
    use litchi_xls::package::Error;
    use litchi_xls::{EncryptionProfile, OpenOptions, Package};
    let bytes = include_bytes!("../../../test-data/microsoft-office/excel/excel-encrypted.xls");
    assert!(matches!(
        Package::from_bytes(bytes),
        Err(Error::Reader(litchi_xls::Error::PasswordRequired))
    ));
    assert!(matches!(
        Package::from_bytes_with_options(bytes, OpenOptions::new().password("wrong")),
        Err(Error::Reader(litchi_xls::Error::InvalidPassword))
    ));
    let package =
        Package::from_bytes_with_options(bytes, OpenOptions::new().password("OfficeFixture42!"))
            .unwrap();
    assert_eq!(
        package.encryption(),
        Some(EncryptionProfile::CryptoApiRc4 { key_bits: 128 })
    );
    let workbook = package.workbook();
    assert_eq!(workbook.worksheet_names().unwrap(), ["Sheet1"]);
    let sheet = workbook.worksheet(0usize).unwrap().unwrap();
    assert_eq!(
        sheet
            .cell_value(litchi_xls::cell_values::Reference::new(0, 0).unwrap())
            .unwrap(),
        Some(CellValue::String("EXCEL_ENCRYPTED_011".into()))
    );
    assert_eq!(
        sheet
            .cell_value(litchi_xls::cell_values::Reference::new(1, 1).unwrap())
            .unwrap(),
        Some(CellValue::Float(42.0))
    );
    assert_eq!(
        sheet
            .cell_value(litchi_xls::cell_values::Reference::new(2, 1).unwrap())
            .unwrap(),
        Some(CellValue::Float(84.0))
    );
    assert_eq!(package.source_bytes().unwrap(), bytes);
    assert!(matches!(package.to_bytes(), Err(Error::EncryptionPolicy)));
    let encrypted = package.to_encrypted_bytes().unwrap();
    assert!(encrypted == bytes);
    let reopened = Package::from_bytes_with_options(
        &encrypted,
        OpenOptions::new().password("OfficeFixture42!"),
    )
    .unwrap();
    assert_eq!(reopened.encryption(), package.encryption());
    assert_eq!(
        reopened
            .workbook()
            .worksheet(0usize)
            .unwrap()
            .unwrap()
            .cell_value(litchi_xls::cell_values::Reference::new(2, 1).unwrap())
            .unwrap(),
        Some(CellValue::Float(84.0))
    );
}

#[test]
fn native_excel_metadata_reader_keeps_every_worksheet() {
    let bytes = include_bytes!("../../../test-data/microsoft-office/excel/excel-metadata.xls");
    let reader = Reader::new(std::io::Cursor::new(bytes.as_slice())).unwrap();
    assert_eq!(
        reader.worksheet_names(),
        [
            "Sheet1",
            "Source_A_015",
            "Source_B_015",
            "Consolidated_015",
            "Print_015"
        ]
    );
    for index in 0..5 {
        reader.worksheet_by_index(index).unwrap();
    }
    let consolidation = reader.xls_worksheet(3).unwrap().consolidation().unwrap();
    assert_eq!(
        consolidation.function(),
        litchi_xls::ConsolidationFunction::Sum
    );
    assert_eq!(consolidation.sources().len(), 2);
    for (source, expected) in consolidation
        .sources()
        .iter()
        .zip(["\u{2}Source_A_015", "\u{2}Source_B_015"])
    {
        let litchi_xls::ConsolidationSource::CellRange { range, file } = source else {
            panic!("expected native cell-range source")
        };
        assert_eq!(file.encoded_path(), expected);
        assert_eq!(
            (
                range.first_row(),
                range.last_row(),
                range.first_column(),
                range.last_column()
            ),
            (0, 2, 0, 1)
        );
    }
    assert_eq!(
        reader
            .xls_worksheet(4)
            .unwrap()
            .page_setup()
            .unwrap()
            .printer_driver_data()[0]
            .len(),
        28
    );
}

#[test]
fn native_excel_workbooks_never_silently_lose_worksheet_substreams() {
    let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/microsoft-office/excel");
    let mut failures = Vec::new();
    for (name, count) in [
        ("excel-values-formulas-chart.xls", 2),
        ("excel-interaction-print.xls", 1),
        ("excel-recorded-macro.xls", 1),
        ("excel-recorded-macro.xlt", 1),
        ("excel-visual-analytics.xls", 3),
        ("excel-conditional-drawings.xls", 1),
        ("excel-links-objects.xls", 1),
        ("excel-pivot-controls.xls", 3),
        ("excel-form-controls.xls", 1),
        ("excel-formula-owners.xls", 2),
        ("excel-connections.xls", 2),
        ("excel-date1904-breaks.xls", 1),
        ("excel-chart-families.xls", 18),
        ("excel-revisions.xls", 2),
        ("excel-metadata.xls", 5),
        ("excel-legacy-sheets.xls", 1),
    ] {
        match Reader::new(std::fs::File::open(root.join(name)).unwrap()) {
            Ok(reader) => {
                if reader.worksheet_names().len() != count {
                    failures.push(format!(
                        "{name}: expected {count} worksheets; got {:?}",
                        reader.worksheet_names()
                    ));
                }
                if matches!(
                    name,
                    "excel-visual-analytics.xls" | "excel-pivot-controls.xls"
                ) {
                    let sort = (0..count)
                        .find_map(|index| reader.xls_worksheet(index).unwrap().sort())
                        .unwrap();
                    assert_eq!(
                        sort.range(),
                        litchi_xls::writer::sort::Range::new(0..=7, 0..=5).unwrap()
                    );
                    assert_eq!(
                        sort.keys()[0].range(),
                        litchi_xls::writer::sort::Range::new(0..=8, 3..=3).unwrap()
                    );
                }
                if name == "excel-connections.xls" {
                    let tables = reader.xls_worksheet(0).unwrap().list_objects();
                    assert_eq!(tables.len(), 1);
                    assert_eq!(
                        tables[0].external_metadata().unwrap().version(),
                        litchi_xls::ExternalTableVersion::Version15
                    );
                }
            },
            Err(error) => failures.push(format!("{name}: {error:?}")),
        }
    }
    assert!(failures.is_empty(), "{}", failures.join("\n"));
}
