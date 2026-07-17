use std::io::Cursor;

use litchi_ole::xls::writer::{
    XlsDataValidation, XlsDataValidationErrorStyle as WriterErrorStyle,
    XlsDataValidationFormulaKind, XlsDataValidationImeMode as WriterImeMode,
    XlsDataValidationOperator, XlsDataValidationOptions, XlsDataValidationRange,
    XlsDataValidationTableOptions, XlsDataValidationType, XlsWriter,
};
use litchi_ole::xls::{
    XlsDataValidationErrorStyle, XlsDataValidationImeMode, XlsDataValidationKind, XlsWorkbook,
};

fn validation(row: u32, validation_type: XlsDataValidationType) -> XlsDataValidation {
    XlsDataValidation {
        first_row: row,
        last_row: row,
        first_col: 0,
        last_col: 0,
        validation_type,
        show_input_message: true,
        input_title: Some("输入".to_string()),
        input_message: Some("Choose a valid value".to_string()),
        show_error_alert: true,
        error_title: Some("Invalid".to_string()),
        error_message: Some("The value is not accepted".to_string()),
    }
}

#[test]
fn complete_validation_family_round_trips_as_inert_metadata() {
    let mut writer = XlsWriter::new();
    let sheet = writer.add_worksheet("Validation").unwrap();
    writer
        .set_data_validation_table_options(
            sheet,
            XlsDataValidationTableOptions {
                window_closed: true,
                x_left: 120,
                y_top: 240,
                dropdown_object_id: Some(12),
            },
        )
        .unwrap();

    let custom = validation(
        4,
        XlsDataValidationType::Custom {
            formula_tokens: vec![0x1d, 1],
        },
    );
    writer
        .add_data_validation_with_options(
            sheet,
            custom,
            &[
                XlsDataValidationRange { first_row: 8, last_row: 9, first_col: 4, last_col: 5 },
                XlsDataValidationRange { first_row: 1, last_row: 2, first_col: 2, last_col: 3 },
                XlsDataValidationRange { first_row: 4, last_row: 4, first_col: 0, last_col: 0 },
            ],
            XlsDataValidationOptions {
                error_style: WriterErrorStyle::Warning,
                allow_blank: false,
                suppress_dropdown: true,
                ime_mode: WriterImeMode::Hiragana,
            },
        )
        .unwrap();
    writer
        .add_data_validation(
            sheet,
            validation(5, XlsDataValidationType::Any),
        )
        .unwrap();
    writer
        .add_data_validation(
            sheet,
            validation(
                6,
                XlsDataValidationType::Decimal {
                    operator: XlsDataValidationOperator::Between,
                    value1: 1.5,
                    value2: Some(9.5),
                },
            ),
        )
        .unwrap();
    writer
        .add_data_validation(
            sheet,
            validation(
                7,
                XlsDataValidationType::Date {
                    operator: XlsDataValidationOperator::GreaterThan,
                    value1: 45_000.0,
                    value2: None,
                },
            ),
        )
        .unwrap();
    writer
        .add_data_validation(
            sheet,
            validation(
                8,
                XlsDataValidationType::Time {
                    operator: XlsDataValidationOperator::LessThan,
                    value1: 0.5,
                    value2: None,
                },
            ),
        )
        .unwrap();
    writer
        .add_data_validation(
            sheet,
            validation(
                9,
                XlsDataValidationType::TextLength {
                    operator: XlsDataValidationOperator::LessThanOrEqual,
                    value1: 20,
                    value2: None,
                },
            ),
        )
        .unwrap();
    writer
        .add_data_validation(
            sheet,
            validation(
                10,
                XlsDataValidationType::List {
                    values: vec!["是".to_string(), "否".to_string()],
                },
            ),
        )
        .unwrap();
    writer
        .add_data_validation(
            sheet,
            validation(
                11,
                XlsDataValidationType::RawFormula {
                    kind: XlsDataValidationFormulaKind::Whole,
                    operator: XlsDataValidationOperator::Equal,
                    formula1_tokens: vec![0x1e, 42, 0],
                    formula2_tokens: None,
                },
            ),
        )
        .unwrap();

    let mut bytes = Cursor::new(Vec::new());
    writer.write_to(&mut bytes).unwrap();
    let workbook = XlsWorkbook::new(Cursor::new(bytes.into_inner())).unwrap();
    let sheet = workbook.xls_worksheet(0).unwrap();
    let settings = sheet.data_validation_settings().unwrap();
    assert!(settings.window_closed());
    assert_eq!(settings.x_left(), 120);
    assert_eq!(settings.y_top(), 240);
    assert_eq!(settings.dropdown_object_id(), Some(12));
    assert_eq!(settings.declared_rule_count(), 8);

    let rules = sheet.data_validations();
    assert_eq!(rules.len(), 8);
    assert_eq!(rules[0].kind(), XlsDataValidationKind::Custom);
    assert_eq!(rules[0].error_style(), XlsDataValidationErrorStyle::Warning);
    assert_eq!(rules[0].ime_mode(), XlsDataValidationImeMode::Hiragana);
    assert!(!rules[0].allow_blank());
    assert!(rules[0].suppress_dropdown());
    assert_eq!(rules[0].formula1().unwrap().tokens(), &[0x1d, 1]);
    assert_eq!(rules[0].ranges().len(), 4);
    assert_eq!(rules[0].ranges()[0].first_row(), 4);
    assert_eq!(rules[0].ranges()[1].first_row(), 8);
    assert_eq!(rules[0].ranges()[2].first_row(), 1);
    assert_eq!(rules[0].ranges()[3].first_row(), 4);
    assert_eq!(rules[1].kind(), XlsDataValidationKind::Any);
    assert_eq!(rules[2].kind(), XlsDataValidationKind::Decimal);
    assert_eq!(rules[3].kind(), XlsDataValidationKind::Date);
    assert_eq!(rules[4].kind(), XlsDataValidationKind::Time);
    assert_eq!(rules[5].kind(), XlsDataValidationKind::TextLength);
    assert_eq!(rules[6].kind(), XlsDataValidationKind::List);
    assert!(rules[6].explicit_list());
    assert_eq!(rules[7].kind(), XlsDataValidationKind::Whole);
    assert_eq!(rules[7].formula1().unwrap().tokens(), &[0x1e, 42, 0]);
}

#[test]
fn malformed_writer_metadata_is_rejected() {
    let mut writer = XlsWriter::new();
    let sheet = writer.add_worksheet("Validation").unwrap();
    assert!(writer
        .set_data_validation_table_options(
            sheet,
            XlsDataValidationTableOptions {
                x_left: 65_536,
                ..XlsDataValidationTableOptions::default()
            },
        )
        .is_err());
    assert!(writer
        .add_data_validation(
            sheet,
            validation(
                0,
                XlsDataValidationType::Custom {
                    formula_tokens: Vec::new(),
                },
            ),
        )
        .is_ok());
    let mut bytes = Cursor::new(Vec::new());
    assert!(writer.write_to(&mut bytes).is_err());
}

#[test]
fn explicit_zero_rule_dval_and_oversized_dv_are_handled_deterministically() {
    let mut writer = XlsWriter::new();
    let sheet = writer.add_worksheet("EmptyValidation").unwrap();
    writer.set_data_validation_table_options(sheet, XlsDataValidationTableOptions {
        window_closed: true, x_left: 7, y_top: 9, dropdown_object_id: None,
    }).unwrap();
    let mut bytes = Cursor::new(Vec::new());
    writer.write_to(&mut bytes).unwrap();
    let workbook = XlsWorkbook::new(Cursor::new(bytes.into_inner())).unwrap();
    let settings = workbook.xls_worksheet(0).unwrap().data_validation_settings().unwrap();
    assert_eq!(settings.declared_rule_count(), 0);
    assert_eq!((settings.x_left(), settings.y_top()), (7, 9));

    let mut writer = XlsWriter::new();
    let sheet = writer.add_worksheet("Oversized").unwrap();
    writer.add_data_validation(sheet, validation(0, XlsDataValidationType::Custom {
        formula_tokens: vec![0x1d; 8_200],
    })).unwrap();
    let mut bytes = Cursor::new(Vec::new());
    assert!(writer.write_to(&mut bytes).is_err());
}
