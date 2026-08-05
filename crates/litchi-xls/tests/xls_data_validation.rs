use std::io::Cursor;

use litchi_xls::writer::{
    DataValidation, DataValidationErrorStyle as WriterErrorStyle, DataValidationFormulaKind,
    DataValidationImeMode as WriterImeMode, DataValidationOperator, DataValidationOptions,
    DataValidationRange, DataValidationTableOptions, DataValidationType, Writer,
};
use litchi_xls::{DataValidationErrorStyle, DataValidationImeMode, DataValidationKind, Workbook};

fn validation(row: u32, validation_type: DataValidationType) -> DataValidation {
    DataValidation {
        range: DataValidationRange::new(row, row, 0, 0).unwrap(),
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
    let mut writer = Writer::new();
    let sheet = writer.add_worksheet("Validation").unwrap();
    writer
        .set_data_validation_table_options(
            sheet,
            DataValidationTableOptions {
                window_closed: true,
                x_left: 120,
                y_top: 240,
                dropdown_object_id: Some(12),
            },
        )
        .unwrap();

    let custom = validation(
        4,
        DataValidationType::Custom {
            formula_tokens: vec![0x1d, 1],
        },
    );
    writer
        .add_data_validation_with_options(
            sheet,
            custom,
            &[
                DataValidationRange::new(8, 9, 4, 5).unwrap(),
                DataValidationRange::new(1, 2, 2, 3).unwrap(),
                DataValidationRange::new(4, 4, 0, 0).unwrap(),
            ],
            DataValidationOptions {
                error_style: WriterErrorStyle::Warning,
                allow_blank: false,
                suppress_dropdown: true,
                ime_mode: WriterImeMode::Hiragana,
            },
        )
        .unwrap();
    writer
        .add_data_validation(sheet, validation(5, DataValidationType::Any))
        .unwrap();
    writer
        .add_data_validation(
            sheet,
            validation(
                6,
                DataValidationType::Decimal {
                    operator: DataValidationOperator::Between,
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
                DataValidationType::Date {
                    operator: DataValidationOperator::GreaterThan,
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
                DataValidationType::Time {
                    operator: DataValidationOperator::LessThan,
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
                DataValidationType::TextLength {
                    operator: DataValidationOperator::LessThanOrEqual,
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
                DataValidationType::List {
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
                DataValidationType::RawFormula {
                    kind: DataValidationFormulaKind::Whole,
                    operator: DataValidationOperator::Equal,
                    formula1_tokens: vec![0x1e, 42, 0],
                    formula2_tokens: None,
                },
            ),
        )
        .unwrap();

    let mut bytes = Cursor::new(Vec::new());
    writer.write_to(&mut bytes).unwrap();
    let workbook = Workbook::new(Cursor::new(bytes.into_inner())).unwrap();
    let sheet = workbook.xls_worksheet(0).unwrap();
    let settings = sheet.data_validation_settings().unwrap();
    assert!(settings.window_closed());
    assert_eq!(settings.x_left(), 120);
    assert_eq!(settings.y_top(), 240);
    assert_eq!(settings.dropdown_object_id(), Some(12));
    assert_eq!(settings.declared_rule_count(), 8);

    let rules = sheet.data_validations();
    assert_eq!(rules.len(), 8);
    assert_eq!(rules[0].kind(), DataValidationKind::Custom);
    assert_eq!(rules[0].error_style(), DataValidationErrorStyle::Warning);
    assert_eq!(rules[0].ime_mode(), DataValidationImeMode::Hiragana);
    assert!(!rules[0].allow_blank());
    assert!(rules[0].suppress_dropdown());
    assert_eq!(rules[0].formula1().unwrap().tokens(), &[0x1d, 1]);
    assert_eq!(rules[0].ranges().len(), 4);
    assert_eq!(rules[0].ranges()[0].first_row(), 4);
    assert_eq!(rules[0].ranges()[1].first_row(), 8);
    assert_eq!(rules[0].ranges()[2].first_row(), 1);
    assert_eq!(rules[0].ranges()[3].first_row(), 4);
    assert_eq!(rules[1].kind(), DataValidationKind::Any);
    assert_eq!(rules[2].kind(), DataValidationKind::Decimal);
    assert_eq!(rules[3].kind(), DataValidationKind::Date);
    assert_eq!(rules[4].kind(), DataValidationKind::Time);
    assert_eq!(rules[5].kind(), DataValidationKind::TextLength);
    assert_eq!(rules[6].kind(), DataValidationKind::List);
    assert!(rules[6].explicit_list());
    assert_eq!(rules[7].kind(), DataValidationKind::Whole);
    assert_eq!(rules[7].formula1().unwrap().tokens(), &[0x1e, 42, 0]);
}

#[test]
fn malformed_writer_metadata_is_rejected() {
    let mut writer = Writer::new();
    let sheet = writer.add_worksheet("Validation").unwrap();
    assert!(
        writer
            .set_data_validation_table_options(
                sheet,
                DataValidationTableOptions {
                    x_left: 65_536,
                    ..DataValidationTableOptions::default()
                },
            )
            .is_err()
    );
    assert!(
        writer
            .add_data_validation(
                sheet,
                validation(
                    0,
                    DataValidationType::Custom {
                        formula_tokens: Vec::new(),
                    },
                ),
            )
            .is_err()
    );
    let mut bytes = Cursor::new(Vec::new());
    writer.write_to(&mut bytes).unwrap();
    let workbook = Workbook::new(Cursor::new(bytes.into_inner())).unwrap();
    assert!(
        workbook
            .xls_worksheet(0)
            .unwrap()
            .data_validations()
            .is_empty()
    );
}

#[test]
fn explicit_zero_rule_dval_and_oversized_dv_are_handled_deterministically() {
    let mut writer = Writer::new();
    let sheet = writer.add_worksheet("EmptyValidation").unwrap();
    writer
        .set_data_validation_table_options(
            sheet,
            DataValidationTableOptions {
                window_closed: true,
                x_left: 7,
                y_top: 9,
                dropdown_object_id: None,
            },
        )
        .unwrap();
    let mut bytes = Cursor::new(Vec::new());
    writer.write_to(&mut bytes).unwrap();
    let workbook = Workbook::new(Cursor::new(bytes.into_inner())).unwrap();
    let settings = workbook
        .xls_worksheet(0)
        .unwrap()
        .data_validation_settings()
        .unwrap();
    assert_eq!(settings.declared_rule_count(), 0);
    assert_eq!((settings.x_left(), settings.y_top()), (7, 9));

    let mut writer = Writer::new();
    let sheet = writer.add_worksheet("Oversized").unwrap();
    writer
        .add_data_validation(
            sheet,
            validation(
                0,
                DataValidationType::Custom {
                    formula_tokens: vec![0x1d; 8_200],
                },
            ),
        )
        .unwrap();
    let mut bytes = Cursor::new(Vec::new());
    assert!(writer.write_to(&mut bytes).is_err());
}
