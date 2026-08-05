//! Regression coverage for the data-validation facade and bounded codecs.

use crate::formula::ParsedFormula;

use super::*;

#[test]
fn test_data_validation_new() {
    let dv = Validation::<ParsedFormula>::new(3, "A1:A10".to_string());
    assert_eq!(dv.validation_type, 3);
    assert_eq!(dv.cell_ranges, "A1:A10");
    // Check defaults
    assert_eq!(dv.operator, 0);
    assert!(dv.formula1.is_none());
    assert!(dv.formula2.is_none());
    assert!(dv.allow_blank);
    assert!(dv.show_dropdown);
    assert!(!dv.show_input_message);
    assert!(dv.show_error_message);
    assert_eq!(dv.error_style, 0);
    assert!(dv.input_title.is_none());
    assert!(dv.input_text.is_none());
    assert!(dv.error_title.is_none());
    assert!(dv.error_text.is_none());
}

#[test]
fn test_data_validation_whole_number() {
    let mut dv = Validation::<ParsedFormula>::new(1, "B1:B20".to_string()); // whole number
    dv.operator = 2; // greater than
    dv.formula1 = Some("10".to_string());
    dv.allow_blank = false;

    assert_eq!(dv.validation_type, 1);
    assert_eq!(dv.operator, 2);
    assert_eq!(dv.formula1, Some("10".to_string()));
    assert!(!dv.allow_blank);
}

#[test]
fn test_data_validation_decimal() {
    let mut dv = Validation::<ParsedFormula>::new(2, "C1:C10".to_string()); // decimal
    dv.operator = 0; // between
    dv.formula1 = Some("0".to_string());
    dv.formula2 = Some("100".to_string());

    assert_eq!(dv.validation_type, 2);
    assert_eq!(dv.operator, 0);
    assert_eq!(dv.formula1, Some("0".to_string()));
    assert_eq!(dv.formula2, Some("100".to_string()));
}

#[test]
fn test_data_validation_list() {
    let mut dv = Validation::<ParsedFormula>::new(3, "D1:D10".to_string()); // list
    dv.formula1 = Some("Yes,No,Maybe".to_string());
    dv.show_dropdown = true;

    assert_eq!(dv.validation_type, 3);
    assert_eq!(dv.formula1, Some("Yes,No,Maybe".to_string()));
    assert!(dv.show_dropdown);
}

#[test]
fn test_data_validation_date() {
    let mut dv = Validation::<ParsedFormula>::new(4, "E1:E10".to_string()); // date
    dv.operator = 4; // greater than
    dv.formula1 = Some("2024-01-01".to_string());

    assert_eq!(dv.validation_type, 4);
    assert_eq!(dv.operator, 4);
}

#[test]
fn test_data_validation_time() {
    let mut dv = Validation::<ParsedFormula>::new(5, "F1:F10".to_string()); // time
    dv.operator = 5; // less than
    dv.formula1 = Some("12:00".to_string());

    assert_eq!(dv.validation_type, 5);
    assert_eq!(dv.operator, 5);
}

#[test]
fn test_data_validation_text_length() {
    let mut dv = Validation::<ParsedFormula>::new(6, "G1:G10".to_string()); // text length
    dv.operator = 6; // greater than or equal
    dv.formula1 = Some("5".to_string());

    assert_eq!(dv.validation_type, 6);
    assert_eq!(dv.formula1, Some("5".to_string()));
}

#[test]
fn test_data_validation_custom() {
    let mut dv = Validation::<ParsedFormula>::new(7, "H1:H10".to_string()); // custom
    dv.formula1 = Some("=A1>0".to_string());

    assert_eq!(dv.validation_type, 7);
    assert_eq!(dv.formula1, Some("=A1>0".to_string()));
}

#[test]
fn test_data_validation_with_messages() {
    let mut dv = Validation::<ParsedFormula>::new(1, "I1:I10".to_string());
    dv.show_input_message = true;
    dv.input_title = Some("Enter value".to_string());
    dv.input_text = Some("Please enter a number greater than 10".to_string());
    dv.show_error_message = true;
    dv.error_style = 0; // stop
    dv.error_title = Some("Invalid input".to_string());
    dv.error_text = Some("The value must be greater than 10".to_string());

    assert!(dv.show_input_message);
    assert_eq!(dv.input_title, Some("Enter value".to_string()));
    assert_eq!(
        dv.input_text,
        Some("Please enter a number greater than 10".to_string())
    );
    assert!(dv.show_error_message);
    assert_eq!(dv.error_style, 0);
    assert_eq!(dv.error_title, Some("Invalid input".to_string()));
    assert_eq!(
        dv.error_text,
        Some("The value must be greater than 10".to_string())
    );
}

#[test]
fn test_data_validation_multiple_ranges() {
    let dv = Validation::<ParsedFormula>::new(3, "A1:A10,C1:C10,E1:E10".to_string());
    assert_eq!(dv.cell_ranges, "A1:A10,C1:C10,E1:E10");
}

#[test]
fn test_data_validation_clone() {
    let dv = Validation::<ParsedFormula>::new(3, "A1:A10".to_string());
    let cloned = dv.clone();
    assert_eq!(cloned.validation_type, dv.validation_type);
    assert_eq!(cloned.cell_ranges, dv.cell_ranges);
}

#[test]
fn parses_collection_settings_and_rejects_reserved_fields() {
    let mut data = Vec::new();
    data.extend_from_slice(&1u16.to_le_bytes());
    data.extend_from_slice(&120u32.to_le_bytes());
    data.extend_from_slice(&240u32.to_le_bytes());
    data.extend_from_slice(&0u32.to_le_bytes());
    data.extend_from_slice(&3u32.to_le_bytes());
    let (settings, count) = parse_collection_settings(&data, false).unwrap();
    assert_eq!(settings.prompt_x, 120);
    assert_eq!(settings.prompt_y, 240);
    assert!(settings.input_prompts_disabled);
    assert_eq!(count, 3);

    data[0] = 2;
    assert!(parse_collection_settings(&data, false).is_err());
}

#[test]
fn validates_dval_list_quotes_and_xml_characters() {
    assert!(validate_dval_list_formula("One,\"Two,Three\",Four").is_ok());
    assert!(validate_dval_list_formula("One,\"Two").is_err());
    assert!(validate_dval_list_formula("One\0Two").is_err());
}
