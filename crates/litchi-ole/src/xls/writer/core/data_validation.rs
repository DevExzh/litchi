use crate::xls::writer::formula::{Ptg, encode_ptg_tokens};
use crate::xls::{XlsError, XlsResult};

/// Data validation operators for numeric constraints.
///
/// This maps directly to Excel's DV operator codes (0..7).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum XlsDataValidationOperator {
    Between,
    NotBetween,
    Equal,
    NotEqual,
    GreaterThan,
    LessThan,
    GreaterThanOrEqual,
    LessThanOrEqual,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum XlsDataValidationErrorStyle { Stop, Warning, Information }

impl XlsDataValidationErrorStyle {
    pub(crate) fn to_biff_code(self) -> u8 {
        match self { Self::Stop => 0, Self::Warning => 1, Self::Information => 2 }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum XlsDataValidationImeMode {
    NoControl, On, Off, Hiragana, WideKatakana, NarrowKatakana,
    FullWidthAlphanumeric, HalfWidthAlphanumeric, FullWidthHangul, HalfWidthHangul,
}

impl XlsDataValidationImeMode {
    pub(crate) fn to_biff_code(self) -> u8 {
        match self {
            Self::NoControl => 0, Self::On => 1, Self::Off => 2, Self::Hiragana => 4,
            Self::WideKatakana => 5, Self::NarrowKatakana => 6,
            Self::FullWidthAlphanumeric => 7, Self::HalfWidthAlphanumeric => 8,
            Self::FullWidthHangul => 9, Self::HalfWidthHangul => 10,
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum XlsDataValidationFormulaKind { Whole, Decimal, Date, Time, TextLength }

impl XlsDataValidationFormulaKind {
    pub(crate) fn to_biff_code(self) -> u8 {
        match self { Self::Whole => 1, Self::Decimal => 2, Self::Date => 4, Self::Time => 5, Self::TextLength => 6 }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct XlsDataValidationRange {
    pub first_row: u32,
    pub last_row: u32,
    pub first_col: u16,
    pub last_col: u16,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct XlsDataValidationOptions {
    pub error_style: XlsDataValidationErrorStyle,
    pub allow_blank: bool,
    pub suppress_dropdown: bool,
    pub ime_mode: XlsDataValidationImeMode,
}

impl Default for XlsDataValidationOptions {
    fn default() -> Self {
        Self { error_style: XlsDataValidationErrorStyle::Stop, allow_blank: true,
            suppress_dropdown: false, ime_mode: XlsDataValidationImeMode::NoControl }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct XlsDataValidationTableOptions {
    pub window_closed: bool,
    pub x_left: u32,
    pub y_top: u32,
    pub dropdown_object_id: Option<u16>,
}

impl Default for XlsDataValidationTableOptions {
    fn default() -> Self {
        Self { window_closed: false, x_left: 0, y_top: 0, dropdown_object_id: None }
    }
}

impl XlsDataValidationOperator {
    pub(crate) fn to_biff_code(self) -> u8 {
        match self {
            Self::Between => 0,
            Self::NotBetween => 1,
            Self::Equal => 2,
            Self::NotEqual => 3,
            Self::GreaterThan => 4,
            Self::LessThan => 5,
            Self::GreaterThanOrEqual => 6,
            Self::LessThanOrEqual => 7,
        }
    }
}

/// Data validation kinds supported by the XLS writer.
///
/// The current implementation focuses on commonly used constraints: whole
/// numbers and explicit lists. Additional variants can be added in
/// a backward-compatible way in the future.
#[derive(Debug, Clone)]
pub enum XlsDataValidationType {
    Any,
    /// Integer ("whole number") constraint.
    Whole {
        operator: XlsDataValidationOperator,
        value1: i64,
        value2: Option<i64>,
    },
    /// Explicit list of allowed string values.
    List { values: Vec<String> },
    Decimal { operator: XlsDataValidationOperator, value1: f64, value2: Option<f64> },
    Date { operator: XlsDataValidationOperator, value1: f64, value2: Option<f64> },
    Time { operator: XlsDataValidationOperator, value1: f64, value2: Option<f64> },
    TextLength { operator: XlsDataValidationOperator, value1: i64, value2: Option<i64> },
    ListFormula { formula_tokens: Vec<u8> },
    Custom { formula_tokens: Vec<u8> },
    RawFormula {
        kind: XlsDataValidationFormulaKind,
        operator: XlsDataValidationOperator,
        formula1_tokens: Vec<u8>,
        formula2_tokens: Option<Vec<u8>>,
    },
}

/// BIFF8-encoded components of a single data validation rule.
#[derive(Debug, Clone)]
pub(crate) struct XlsDataValidationBiffPayload {
    pub data_type: u8,
    pub operator: u8,
    pub is_explicit_list: bool,
    pub formula1: Option<Vec<u8>>,
    pub formula2: Option<Vec<u8>>,
}

impl XlsDataValidationType {
    /// Convert this validation type into BIFF8 DV payload components.
    pub(crate) fn to_biff_payload(&self) -> XlsResult<XlsDataValidationBiffPayload> {
        match self {
            XlsDataValidationType::Any => Ok(XlsDataValidationBiffPayload {
                data_type: 0, operator: 0, is_explicit_list: false,
                formula1: None, formula2: None,
            }),
            XlsDataValidationType::Whole {
                operator,
                value1,
                value2,
            } => {
                let data_type = 0x01u8; // INTEGER
                let op = operator.to_biff_code();

                // Encode numeric bounds as simple PtgNum tokens.
                let f1_tokens = vec![Ptg::PtgNum(*value1 as f64)];
                let formula1 = Some(encode_ptg_tokens(&f1_tokens));

                let formula2 = if let Some(v2) = value2 {
                    let f2_tokens = vec![Ptg::PtgNum(*v2 as f64)];
                    Some(encode_ptg_tokens(&f2_tokens))
                } else {
                    // Between / NotBetween require a second bound.
                    match operator {
                        XlsDataValidationOperator::Between
                        | XlsDataValidationOperator::NotBetween => {
                            return Err(XlsError::InvalidData(
                                "Data validation: BETWEEN/NOT BETWEEN require a second bound"
                                    .to_string(),
                            ));
                        },
                        _ => None,
                    }
                };
                Ok(XlsDataValidationBiffPayload {
                    data_type,
                    operator: op,
                    is_explicit_list: false,
                    formula1,
                    formula2,
                })
            },
            XlsDataValidationType::List { values } => {
                if values.is_empty() {
                    return Err(XlsError::InvalidData(
                        "Data validation list must contain at least one value".to_string(),
                    ));
                }

                // Join values with NUL separators as POI does when encoding
                // explicit list validations.
                let joined = values.join("\u{0000}");

                if values.iter().any(|value| value.contains('\0')) {
                    return Err(XlsError::InvalidData(
                        "Data validation list values must not contain NUL".to_string(),
                    ));
                }
                if joined.encode_utf16().count() > 255 {
                    return Err(XlsError::InvalidData(
                        "XLS data validation list source exceeds 255 characters".to_string(),
                    ));
                }

                let tokens = vec![Ptg::PtgStr(joined)];
                let formula1 = Some(encode_ptg_tokens(&tokens));

                // LIST uses operator IGNORED (0) and marks explicit list formula.
                Ok(XlsDataValidationBiffPayload {
                    data_type: 0x03,
                    operator: 0,
                    is_explicit_list: true,
                    formula1,
                    formula2: None,
                })
            },
            XlsDataValidationType::Decimal { operator, value1, value2 }
            | XlsDataValidationType::Date { operator, value1, value2 }
            | XlsDataValidationType::Time { operator, value1, value2 } => {
                if !value1.is_finite() || value2.is_some_and(|value| !value.is_finite()) {
                    return Err(XlsError::InvalidData("Data validation value must be finite".to_string()));
                }
                let data_type = match self {
                    XlsDataValidationType::Decimal { .. } => 2,
                    XlsDataValidationType::Date { .. } => 4,
                    _ => 5,
                };
                numeric_payload(data_type, *operator, *value1, *value2)
            },
            XlsDataValidationType::TextLength { operator, value1, value2 } => {
                numeric_payload(6, *operator, *value1 as f64, value2.map(|value| value as f64))
            },
            XlsDataValidationType::ListFormula { formula_tokens } => {
                raw_payload(3, XlsDataValidationOperator::Equal, formula_tokens, None, false)
            },
            XlsDataValidationType::Custom { formula_tokens } => {
                raw_payload(7, XlsDataValidationOperator::Equal, formula_tokens, None, false)
            },
            XlsDataValidationType::RawFormula { kind, operator, formula1_tokens, formula2_tokens } => {
                raw_payload(kind.to_biff_code(), *operator, formula1_tokens,
                    formula2_tokens.as_deref(), true)
            },
        }
    }
}

fn numeric_payload(
    data_type: u8,
    operator: XlsDataValidationOperator,
    value1: f64,
    value2: Option<f64>,
) -> XlsResult<XlsDataValidationBiffPayload> {
    let formula1 = Some(encode_ptg_tokens(&[Ptg::PtgNum(value1)]));
    let formula2 = value2.map(|value| encode_ptg_tokens(&[Ptg::PtgNum(value)]));
    let needs_two = matches!(operator, XlsDataValidationOperator::Between | XlsDataValidationOperator::NotBetween);
    if needs_two != formula2.is_some() {
        return Err(XlsError::InvalidData(
            "Data validation formula count does not match its operator".to_string(),
        ));
    }
    Ok(XlsDataValidationBiffPayload { data_type, operator: operator.to_biff_code(),
        is_explicit_list: false, formula1, formula2 })
}

fn raw_payload(
    data_type: u8,
    operator: XlsDataValidationOperator,
    formula1: &[u8],
    formula2: Option<&[u8]>,
    operator_controls_count: bool,
) -> XlsResult<XlsDataValidationBiffPayload> {
    if formula1.is_empty() || formula1.len() > u16::MAX as usize
        || formula2.is_some_and(|tokens| tokens.is_empty() || tokens.len() > u16::MAX as usize)
    {
        return Err(XlsError::InvalidData("Data validation formula tokens have invalid length".to_string()));
    }
    if operator_controls_count {
        let needs_two = matches!(operator, XlsDataValidationOperator::Between | XlsDataValidationOperator::NotBetween);
        if needs_two != formula2.is_some() {
            return Err(XlsError::InvalidData("Data validation formula count does not match its operator".to_string()));
        }
    } else if formula2.is_some() {
        return Err(XlsError::InvalidData("This validation type accepts one formula".to_string()));
    }
    Ok(XlsDataValidationBiffPayload { data_type, operator: operator.to_biff_code(),
        is_explicit_list: false, formula1: Some(formula1.to_vec()),
        formula2: formula2.map(<[u8]>::to_vec) })
}

/// Data validation rule applied to a rectangular cell range in a worksheet.
///
/// Row and column indices are 0-based and inclusive at both ends, matching
/// the rest of the XLS writer APIs.
#[derive(Debug, Clone)]
pub struct XlsDataValidation {
    pub first_row: u32,
    pub last_row: u32,
    pub first_col: u16,
    pub last_col: u16,
    pub validation_type: XlsDataValidationType,
    pub show_input_message: bool,
    pub input_title: Option<String>,
    pub input_message: Option<String>,
    pub show_error_alert: bool,
    pub error_title: Option<String>,
    pub error_message: Option<String>,
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_xls_data_validation_operator_to_biff_code() {
        assert_eq!(XlsDataValidationOperator::Between.to_biff_code(), 0);
        assert_eq!(XlsDataValidationOperator::NotBetween.to_biff_code(), 1);
        assert_eq!(XlsDataValidationOperator::Equal.to_biff_code(), 2);
        assert_eq!(XlsDataValidationOperator::NotEqual.to_biff_code(), 3);
        assert_eq!(XlsDataValidationOperator::GreaterThan.to_biff_code(), 4);
        assert_eq!(XlsDataValidationOperator::LessThan.to_biff_code(), 5);
        assert_eq!(
            XlsDataValidationOperator::GreaterThanOrEqual.to_biff_code(),
            6
        );
        assert_eq!(XlsDataValidationOperator::LessThanOrEqual.to_biff_code(), 7);
    }

    #[test]
    fn test_whole_to_biff_payload_greater_than() {
        let validation = XlsDataValidationType::Whole {
            operator: XlsDataValidationOperator::GreaterThan,
            value1: 10,
            value2: None,
        };
        let payload = validation.to_biff_payload().unwrap();
        assert_eq!(payload.data_type, 0x01);
        assert_eq!(payload.operator, 4);
        assert!(!payload.is_explicit_list);
        assert!(payload.formula1.is_some());
        assert!(payload.formula2.is_none());
    }

    #[test]
    fn test_whole_to_biff_payload_between() {
        let validation = XlsDataValidationType::Whole {
            operator: XlsDataValidationOperator::Between,
            value1: 1,
            value2: Some(100),
        };
        let payload = validation.to_biff_payload().unwrap();
        assert_eq!(payload.data_type, 0x01);
        assert_eq!(payload.operator, 0);
        assert!(!payload.is_explicit_list);
        assert!(payload.formula1.is_some());
        assert!(payload.formula2.is_some());
    }

    #[test]
    fn test_whole_to_biff_payload_between_missing_value2() {
        let validation = XlsDataValidationType::Whole {
            operator: XlsDataValidationOperator::Between,
            value1: 1,
            value2: None,
        };
        let result = validation.to_biff_payload();
        assert!(result.is_err());
    }

    #[test]
    fn test_list_to_biff_payload() {
        let validation = XlsDataValidationType::List {
            values: vec!["Yes".to_string(), "No".to_string(), "Maybe".to_string()],
        };
        let payload = validation.to_biff_payload().unwrap();
        assert_eq!(payload.data_type, 0x03);
        assert_eq!(payload.operator, 0);
        assert!(payload.is_explicit_list);
        assert!(payload.formula1.is_some());
        assert!(payload.formula2.is_none());
    }

    #[test]
    fn test_list_to_biff_payload_empty() {
        let validation = XlsDataValidationType::List { values: vec![] };
        let result = validation.to_biff_payload();
        assert!(result.is_err());
    }

    #[test]
    fn test_list_to_biff_payload_non_ascii() {
        let validation = XlsDataValidationType::List {
            values: vec!["是".to_string(), "否".to_string()],
        };
        let payload = validation.to_biff_payload().unwrap();
        assert!(payload.formula1.is_some());
    }

    #[test]
    fn test_list_to_biff_payload_too_long() {
        let long_value = "a".repeat(256);
        let validation = XlsDataValidationType::List {
            values: vec![long_value],
        };
        let result = validation.to_biff_payload();
        assert!(result.is_err());
    }

    #[test]
    fn test_xls_data_validation_struct() {
        let dv = XlsDataValidation {
            first_row: 0,
            last_row: 9,
            first_col: 0,
            last_col: 1,
            validation_type: XlsDataValidationType::List {
                values: vec!["A".to_string(), "B".to_string()],
            },
            show_input_message: true,
            input_title: Some("Input".to_string()),
            input_message: Some("Choose A or B".to_string()),
            show_error_alert: true,
            error_title: Some("Error".to_string()),
            error_message: Some("Invalid choice".to_string()),
        };
        assert_eq!(dv.first_row, 0);
        assert_eq!(dv.last_row, 9);
        assert_eq!(dv.first_col, 0);
        assert_eq!(dv.last_col, 1);
        assert!(dv.show_input_message);
        assert!(dv.show_error_alert);
    }

    #[test]
    fn test_xls_data_validation_clone() {
        let dv = XlsDataValidation {
            first_row: 0,
            last_row: 9,
            first_col: 0,
            last_col: 1,
            validation_type: XlsDataValidationType::Whole {
                operator: XlsDataValidationOperator::GreaterThan,
                value1: 10,
                value2: None,
            },
            show_input_message: false,
            input_title: None,
            input_message: None,
            show_error_alert: true,
            error_title: None,
            error_message: None,
        };
        let cloned = dv.clone();
        assert_eq!(cloned.first_row, dv.first_row);
        assert_eq!(cloned.last_row, dv.last_row);
    }
}
