use crate::writer::formula::{Ptg, encode_ptg_tokens};
use crate::{Error, Result};

/// Data validation operators for numeric constraints.
///
/// This maps directly to Excel's DV operator codes (0..7).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DataValidationOperator {
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
pub enum DataValidationErrorStyle {
    Stop,
    Warning,
    Information,
}

impl DataValidationErrorStyle {
    pub(crate) fn to_biff_code(self) -> u8 {
        match self {
            Self::Stop => 0,
            Self::Warning => 1,
            Self::Information => 2,
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DataValidationImeMode {
    NoControl,
    On,
    Off,
    Hiragana,
    WideKatakana,
    NarrowKatakana,
    FullWidthAlphanumeric,
    HalfWidthAlphanumeric,
    FullWidthHangul,
    HalfWidthHangul,
}

impl DataValidationImeMode {
    pub(crate) fn to_biff_code(self) -> u8 {
        match self {
            Self::NoControl => 0,
            Self::On => 1,
            Self::Off => 2,
            Self::Hiragana => 4,
            Self::WideKatakana => 5,
            Self::NarrowKatakana => 6,
            Self::FullWidthAlphanumeric => 7,
            Self::HalfWidthAlphanumeric => 8,
            Self::FullWidthHangul => 9,
            Self::HalfWidthHangul => 10,
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DataValidationFormulaKind {
    Whole,
    Decimal,
    Date,
    Time,
    TextLength,
}

impl DataValidationFormulaKind {
    pub(crate) fn to_biff_code(self) -> u8 {
        match self {
            Self::Whole => 1,
            Self::Decimal => 2,
            Self::Date => 4,
            Self::Time => 5,
            Self::TextLength => 6,
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DataValidationRange {
    first_row: u16,
    last_row: u16,
    first_col: u8,
    last_col: u8,
}

impl DataValidationRange {
    /// Create a checked, inclusive BIFF8 cell range from zero-based indices.
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn new(first_row: u32, last_row: u32, first_col: u16, last_col: u16) -> Result<Self> {
        let invalid = || {
            Error::InvalidCellReference(format!(
                "range ({first_row}, {first_col})..=({last_row}, {last_col}) is outside the BIFF8 grid"
            ))
        };
        let first_row = u16::try_from(first_row).map_err(|_error| invalid())?;
        let last_row = u16::try_from(last_row).map_err(|_error| invalid())?;
        let first_col = u8::try_from(first_col).map_err(|_error| invalid())?;
        let last_col = u8::try_from(last_col).map_err(|_error| invalid())?;
        if first_row > last_row || first_col > last_col {
            return Err(invalid());
        }
        Ok(Self {
            first_row,
            last_row,
            first_col,
            last_col,
        })
    }

    #[must_use]
    pub const fn first_row(self) -> u16 {
        self.first_row
    }

    #[must_use]
    pub const fn last_row(self) -> u16 {
        self.last_row
    }

    #[must_use]
    pub const fn first_col(self) -> u8 {
        self.first_col
    }

    #[must_use]
    pub const fn last_col(self) -> u8 {
        self.last_col
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DataValidationOptions {
    pub error_style: DataValidationErrorStyle,
    pub allow_blank: bool,
    pub suppress_dropdown: bool,
    pub ime_mode: DataValidationImeMode,
}

impl Default for DataValidationOptions {
    fn default() -> Self {
        Self {
            error_style: DataValidationErrorStyle::Stop,
            allow_blank: true,
            suppress_dropdown: false,
            ime_mode: DataValidationImeMode::NoControl,
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct DataValidationTableOptions {
    pub window_closed: bool,
    pub x_left: u32,
    pub y_top: u32,
    pub dropdown_object_id: Option<u16>,
}

impl DataValidationOperator {
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
pub enum DataValidationType {
    Any,
    /// Integer ("whole number") constraint.
    Whole {
        operator: DataValidationOperator,
        value1: i64,
        value2: Option<i64>,
    },
    /// Explicit list of allowed string values.
    List {
        values: Vec<String>,
    },
    Decimal {
        operator: DataValidationOperator,
        value1: f64,
        value2: Option<f64>,
    },
    Date {
        operator: DataValidationOperator,
        value1: f64,
        value2: Option<f64>,
    },
    Time {
        operator: DataValidationOperator,
        value1: f64,
        value2: Option<f64>,
    },
    TextLength {
        operator: DataValidationOperator,
        value1: i64,
        value2: Option<i64>,
    },
    ListFormula {
        formula_tokens: Vec<u8>,
    },
    Custom {
        formula_tokens: Vec<u8>,
    },
    RawFormula {
        kind: DataValidationFormulaKind,
        operator: DataValidationOperator,
        formula1_tokens: Vec<u8>,
        formula2_tokens: Option<Vec<u8>>,
    },
}

/// BIFF8-encoded components of a single data validation rule.
#[derive(Debug, Clone)]
pub(crate) struct DataValidationBiffPayload {
    pub data_type: u8,
    pub operator: u8,
    pub is_explicit_list: bool,
    pub formula1: Option<Vec<u8>>,
    pub formula2: Option<Vec<u8>>,
}

impl DataValidationType {
    /// Convert this validation type into BIFF8 DV payload components.
    pub(crate) fn to_biff_payload(&self) -> Result<DataValidationBiffPayload> {
        match self {
            DataValidationType::Any => Ok(DataValidationBiffPayload {
                data_type: 0,
                operator: 0,
                is_explicit_list: false,
                formula1: None,
                formula2: None,
            }),
            DataValidationType::Whole {
                operator,
                value1,
                value2,
            } => {
                let data_type = 0x01u8; // INTEGER
                let op = operator.to_biff_code();

                // Encode numeric bounds as simple PtgNum tokens.
                let f1_tokens = vec![Ptg::Num(crate::utils::approximate_i64_as_f64(*value1))];
                let formula1 = Some(encode_ptg_tokens(&f1_tokens));

                let formula2 = if let Some(v2) = value2 {
                    let f2_tokens = vec![Ptg::Num(crate::utils::approximate_i64_as_f64(*v2))];
                    Some(encode_ptg_tokens(&f2_tokens))
                } else {
                    // Between / NotBetween require a second bound.
                    match operator {
                        DataValidationOperator::Between | DataValidationOperator::NotBetween => {
                            return Err(Error::InvalidData(
                                "Data validation: BETWEEN/NOT BETWEEN require a second bound"
                                    .to_string(),
                            ));
                        },
                        _ => None,
                    }
                };
                Ok(DataValidationBiffPayload {
                    data_type,
                    operator: op,
                    is_explicit_list: false,
                    formula1,
                    formula2,
                })
            },
            DataValidationType::List { values } => {
                if values.is_empty() {
                    return Err(Error::InvalidData(
                        "Data validation list must contain at least one value".to_string(),
                    ));
                }

                // Join values with NUL separators as POI does when encoding
                // explicit list validations.
                let joined = values.join("\u{0000}");

                if values.iter().any(|value| value.contains('\0')) {
                    return Err(Error::InvalidData(
                        "Data validation list values must not contain NUL".to_string(),
                    ));
                }
                if joined.encode_utf16().count() > 255 {
                    return Err(Error::InvalidData(
                        "XLS data validation list source exceeds 255 characters".to_string(),
                    ));
                }

                let tokens = vec![Ptg::Str(joined)];
                let formula1 = Some(encode_ptg_tokens(&tokens));

                // LIST uses operator IGNORED (0) and marks explicit list formula.
                Ok(DataValidationBiffPayload {
                    data_type: 0x03,
                    operator: 0,
                    is_explicit_list: true,
                    formula1,
                    formula2: None,
                })
            },
            DataValidationType::Decimal {
                operator,
                value1,
                value2,
            }
            | DataValidationType::Date {
                operator,
                value1,
                value2,
            }
            | DataValidationType::Time {
                operator,
                value1,
                value2,
            } => {
                if !value1.is_finite() || value2.is_some_and(|value| !value.is_finite()) {
                    return Err(Error::InvalidData(
                        "Data validation value must be finite".to_string(),
                    ));
                }
                let data_type = match self {
                    DataValidationType::Decimal { .. } => 2,
                    DataValidationType::Date { .. } => 4,
                    _ => 5,
                };
                numeric_payload(data_type, *operator, *value1, *value2)
            },
            DataValidationType::TextLength {
                operator,
                value1,
                value2,
            } => numeric_payload(
                6,
                *operator,
                crate::utils::approximate_i64_as_f64(*value1),
                value2.map(crate::utils::approximate_i64_as_f64),
            ),
            DataValidationType::ListFormula { formula_tokens } => raw_payload(
                3,
                DataValidationOperator::Equal,
                formula_tokens,
                None,
                false,
            ),
            DataValidationType::Custom { formula_tokens } => raw_payload(
                7,
                DataValidationOperator::Equal,
                formula_tokens,
                None,
                false,
            ),
            DataValidationType::RawFormula {
                kind,
                operator,
                formula1_tokens,
                formula2_tokens,
            } => raw_payload(
                kind.to_biff_code(),
                *operator,
                formula1_tokens,
                formula2_tokens.as_deref(),
                true,
            ),
        }
    }
}

fn numeric_payload(
    data_type: u8,
    operator: DataValidationOperator,
    value1: f64,
    value2: Option<f64>,
) -> Result<DataValidationBiffPayload> {
    let formula1 = Some(encode_ptg_tokens(&[Ptg::Num(value1)]));
    let formula2 = value2.map(|value| encode_ptg_tokens(&[Ptg::Num(value)]));
    let needs_two = matches!(
        operator,
        DataValidationOperator::Between | DataValidationOperator::NotBetween
    );
    if needs_two != formula2.is_some() {
        return Err(Error::InvalidData(
            "Data validation formula count does not match its operator".to_string(),
        ));
    }
    Ok(DataValidationBiffPayload {
        data_type,
        operator: operator.to_biff_code(),
        is_explicit_list: false,
        formula1,
        formula2,
    })
}

fn raw_payload(
    data_type: u8,
    operator: DataValidationOperator,
    formula1: &[u8],
    formula2: Option<&[u8]>,
    operator_controls_count: bool,
) -> Result<DataValidationBiffPayload> {
    if formula1.is_empty()
        || formula1.len() > u16::MAX as usize
        || formula2.is_some_and(|tokens| tokens.is_empty() || tokens.len() > u16::MAX as usize)
    {
        return Err(Error::InvalidData(
            "Data validation formula tokens have invalid length".to_string(),
        ));
    }
    if operator_controls_count {
        let needs_two = matches!(
            operator,
            DataValidationOperator::Between | DataValidationOperator::NotBetween
        );
        if needs_two != formula2.is_some() {
            return Err(Error::InvalidData(
                "Data validation formula count does not match its operator".to_string(),
            ));
        }
    } else if formula2.is_some() {
        return Err(Error::InvalidData(
            "This validation type accepts one formula".to_string(),
        ));
    }
    Ok(DataValidationBiffPayload {
        data_type,
        operator: operator.to_biff_code(),
        is_explicit_list: false,
        formula1: Some(formula1.to_vec()),
        formula2: formula2.map(<[u8]>::to_vec),
    })
}

/// Data validation rule applied to a rectangular cell range in a worksheet.
///
/// Row and column indices are 0-based and inclusive at both ends, matching
/// the rest of the XLS writer APIs.
#[derive(Debug, Clone)]
pub struct DataValidation {
    pub range: DataValidationRange,
    pub validation_type: DataValidationType,
    pub show_input_message: bool,
    pub input_title: Option<String>,
    pub input_message: Option<String>,
    pub show_error_alert: bool,
    pub error_title: Option<String>,
    pub error_message: Option<String>,
}

impl DataValidation {
    /// Create a validation rule with messages and alerts disabled by default.
    #[must_use]
    pub fn new(range: DataValidationRange, validation_type: DataValidationType) -> Self {
        Self {
            range,
            validation_type,
            show_input_message: false,
            input_title: None,
            input_message: None,
            show_error_alert: false,
            error_title: None,
            error_message: None,
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_xls_data_validation_operator_to_biff_code() {
        assert_eq!(DataValidationOperator::Between.to_biff_code(), 0);
        assert_eq!(DataValidationOperator::NotBetween.to_biff_code(), 1);
        assert_eq!(DataValidationOperator::Equal.to_biff_code(), 2);
        assert_eq!(DataValidationOperator::NotEqual.to_biff_code(), 3);
        assert_eq!(DataValidationOperator::GreaterThan.to_biff_code(), 4);
        assert_eq!(DataValidationOperator::LessThan.to_biff_code(), 5);
        assert_eq!(DataValidationOperator::GreaterThanOrEqual.to_biff_code(), 6);
        assert_eq!(DataValidationOperator::LessThanOrEqual.to_biff_code(), 7);
    }

    #[test]
    fn test_whole_to_biff_payload_greater_than() {
        let validation = DataValidationType::Whole {
            operator: DataValidationOperator::GreaterThan,
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
        let validation = DataValidationType::Whole {
            operator: DataValidationOperator::Between,
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
        let validation = DataValidationType::Whole {
            operator: DataValidationOperator::Between,
            value1: 1,
            value2: None,
        };
        let result = validation.to_biff_payload();
        assert!(result.is_err());
    }

    #[test]
    fn test_list_to_biff_payload() {
        let validation = DataValidationType::List {
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
        let validation = DataValidationType::List { values: vec![] };
        let result = validation.to_biff_payload();
        assert!(result.is_err());
    }

    #[test]
    fn test_list_to_biff_payload_non_ascii() {
        let validation = DataValidationType::List {
            values: vec!["是".to_string(), "否".to_string()],
        };
        let payload = validation.to_biff_payload().unwrap();
        assert!(payload.formula1.is_some());
    }

    #[test]
    fn test_list_to_biff_payload_too_long() {
        let long_value = "a".repeat(256);
        let validation = DataValidationType::List {
            values: vec![long_value],
        };
        let result = validation.to_biff_payload();
        assert!(result.is_err());
    }

    #[test]
    fn test_xls_data_validation_struct() {
        let dv = DataValidation {
            range: DataValidationRange::new(0, 9, 0, 1).unwrap(),
            validation_type: DataValidationType::List {
                values: vec!["A".to_string(), "B".to_string()],
            },
            show_input_message: true,
            input_title: Some("Input".to_string()),
            input_message: Some("Choose A or B".to_string()),
            show_error_alert: true,
            error_title: Some("Error".to_string()),
            error_message: Some("Invalid choice".to_string()),
        };
        assert_eq!(dv.range.first_row(), 0);
        assert_eq!(dv.range.last_row(), 9);
        assert_eq!(dv.range.first_col(), 0);
        assert_eq!(dv.range.last_col(), 1);
        assert!(dv.show_input_message);
        assert!(dv.show_error_alert);
    }

    #[test]
    fn test_xls_data_validation_clone() {
        let dv = DataValidation {
            range: DataValidationRange::new(0, 9, 0, 1).unwrap(),
            validation_type: DataValidationType::Whole {
                operator: DataValidationOperator::GreaterThan,
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
        assert_eq!(cloned.range, dv.range);
    }
}
