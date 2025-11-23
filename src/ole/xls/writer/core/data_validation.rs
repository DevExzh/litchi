use crate::ole::xls::writer::formula::{Ptg, encode_ptg_tokens};
use crate::ole::xls::{XlsError, XlsResult};

/// Data validation operators for numeric constraints.
///
/// This maps directly to Excel's DV operator codes (0..7).
#[derive(Debug, Clone, Copy)]
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
    /// Integer ("whole number") constraint.
    Whole {
        operator: XlsDataValidationOperator,
        value1: i64,
        value2: Option<i64>,
    },
    /// Explicit list of allowed string values.
    List { values: Vec<String> },
}

impl XlsDataValidationType {
    /// Convert this validation type into BIFF8 DV payload components.
    ///
    /// Returns `(data_type, operator, is_explicit_list, formula1_bytes, formula2_bytes)`.
    pub(crate) fn to_biff_payload(
        &self,
    ) -> XlsResult<(u8, u8, bool, Option<Vec<u8>>, Option<Vec<u8>>)> {
        match self {
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

                Ok((data_type, op, false, formula1, formula2))
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

                if !joined.is_ascii() {
                    return Err(XlsError::Encoding(
                        "XLS data validation list values must be ASCII".to_string(),
                    ));
                }
                if joined.len() > 255 {
                    return Err(XlsError::InvalidData(
                        "XLS data validation list source exceeds 255 characters".to_string(),
                    ));
                }

                let tokens = vec![Ptg::PtgStr(joined)];
                let formula1 = Some(encode_ptg_tokens(&tokens));

                // LIST uses operator IGNORED (0) and marks explicit list formula.
                Ok((0x03, 0, true, formula1, None))
            },
        }
    }
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
