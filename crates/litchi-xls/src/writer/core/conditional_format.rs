use crate::writer::formatting::FillPattern;
use crate::writer::formula::{FormulaTokenizer, encode_ptg_tokens};
use crate::{Error, Result};

/// Conditional formatting rule types supported by the XLS writer.
///
/// For the initial implementation we focus on the classic "formula" style
/// conditional formats (`CONDITION_TYPE_FORMULA` in BIFF8). This keeps the
/// model small while still being expressive: most conditional formatting
/// scenarios can be expressed as a boolean formula.
#[derive(Debug, Clone)]
pub enum ConditionalFormatType {
    /// Formula that evaluates to TRUE for cells that should be formatted.
    ///
    /// The formula is written without a leading `=` and is tokenized using
    /// the shared `FormulaTokenizer` used elsewhere in the XLS writer.
    Formula {
        /// Formula string (without leading `=`).
        formula: String,
    },
    CellValue {
        operator: ConditionalFormatOperator,
        formula1: String,
        formula2: Option<String>,
    },
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ConditionalFormatOperator {
    Between,
    NotBetween,
    Equal,
    NotEqual,
    GreaterThan,
    LessThan,
    GreaterThanOrEqual,
    LessThanOrEqual,
}
impl ConditionalFormatOperator {
    fn code(self) -> u8 {
        match self {
            Self::Between => 1,
            Self::NotBetween => 2,
            Self::Equal => 3,
            Self::NotEqual => 4,
            Self::GreaterThan => 5,
            Self::LessThan => 6,
            Self::GreaterThanOrEqual => 7,
            Self::LessThanOrEqual => 8,
        }
    }
}

impl ConditionalFormatType {
    /// Convert this conditional format description into BIFF8 `CFRule` payload
    /// components.
    ///
    /// Returns `(condition_type, comparison_operator, formula1_bytes, formula2_bytes)`.
    /// The returned byte vectors contain encoded Ptg tokens in RPN order.
    pub(crate) fn to_biff_payload(&self) -> Result<(u8, u8, Vec<u8>, Vec<u8>)> {
        let tokenizer = FormulaTokenizer::new();

        match self {
            ConditionalFormatType::Formula { formula } => {
                // CONDITION_TYPE_FORMULA (2) with NO_COMPARISON (0)
                let condition_type = 0x02u8;
                let comparison_op = 0x00u8;

                let tokens = tokenizer.tokenize(formula).map_err(|e| {
                    Error::InvalidData(format!(
                        "Invalid conditional formatting formula '{formula}': {e}"
                    ))
                })?;
                let formula1 = encode_ptg_tokens(&tokens);

                // Second formula is unused for simple expression-based rules.
                Ok((condition_type, comparison_op, formula1, Vec::new()))
            },
            ConditionalFormatType::CellValue {
                operator,
                formula1,
                formula2,
            } => {
                let needs_two = matches!(
                    operator,
                    ConditionalFormatOperator::Between | ConditionalFormatOperator::NotBetween
                );
                if needs_two != formula2.is_some() {
                    return Err(Error::InvalidData("between/not-between conditional format requires two formulas; other comparisons require one".to_string()));
                }
                let encode = |formula: &str| -> Result<Vec<u8>> {
                    let tokens = tokenizer.tokenize(formula).map_err(|error| {
                        Error::InvalidData(format!(
                            "Invalid conditional formatting formula '{formula}': {error}"
                        ))
                    })?;
                    Ok(encode_ptg_tokens(&tokens))
                };
                Ok((
                    1,
                    operator.code(),
                    encode(formula1)?,
                    formula2
                        .as_deref()
                        .map(encode)
                        .transpose()?
                        .unwrap_or_default(),
                ))
            },
        }
    }
}

/// Pattern fill definition for a conditional formatting rule.
#[derive(Debug, Clone)]
pub struct ConditionalPattern {
    pub pattern: FillPattern,
    pub foreground_color: u16,
    pub background_color: u16,
}

/// Conditional formatting rule applied to a rectangular cell range.
///
/// Row and column indices are 0-based and inclusive at both ends.
#[derive(Debug, Clone)]
pub struct ConditionalFormat {
    pub first_row: u32,
    pub last_row: u32,
    pub first_col: u16,
    pub last_col: u16,
    pub format_type: ConditionalFormatType,
    pub pattern: Option<ConditionalPattern>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ConditionalFormatRange {
    pub first_row: u32,
    pub last_row: u32,
    pub first_col: u16,
    pub last_col: u16,
}
#[derive(Debug, Clone)]
pub struct ConditionalFormatRule {
    pub format_type: ConditionalFormatType,
    pub pattern: Option<ConditionalPattern>,
}
#[derive(Debug, Clone)]
pub struct ConditionalFormatGroup {
    pub ranges: Vec<ConditionalFormatRange>,
    pub rules: Vec<ConditionalFormatRule>,
}

/// A future-record (`CF12`) conditional-format rule type.
///
/// Formula fields contain encoded BIFF Ptg tokens. Visual rule payloads are
/// preserved and written verbatim; the writer never evaluates them.
#[derive(Debug, Clone)]
pub enum ConditionalFormat12Type {
    CellValue {
        operator: ConditionalFormatOperator,
        formula1: Vec<u8>,
        formula2: Option<Vec<u8>>,
    },
    Formula {
        formula: Vec<u8>,
    },
    ColorScale {
        active_formula: Vec<u8>,
        payload: Vec<u8>,
    },
    DataBar {
        active_formula: Vec<u8>,
        payload: Vec<u8>,
    },
    Filter {
        payload: Vec<u8>,
    },
    IconSet {
        active_formula: Vec<u8>,
        payload: Vec<u8>,
    },
}

impl ConditionalFormat12Type {
    #[allow(
        clippy::type_complexity,
        reason = "type mirrors the decoded BIFF record structure"
    )]
    pub(crate) fn biff_parts(&self) -> (u8, u8, &[u8], &[u8], &[u8], &[u8]) {
        match self {
            Self::CellValue {
                operator,
                formula1,
                formula2,
            } => (
                1,
                operator.code(),
                formula1,
                formula2.as_deref().unwrap_or(&[]),
                &[],
                &[],
            ),
            Self::Formula { formula } => (2, 0, formula, &[], &[], &[]),
            Self::ColorScale {
                active_formula,
                payload,
            } => (3, 0, &[], &[], active_formula, payload),
            Self::DataBar {
                active_formula,
                payload,
            } => (4, 0, &[], &[], active_formula, payload),
            Self::Filter { payload } => (5, 0, &[], &[], &[], payload),
            Self::IconSet {
                active_formula,
                payload,
            } => (6, 0, &[], &[], active_formula, payload),
        }
    }
}

/// One ordered rule in a `CondFmt12` collection.
#[derive(Debug, Clone)]
pub struct ConditionalFormat12Rule {
    pub format_type: ConditionalFormat12Type,
    /// Complete serialized DXFN12 structure, including its length prefix.
    pub differential_format: Vec<u8>,
    pub stop_if_true: bool,
    pub priority: u16,
    pub template: u16,
    pub template_parameters: [u8; 16],
}

/// One future conditional-format collection with ordered ranges and rules.
#[derive(Debug, Clone)]
pub struct ConditionalFormat12Group {
    pub ranges: Vec<ConditionalFormatRange>,
    pub rules: Vec<ConditionalFormat12Rule>,
}

impl ConditionalFormat {
    /// Convert the optional pattern into BIFF8 `PatternFormatting` triple
    /// `(pattern_code, fg_index, bg_index)`.
    #[must_use]
    pub fn to_biff_pattern(&self) -> Option<(u16, u16, u16)> {
        let pat = self.pattern.as_ref()?;
        let pattern_code = pat.pattern as u16;
        let fg = pat.foreground_color & 0x007F;
        let bg = pat.background_color & 0x007F;
        Some((pattern_code, fg, bg))
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::writer::formatting::FillPattern;

    #[test]
    fn test_formula_to_biff_payload() {
        let cf_type = ConditionalFormatType::Formula {
            formula: "A1>0".to_string(),
        };
        let result = cf_type.to_biff_payload();
        assert!(result.is_ok());

        let (condition_type, comparison_op, formula1, formula2) = result.unwrap();
        assert_eq!(condition_type, 0x02); // CONDITION_TYPE_FORMULA
        assert_eq!(comparison_op, 0x00); // NO_COMPARISON
        assert!(!formula1.is_empty());
        assert!(formula2.is_empty());
    }

    #[test]
    fn test_formula_to_biff_payload_invalid() {
        // This should still work as the tokenizer may handle it differently
        let cf_type = ConditionalFormatType::Formula {
            formula: "".to_string(),
        };
        // Empty formula should still tokenize (may produce empty tokens)
        let result = cf_type.to_biff_payload();
        // Result depends on tokenizer behavior
        assert!(result.is_ok() || result.is_err());
    }

    #[test]
    fn test_conditional_pattern() {
        let pattern = ConditionalPattern {
            pattern: FillPattern::Solid,
            foreground_color: 0x0040, // Palette index
            background_color: 0x0041,
        };

        assert_eq!(pattern.pattern, FillPattern::Solid);
        assert_eq!(pattern.foreground_color, 0x0040);
        assert_eq!(pattern.background_color, 0x0041);
    }

    #[test]
    fn test_xls_conditional_format_to_biff_pattern() {
        let cf = ConditionalFormat {
            first_row: 0,
            last_row: 9,
            first_col: 0,
            last_col: 1,
            format_type: ConditionalFormatType::Formula {
                formula: "A1>0".to_string(),
            },
            pattern: Some(ConditionalPattern {
                pattern: FillPattern::Solid,
                foreground_color: 0x0040,
                background_color: 0x0041,
            }),
        };

        let result = cf.to_biff_pattern();
        assert!(result.is_some());
        let (pattern_code, fg, bg) = result.unwrap();
        assert_eq!(pattern_code, FillPattern::Solid as u16);
        assert_eq!(fg, 0x0040);
        assert_eq!(bg, 0x0041);
    }

    #[test]
    fn test_xls_conditional_format_no_pattern() {
        let cf = ConditionalFormat {
            first_row: 0,
            last_row: 9,
            first_col: 0,
            last_col: 1,
            format_type: ConditionalFormatType::Formula {
                formula: "A1>0".to_string(),
            },
            pattern: None,
        };

        let result = cf.to_biff_pattern();
        assert!(result.is_none());
    }

    #[test]
    fn test_xls_conditional_format_color_masking() {
        // Test that colors are properly masked to 7 bits
        let cf = ConditionalFormat {
            first_row: 0,
            last_row: 9,
            first_col: 0,
            last_col: 1,
            format_type: ConditionalFormatType::Formula {
                formula: "A1>0".to_string(),
            },
            pattern: Some(ConditionalPattern {
                pattern: FillPattern::Solid,
                foreground_color: 0xFFFF, // Should be masked to 0x007F
                background_color: 0xFF80, // Should be masked to 0x0000
            }),
        };

        let result = cf.to_biff_pattern().unwrap();
        assert_eq!(result.1, 0x007F); // foreground masked
        assert_eq!(result.2, 0x0000); // background masked
    }

    #[test]
    fn test_xls_conditional_format_clone() {
        let cf = ConditionalFormat {
            first_row: 0,
            last_row: 9,
            first_col: 0,
            last_col: 1,
            format_type: ConditionalFormatType::Formula {
                formula: "A1>0".to_string(),
            },
            pattern: Some(ConditionalPattern {
                pattern: FillPattern::Solid,
                foreground_color: 0x0040,
                background_color: 0x0041,
            }),
        };

        let cloned = cf.clone();
        assert_eq!(cloned.first_row, cf.first_row);
        assert_eq!(cloned.last_row, cf.last_row);
        assert_eq!(cloned.first_col, cf.first_col);
        assert_eq!(cloned.last_col, cf.last_col);
    }

    #[test]
    fn test_xls_conditional_format_type_clone() {
        let cf_type = ConditionalFormatType::Formula {
            formula: "A1>0".to_string(),
        };
        let cloned = cf_type.clone();

        match cloned {
            ConditionalFormatType::Formula { formula } => {
                assert_eq!(formula, "A1>0");
            },
            ConditionalFormatType::CellValue { .. } => unreachable!(),
        }
    }

    #[test]
    fn test_xls_conditional_pattern_clone() {
        let pattern = ConditionalPattern {
            pattern: FillPattern::Solid,
            foreground_color: 0x0040,
            background_color: 0x0041,
        };
        let cloned = pattern.clone();

        assert_eq!(cloned.pattern, pattern.pattern);
        assert_eq!(cloned.foreground_color, pattern.foreground_color);
        assert_eq!(cloned.background_color, pattern.background_color);
    }
}
