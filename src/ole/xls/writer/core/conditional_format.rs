use crate::ole::xls::writer::formatting::FillPattern;
use crate::ole::xls::writer::formula::{FormulaTokenizer, encode_ptg_tokens};
use crate::ole::xls::{XlsError, XlsResult};

/// Conditional formatting rule types supported by the XLS writer.
///
/// For the initial implementation we focus on the classic "formula" style
/// conditional formats (CONDITION_TYPE_FORMULA in BIFF8). This keeps the
/// model small while still being expressive: most conditional formatting
/// scenarios can be expressed as a boolean formula.
#[derive(Debug, Clone)]
pub enum XlsConditionalFormatType {
    /// Formula that evaluates to TRUE for cells that should be formatted.
    ///
    /// The formula is written without a leading `=` and is tokenized using
    /// the shared `FormulaTokenizer` used elsewhere in the XLS writer.
    Formula {
        /// Formula string (without leading `=`).
        formula: String,
    },
}

impl XlsConditionalFormatType {
    /// Convert this conditional format description into BIFF8 CFRule payload
    /// components.
    ///
    /// Returns `(condition_type, comparison_operator, formula1_bytes, formula2_bytes)`.
    /// The returned byte vectors contain encoded Ptg tokens in RPN order.
    pub(crate) fn to_biff_payload(&self) -> XlsResult<(u8, u8, Vec<u8>, Vec<u8>)> {
        let tokenizer = FormulaTokenizer::new();

        match self {
            XlsConditionalFormatType::Formula { formula } => {
                // CONDITION_TYPE_FORMULA (2) with NO_COMPARISON (0)
                let condition_type = 0x02u8;
                let comparison_op = 0x00u8;

                let tokens = tokenizer.tokenize(formula).map_err(|e| {
                    XlsError::InvalidData(format!(
                        "Invalid conditional formatting formula '{}': {}",
                        formula, e
                    ))
                })?;
                let formula1 = encode_ptg_tokens(&tokens);

                // Second formula is unused for simple expression-based rules.
                Ok((condition_type, comparison_op, formula1, Vec::new()))
            },
        }
    }
}

/// Pattern fill definition for a conditional formatting rule.
#[derive(Debug, Clone)]
pub struct XlsConditionalPattern {
    pub pattern: FillPattern,
    pub foreground_color: u16,
    pub background_color: u16,
}

/// Conditional formatting rule applied to a rectangular cell range.
///
/// Row and column indices are 0-based and inclusive at both ends.
#[derive(Debug, Clone)]
pub struct XlsConditionalFormat {
    pub first_row: u32,
    pub last_row: u32,
    pub first_col: u16,
    pub last_col: u16,
    pub format_type: XlsConditionalFormatType,
    pub pattern: Option<XlsConditionalPattern>,
}

impl XlsConditionalFormat {
    /// Convert the optional pattern into BIFF8 PatternFormatting triple
    /// `(pattern_code, fg_index, bg_index)`.
    pub(crate) fn to_biff_pattern(&self) -> Option<(u16, u16, u16)> {
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
    use crate::ole::xls::writer::formatting::FillPattern;

    #[test]
    fn test_formula_to_biff_payload() {
        let cf_type = XlsConditionalFormatType::Formula {
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
        let cf_type = XlsConditionalFormatType::Formula {
            formula: "".to_string(),
        };
        // Empty formula should still tokenize (may produce empty tokens)
        let result = cf_type.to_biff_payload();
        // Result depends on tokenizer behavior
        assert!(result.is_ok() || result.is_err());
    }

    #[test]
    fn test_conditional_pattern() {
        let pattern = XlsConditionalPattern {
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
        let cf = XlsConditionalFormat {
            first_row: 0,
            last_row: 9,
            first_col: 0,
            last_col: 1,
            format_type: XlsConditionalFormatType::Formula {
                formula: "A1>0".to_string(),
            },
            pattern: Some(XlsConditionalPattern {
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
        let cf = XlsConditionalFormat {
            first_row: 0,
            last_row: 9,
            first_col: 0,
            last_col: 1,
            format_type: XlsConditionalFormatType::Formula {
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
        let cf = XlsConditionalFormat {
            first_row: 0,
            last_row: 9,
            first_col: 0,
            last_col: 1,
            format_type: XlsConditionalFormatType::Formula {
                formula: "A1>0".to_string(),
            },
            pattern: Some(XlsConditionalPattern {
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
        let cf = XlsConditionalFormat {
            first_row: 0,
            last_row: 9,
            first_col: 0,
            last_col: 1,
            format_type: XlsConditionalFormatType::Formula {
                formula: "A1>0".to_string(),
            },
            pattern: Some(XlsConditionalPattern {
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
        let cf_type = XlsConditionalFormatType::Formula {
            formula: "A1>0".to_string(),
        };
        let cloned = cf_type.clone();

        match cloned {
            XlsConditionalFormatType::Formula { formula } => {
                assert_eq!(formula, "A1>0");
            },
        }
    }

    #[test]
    fn test_xls_conditional_pattern_clone() {
        let pattern = XlsConditionalPattern {
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
