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
