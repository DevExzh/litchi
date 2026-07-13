//! Data validation support for XLSB.

use crate::xlsb::error::{XlsbError, XlsbResult};
use crate::xlsb::formula::{
    CellParsedFormula, FormulaConverter, FormulaParser, FormulaResolutionContext,
    MAX_CELL_FORMULA_BYTES,
};
use crate::xlsb::utils::cell_reference;
use litchi_core::binary;

/// Worksheet-level UI settings stored by a data-validation collection.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct DataValidationSettings {
    /// Whether every validation input prompt is disabled on the sheet.
    pub input_prompts_disabled: bool,
    /// Horizontal prompt-window position in pixels.
    pub prompt_x: u16,
    /// Vertical prompt-window position in pixels.
    pub prompt_y: u16,
}

/// Binary record family used to store a validation rule.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub enum DataValidationRecordKind {
    /// `BrtDVal`, limited to 8,191 target ranges.
    #[default]
    Classic,
    /// Office 2013 `BrtDVal14`, which permits more target ranges.
    Extension14,
}

/// Data validation rule
///
/// Represents data validation constraints on a cell or range.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DataValidation {
    /// Type of validation (0=none, 1=whole, 2=decimal, 3=list, 4=date, 5=time, 6=text length, 7=custom)
    pub validation_type: u8,
    /// Operator (0=between, 1=not between, 2=equal, 3=not equal, 4=greater than, 5=less than, 6=greater or equal, 7=less or equal)
    pub operator: u8,
    /// First formula (constraint)
    pub formula1: Option<String>,
    /// Second formula (for between/not between)
    pub formula2: Option<String>,
    /// Original first binary formula, retained for lossless round-tripping.
    pub formula1_binary: Option<CellParsedFormula>,
    /// Original second binary formula, retained for lossless round-tripping.
    pub formula2_binary: Option<CellParsedFormula>,
    /// Allow blank cells
    pub allow_blank: bool,
    /// Show dropdown (for list validation)
    pub show_dropdown: bool,
    /// Show input message
    pub show_input_message: bool,
    /// Show error message
    pub show_error_message: bool,
    /// Error style (0=stop, 1=warning, 2=information)
    pub error_style: u8,
    /// Input Method Editor mode (0=no control, 1..=10 are XLSB IME modes).
    pub ime_mode: u8,
    /// Compatibility bit used by Excel and LibreOffice for inline list strings.
    pub string_list: bool,
    /// Optional `BrtDValList` text which overrides the first binary formula.
    pub list_formula: Option<String>,
    /// Input message title
    pub input_title: Option<String>,
    /// Input message text
    pub input_text: Option<String>,
    /// Error message title
    pub error_title: Option<String>,
    /// Error message text
    pub error_text: Option<String>,
    /// Cell ranges (e.g., "A1:B2,C3:D4")
    pub cell_ranges: String,
    /// Binary record family from which this rule was read or should be written.
    pub record_kind: DataValidationRecordKind,
}

impl DataValidation {
    /// Create a new data validation rule
    pub fn new(validation_type: u8, cell_ranges: String) -> Self {
        DataValidation {
            validation_type,
            operator: 0,
            formula1: None,
            formula2: None,
            formula1_binary: None,
            formula2_binary: None,
            allow_blank: true,
            show_dropdown: true,
            show_input_message: false,
            show_error_message: true,
            error_style: 0,
            ime_mode: 0,
            string_list: validation_type == 3,
            list_formula: None,
            input_title: None,
            input_text: None,
            error_title: None,
            error_text: None,
            cell_ranges,
            record_kind: DataValidationRecordKind::Classic,
        }
    }

    pub(crate) fn parse_classic(
        data: &[u8],
        list_formula: Option<String>,
        formula_context: &FormulaResolutionContext,
    ) -> XlsbResult<Self> {
        let mut cursor = ValidationCursor::new(data, "BrtDVal");
        let flags = cursor.read_u32()?;
        let validation_type = (flags & 0x0f) as u8;
        let error_style = ((flags >> 4) & 0x07) as u8;
        let string_list = flags & 0x80 != 0;
        let ime_mode = ((flags >> 10) & 0xff) as u8;
        let operator = ((flags >> 20) & 0x0f) as u8;
        if validation_type > 7 {
            return Err(invalid(
                "BrtDVal",
                format!("validation type {validation_type}"),
            ));
        }
        if error_style > 2 {
            return Err(invalid("BrtDVal", format!("error style {error_style}")));
        }
        if ime_mode > 10 {
            return Err(invalid("BrtDVal", format!("IME mode {ime_mode}")));
        }
        if !matches!(validation_type, 0 | 3 | 7) && operator > 7 {
            return Err(invalid("BrtDVal", format!("operator {operator}")));
        }
        if flags & 0xff00_0000 != 0 {
            return Err(invalid(
                "BrtDVal",
                format!("reserved flag bits are set in 0x{flags:08X}"),
            ));
        }

        let ranges = cursor.read_ranges(1, 8_191)?;
        let error_title = cursor.read_nullable_string(32, "error title")?;
        let error_text = cursor.read_nullable_string(225, "error message")?;
        let input_title = cursor.read_nullable_string(32, "prompt title")?;
        let input_text = cursor.read_nullable_string(255, "prompt message")?;
        let formula1_binary = cursor.read_classic_formula()?;
        let formula2_binary = cursor.read_classic_formula()?;
        cursor.finish()?;

        validate_formula_presence(
            "BrtDVal",
            validation_type,
            operator,
            formula1_binary.is_some(),
            formula2_binary.is_some(),
        )?;
        let base = (ranges[0].0, ranges[0].2);
        let parsed_formula1 = render_formula(formula1_binary.as_ref(), base, formula_context)?;
        let formula1 = list_formula.clone().or(parsed_formula1);
        let formula2 = render_formula(formula2_binary.as_ref(), base, formula_context)?;

        Ok(Self {
            validation_type,
            operator,
            formula1,
            formula2,
            formula1_binary,
            formula2_binary,
            allow_blank: flags & 0x100 != 0,
            show_dropdown: flags & 0x200 == 0,
            show_input_message: flags & 0x0004_0000 != 0,
            show_error_message: flags & 0x0008_0000 != 0,
            error_style,
            ime_mode,
            string_list,
            list_formula,
            input_title,
            input_text,
            error_title,
            error_text,
            cell_ranges: format_ranges(&ranges),
            record_kind: DataValidationRecordKind::Classic,
        })
    }

    pub(crate) fn parse_extension14(
        data: &[u8],
        formula_context: &FormulaResolutionContext,
    ) -> XlsbResult<Self> {
        let mut cursor = ValidationCursor::new(data, "BrtDVal14");
        let header_flags = cursor.read_u32()?;
        if header_flags & !0x06 != 0 || header_flags & 0x02 == 0 {
            return Err(invalid(
                "BrtDVal14",
                format!("invalid FRTHeader flags 0x{header_flags:08X}"),
            ));
        }

        let sqref_count = cursor.read_u32()?;
        if sqref_count != 1 {
            return Err(invalid(
                "BrtDVal14",
                format!("FRTSqrefs count {sqref_count} is not 1"),
            ));
        }
        let sqref_flags = cursor.read_u32()?;
        if sqref_flags & 0x02 == 0 || sqref_flags & !0x0001_000f != 0 {
            return Err(invalid(
                "BrtDVal14",
                format!("invalid FRTSqref flags 0x{sqref_flags:08X}"),
            ));
        }
        let ranges = cursor.read_ranges(1, i32::MAX as usize)?;

        let mut formulas = Vec::new();
        if header_flags & 0x04 != 0 {
            let formula_count = usize::try_from(cursor.read_u32()?)
                .map_err(|_| invalid("BrtDVal14", "formula count overflow"))?;
            if formula_count > 2 {
                return Err(invalid(
                    "BrtDVal14",
                    format!("FRT formula count {formula_count} exceeds 2"),
                ));
            }
            formulas.reserve(formula_count);
            for _ in 0..formula_count {
                formulas.push(cursor.read_frt_formula()?);
            }
        }

        let flags = cursor.read_u32()?;
        let validation_type = (flags & 0x0f) as u8;
        let error_style = ((flags >> 4) & 0x07) as u8;
        let string_list = flags & 0x80 != 0;
        let ime_mode = ((flags >> 10) & 0xff) as u8;
        let operator = ((flags >> 20) & 0x0f) as u8;
        let has_first = flags & 0x0100_0000 != 0;
        let has_second = flags & 0x0200_0000 != 0;
        if validation_type > 7 {
            return Err(invalid(
                "BrtDVal14",
                format!("validation type {validation_type}"),
            ));
        }
        if error_style > 2 {
            return Err(invalid("BrtDVal14", format!("error style {error_style}")));
        }
        if ime_mode > 10 {
            return Err(invalid("BrtDVal14", format!("IME mode {ime_mode}")));
        }
        if !matches!(validation_type, 0 | 3 | 7) && operator > 7 {
            return Err(invalid("BrtDVal14", format!("operator {operator}")));
        }
        if flags & 0xfc00_0000 != 0 {
            return Err(invalid(
                "BrtDVal14",
                format!("reserved flag bits are set in 0x{flags:08X}"),
            ));
        }
        let expected_formula_count = usize::from(has_first) + usize::from(has_second);
        if formulas.len() != expected_formula_count
            || (header_flags & 0x04 != 0) != (expected_formula_count != 0)
        {
            return Err(invalid(
                "BrtDVal14",
                "FRTHeader formula metadata does not match DVal14 flags",
            ));
        }
        validate_formula_presence(
            "BrtDVal14",
            validation_type,
            operator,
            has_first,
            has_second,
        )?;

        let error_title = cursor.read_nullable_string(32, "error title")?;
        let error_text = cursor.read_nullable_string(225, "error message")?;
        let input_title = cursor.read_nullable_string(32, "prompt title")?;
        let input_text = cursor.read_nullable_string(255, "prompt message")?;
        cursor.finish()?;

        let mut formulas = formulas.into_iter();
        let formula1_binary = if has_first { formulas.next() } else { None };
        let formula2_binary = if has_second { formulas.next() } else { None };
        let base = (ranges[0].0, ranges[0].2);
        let formula1 = render_formula(formula1_binary.as_ref(), base, formula_context)?;
        let formula2 = render_formula(formula2_binary.as_ref(), base, formula_context)?;

        Ok(Self {
            validation_type,
            operator,
            formula1,
            formula2,
            formula1_binary,
            formula2_binary,
            allow_blank: flags & 0x100 != 0,
            show_dropdown: flags & 0x200 == 0,
            show_input_message: flags & 0x0004_0000 != 0,
            show_error_message: flags & 0x0008_0000 != 0,
            error_style,
            ime_mode,
            string_list,
            list_formula: None,
            input_title,
            input_text,
            error_title,
            error_text,
            cell_ranges: format_ranges(&ranges),
            record_kind: DataValidationRecordKind::Extension14,
        })
    }
}

pub(crate) fn parse_collection_settings(
    data: &[u8],
    extension14: bool,
) -> XlsbResult<(DataValidationSettings, u32)> {
    let expected = if extension14 { 22 } else { 18 };
    if data.len() != expected {
        return Err(XlsbError::InvalidLength {
            expected,
            found: data.len(),
        });
    }
    let offset = usize::from(extension14) * 4;
    if extension14 && binary::read_u32_le_at(data, 0)? != 0 {
        return Err(invalid("BrtBeginDVals14", "nonzero FRTBlank header"));
    }
    let flags = binary::read_u16_le_at(data, offset)?;
    let prompt_x = binary::read_u32_le_at(data, offset + 2)?;
    let prompt_y = binary::read_u32_le_at(data, offset + 6)?;
    let unused = binary::read_u32_le_at(data, offset + 10)?;
    let count = binary::read_u32_le_at(data, offset + 14)?;
    if flags & !1 != 0 || prompt_x > u16::MAX.into() || prompt_y > u16::MAX.into() || unused != 0 {
        return Err(invalid(
            if extension14 {
                "BrtBeginDVals14"
            } else {
                "BrtBeginDVals"
            },
            "invalid reserved fields or prompt coordinates",
        ));
    }
    if count > 65_534 {
        return Err(invalid(
            if extension14 {
                "BrtBeginDVals14"
            } else {
                "BrtBeginDVals"
            },
            format!("validation count {count} exceeds 65,534"),
        ));
    }
    Ok((
        DataValidationSettings {
            input_prompts_disabled: flags & 1 != 0,
            prompt_x: prompt_x as u16,
            prompt_y: prompt_y as u16,
        },
        count,
    ))
}

pub(crate) fn parse_dval_list(data: &[u8]) -> XlsbResult<String> {
    let mut cursor = ValidationCursor::new(data, "BrtDValList");
    let value = cursor
        .read_nullable_string(usize::MAX, "list formula")?
        .ok_or_else(|| invalid("BrtDValList", "NULL list formula"))?;
    cursor.finish()?;
    validate_dval_list_formula(&value)?;
    Ok(value)
}

pub(crate) fn validate_dval_list_formula(value: &str) -> XlsbResult<()> {
    if value.encode_utf16().count() >= u32::MAX as usize {
        return Err(invalid(
            "BrtDValList",
            "list formula is too long for XLWideString",
        ));
    }
    if value.chars().any(|character| {
        !matches!(character, '\u{9}' | '\u{a}' | '\u{d}')
            && (character < '\u{20}' || matches!(character, '\u{fffe}' | '\u{ffff}'))
    }) {
        return Err(invalid(
            "BrtDValList",
            "list formula contains a character forbidden by XML Char",
        ));
    }
    let mut quoted = false;
    let mut characters = value.chars().peekable();
    while let Some(character) = characters.next() {
        if character != '"' {
            continue;
        }
        if characters.peek() == Some(&'"') {
            let _ = characters.next();
        } else {
            quoted = !quoted;
        }
    }
    if quoted {
        return Err(invalid(
            "BrtDValList",
            "list formula contains an unmatched double quote",
        ));
    }
    Ok(())
}

fn validate_formula_presence(
    record: &'static str,
    validation_type: u8,
    operator: u8,
    first: bool,
    second: bool,
) -> XlsbResult<()> {
    let expected = match validation_type {
        0 => (false, false),
        3 | 7 => (true, false),
        _ if operator <= 1 => (true, true),
        _ => (true, false),
    };
    if (first, second) != expected {
        return Err(invalid(
            record,
            format!(
                "formula presence {first}/{second} does not match required {}/{}",
                expected.0, expected.1
            ),
        ));
    }
    Ok(())
}

fn render_formula(
    formula: Option<&CellParsedFormula>,
    base: (u32, u32),
    context: &FormulaResolutionContext,
) -> XlsbResult<Option<String>> {
    let Some(formula) = formula else {
        return Ok(None);
    };
    let tokens =
        FormulaParser::with_base_cell_and_extra(&formula.rgce, &formula.rgcb, base.0, base.1)
            .parse()?;
    Ok(Some(FormulaConverter::try_tokens_to_string_with_context(
        &tokens, context,
    )?))
}

fn format_ranges(ranges: &[(u32, u32, u32, u32)]) -> String {
    ranges
        .iter()
        .map(|&(first_row, last_row, first_col, last_col)| {
            let first = cell_reference(first_row, first_col);
            let last = cell_reference(last_row, last_col);
            if first == last {
                first
            } else {
                format!("{first}:{last}")
            }
        })
        .collect::<Vec<_>>()
        .join(" ")
}

fn invalid(typ: &'static str, val: impl Into<String>) -> XlsbError {
    XlsbError::Unrecognized {
        typ: typ.to_string(),
        val: val.into(),
    }
}

struct ValidationCursor<'a> {
    data: &'a [u8],
    offset: usize,
    record: &'static str,
}

impl<'a> ValidationCursor<'a> {
    fn new(data: &'a [u8], record: &'static str) -> Self {
        Self {
            data,
            offset: 0,
            record,
        }
    }

    fn read_u32(&mut self) -> XlsbResult<u32> {
        let value = binary::read_u32_le_at(self.data, self.offset)?;
        self.offset += 4;
        Ok(value)
    }

    fn take(&mut self, len: usize) -> XlsbResult<&'a [u8]> {
        let end = self
            .offset
            .checked_add(len)
            .ok_or_else(|| invalid(self.record, "field length overflow"))?;
        let value = self
            .data
            .get(self.offset..end)
            .ok_or(XlsbError::InvalidLength {
                expected: end,
                found: self.data.len(),
            })?;
        self.offset = end;
        Ok(value)
    }

    fn read_ranges(
        &mut self,
        minimum: usize,
        maximum: usize,
    ) -> XlsbResult<Vec<(u32, u32, u32, u32)>> {
        let raw_count = self.read_u32()? as i32;
        let count =
            usize::try_from(raw_count).map_err(|_| invalid(self.record, "NULL range list"))?;
        if !(minimum..=maximum).contains(&count) {
            return Err(invalid(
                self.record,
                format!("range count {count} is outside {minimum}..={maximum}"),
            ));
        }
        if count > self.data.len().saturating_sub(self.offset) / 16 {
            return Err(XlsbError::InvalidLength {
                expected: self.offset.saturating_add(count.saturating_mul(16)),
                found: self.data.len(),
            });
        }
        let mut ranges = Vec::with_capacity(count);
        for _ in 0..count {
            let first_row = self.read_u32()?;
            let last_row = self.read_u32()?;
            let first_col = self.read_u32()?;
            let last_col = self.read_u32()?;
            if first_row > last_row
                || first_col > last_col
                || last_row >= 1_048_576
                || last_col >= 16_384
            {
                return Err(invalid(self.record, "invalid validation target range"));
            }
            ranges.push((first_row, last_row, first_col, last_col));
        }
        Ok(ranges)
    }

    fn read_nullable_string(
        &mut self,
        maximum: usize,
        field: &'static str,
    ) -> XlsbResult<Option<String>> {
        let count = self.read_u32()?;
        if count == u32::MAX {
            return Ok(None);
        }
        let count =
            usize::try_from(count).map_err(|_| invalid(self.record, "string length overflow"))?;
        if count > maximum {
            return Err(invalid(
                self.record,
                format!("{field} length {count} exceeds {maximum}"),
            ));
        }
        let bytes = self.take(
            count
                .checked_mul(2)
                .ok_or_else(|| invalid(self.record, "string size overflow"))?,
        )?;
        let units = bytes
            .chunks_exact(2)
            .map(|chunk| u16::from_le_bytes([chunk[0], chunk[1]]))
            .collect::<Vec<_>>();
        String::from_utf16(&units)
            .map(Some)
            .map_err(|error| XlsbError::Encoding(format!("invalid {field} UTF-16: {error}")))
    }

    fn read_classic_formula(&mut self) -> XlsbResult<Option<CellParsedFormula>> {
        let cce = usize::try_from(self.read_u32()?)
            .map_err(|_| invalid(self.record, "formula token length overflow"))?;
        if cce > MAX_CELL_FORMULA_BYTES {
            return Err(XlsbError::InvalidFormula(format!(
                "data-validation formula token length {cce} exceeds {MAX_CELL_FORMULA_BYTES}"
            )));
        }
        let rgce = self.take(cce)?.to_vec();
        let cb = usize::try_from(self.read_u32()?)
            .map_err(|_| invalid(self.record, "formula ancillary length overflow"))?;
        let rgcb = self.take(cb)?.to_vec();
        if cce == 0 {
            if cb != 0 {
                return Err(XlsbError::InvalidFormula(
                    "empty data-validation formula has ancillary bytes".to_string(),
                ));
            }
            Ok(None)
        } else {
            Ok(Some(CellParsedFormula { rgce, rgcb }))
        }
    }

    fn read_frt_formula(&mut self) -> XlsbResult<CellParsedFormula> {
        let flags = self.read_u32()?;
        if flags != 2 {
            return Err(invalid(
                self.record,
                format!("invalid FRTFormula flags 0x{flags:08X}"),
            ));
        }
        let cce = usize::try_from(self.read_u32()?)
            .map_err(|_| invalid(self.record, "formula token length overflow"))?;
        let cb = usize::try_from(self.read_u32()?)
            .map_err(|_| invalid(self.record, "formula ancillary length overflow"))?;
        if cce == 0 || cce > MAX_CELL_FORMULA_BYTES {
            return Err(XlsbError::InvalidFormula(format!(
                "FRT data-validation formula token length {cce} is outside 1..={MAX_CELL_FORMULA_BYTES}"
            )));
        }
        let rgce = self.take(cce)?.to_vec();
        let rgcb = self.take(cb)?.to_vec();
        Ok(CellParsedFormula { rgce, rgcb })
    }

    fn finish(self) -> XlsbResult<()> {
        if self.offset != self.data.len() {
            return Err(XlsbError::InvalidLength {
                expected: self.offset,
                found: self.data.len(),
            });
        }
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_data_validation_new() {
        let dv = DataValidation::new(3, "A1:A10".to_string());
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
        let mut dv = DataValidation::new(1, "B1:B20".to_string()); // whole number
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
        let mut dv = DataValidation::new(2, "C1:C10".to_string()); // decimal
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
        let mut dv = DataValidation::new(3, "D1:D10".to_string()); // list
        dv.formula1 = Some("Yes,No,Maybe".to_string());
        dv.show_dropdown = true;

        assert_eq!(dv.validation_type, 3);
        assert_eq!(dv.formula1, Some("Yes,No,Maybe".to_string()));
        assert!(dv.show_dropdown);
    }

    #[test]
    fn test_data_validation_date() {
        let mut dv = DataValidation::new(4, "E1:E10".to_string()); // date
        dv.operator = 4; // greater than
        dv.formula1 = Some("2024-01-01".to_string());

        assert_eq!(dv.validation_type, 4);
        assert_eq!(dv.operator, 4);
    }

    #[test]
    fn test_data_validation_time() {
        let mut dv = DataValidation::new(5, "F1:F10".to_string()); // time
        dv.operator = 5; // less than
        dv.formula1 = Some("12:00".to_string());

        assert_eq!(dv.validation_type, 5);
        assert_eq!(dv.operator, 5);
    }

    #[test]
    fn test_data_validation_text_length() {
        let mut dv = DataValidation::new(6, "G1:G10".to_string()); // text length
        dv.operator = 6; // greater than or equal
        dv.formula1 = Some("5".to_string());

        assert_eq!(dv.validation_type, 6);
        assert_eq!(dv.formula1, Some("5".to_string()));
    }

    #[test]
    fn test_data_validation_custom() {
        let mut dv = DataValidation::new(7, "H1:H10".to_string()); // custom
        dv.formula1 = Some("=A1>0".to_string());

        assert_eq!(dv.validation_type, 7);
        assert_eq!(dv.formula1, Some("=A1>0".to_string()));
    }

    #[test]
    fn test_data_validation_with_messages() {
        let mut dv = DataValidation::new(1, "I1:I10".to_string());
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
        let dv = DataValidation::new(3, "A1:A10,C1:C10,E1:E10".to_string());
        assert_eq!(dv.cell_ranges, "A1:A10,C1:C10,E1:E10");
    }

    #[test]
    fn test_data_validation_clone() {
        let dv = DataValidation::new(3, "A1:A10".to_string());
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
}
