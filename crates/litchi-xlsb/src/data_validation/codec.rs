//! BIFF12 wire codecs for XLSB data-validation records.

use super::model::{FormulaBinary, RecordKind, Settings, Validation};
use super::{Error, Result};
use crate::formula::{Compiler, MAX_CELL_FORMULA_BYTES, Parser, Resolution};

impl<F: FormulaBinary> Validation<F> {
    pub fn parse_classic<R: Resolution>(
        data: &[u8],
        list_formula: Option<String>,
        formula_context: &R,
    ) -> Result<Self> {
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
            record_kind: RecordKind::Classic,
        })
    }

    pub fn parse_extension14<R: Resolution>(data: &[u8], formula_context: &R) -> Result<Self> {
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
            record_kind: RecordKind::Extension14,
        })
    }
}

pub fn parse_collection_settings(data: &[u8], extension14: bool) -> Result<(Settings, u32)> {
    let expected = if extension14 { 22 } else { 18 };
    if data.len() != expected {
        return Err(Error::InvalidLength {
            expected,
            found: data.len(),
        });
    }
    let offset = usize::from(extension14) * 4;
    if extension14 && read_u32_le_at(data, 0)? != 0 {
        return Err(invalid("BrtBeginDVals14", "nonzero FRTBlank header"));
    }
    let flags = read_u16_le_at(data, offset)?;
    let prompt_x = read_u32_le_at(data, offset + 2)?;
    let prompt_y = read_u32_le_at(data, offset + 6)?;
    let unused = read_u32_le_at(data, offset + 10)?;
    let count = read_u32_le_at(data, offset + 14)?;
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
        Settings {
            input_prompts_disabled: flags & 1 != 0,
            prompt_x: prompt_x as u16,
            prompt_y: prompt_y as u16,
        },
        count,
    ))
}

pub fn parse_dval_list(data: &[u8]) -> Result<String> {
    let mut cursor = ValidationCursor::new(data, "BrtDValList");
    let value = cursor
        .read_nullable_string(usize::MAX, "list formula")?
        .ok_or_else(|| invalid("BrtDValList", "NULL list formula"))?;
    cursor.finish()?;
    validate_dval_list_formula(&value)?;
    Ok(value)
}

pub fn validate_dval_list_formula(value: &str) -> Result<()> {
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
) -> Result<()> {
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

fn render_formula<F: FormulaBinary, R: Resolution>(
    formula: Option<&F>,
    base: (u32, u32),
    context: &R,
) -> Result<Option<String>> {
    let Some(formula) = formula else {
        return Ok(None);
    };
    let tokens =
        Parser::with_base_cell_and_extra(formula.rgce(), formula.rgcb(), base.0, base.1).parse()?;
    Ok(Some(Compiler::try_tokens_to_string_with_resolution(
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

fn invalid(typ: &'static str, val: impl Into<String>) -> Error {
    Error::Unrecognized {
        typ: typ.to_string(),
        val: val.into(),
    }
}

fn read_u16_le_at(data: &[u8], offset: usize) -> Result<u16> {
    let bytes = data
        .get(offset..offset.saturating_add(2))
        .ok_or(Error::InvalidLength {
            expected: offset.saturating_add(2),
            found: data.len(),
        })?;
    Ok(u16::from_le_bytes([bytes[0], bytes[1]]))
}

fn read_u32_le_at(data: &[u8], offset: usize) -> Result<u32> {
    let bytes = data
        .get(offset..offset.saturating_add(4))
        .ok_or(Error::InvalidLength {
            expected: offset.saturating_add(4),
            found: data.len(),
        })?;
    Ok(u32::from_le_bytes([bytes[0], bytes[1], bytes[2], bytes[3]]))
}

fn column_name(mut column: u32) -> String {
    let mut value = String::new();
    loop {
        value.insert(0, char::from(b'A' + (column % 26) as u8));
        if column < 26 {
            break;
        }
        column = column / 26 - 1;
    }
    value
}

fn cell_reference(row: u32, column: u32) -> String {
    format!("{}{}", column_name(column), row + 1)
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

    fn read_u32(&mut self) -> Result<u32> {
        let bytes = self.take(4)?;
        Ok(u32::from_le_bytes([bytes[0], bytes[1], bytes[2], bytes[3]]))
    }

    fn take(&mut self, len: usize) -> Result<&'a [u8]> {
        let end = self
            .offset
            .checked_add(len)
            .ok_or_else(|| invalid(self.record, "field length overflow"))?;
        let value = self
            .data
            .get(self.offset..end)
            .ok_or(Error::InvalidLength {
                expected: end,
                found: self.data.len(),
            })?;
        self.offset = end;
        Ok(value)
    }

    fn read_ranges(&mut self, minimum: usize, maximum: usize) -> Result<Vec<(u32, u32, u32, u32)>> {
        let raw_count = self.read_u32()?;
        let raw_count = i32::try_from(raw_count)
            .map_err(|_| invalid(self.record, "range count exceeds signed int32"))?;
        let count =
            usize::try_from(raw_count).map_err(|_| invalid(self.record, "NULL range list"))?;
        if !(minimum..=maximum).contains(&count) {
            return Err(invalid(
                self.record,
                format!("range count {count} is outside {minimum}..={maximum}"),
            ));
        }
        if count > self.data.len().saturating_sub(self.offset) / 16 {
            return Err(Error::InvalidLength {
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
    ) -> Result<Option<String>> {
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
            .map_err(|error| Error::Encoding(format!("invalid {field} UTF-16: {error}")))
    }

    fn read_classic_formula<F: FormulaBinary>(&mut self) -> Result<Option<F>> {
        let cce = usize::try_from(self.read_u32()?)
            .map_err(|_| invalid(self.record, "formula token length overflow"))?;
        if cce > MAX_CELL_FORMULA_BYTES {
            return Err(Error::InvalidFormula(format!(
                "data-validation formula token length {cce} exceeds {MAX_CELL_FORMULA_BYTES}"
            )));
        }
        let rgce = self.take(cce)?.to_vec();
        let cb = usize::try_from(self.read_u32()?)
            .map_err(|_| invalid(self.record, "formula ancillary length overflow"))?;
        let rgcb = self.take(cb)?.to_vec();
        if cce == 0 {
            if cb != 0 {
                return Err(Error::InvalidFormula(
                    "empty data-validation formula has ancillary bytes".to_string(),
                ));
            }
            Ok(None)
        } else {
            Ok(Some(F::from_parts(rgce, rgcb)))
        }
    }

    fn read_frt_formula<F: FormulaBinary>(&mut self) -> Result<F> {
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
            return Err(Error::InvalidFormula(format!(
                "FRT data-validation formula token length {cce} is outside 1..={MAX_CELL_FORMULA_BYTES}"
            )));
        }
        let rgce = self.take(cce)?.to_vec();
        let rgcb = self.take(cb)?.to_vec();
        Ok(F::from_parts(rgce, rgcb))
    }

    fn finish(self) -> Result<()> {
        if self.offset != self.data.len() {
            return Err(Error::InvalidLength {
                expected: self.offset,
                found: self.data.len(),
            });
        }
        Ok(())
    }
}
