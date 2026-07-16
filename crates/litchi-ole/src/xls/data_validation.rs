//! BIFF8 worksheet data-validation records.

use super::{XlsError, XlsResult};

pub(crate) const DVAL_RECORD_TYPE: u16 = 0x01B2;
pub(crate) const DV_RECORD_TYPE: u16 = 0x01BE;
const MAX_VALIDATION_RANGES: usize = 432;

/// Worksheet-level settings declared by a BIFF8 `DVAL` record.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XlsDataValidationSettings {
    window_closed: bool,
    x_left: u32,
    y_top: u32,
    dropdown_object_id: Option<u16>,
    declared_rule_count: u16,
}

impl XlsDataValidationSettings {
    pub fn window_closed(&self) -> bool { self.window_closed }
    pub fn x_left(&self) -> u32 { self.x_left }
    pub fn y_top(&self) -> u32 { self.y_top }
    pub fn dropdown_object_id(&self) -> Option<u16> { self.dropdown_object_id }
    pub fn declared_rule_count(&self) -> u16 { self.declared_rule_count }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum XlsDataValidationKind { Any, Whole, Decimal, List, Date, Time, TextLength, Custom }

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum XlsDataValidationErrorStyle { Stop, Warning, Information }

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum XlsDataValidationOperator {
    Between, NotBetween, Equal, NotEqual, GreaterThan, LessThan, GreaterOrEqual, LessOrEqual,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum XlsDataValidationImeMode {
    NoControl, On, Off, Hiragana, WideKatakana, NarrowKatakana,
    FullWidthAlphanumeric, HalfWidthAlphanumeric, FullWidthHangul, HalfWidthHangul,
}

/// An unevaluated BIFF formula token stream from a `DV` record.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XlsDataValidationFormula { tokens: Vec<u8> }

impl XlsDataValidationFormula {
    pub fn tokens(&self) -> &[u8] { &self.tokens }
}

/// An inclusive BIFF8 cell range targeted by a validation rule.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct XlsDataValidationRange {
    first_row: u16,
    last_row: u16,
    first_column: u8,
    last_column: u8,
}

impl XlsDataValidationRange {
    pub fn first_row(&self) -> u16 { self.first_row }
    pub fn last_row(&self) -> u16 { self.last_row }
    pub fn first_column(&self) -> u8 { self.first_column }
    pub fn last_column(&self) -> u8 { self.last_column }
}

/// One BIFF8 worksheet data-validation rule.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XlsDataValidationRule {
    kind: XlsDataValidationKind,
    error_style: XlsDataValidationErrorStyle,
    explicit_list: bool,
    allow_blank: bool,
    suppress_dropdown: bool,
    ime_mode: XlsDataValidationImeMode,
    show_input_message: bool,
    show_error_message: bool,
    operator: XlsDataValidationOperator,
    prompt_title: Option<String>,
    error_title: Option<String>,
    prompt: Option<String>,
    error: Option<String>,
    formula1: Option<XlsDataValidationFormula>,
    formula2: Option<XlsDataValidationFormula>,
    ranges: Vec<XlsDataValidationRange>,
}

impl XlsDataValidationRule {
    pub fn kind(&self) -> XlsDataValidationKind { self.kind }
    pub fn error_style(&self) -> XlsDataValidationErrorStyle { self.error_style }
    pub fn explicit_list(&self) -> bool { self.explicit_list }
    pub fn allow_blank(&self) -> bool { self.allow_blank }
    pub fn suppress_dropdown(&self) -> bool { self.suppress_dropdown }
    pub fn ime_mode(&self) -> XlsDataValidationImeMode { self.ime_mode }
    pub fn show_input_message(&self) -> bool { self.show_input_message }
    pub fn show_error_message(&self) -> bool { self.show_error_message }
    pub fn operator(&self) -> XlsDataValidationOperator { self.operator }
    pub fn prompt_title(&self) -> Option<&str> { self.prompt_title.as_deref() }
    pub fn error_title(&self) -> Option<&str> { self.error_title.as_deref() }
    pub fn prompt(&self) -> Option<&str> { self.prompt.as_deref() }
    pub fn error(&self) -> Option<&str> { self.error.as_deref() }
    pub fn formula1(&self) -> Option<&XlsDataValidationFormula> { self.formula1.as_ref() }
    pub fn formula2(&self) -> Option<&XlsDataValidationFormula> { self.formula2.as_ref() }
    pub fn ranges(&self) -> &[XlsDataValidationRange] { &self.ranges }
}

pub(crate) fn parse_dval(data: &[u8]) -> XlsResult<XlsDataValidationSettings> {
    if data.len() != 18 {
        return invalid(format!("DVAL payload must be exactly 18 bytes, got {}", data.len()));
    }
    let options = u16::from_le_bytes([data[0], data[1]]);
    if options & !0x0005 != 0 {
        return invalid(format!("DVAL contains reserved option bits: {options:#06x}"));
    }
    let x_left = read_u32(data, 2)?;
    let y_top = read_u32(data, 6)?;
    if x_left > 65_535 || y_top > 65_535 {
        return invalid("DVAL window coordinates exceed 65535".to_string());
    }
    let object_id = i32::from_le_bytes(data[10..14].try_into().unwrap());
    let dropdown_object_id = match object_id {
        -1 => None,
        1..=32_767 => Some(object_id as u16),
        _ => return invalid(format!("invalid DVAL dropdown object id: {object_id}")),
    };
    let rule_count = read_u32(data, 14)?;
    if rule_count > 65_534 {
        return invalid(format!("DVAL rule count exceeds 65534: {rule_count}"));
    }
    Ok(XlsDataValidationSettings {
        window_closed: options & 1 != 0,
        x_left,
        y_top,
        dropdown_object_id,
        declared_rule_count: rule_count as u16,
    })
}

pub(crate) fn parse_dv(data: &[u8]) -> XlsResult<XlsDataValidationRule> {
    let mut cursor = Cursor::new(data);
    let options = cursor.u32()?;
    if options & 0xFF00_0000 != 0 {
        return invalid(format!("DV contains reserved option bits: {options:#010x}"));
    }
    let kind = match options & 0x0F {
        0 => XlsDataValidationKind::Any,
        1 => XlsDataValidationKind::Whole,
        2 => XlsDataValidationKind::Decimal,
        3 => XlsDataValidationKind::List,
        4 => XlsDataValidationKind::Date,
        5 => XlsDataValidationKind::Time,
        6 => XlsDataValidationKind::TextLength,
        7 => XlsDataValidationKind::Custom,
        value => return invalid(format!("invalid DV validation type: {value}")),
    };
    let error_style = match (options >> 4) & 0x07 {
        0 => XlsDataValidationErrorStyle::Stop,
        1 => XlsDataValidationErrorStyle::Warning,
        2 => XlsDataValidationErrorStyle::Information,
        value => return invalid(format!("invalid DV error style: {value}")),
    };
    let ime_mode = match (options >> 10) & 0xFF {
        0 => XlsDataValidationImeMode::NoControl,
        1 => XlsDataValidationImeMode::On,
        2 => XlsDataValidationImeMode::Off,
        4 => XlsDataValidationImeMode::Hiragana,
        5 => XlsDataValidationImeMode::WideKatakana,
        6 => XlsDataValidationImeMode::NarrowKatakana,
        7 => XlsDataValidationImeMode::FullWidthAlphanumeric,
        8 => XlsDataValidationImeMode::HalfWidthAlphanumeric,
        9 => XlsDataValidationImeMode::FullWidthHangul,
        10 => XlsDataValidationImeMode::HalfWidthHangul,
        value => return invalid(format!("invalid DV IME mode: {value}")),
    };
    let operator = match (options >> 20) & 0x0F {
        0 => XlsDataValidationOperator::Between,
        1 => XlsDataValidationOperator::NotBetween,
        2 => XlsDataValidationOperator::Equal,
        3 => XlsDataValidationOperator::NotEqual,
        4 => XlsDataValidationOperator::GreaterThan,
        5 => XlsDataValidationOperator::LessThan,
        6 => XlsDataValidationOperator::GreaterOrEqual,
        7 => XlsDataValidationOperator::LessOrEqual,
        value => return invalid(format!("invalid DV operator: {value}")),
    };

    let prompt_title = cursor.unicode_string(32)?;
    let error_title = cursor.unicode_string(32)?;
    let prompt = cursor.unicode_string(255)?;
    let error = cursor.unicode_string(225)?;
    let formula1 = cursor.formula()?;
    let formula2 = cursor.formula()?;
    let needs_two = !matches!(kind, XlsDataValidationKind::Any | XlsDataValidationKind::List | XlsDataValidationKind::Custom)
        && matches!(operator, XlsDataValidationOperator::Between | XlsDataValidationOperator::NotBetween);
    match kind {
        XlsDataValidationKind::Any if formula1.is_some() || formula2.is_some() =>
            return invalid("DV type Any must not contain formulas".to_string()),
        XlsDataValidationKind::Any => {}
        _ if needs_two && (formula1.is_none() || formula2.is_none()) =>
            return invalid("DV Between/NotBetween rule requires two formulas".to_string()),
        _ if !needs_two && (formula1.is_none() || formula2.is_some()) =>
            return invalid("DV rule requires exactly one formula".to_string()),
        _ => {}
    }

    let range_count = cursor.u16()? as usize;
    if !(1..=MAX_VALIDATION_RANGES).contains(&range_count) {
        return invalid(format!("DV range count must be 1..={MAX_VALIDATION_RANGES}, got {range_count}"));
    }
    let bytes_needed = range_count.checked_mul(8)
        .ok_or_else(|| XlsError::InvalidData("DV range size overflow".to_string()))?;
    if cursor.remaining() != bytes_needed {
        return invalid(format!("DV range list length mismatch: expected {bytes_needed} bytes, got {}", cursor.remaining()));
    }
    let mut ranges = Vec::with_capacity(range_count);
    for _ in 0..range_count {
        let first_row = cursor.u16()?;
        let last_row = cursor.u16()?;
        let first_column = cursor.u16()?;
        let last_column = cursor.u16()?;
        if first_row > last_row || first_column > last_column || last_column > 255 {
            return invalid("DV contains an invalid or out-of-range cell range".to_string());
        }
        ranges.push(XlsDataValidationRange {
            first_row,
            last_row,
            first_column: first_column as u8,
            last_column: last_column as u8,
        });
    }
    Ok(XlsDataValidationRule {
        kind,
        error_style,
        explicit_list: options & 0x80 != 0,
        allow_blank: options & 0x100 != 0,
        suppress_dropdown: options & 0x200 != 0,
        ime_mode,
        show_input_message: options & 0x0004_0000 != 0,
        show_error_message: options & 0x0008_0000 != 0,
        operator,
        prompt_title,
        error_title,
        prompt,
        error,
        formula1,
        formula2,
        ranges,
    })
}

fn invalid<T>(message: String) -> XlsResult<T> { Err(XlsError::InvalidData(message)) }

fn read_u32(data: &[u8], offset: usize) -> XlsResult<u32> {
    let bytes = data.get(offset..offset + 4)
        .ok_or_else(|| XlsError::InvalidData("truncated BIFF data-validation record".to_string()))?;
    Ok(u32::from_le_bytes(bytes.try_into().unwrap()))
}

struct Cursor<'a> { data: &'a [u8], position: usize }

impl<'a> Cursor<'a> {
    fn new(data: &'a [u8]) -> Self { Self { data, position: 0 } }
    fn remaining(&self) -> usize { self.data.len().saturating_sub(self.position) }
    fn take(&mut self, count: usize) -> XlsResult<&'a [u8]> {
        let end = self.position.checked_add(count)
            .ok_or_else(|| XlsError::InvalidData("DV field size overflow".to_string()))?;
        let bytes = self.data.get(self.position..end)
            .ok_or_else(|| XlsError::InvalidData("truncated DV record".to_string()))?;
        self.position = end;
        Ok(bytes)
    }
    fn u16(&mut self) -> XlsResult<u16> {
        Ok(u16::from_le_bytes(self.take(2)?.try_into().unwrap()))
    }
    fn u32(&mut self) -> XlsResult<u32> {
        Ok(u32::from_le_bytes(self.take(4)?.try_into().unwrap()))
    }
    fn unicode_string(&mut self, max_units: usize) -> XlsResult<Option<String>> {
        let units = self.u16()? as usize;
        if units > max_units {
            return invalid(format!("DV string exceeds its {max_units}-character limit"));
        }
        let flags = self.take(1)?[0];
        if flags & !0x0D != 0 {
            return invalid(format!("DV string contains reserved flags: {flags:#04x}"));
        }
        let rich_runs = if flags & 0x08 != 0 { self.u16()? as usize } else { 0 };
        let extension_size = if flags & 0x04 != 0 { self.u32()? as usize } else { 0 };
        let wide = flags & 0x01 != 0;
        let character_bytes = units.checked_mul(if wide { 2 } else { 1 })
            .ok_or_else(|| XlsError::InvalidData("DV string size overflow".to_string()))?;
        let characters = self.take(character_bytes)?;
        let value = if wide {
            let utf16: Vec<u16> = characters.chunks_exact(2)
                .map(|bytes| u16::from_le_bytes([bytes[0], bytes[1]]))
                .collect();
            String::from_utf16(&utf16)
                .map_err(|_| XlsError::InvalidData("DV string contains invalid UTF-16".to_string()))?
        } else {
            characters.iter().map(|&byte| char::from(byte)).collect()
        };
        let formatting_size = rich_runs.checked_mul(4)
            .ok_or_else(|| XlsError::InvalidData("DV rich-text size overflow".to_string()))?;
        self.take(formatting_size)?;
        self.take(extension_size)?;
        Ok(if value == "\0" { None } else { Some(value) })
    }
    fn formula(&mut self) -> XlsResult<Option<XlsDataValidationFormula>> {
        let size = self.u16()? as usize;
        self.take(2)?;
        let tokens = self.take(size)?;
        Ok((!tokens.is_empty()).then(|| XlsDataValidationFormula { tokens: tokens.to_vec() }))
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn string(data: &mut Vec<u8>, value: &str) {
        data.extend_from_slice(&(value.len() as u16).to_le_bytes());
        data.push(0);
        data.extend_from_slice(value.as_bytes());
    }
    fn formula(data: &mut Vec<u8>, tokens: &[u8]) {
        data.extend_from_slice(&(tokens.len() as u16).to_le_bytes());
        data.extend_from_slice(&0u16.to_le_bytes());
        data.extend_from_slice(tokens);
    }
    fn valid_dv() -> Vec<u8> {
        let options = 1u32 | (1 << 4) | (1 << 8) | (1 << 18) | (1 << 19);
        let mut data = options.to_le_bytes().to_vec();
        string(&mut data, "Input");
        string(&mut data, "Error");
        string(&mut data, "Enter a value");
        string(&mut data, "Invalid");
        formula(&mut data, &[0x1E, 1, 0]);
        formula(&mut data, &[0x1E, 10, 0]);
        data.extend_from_slice(&1u16.to_le_bytes());
        for value in [2u16, 4, 3, 5] { data.extend_from_slice(&value.to_le_bytes()); }
        data
    }

    #[test]
    fn parses_dval_and_dv_with_raw_formulas() {
        let mut dval = Vec::new();
        dval.extend_from_slice(&1u16.to_le_bytes());
        dval.extend_from_slice(&10u32.to_le_bytes());
        dval.extend_from_slice(&20u32.to_le_bytes());
        dval.extend_from_slice(&(-1i32).to_le_bytes());
        dval.extend_from_slice(&1u32.to_le_bytes());
        let settings = parse_dval(&dval).unwrap();
        assert!(settings.window_closed());
        assert_eq!(settings.declared_rule_count(), 1);

        let rule = parse_dv(&valid_dv()).unwrap();
        assert_eq!(rule.kind(), XlsDataValidationKind::Whole);
        assert_eq!(rule.error_style(), XlsDataValidationErrorStyle::Warning);
        assert_eq!(rule.formula1().unwrap().tokens(), &[0x1E, 1, 0]);
        assert_eq!(rule.formula2().unwrap().tokens(), &[0x1E, 10, 0]);
        assert_eq!(rule.ranges()[0].first_row(), 2);
        assert_eq!(rule.ranges()[0].last_column(), 5);
    }

    #[test]
    fn rejects_reserved_bits_and_malformed_rule_shape() {
        let mut dval = [0u8; 18];
        dval[0] = 2;
        assert!(parse_dval(&dval).is_err());
        let mut dv = valid_dv();
        dv[3] = 0x80;
        assert!(parse_dv(&dv).is_err());
        let mut dv = valid_dv();
        let end = dv.len();
        dv[end - 8..end - 6].copy_from_slice(&5u16.to_le_bytes());
        dv[end - 6..end - 4].copy_from_slice(&4u16.to_le_bytes());
        assert!(parse_dv(&dv).is_err());
    }

    #[test]
    fn enforces_formula_cardinality_and_range_limit() {
        let mut dv = valid_dv();
        dv[0] = 0;
        assert!(parse_dv(&dv).is_err());
        let mut data = 3u32.to_le_bytes().to_vec();
        for _ in 0..4 { string(&mut data, "\0"); }
        formula(&mut data, &[0x17, 1, 0, b'A']);
        formula(&mut data, &[]);
        data.extend_from_slice(&0u16.to_le_bytes());
        assert!(parse_dv(&data).is_err());
    }
}
