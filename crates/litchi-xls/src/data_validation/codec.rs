//! BIFF8 DVAL and DV payload codecs.

use super::model::{ErrorStyle, Formula, ImeMode, Kind, Operator, Range, Rule, Settings};
use crate::formula::{FormulaContext, render_formula};
use crate::{Error, Result};

const MAX_VALIDATION_RANGES: usize = 432;

pub(crate) fn parse_dval(data: &[u8]) -> Result<Settings> {
    if data.len() != 18 {
        return invalid(format!(
            "DVAL payload must be exactly 18 bytes, got {}",
            data.len()
        ));
    }
    let mut cursor = Cursor::new(data);
    let options = cursor.u16()?;
    if options & !0x0005 != 0 {
        return invalid(format!(
            "DVAL contains reserved option bits: {options:#06x}"
        ));
    }
    let x_left = cursor.u32()?;
    let y_top = cursor.u32()?;
    if x_left > 65_535 || y_top > 65_535 {
        return invalid("DVAL window coordinates exceed 65535".to_string());
    }
    let object_id = cursor.i32()?;
    let dropdown_object_id = match object_id {
        -1 => None,
        1..=32_767 => Some(u16::try_from(object_id).map_err(|_| {
            Error::InvalidData(format!("invalid DVAL dropdown object id: {object_id}"))
        })?),
        _ => return invalid(format!("invalid DVAL dropdown object id: {object_id}")),
    };
    let rule_count = cursor.u32()?;
    if rule_count > 65_534 {
        return invalid(format!("DVAL rule count exceeds 65534: {rule_count}"));
    }
    Ok(Settings {
        window_closed: options & 1 != 0,
        x_left,
        y_top,
        dropdown_object_id,
        declared_rule_count: u16::try_from(rule_count).map_err(|_| {
            Error::InvalidData(format!("DVAL rule count exceeds 65534: {rule_count}"))
        })?,
    })
}

pub(crate) fn parse_dv(data: &[u8], formula_context: Option<&FormulaContext>) -> Result<Rule> {
    let mut cursor = Cursor::new(data);
    let options = cursor.u32()?;
    if options & 0xFF00_0000 != 0 {
        return invalid(format!("DV contains reserved option bits: {options:#010x}"));
    }
    let kind = match options & 0x0F {
        0 => Kind::Any,
        1 => Kind::Whole,
        2 => Kind::Decimal,
        3 => Kind::List,
        4 => Kind::Date,
        5 => Kind::Time,
        6 => Kind::TextLength,
        7 => Kind::Custom,
        value => return invalid(format!("invalid DV validation type: {value}")),
    };
    let error_style = match (options >> 4) & 0x07 {
        0 => ErrorStyle::Stop,
        1 => ErrorStyle::Warning,
        2 => ErrorStyle::Information,
        value => return invalid(format!("invalid DV error style: {value}")),
    };
    let ime_mode = match (options >> 10) & 0xFF {
        0 => ImeMode::NoControl,
        1 => ImeMode::On,
        2 => ImeMode::Off,
        4 => ImeMode::Hiragana,
        5 => ImeMode::WideKatakana,
        6 => ImeMode::NarrowKatakana,
        7 => ImeMode::FullWidthAlphanumeric,
        8 => ImeMode::HalfWidthAlphanumeric,
        9 => ImeMode::FullWidthHangul,
        10 => ImeMode::HalfWidthHangul,
        value => return invalid(format!("invalid DV IME mode: {value}")),
    };
    let raw_operator = (options >> 20) & 0x0F;
    let operator = if matches!(kind, Kind::Any | Kind::List | Kind::Custom) {
        // This field is undefined and MUST be ignored for operator-less validation kinds.
        Operator::Between
    } else {
        match raw_operator {
            0 => Operator::Between,
            1 => Operator::NotBetween,
            2 => Operator::Equal,
            3 => Operator::NotEqual,
            4 => Operator::GreaterThan,
            5 => Operator::LessThan,
            6 => Operator::GreaterOrEqual,
            7 => Operator::LessOrEqual,
            value => return invalid(format!("invalid DV operator: {value}")),
        }
    };

    let prompt_title = cursor.unicode_string(32)?;
    let error_title = cursor.unicode_string(32)?;
    let prompt = cursor.unicode_string(255)?;
    let error = cursor.unicode_string(225)?;
    let mut formula1 = cursor.formula()?;
    let mut formula2 = cursor.formula()?;
    for formula in [&mut formula1, &mut formula2].into_iter().flatten() {
        formula.rendered = render_formula(&formula.tokens, formula_context);
    }
    let needs_two = !matches!(kind, Kind::Any | Kind::List | Kind::Custom)
        && matches!(operator, Operator::Between | Operator::NotBetween);
    match kind {
        Kind::Any if formula1.is_some() || formula2.is_some() => {
            return invalid("DV type Any must not contain formulas".to_string());
        },
        Kind::Any => {},
        _ if needs_two && (formula1.is_none() || formula2.is_none()) => {
            return invalid("DV Between/NotBetween rule requires two formulas".to_string());
        },
        _ if !needs_two && (formula1.is_none() || formula2.is_some()) => {
            return invalid("DV rule requires exactly one formula".to_string());
        },
        _ => {},
    }

    let range_count = usize::from(cursor.u16()?);
    if !(1..=MAX_VALIDATION_RANGES).contains(&range_count) {
        return invalid(format!(
            "DV range count must be 1..={MAX_VALIDATION_RANGES}, got {range_count}"
        ));
    }
    let bytes_needed = range_count
        .checked_mul(8)
        .ok_or_else(|| Error::InvalidData("DV range size overflow".to_string()))?;
    if cursor.remaining() != bytes_needed {
        return invalid(format!(
            "DV range list length mismatch: expected {bytes_needed} bytes, got {}",
            cursor.remaining()
        ));
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
        let first_column = u8::try_from(first_column).map_err(|_| {
            Error::InvalidData("DV contains an invalid or out-of-range cell range".to_string())
        })?;
        let last_column = u8::try_from(last_column).map_err(|_| {
            Error::InvalidData("DV contains an invalid or out-of-range cell range".to_string())
        })?;
        ranges.push(Range {
            first_row,
            last_row,
            first_column,
            last_column,
        });
    }
    Ok(Rule {
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

fn invalid<T>(message: String) -> Result<T> {
    Err(Error::InvalidData(message))
}

struct Cursor<'a> {
    data: &'a [u8],
    position: usize,
}

impl<'a> Cursor<'a> {
    fn new(data: &'a [u8]) -> Self {
        Self { data, position: 0 }
    }

    fn remaining(&self) -> usize {
        self.data.len().saturating_sub(self.position)
    }

    fn take(&mut self, count: usize) -> Result<&'a [u8]> {
        let end = self
            .position
            .checked_add(count)
            .ok_or_else(|| Error::InvalidData("DV field size overflow".to_string()))?;
        let bytes = self
            .data
            .get(self.position..end)
            .ok_or_else(|| Error::InvalidData("truncated DV record".to_string()))?;
        self.position = end;
        Ok(bytes)
    }

    fn take_array<const N: usize>(&mut self) -> Result<[u8; N]> {
        let bytes: [u8; N] = self
            .take(N)?
            .try_into()
            .map_err(|_| Error::InvalidData("truncated DV record".to_string()))?;
        Ok(bytes)
    }

    fn u16(&mut self) -> Result<u16> {
        Ok(u16::from_le_bytes(self.take_array()?))
    }

    fn u32(&mut self) -> Result<u32> {
        Ok(u32::from_le_bytes(self.take_array()?))
    }

    fn i32(&mut self) -> Result<i32> {
        Ok(i32::from_le_bytes(self.take_array()?))
    }

    fn unicode_string(&mut self, max_units: usize) -> Result<Option<String>> {
        let units = usize::from(self.u16()?);
        if units > max_units {
            return invalid(format!("DV string exceeds its {max_units}-character limit"));
        }
        let [flags] = self.take_array()?;
        if flags & !0x0D != 0 {
            return invalid(format!("DV string contains reserved flags: {flags:#04x}"));
        }
        let rich_runs = if flags & 0x08 != 0 {
            usize::from(self.u16()?)
        } else {
            0
        };
        let extension_size = if flags & 0x04 != 0 {
            usize::try_from(self.u32()?)
                .map_err(|_| Error::InvalidData("DV string size overflow".to_string()))?
        } else {
            0
        };
        let wide = flags & 0x01 != 0;
        let character_bytes = units
            .checked_mul(if wide { 2 } else { 1 })
            .ok_or_else(|| Error::InvalidData("DV string size overflow".to_string()))?;
        let characters = self.take(character_bytes)?;
        let value = if wide {
            let mut chunks = characters.chunks_exact(2);
            let mut utf16 = Vec::with_capacity(units);
            for bytes in &mut chunks {
                let bytes: [u8; 2] = bytes.try_into().map_err(|_| {
                    Error::InvalidData("DV string contains invalid UTF-16".to_string())
                })?;
                utf16.push(u16::from_le_bytes(bytes));
            }
            if !chunks.remainder().is_empty() {
                return invalid("DV string contains invalid UTF-16".to_string());
            }
            String::from_utf16(&utf16)
                .map_err(|_| Error::InvalidData("DV string contains invalid UTF-16".to_string()))?
        } else {
            characters.iter().map(|&byte| char::from(byte)).collect()
        };
        let formatting_size = rich_runs
            .checked_mul(4)
            .ok_or_else(|| Error::InvalidData("DV rich-text size overflow".to_string()))?;
        self.take(formatting_size)?;
        self.take(extension_size)?;
        Ok(if value == "\0" { None } else { Some(value) })
    }

    fn formula(&mut self) -> Result<Option<Formula>> {
        let size = self.u16()? as usize;
        self.take(2)?;
        let tokens = self.take(size)?;
        Ok((!tokens.is_empty()).then(|| Formula {
            tokens: tokens.to_vec(),
            rendered: None,
        }))
    }
}
