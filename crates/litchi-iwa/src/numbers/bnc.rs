//! Binary Numbers Cell (BNC) value storage.

use std::collections::BTreeMap;

use crate::{Error, Result};

pub(crate) const DECIMAL_FLAG: u32 = 0x000001;
pub(crate) const NUMBER_FLAG: u32 = 0x000002;
pub(crate) const DATE_FLAG: u32 = 0x000004;
pub(crate) const STRING_FLAG: u32 = 0x000008;
pub(crate) const RICH_TEXT_FLAG: u32 = 0x000010;
pub(crate) const FORMULA_FLAG: u32 = 0x000200;
pub(crate) const FORMULA_ERROR_FLAG: u32 = 0x000800;
pub(crate) const COMMENT_FLAG: u32 = 0x080000;

const VALUE_FLAGS: u32 = DECIMAL_FLAG
    | NUMBER_FLAG
    | DATE_FLAG
    | STRING_FLAG
    | RICH_TEXT_FLAG
    | FORMULA_FLAG
    | FORMULA_ERROR_FLAG;

pub(crate) const FIELD_LAYOUT: &[(u32, usize)] = &[
    (0x000001, 16),
    (0x000002, 8),
    (0x000004, 8),
    (0x000008, 4),
    (0x000010, 4),
    (0x000020, 4),
    (0x000040, 4),
    (0x000080, 4),
    (0x000100, 4),
    (0x000200, 4),
    (0x000400, 4),
    (0x000800, 4),
    (0x001000, 4),
    (0x002000, 4),
    (0x004000, 4),
    (0x008000, 4),
    (0x010000, 4),
    (0x020000, 4),
    (0x040000, 4),
    (0x080000, 4),
    (0x100000, 4),
];

#[derive(Debug, Clone)]
pub(crate) struct BncCell {
    prefix: [u8; 8],
    fields: BTreeMap<u32, Vec<u8>>,
    tail: Vec<u8>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum StoredValue {
    Empty,
    Number,
    Text(u32),
    Formula(u32),
    RichText(u32),
    Date,
    Boolean,
    Duration,
    Error,
    Unsupported(u8),
}

impl BncCell {
    pub(crate) fn parse(data: &[u8]) -> Result<Self> {
        if data.len() < 12 {
            return Err(Error::ParseError(
                "Truncated Numbers BNC cell header".to_string(),
            ));
        }
        if data[0] != 5 {
            return Err(Error::ParseError(format!(
                "Numbers cell storage version {} is not writable BNC v5",
                data[0]
            )));
        }

        let mut prefix = [0; 8];
        prefix.copy_from_slice(&data[..8]);
        let mut flag_bytes = [0; 4];
        flag_bytes.copy_from_slice(&data[8..12]);
        let flags = u32::from_le_bytes(flag_bytes);
        let known_flags = FIELD_LAYOUT.iter().fold(0, |mask, (flag, _)| mask | flag);
        if flags & !known_flags != 0 {
            return Err(Error::ParseError(format!(
                "Numbers BNC cell uses unknown flags 0x{:08x}",
                flags & !known_flags
            )));
        }

        let mut cursor = 12usize;
        let mut fields = BTreeMap::new();
        for &(flag, size) in FIELD_LAYOUT {
            if flags & flag == 0 {
                continue;
            }
            let end = cursor.checked_add(size).ok_or_else(|| {
                Error::ParseError("Numbers BNC field offset overflow".to_string())
            })?;
            let bytes = data.get(cursor..end).ok_or_else(|| {
                Error::ParseError(format!("Truncated Numbers BNC field 0x{flag:08x}"))
            })?;
            fields.insert(flag, bytes.to_vec());
            cursor = end;
        }

        Ok(Self {
            prefix,
            fields,
            tail: data[cursor..].to_vec(),
        })
    }

    pub(crate) fn minimal() -> Self {
        let mut prefix = [0; 8];
        prefix[0] = 5;
        Self {
            prefix,
            fields: BTreeMap::new(),
            tail: Vec::new(),
        }
    }

    pub(crate) fn stored_value(&self) -> StoredValue {
        if let Some(identifier) = self.u32_field(FORMULA_FLAG) {
            return StoredValue::Formula(identifier);
        }
        match self.prefix[1] {
            0 => StoredValue::Empty,
            2 | 10 => StoredValue::Number,
            3 => self
                .u32_field(STRING_FLAG)
                .map_or(StoredValue::Empty, StoredValue::Text),
            5 => StoredValue::Date,
            6 => StoredValue::Boolean,
            7 => StoredValue::Duration,
            8 => StoredValue::Error,
            9 => {
                if let Some(identifier) = self.u32_field(RICH_TEXT_FLAG) {
                    StoredValue::RichText(identifier)
                } else if let Some(identifier) = self.u32_field(STRING_FLAG) {
                    StoredValue::Text(identifier)
                } else {
                    StoredValue::Number
                }
            },
            other => StoredValue::Unsupported(other),
        }
    }

    pub(crate) fn set_number(&mut self, value: f64) -> Result<()> {
        if !value.is_finite() {
            return Err(Error::ParseError(
                "Numbers cells cannot store a non-finite numeric value".to_string(),
            ));
        }
        self.replace_value(2, DECIMAL_FLAG, decimal128_le(value)?.to_vec());
        Ok(())
    }

    pub(crate) fn set_boolean(&mut self, value: bool) {
        self.replace_value(
            6,
            NUMBER_FLAG,
            (if value { 1.0f64 } else { 0.0f64 }).to_le_bytes().to_vec(),
        );
    }

    pub(crate) fn set_duration(&mut self, value: f64) -> Result<()> {
        if !value.is_finite() {
            return Err(Error::ParseError(
                "Numbers cells cannot store a non-finite duration".to_string(),
            ));
        }
        self.replace_value(7, NUMBER_FLAG, value.to_le_bytes().to_vec());
        Ok(())
    }

    pub(crate) fn set_date(&mut self, value: f64) -> Result<()> {
        if !value.is_finite() {
            return Err(Error::ParseError(
                "Numbers cells cannot store a non-finite date".to_string(),
            ));
        }
        self.replace_value(5, DATE_FLAG, value.to_le_bytes().to_vec());
        Ok(())
    }

    pub(crate) fn set_string(&mut self, identifier: u32) {
        self.replace_value(3, STRING_FLAG, identifier.to_le_bytes().to_vec());
    }

    pub(crate) fn set_rich_text(&mut self, identifier: u32) {
        self.replace_value(9, RICH_TEXT_FLAG, identifier.to_le_bytes().to_vec());
    }

    pub(crate) fn set_formula_reference(&mut self, identifier: u32) {
        // Formula references coexist with the cached result value and its cell
        // type in app-generated BNC. The caller seeds a numeric cache before
        // attaching the formula when the target cell was empty.
        if self.prefix[1] == 0 {
            self.prefix[1] = 2;
            self.fields
                .insert(NUMBER_FLAG, 0.0f64.to_le_bytes().to_vec());
        }
        self.fields
            .insert(FORMULA_FLAG, identifier.to_le_bytes().to_vec());
        self.fields.remove(&FORMULA_ERROR_FLAG);
    }

    pub(crate) fn formula_error_identifier(&self) -> Option<u32> {
        self.u32_field(FORMULA_ERROR_FLAG)
    }

    pub(crate) fn comment_identifier(&self) -> Option<u32> {
        self.u32_field(COMMENT_FLAG)
    }

    pub(crate) fn set_comment_identifier(&mut self, identifier: Option<u32>) {
        if let Some(identifier) = identifier {
            self.fields
                .insert(COMMENT_FLAG, identifier.to_le_bytes().to_vec());
        } else {
            self.fields.remove(&COMMENT_FLAG);
        }
    }

    pub(crate) fn clear_value_preserving_metadata(&mut self) {
        self.prefix[1] = 0;
        self.fields.retain(|field, _| VALUE_FLAGS & field == 0);
    }

    pub(crate) fn encode(&self) -> Vec<u8> {
        let flags = self.fields.keys().fold(0u32, |mask, flag| mask | flag);
        let field_len = self.fields.values().map(Vec::len).sum::<usize>();
        let mut output = Vec::with_capacity(12 + field_len + self.tail.len());
        output.extend_from_slice(&self.prefix);
        output.extend_from_slice(&flags.to_le_bytes());
        for (flag, _) in FIELD_LAYOUT {
            if let Some(value) = self.fields.get(flag) {
                output.extend_from_slice(value);
            }
        }
        output.extend_from_slice(&self.tail);
        output
    }

    fn replace_value(&mut self, cell_type: u8, flag: u32, value: Vec<u8>) {
        self.prefix[1] = cell_type;
        self.fields.retain(|field, _| VALUE_FLAGS & field == 0);
        self.fields.insert(flag, value);
    }

    fn u32_field(&self, flag: u32) -> Option<u32> {
        let bytes: [u8; 4] = self.fields.get(&flag)?.as_slice().try_into().ok()?;
        Some(u32::from_le_bytes(bytes))
    }
}

/// Encode the finite `f64`'s shortest round-tripping decimal spelling into
/// the little-endian IEEE 754 decimal128 layout used by Numbers BNC cells and
/// formula AST compatibility fields.
pub(crate) fn decimal128_le(value: f64) -> Result<[u8; 16]> {
    if !value.is_finite() {
        return Err(Error::ParseError(
            "Numbers cannot encode a non-finite decimal value".to_owned(),
        ));
    }
    let negative = value.is_sign_negative();
    let magnitude = value.abs();
    let spelling = if magnitude == 0.0 {
        "0".to_owned()
    } else {
        magnitude.to_string()
    };
    let (mantissa, explicit_exponent) = spelling
        .split_once(['e', 'E'])
        .map_or((spelling.as_str(), 0), |(mantissa, exponent)| {
            (mantissa, exponent.parse::<i32>().unwrap_or(i32::MIN))
        });
    if explicit_exponent == i32::MIN {
        return Err(Error::ParseError(format!(
            "Could not encode Numbers decimal {spelling:?}"
        )));
    }
    let fractional_digits = mantissa
        .split_once('.')
        .map_or(0usize, |(_, fraction)| fraction.len());
    let mut digits = mantissa
        .bytes()
        .filter(|byte| *byte != b'.')
        .collect::<Vec<_>>();
    while digits.len() > 1 && digits.first() == Some(&b'0') {
        digits.remove(0);
    }
    let mut trailing_zeroes = 0i32;
    while digits.len() > 1 && digits.last() == Some(&b'0') {
        digits.pop();
        trailing_zeroes += 1;
    }
    let digits = std::str::from_utf8(&digits)
        .map_err(|_| Error::ParseError(format!("Could not encode Numbers decimal {spelling:?}")))?;
    let coefficient = digits
        .parse::<u128>()
        .map_err(|_| Error::ParseError(format!("Could not encode Numbers decimal {spelling:?}")))?;
    if coefficient >= (1u128 << 113) {
        return Err(Error::ParseError(
            "Numbers decimal coefficient exceeds 113 bits".to_owned(),
        ));
    }
    let fractional_digits = i32::try_from(fractional_digits)
        .map_err(|_| Error::ParseError("Numbers decimal exponent overflow".to_owned()))?;
    let exponent = explicit_exponent
        .checked_sub(fractional_digits)
        .and_then(|value| value.checked_add(trailing_zeroes))
        .ok_or_else(|| Error::ParseError("Numbers decimal exponent overflow".to_owned()))?;
    let biased_exponent = exponent
        .checked_add(0x1820)
        .filter(|value| (0..=0x3fff).contains(value))
        .ok_or_else(|| Error::ParseError("Numbers decimal exponent is out of range".to_owned()))?;
    let mut encoded = coefficient | ((biased_exponent as u128) << 113);
    if negative {
        encoded |= 1u128 << 127;
    }
    Ok(encoded.to_le_bytes())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn changes_value_without_changing_style_fields() {
        let original = hex("050300000000000008100200040000000500000001000000");
        let mut cell = BncCell::parse(&original).unwrap();
        assert_eq!(cell.stored_value(), StoredValue::Text(4));

        cell.set_number(42.5).unwrap();
        let encoded = cell.encode();
        let reparsed = BncCell::parse(&encoded).unwrap();
        assert_eq!(reparsed.stored_value(), StoredValue::Number);
        assert_eq!(reparsed.fields[&0x001000], 5u32.to_le_bytes());
        assert_eq!(reparsed.fields[&0x020000], 1u32.to_le_bytes());
        assert_eq!(reparsed.fields[&DECIMAL_FLAG], decimal128_le(42.5).unwrap());
    }

    #[test]
    fn rejects_unknown_flags_and_non_finite_numbers() {
        let mut data = vec![5, 2, 0, 0, 0, 0, 0, 0];
        data.extend_from_slice(&0x8000_0000u32.to_le_bytes());
        assert!(BncCell::parse(&data).is_err());
        assert!(BncCell::minimal().set_number(f64::NAN).is_err());
    }

    #[test]
    fn value_and_formula_replacement_clear_cached_formula_error_ids() {
        let mut cell = BncCell::minimal();
        cell.prefix[1] = 8;
        cell.fields
            .insert(FORMULA_ERROR_FLAG, 17u32.to_le_bytes().to_vec());
        assert_eq!(cell.formula_error_identifier(), Some(17));

        cell.set_number(1.0).unwrap();
        assert_eq!(cell.formula_error_identifier(), None);
        cell.fields
            .insert(FORMULA_ERROR_FLAG, 18u32.to_le_bytes().to_vec());
        cell.set_formula_reference(3);
        assert_eq!(cell.formula_error_identifier(), None);
        assert_eq!(cell.stored_value(), StoredValue::Formula(3));
    }

    #[test]
    fn comments_are_orthogonal_to_cell_values() {
        let mut cell = BncCell::minimal();
        cell.set_comment_identifier(Some(9));
        cell.set_string(3);
        assert_eq!(cell.comment_identifier(), Some(9));
        assert_eq!(cell.stored_value(), StoredValue::Text(3));

        cell.clear_value_preserving_metadata();
        assert_eq!(cell.stored_value(), StoredValue::Empty);
        assert_eq!(cell.comment_identifier(), Some(9));
        cell.set_comment_identifier(None);
        assert_eq!(cell.comment_identifier(), None);
    }

    fn hex(value: &str) -> Vec<u8> {
        value
            .as_bytes()
            .chunks_exact(2)
            .map(|pair| u8::from_str_radix(std::str::from_utf8(pair).unwrap(), 16).unwrap())
            .collect()
    }
}
