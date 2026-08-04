//! Semantic XLSB defined-name values.

use super::{Error, Result};

/// A workbook or sheet-scoped defined name.
#[derive(Debug, Clone)]
pub struct Definition {
    /// Defined-name text.
    pub name: String,
    /// Raw `NameParsedFormula.rgce` bytes.
    pub formula: Option<Vec<u8>>,
    /// Zero-based workbook sheet index, or `None` for workbook scope.
    pub sheet_id: Option<u32>,
    /// Whether the name is hidden.
    pub hidden: bool,
    /// Whether the name is a macro/function definition.
    pub function: bool,
}

impl Definition {
    /// Create a workbook or sheet-scoped defined name without a formula.
    #[must_use]
    pub fn new(name: String, sheet_id: Option<u32>) -> Self {
        Self {
            name,
            formula: None,
            sheet_id,
            hidden: false,
            function: false,
        }
    }

    /// Set the raw `NameParsedFormula.rgce` bytes.
    #[must_use]
    pub fn with_formula(mut self, formula: Vec<u8>) -> Self {
        self.formula = Some(formula);
        self
    }

    /// Set the hidden flag.
    #[must_use]
    pub fn with_hidden(mut self, hidden: bool) -> Self {
        self.hidden = hidden;
        self
    }
}

/// Validate the `XLNameWideString` grammar used by XLSB defined names.
pub fn validate_name(name: &str) -> Result<()> {
    let utf16_len = name.encode_utf16().count();
    if utf16_len == 0 || utf16_len > 255 {
        return Err(Error::InvalidFormula(format!(
            "defined name length {utf16_len} is outside 1..=255"
        )));
    }
    let mut chars = name.chars();
    let first = chars.next().expect("checked non-empty defined name");
    if !is_name_start(first) || !chars.all(is_name_character) {
        return Err(Error::InvalidFormula(format!(
            "defined name {name:?} does not follow XLNameWideString grammar"
        )));
    }
    if name.eq_ignore_ascii_case("TRUE") || name.eq_ignore_ascii_case("FALSE") {
        return Err(Error::InvalidFormula(format!(
            "defined name {name:?} is a reserved Boolean literal"
        )));
    }
    if is_a1_reference(name) || starts_with_r1c1_reference(name) {
        return Err(Error::InvalidFormula(format!(
            "defined name {name:?} conflicts with a cell reference"
        )));
    }
    Ok(())
}

fn is_name_start(value: char) -> bool {
    value == '_' || value == '\\' || value.is_ascii_alphabetic() || value.is_alphabetic()
}

fn is_name_character(value: char) -> bool {
    is_name_start(value) || matches!(value, '?' | '\u{061F}' | '.') || value.is_numeric()
}

fn is_a1_reference(value: &str) -> bool {
    let bytes = value.as_bytes();
    let split = bytes
        .iter()
        .position(u8::is_ascii_digit)
        .unwrap_or(bytes.len());
    if split == 0 || split > 3 || split == bytes.len() {
        return false;
    }
    if !bytes[..split].iter().all(u8::is_ascii_alphabetic)
        || !bytes[split..].iter().all(u8::is_ascii_digit)
    {
        return false;
    }
    let mut column = 0u32;
    for byte in bytes[..split].iter().map(u8::to_ascii_uppercase) {
        let Some(next) = column
            .checked_mul(26)
            .and_then(|column| column.checked_add(u32::from(byte - b'A' + 1)))
        else {
            return false;
        };
        column = next;
    }
    let Some(row) = value[split..].parse::<u32>().ok() else {
        return false;
    };
    column <= 16_384 && (1..=1_048_576).contains(&row)
}

fn starts_with_r1c1_reference(value: &str) -> bool {
    let bytes = value.as_bytes();
    let Some(first) = bytes.first().copied().map(|byte| byte.to_ascii_uppercase()) else {
        return false;
    };
    match first {
        b'R' => numeric_reference_prefix(bytes, 1, 1_048_576).is_some(),
        b'C' => numeric_reference_prefix(bytes, 1, 16_384).is_some(),
        _ => false,
    }
}

fn numeric_reference_prefix(bytes: &[u8], offset: usize, maximum: u32) -> Option<usize> {
    let end = bytes[offset..]
        .iter()
        .position(|byte| !byte.is_ascii_digit())
        .map_or(bytes.len(), |length| offset + length);
    if end == offset {
        return None;
    }
    let value = std::str::from_utf8(&bytes[offset..end])
        .ok()?
        .parse::<u32>()
        .ok()?;
    (1..=maximum).contains(&value).then_some(end)
}
