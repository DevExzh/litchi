//! Named range support for XLSB

use crate::xlsb::error::{XlsbError, XlsbResult};
use crate::xlsb::formula::CellParsedFormula;
use crate::xlsb::formula::ptg_types;
use crate::xlsb::records::decode_string;
use litchi_core::binary;

/// Named range definition
///
/// Represents a defined name (named range) in the workbook.
#[derive(Debug, Clone)]
pub struct NamedRange {
    /// Name of the range
    pub name: String,
    /// Formula defining the range (the raw `NameParsedFormula.rgce` bytes)
    pub formula: Option<Vec<u8>>,
    /// Sheet ID (None for global scope)
    pub sheet_id: Option<u32>,
    /// Whether the name is hidden
    pub hidden: bool,
    /// Whether the name is a function
    pub function: bool,
}

impl NamedRange {
    /// Create a new named range
    ///
    /// # Example
    ///
    /// ```rust
    /// use litchi_ooxml::xlsb::named_ranges::NamedRange;
    ///
    /// let range = NamedRange::new("MyRange".to_string(), None);
    /// ```
    pub fn new(name: String, sheet_id: Option<u32>) -> Self {
        NamedRange {
            name,
            formula: None,
            sheet_id,
            hidden: false,
            function: false,
        }
    }

    /// Set formula bytes
    pub fn with_formula(mut self, formula: Vec<u8>) -> Self {
        self.formula = Some(formula);
        self
    }

    /// Set hidden flag
    pub fn with_hidden(mut self, hidden: bool) -> Self {
        self.hidden = hidden;
        self
    }

    /// Create a 3D area formula token stream for a workbook-local sheet range.
    ///
    /// The `sheet_id` is the zero-based workbook sheet index. XLSB formulas reference
    /// the workbook's self extern-sheet table, which reserves the first two
    /// entries for workbook and `#REF!`, so sheet references start at index 2.
    pub fn create_area3d_formula(
        sheet_id: u32,
        first_row: u32,
        last_row: u32,
        first_col: u16,
        last_col: u16,
    ) -> XlsbResult<Vec<u8>> {
        if first_row > last_row
            || last_row >= 1_048_576
            || first_col > last_col
            || last_col >= 16_384
        {
            return Err(XlsbError::InvalidCellReference(format!(
                "named range ({first_row}, {first_col})..=({last_row}, {last_col})"
            )));
        }
        let sheet_index = u16::try_from(sheet_id)
            .ok()
            .and_then(|value| value.checked_add(2))
            .ok_or_else(|| {
                XlsbError::InvalidFormula(format!(
                    "sheet index {sheet_id} cannot be represented in the extern-sheet table"
                ))
            })?;
        let mut formula = Vec::with_capacity(15);
        formula.push(ptg_types::PTG_AREA_3D);
        formula.extend_from_slice(&sheet_index.to_le_bytes());
        formula.extend_from_slice(&first_row.to_le_bytes());
        formula.extend_from_slice(&last_row.to_le_bytes());
        formula.extend_from_slice(&first_col.to_le_bytes());
        formula.extend_from_slice(&last_col.to_le_bytes());
        Ok(formula)
    }

    /// Parse from XLSB BrtName record
    pub fn parse(data: &[u8]) -> XlsbResult<Self> {
        if data.len() < 13 {
            return Err(XlsbError::InvalidLength {
                expected: 13,
                found: data.len(),
            });
        }

        let flags = binary::read_u32_le_at(data, 0)?;
        if flags & 0xFFFC_0000 != 0 {
            return Err(XlsbError::InvalidFormula(format!(
                "BrtName reserved flags are nonzero: 0x{flags:08X}"
            )));
        }
        let hidden = (flags & 0x0001) != 0;
        let f_func = (flags & 0x0002) != 0;
        let f_ob = (flags & 0x0004) != 0;
        let f_proc = (flags & 0x0008) != 0;
        let function = f_proc;
        if (f_func || f_ob) && !f_proc {
            return Err(XlsbError::InvalidFormula(
                "BrtName macro type requires fProc".to_string(),
            ));
        }
        let function_group = (flags >> 6) & 0x01FF;
        if !f_proc && function_group != 0 {
            return Err(XlsbError::InvalidFormula(
                "BrtName non-macro has a function group".to_string(),
            ));
        }
        let ch_key = data[4];
        if (f_func || !f_proc) && ch_key != 0 {
            return Err(XlsbError::InvalidFormula(format!(
                "BrtName has invalid macro shortcut key 0x{ch_key:02X}"
            )));
        }
        if f_proc && !f_func && ch_key < 0x20 {
            return Err(XlsbError::InvalidFormula(format!(
                "BrtName has invalid macro shortcut key 0x{ch_key:02X}"
            )));
        }

        // Sheet ID (-1 for global scope, otherwise sheet-specific)
        let sheet_id_raw = binary::read_u32_le_at(data, 5)? as i32;
        let sheet_id = if sheet_id_raw == -1 {
            None
        } else {
            Some(sheet_id_raw as u32)
        };

        let mut offset = 9;

        // Read name
        let (name, consumed) = decode_string(&data[offset..])?;
        offset += consumed;
        validate_defined_name(&name)?;

        // Read formula if present
        let (parsed_formula, consumed) = CellParsedFormula::parse(&data[offset..])?;
        offset += consumed;
        if data.len() < offset + 4 {
            return Err(XlsbError::InvalidLength {
                expected: offset + 4,
                found: data.len(),
            });
        }
        let (comment_len, consumed) = parse_nullable_wide_string(&data[offset..])?;
        offset += consumed;
        if comment_len.is_some_and(|length| length >= 256) {
            return Err(XlsbError::InvalidFormula(
                "BrtName comment exceeds 255 characters".to_string(),
            ));
        }
        if f_proc {
            for index in 0..4 {
                let (length, consumed) = parse_nullable_wide_string(&data[offset..])?;
                offset += consumed;
                if matches!(index, 0 | 3) && length.is_some() {
                    return Err(XlsbError::InvalidFormula(
                        "BrtName macro unused string is not NULL".to_string(),
                    ));
                }
                if matches!(index, 1 | 2) && length.is_some_and(|length| length >= 32_768) {
                    return Err(XlsbError::InvalidFormula(
                        "BrtName macro description exceeds 32,767 characters".to_string(),
                    ));
                }
            }
        }
        if offset != data.len() {
            return Err(XlsbError::InvalidFormula(format!(
                "BrtName has {} trailing bytes",
                data.len() - offset
            )));
        }
        let formula = Some(parsed_formula.rgce);

        Ok(NamedRange {
            name,
            formula,
            sheet_id,
            hidden,
            function,
        })
    }
}

/// Validate the `XLNameWideString` grammar used by XLSB defined names.
pub(crate) fn validate_defined_name(name: &str) -> XlsbResult<()> {
    let utf16_len = name.encode_utf16().count();
    if utf16_len == 0 || utf16_len > 255 {
        return Err(XlsbError::InvalidFormula(format!(
            "defined name length {utf16_len} is outside 1..=255"
        )));
    }
    let mut chars = name.chars();
    let first = chars.next().expect("checked non-empty defined name");
    if !is_name_start(first) || !chars.all(is_name_character) {
        return Err(XlsbError::InvalidFormula(format!(
            "defined name {name:?} does not follow XLNameWideString grammar"
        )));
    }
    if name.eq_ignore_ascii_case("TRUE") || name.eq_ignore_ascii_case("FALSE") {
        return Err(XlsbError::InvalidFormula(format!(
            "defined name {name:?} is a reserved Boolean literal"
        )));
    }
    if is_a1_reference(name) || starts_with_r1c1_reference(name) {
        return Err(XlsbError::InvalidFormula(format!(
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
    (value >= 1 && value <= maximum).then_some(end)
}

fn parse_nullable_wide_string(data: &[u8]) -> XlsbResult<(Option<usize>, usize)> {
    if data.len() < 4 {
        return Err(XlsbError::InvalidLength {
            expected: 4,
            found: data.len(),
        });
    }
    let length = binary::read_u32_le_at(data, 0)?;
    if length == u32::MAX {
        return Ok((None, 4));
    }
    let length = usize::try_from(length)
        .map_err(|_| XlsbError::Encoding("nullable string length overflow".to_string()))?;
    let consumed = length
        .checked_mul(2)
        .and_then(|byte_len| byte_len.checked_add(4))
        .ok_or_else(|| XlsbError::Encoding("nullable string length overflow".to_string()))?;
    if data.len() < consumed {
        return Err(XlsbError::InvalidLength {
            expected: consumed,
            found: data.len(),
        });
    }
    Ok((Some(length), consumed))
}

/// Create a 3D area formula token stream for a workbook-local sheet range.
pub fn create_area3d_formula(
    sheet_id: u32,
    first_row: u32,
    last_row: u32,
    first_col: u16,
    last_col: u16,
) -> XlsbResult<Vec<u8>> {
    NamedRange::create_area3d_formula(sheet_id, first_row, last_row, first_col, last_col)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn name_record(flags: u32, ch_key: u8, name: &str) -> Vec<u8> {
        let mut data = Vec::new();
        data.extend_from_slice(&flags.to_le_bytes());
        data.push(ch_key);
        data.extend_from_slice(&u32::MAX.to_le_bytes());
        let utf16: Vec<_> = name.encode_utf16().collect();
        data.extend_from_slice(&(utf16.len() as u32).to_le_bytes());
        for code_unit in utf16 {
            data.extend_from_slice(&code_unit.to_le_bytes());
        }
        data.extend_from_slice(
            &CellParsedFormula {
                rgce: vec![ptg_types::PTG_INT, 1, 0],
                rgcb: Vec::new(),
            }
            .to_bytes()
            .unwrap(),
        );
        data.extend_from_slice(&u32::MAX.to_le_bytes());
        data
    }

    #[test]
    fn test_named_range_builder() {
        let range = NamedRange::new("MyRange".to_string(), None)
            .with_hidden(true)
            .with_formula(vec![1, 2, 3]);

        assert_eq!(range.name, "MyRange");
        assert!(range.hidden);
        assert_eq!(range.formula, Some(vec![1, 2, 3]));
    }

    #[test]
    fn test_create_area3d_formula() {
        let formula = NamedRange::create_area3d_formula(0, 1, 3, 1, 1).unwrap();
        assert_eq!(formula[0], ptg_types::PTG_AREA_3D);
        assert_eq!(u16::from_le_bytes([formula[1], formula[2]]), 2);
        assert_eq!(u32::from_le_bytes(formula[3..7].try_into().unwrap()), 1);
        assert_eq!(u32::from_le_bytes(formula[7..11].try_into().unwrap()), 3);
        assert_eq!(u16::from_le_bytes([formula[11], formula[12]]), 1);
        assert_eq!(u16::from_le_bytes([formula[13], formula[14]]), 1);
    }

    #[test]
    fn validates_defined_name_grammar() {
        for name in ["SalesData", "_rate.2026", "\\Print_Area", "数据1"] {
            validate_defined_name(name).unwrap();
        }
        for name in [
            "",
            "1Sales",
            "Sales Data",
            "TRUE",
            "xfd1048576",
            "R1total",
            "C16384x",
        ] {
            assert!(validate_defined_name(name).is_err(), "accepted {name:?}");
        }
        validate_defined_name("XFE1").unwrap();
        validate_defined_name("R1048577total").unwrap();
        validate_defined_name("C16385x").unwrap();
    }

    #[test]
    fn parses_complete_brt_name_and_rejects_malformed_records() {
        let record = name_record(1, 0, "SalesData");
        let parsed = NamedRange::parse(&record).unwrap();
        assert_eq!(parsed.name, "SalesData");
        assert!(parsed.hidden);
        assert!(!parsed.function);
        assert_eq!(parsed.formula, Some(vec![ptg_types::PTG_INT, 1, 0]));

        let mut reserved = record.clone();
        reserved[3] = 0x80;
        assert!(NamedRange::parse(&reserved).is_err());

        let mut shortcut = record.clone();
        shortcut[4] = b'A';
        assert!(NamedRange::parse(&shortcut).is_err());

        let mut trailing = record.clone();
        trailing.push(0);
        assert!(NamedRange::parse(&trailing).is_err());

        let mut macro_record = name_record(0x0008, b'A', "MacroName");
        for _ in 0..4 {
            macro_record.extend_from_slice(&u32::MAX.to_le_bytes());
        }
        assert!(NamedRange::parse(&macro_record).unwrap().function);

        let mut truncated_comment = record;
        let comment_offset = truncated_comment.len() - 4;
        truncated_comment[comment_offset..].copy_from_slice(&1_u32.to_le_bytes());
        assert!(NamedRange::parse(&truncated_comment).is_err());
    }
}
