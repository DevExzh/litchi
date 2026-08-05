//! Bounded XLSB Brt* conditional-formatting record codec.
//!
//! Every reader consumes a bounded payload and every writer validates the
//! semantic model before emitting a record.

use crate::formula::{MAX_CELL_FORMULA_BYTES, ParsedFormula, Resolution};
use crate::raw::{Writer, kind};
use std::collections::HashSet;
use std::io::Write;

use super::super::model::*;
use super::semantic::{
    EmptyFormulaResolution, effective_cfvo_formula, effective_rule_formulas,
    effective_rule_parameter, format_number, icon_count, icon_count14, render_formula,
    validate_boundary_thresholds, validate_data_bar14, validate_extension_links,
    validate_extension14_template, validate_formula_count, validate_formula_slots,
    validate_icon_set14, validate_parameter_and_flags, validate_rule_metadata, validate_rule_text,
    validate_scale_thresholds, validate_scale_thresholds14, validate_template,
};
use super::{Error, Result, invalid};

// -----------------------------------------------------------------------------
// Owner-local range and Future Record Type helpers.
//
// These helpers intentionally live at the XLSB boundary. They model the
// BinRangeList and FRTHeader structures used by [MS-XLSB] §§2.2.6.2.1,
// 2.4.23--2.4.24, 2.4.43--2.4.44, 2.4.91--2.4.92, 2.4.332--2.4.335,
// 2.4.380--2.4.381, 2.4.399--2.4.400, 2.4.445--2.4.446, 2.5.19--2.5.20,
// and 2.5.98.7.
// -----------------------------------------------------------------------------

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct CellRange {
    row_first: u32,
    row_last: u32,
    col_first: u32,
    col_last: u32,
}

impl CellRange {
    fn new(row_first: u32, row_last: u32, col_first: u32, col_last: u32) -> Self {
        Self {
            row_first,
            row_last,
            col_first,
            col_last,
        }
    }

    fn parse(value: &str) -> Result<Self> {
        let value = value.trim();
        let (first, last) = value.split_once(':').unwrap_or((value, value));
        let (row_first, col_first) = parse_cell_reference(first)?;
        let (row_last, col_last) = parse_cell_reference(last)?;
        Ok(Self::new(row_first, row_last, col_first, col_last))
    }

    fn write<W: Write>(&self, writer: &mut W) -> Result<()> {
        writer.write_all(&self.row_first.to_le_bytes())?;
        writer.write_all(&self.row_last.to_le_bytes())?;
        writer.write_all(&self.col_first.to_le_bytes())?;
        writer.write_all(&self.col_last.to_le_bytes())?;
        Ok(())
    }
}

fn parse_range_list(value: &str) -> Result<Vec<CellRange>> {
    value
        .split([',', ' '])
        .filter(|part| !part.is_empty())
        .map(CellRange::parse)
        .collect()
}

fn write_bin_range_list<W: Write>(ranges: &[CellRange], writer: &mut W) -> Result<()> {
    let count = i32::try_from(ranges.len())
        .map_err(|_| invalid("BinRangeList", "range count overflows i32"))?;
    writer.write_all(&count.to_le_bytes())?;
    for range in ranges {
        range.write(writer)?;
    }
    Ok(())
}

fn column_index_to_name(mut column: u32) -> String {
    if column == 0 {
        return String::new();
    }
    let mut result = String::new();
    while column > 0 {
        column -= 1;
        result.insert(0, char::from(b'A' + (column % 26) as u8));
        column /= 26;
    }
    result
}

fn cell_reference(row: u32, column: u32) -> String {
    let Some(column) = column.checked_add(1) else {
        return format!("R{row}C{column}");
    };
    let Some(row) = row.checked_add(1) else {
        return format!("R{row}C{column}");
    };
    format!("{}{}", column_index_to_name(column), row)
}

fn parse_cell_reference(value: &str) -> Result<(u32, u32)> {
    let normalized = value.trim().to_ascii_uppercase();
    let mut column = String::new();
    let mut row = String::new();
    let mut digit_seen = false;
    for character in normalized.chars() {
        if character.is_ascii_alphabetic() {
            if digit_seen {
                return Err(Error::InvalidCellReference(normalized));
            }
            column.push(character);
        } else if character.is_ascii_digit() {
            digit_seen = true;
            row.push(character);
        } else {
            return Err(invalid("cell reference", normalized));
        }
    }
    if column.is_empty() || row.is_empty() {
        return Err(Error::InvalidCellReference(normalized));
    }
    let mut column_index = 0_u32;
    for character in column.bytes() {
        if !character.is_ascii_uppercase() {
            return Err(Error::InvalidCellReference(normalized));
        }
        column_index = column_index
            .checked_mul(26)
            .and_then(|value| value.checked_add(u32::from(character - b'A' + 1)))
            .ok_or_else(|| Error::InvalidCellReference(normalized.clone()))?;
    }
    let row_index = row
        .parse::<u32>()
        .map_err(|_| Error::InvalidCellReference(normalized.clone()))?;
    if row_index == 0 || column_index == 0 {
        return Err(Error::InvalidCellReference(normalized));
    }
    Ok((row_index - 1, column_index - 1))
}

type FrtRange = (u32, u32, u32, u32);

fn parse_sqref_header(
    data: &[u8],
    record: &'static str,
    maximum_ranges: usize,
) -> Result<(Vec<FrtRange>, usize)> {
    let mut cursor = FrtCursor::new(data, record);
    if cursor.read_u32()? != 0x02 {
        return Err(invalid(record, "FRTHeader is not sqref-only"));
    }
    if cursor.read_u32()? != 1 {
        return Err(invalid(record, "FRTSqrefs count is not 1"));
    }
    let flags = cursor.read_u32()?;
    if flags & 0x02 == 0 || flags & !0x0001_000f != 0 {
        return Err(invalid(
            record,
            format!("invalid FRTSqref flags 0x{flags:08X}"),
        ));
    }
    let count = usize::try_from(cursor.read_u32()? as i32)
        .map_err(|_| invalid(record, "NULL range collection"))?;
    if count == 0 || count > maximum_ranges || count > cursor.remaining() / 16 {
        return Err(invalid(record, format!("invalid range count {count}")));
    }
    let mut ranges = Vec::with_capacity(count);
    for _ in 0..count {
        let row_first = cursor.read_u32()?;
        let row_last = cursor.read_u32()?;
        let col_first = cursor.read_u32()?;
        let col_last = cursor.read_u32()?;
        if row_first > row_last
            || row_last >= 1_048_576
            || col_first > col_last
            || col_last >= 16_384
        {
            return Err(invalid(record, "invalid FRT target range"));
        }
        ranges.push((row_first, row_last, col_first, col_last));
    }
    Ok((ranges, cursor.offset))
}

fn serialize_sqref_header(ranges: &[FrtRange]) -> Result<Vec<u8>> {
    if ranges.is_empty() || ranges.len() > i32::MAX as usize {
        return Err(invalid(
            "FRTHeader",
            format!("invalid range count {}", ranges.len()),
        ));
    }
    let mut data = Vec::with_capacity(16 + ranges.len() * 16);
    data.extend_from_slice(&2u32.to_le_bytes());
    data.extend_from_slice(&1u32.to_le_bytes());
    data.extend_from_slice(&2u32.to_le_bytes());
    data.extend_from_slice(&(ranges.len() as u32).to_le_bytes());
    for &(row_first, row_last, col_first, col_last) in ranges {
        if row_first > row_last
            || row_last >= 1_048_576
            || col_first > col_last
            || col_last >= 16_384
        {
            return Err(invalid("FRTHeader", "invalid FRT target range"));
        }
        for value in [row_first, row_last, col_first, col_last] {
            data.extend_from_slice(&value.to_le_bytes());
        }
    }
    Ok(data)
}

pub(super) fn parse_formula_header(
    data: &[u8],
    record: &'static str,
    maximum_formulas: usize,
) -> Result<(Vec<ParsedFormula>, usize)> {
    let mut cursor = FrtCursor::new(data, record);
    let flags = cursor.read_u32()?;
    if flags & !0x04 != 0 {
        return Err(invalid(
            record,
            format!("invalid FRTHeader flags 0x{flags:08X}"),
        ));
    }
    let mut formulas = Vec::new();
    if flags & 0x04 != 0 {
        let count = usize::try_from(cursor.read_u32()?)
            .map_err(|_| invalid(record, "FRT formula count overflow"))?;
        if count == 0 || count > maximum_formulas {
            return Err(invalid(
                record,
                format!("FRT formula count {count} is outside 1..={maximum_formulas}"),
            ));
        }
        formulas
            .try_reserve(count)
            .map_err(|_| Error::Unrecognized {
                typ: record.to_string(),
                val: "formula allocation exceeds bounded capacity".to_string(),
            })?;
        for _ in 0..count {
            formulas.push(cursor.read_formula()?);
        }
    }
    Ok((formulas, cursor.offset))
}

fn serialize_formula_header(
    formulas: &[ParsedFormula],
    maximum_formulas: usize,
) -> Result<Vec<u8>> {
    if formulas.len() > maximum_formulas {
        return Err(invalid(
            "FRTHeader",
            format!(
                "formula count {} exceeds {maximum_formulas}",
                formulas.len()
            ),
        ));
    }
    let mut data = Vec::new();
    data.extend_from_slice(&if formulas.is_empty() { 0u32 } else { 4u32 }.to_le_bytes());
    if formulas.is_empty() {
        return Ok(data);
    }
    data.extend_from_slice(
        &u32::try_from(formulas.len())
            .map_err(|_| invalid("FRTHeader", "formula count overflow"))?
            .to_le_bytes(),
    );
    for formula in formulas {
        if formula.rgce.is_empty() || formula.rgce.len() > MAX_CELL_FORMULA_BYTES {
            return Err(Error::InvalidFormula(format!(
                "FRT formula token length {} is outside 1..={MAX_CELL_FORMULA_BYTES}",
                formula.rgce.len()
            )));
        }
        data.extend_from_slice(&2u32.to_le_bytes());
        data.extend_from_slice(
            &u32::try_from(formula.rgce.len())
                .map_err(|_| invalid("FRTFormula", "token length overflow"))?
                .to_le_bytes(),
        );
        data.extend_from_slice(
            &u32::try_from(formula.rgcb.len())
                .map_err(|_| invalid("FRTFormula", "ancillary length overflow"))?
                .to_le_bytes(),
        );
        data.extend_from_slice(&formula.rgce);
        data.extend_from_slice(&formula.rgcb);
    }
    Ok(data)
}

struct FrtCursor<'a> {
    data: &'a [u8],
    offset: usize,
    record: &'static str,
}

impl<'a> FrtCursor<'a> {
    pub(super) fn new(data: &'a [u8], record: &'static str) -> Self {
        Self {
            data,
            offset: 0,
            record,
        }
    }

    fn take(&mut self, count: usize) -> Result<&'a [u8]> {
        let end = self
            .offset
            .checked_add(count)
            .ok_or_else(|| invalid(self.record, "field size overflow"))?;
        let bytes = self
            .data
            .get(self.offset..end)
            .ok_or(Error::InvalidLength {
                expected: end,
                found: self.data.len(),
            })?;
        self.offset = end;
        Ok(bytes)
    }

    fn read_u32(&mut self) -> Result<u32> {
        let bytes = self.take(4)?;
        Ok(u32::from_le_bytes(
            bytes.try_into().expect("four-byte field"),
        ))
    }

    fn remaining(&self) -> usize {
        self.data.len().saturating_sub(self.offset)
    }

    fn read_formula(&mut self) -> Result<ParsedFormula> {
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
                "FRT formula token length {cce} is outside 1..={MAX_CELL_FORMULA_BYTES}"
            )));
        }
        Ok(ParsedFormula {
            rgce: self.take(cce)?.to_vec(),
            rgcb: self.take(cb)?.to_vec(),
        })
    }
}

pub fn parse_rule_extension_guid(data: &[u8]) -> Result<[u8; 16]> {
    if data.len() != 20 {
        return Err(Error::InvalidLength {
            expected: 20,
            found: data.len(),
        });
    }
    if data[..4] != [0; 4] {
        return Err(invalid("BrtCFRuleExt", "nonzero FRTBlank"));
    }
    Ok(data[4..].try_into().expect("sixteen-byte GUID"))
}

pub fn serialize_rule_extension_guid(guid: [u8; 16]) -> [u8; 20] {
    let mut data = [0; 20];
    data[4..].copy_from_slice(&guid);
    data
}

impl Value {
    pub fn parse(data: &[u8]) -> Result<Self> {
        if data.len() < 24 {
            return Err(Error::InvalidLength {
                expected: 24,
                found: data.len(),
            });
        }
        let context = EmptyFormulaResolution;
        Self::parse_with_context(data, (0, 0), &context)
    }

    pub fn parse_with_context(
        data: &[u8],
        base: (u32, u32),
        context: &impl Resolution,
    ) -> Result<Self> {
        let mut cursor = CfCursor::new(data, "BrtCFVO");
        let cfvo_type = u8::try_from(cursor.read_u32()?)
            .map_err(|_| invalid("BrtCFVO", "CFVO type overflow"))?;
        if !matches!(cfvo_type, 1 | 2 | 3 | 4 | 5 | 7) {
            return Err(invalid("BrtCFVO", format!("invalid type {cfvo_type}")));
        }
        let numeric_value = cursor.read_f64()?;
        if !numeric_value.is_finite() {
            return Err(invalid("BrtCFVO", "non-finite numeric parameter"));
        }
        if matches!(cfvo_type, 4 | 5) && !(0.0..=100.0).contains(&numeric_value) {
            return Err(invalid(
                "BrtCFVO",
                format!("percentage parameter {numeric_value} outside 0..=100"),
            ));
        }
        let save_greater_than_or_equal = cursor.read_bool32()?;
        let greater_than_or_equal = cursor.read_bool32()?;
        let declared_formula_size = cursor.read_u32()? as usize;
        let formula_binary = if declared_formula_size == 0 {
            None
        } else {
            let formula = cursor.read_formula()?;
            if formula.rgce.len() != declared_formula_size {
                return Err(invalid(
                    "BrtCFVO",
                    "declared formula size does not match token stream",
                ));
            }
            Some(formula)
        };
        cursor.finish()?;
        if matches!(cfvo_type, 2 | 3) && formula_binary.is_some() {
            return Err(invalid("BrtCFVO", "min/max threshold contains a formula"));
        }
        if cfvo_type == 7 && formula_binary.is_none() {
            return Err(invalid("BrtCFVO", "formula threshold omits its formula"));
        }
        let value = if let Some(formula) = &formula_binary {
            Some(render_formula(formula, base, context)?)
        } else if matches!(cfvo_type, 1 | 4 | 5) {
            Some(format_number(numeric_value))
        } else {
            None
        };
        Ok(Self {
            cfvo_type,
            value,
            numeric_value,
            save_greater_than_or_equal,
            greater_than_or_equal,
            formula_binary,
        })
    }

    /// Parse an Office 2013 `BrtCFVO14` record.
    pub fn parse_extension14(data: &[u8]) -> Result<Self> {
        let context = EmptyFormulaResolution;
        Self::parse_extension14_with_context(data, (0, 0), &context)
    }

    pub fn parse_extension14_with_context(
        data: &[u8],
        base: (u32, u32),
        context: &impl Resolution,
    ) -> Result<Self> {
        let (formulas, header_size) = parse_formula_header(data, "BrtCFVO14", 1)?;
        let mut cursor = CfCursor::new(&data[header_size..], "BrtCFVO14");
        let cfvo_type = u8::try_from(cursor.read_u32()?)
            .map_err(|_| invalid("BrtCFVO14", "CFVO type overflow"))?;
        if !matches!(cfvo_type, 1 | 2 | 3 | 4 | 5 | 7 | 8 | 9) {
            return Err(invalid("BrtCFVO14", format!("invalid type {cfvo_type}")));
        }
        let numeric_value = cursor.read_f64()?;
        if !numeric_value.is_finite() {
            return Err(invalid("BrtCFVO14", "non-finite numeric parameter"));
        }
        let save_greater_than_or_equal = cursor.read_bool32()?;
        let greater_than_or_equal = cursor.read_bool32()?;
        let declared_formula_size = cursor.read_u32()? as usize;
        cursor.finish()?;
        let formula_binary = formulas.into_iter().next();
        if formula_binary
            .as_ref()
            .map_or(0, |formula| formula.rgce.len())
            != declared_formula_size
        {
            return Err(invalid(
                "BrtCFVO14",
                "FRT formula and declared token size disagree",
            ));
        }
        if matches!(cfvo_type, 2 | 3 | 8 | 9) && formula_binary.is_some() {
            return Err(invalid(
                "BrtCFVO14",
                "automatic/min/max threshold contains a formula",
            ));
        }
        if cfvo_type == 7 && formula_binary.is_none() {
            return Err(invalid("BrtCFVO14", "formula threshold omits its formula"));
        }
        if formula_binary.is_none()
            && matches!(cfvo_type, 4 | 5)
            && !(0.0..=100.0).contains(&numeric_value)
        {
            return Err(invalid(
                "BrtCFVO14",
                format!("percentage parameter {numeric_value} outside 0..=100"),
            ));
        }
        let value = if let Some(formula) = &formula_binary {
            Some(render_formula(formula, base, context)?)
        } else if matches!(cfvo_type, 1 | 4 | 5) {
            Some(format_number(numeric_value))
        } else {
            None
        };
        Ok(Self {
            cfvo_type,
            value,
            numeric_value,
            save_greater_than_or_equal,
            greater_than_or_equal,
            formula_binary,
        })
    }

    /// Serialize an Office 2013 `BrtCFVO14` payload using its binary formula.
    pub fn serialize_extension14(&self) -> Result<Vec<u8>> {
        self.serialize_extension14_with(
            self.formula_binary.as_ref(),
            self.numeric_value,
            self.save_greater_than_or_equal,
        )
    }

    fn serialize_extension14_with(
        &self,
        formula_binary: Option<&ParsedFormula>,
        numeric_value: f64,
        save_greater_than_or_equal: bool,
    ) -> Result<Vec<u8>> {
        if !matches!(self.cfvo_type, 1 | 2 | 3 | 4 | 5 | 7 | 8 | 9) {
            return Err(invalid(
                "BrtCFVO14",
                format!("invalid type {}", self.cfvo_type),
            ));
        }
        if !numeric_value.is_finite() {
            return Err(invalid("BrtCFVO14", "non-finite numeric parameter"));
        }
        if formula_binary.is_none()
            && matches!(self.cfvo_type, 4 | 5)
            && !(0.0..=100.0).contains(&numeric_value)
        {
            return Err(invalid(
                "BrtCFVO14",
                format!("percentage parameter {} outside 0..=100", numeric_value),
            ));
        }
        if matches!(self.cfvo_type, 2 | 3 | 8 | 9) && formula_binary.is_some() {
            return Err(invalid(
                "BrtCFVO14",
                "automatic/min/max threshold contains a formula",
            ));
        }
        if self.cfvo_type == 7 && formula_binary.is_none() {
            return Err(invalid("BrtCFVO14", "formula threshold omits its formula"));
        }
        let formulas = formula_binary.map_or(&[][..], std::slice::from_ref);
        let mut data = serialize_formula_header(formulas, 1)?;
        data.extend_from_slice(&u32::from(self.cfvo_type).to_le_bytes());
        data.extend_from_slice(&numeric_value.to_le_bytes());
        data.extend_from_slice(&u32::from(save_greater_than_or_equal).to_le_bytes());
        data.extend_from_slice(&u32::from(self.greater_than_or_equal).to_le_bytes());
        data.extend_from_slice(
            &u32::try_from(formula_binary.map_or(0, |formula| formula.rgce.len()))
                .map_err(|_| Error::InvalidFormula("formula is too large".to_string()))?
                .to_le_bytes(),
        );
        Ok(data)
    }
}

impl Color {
    pub fn theme(index: u8, tint: i16) -> Result<Self> {
        if index > 0x0b {
            return Err(invalid("BrtColor", format!("theme color index {index}")));
        }
        let tint_bytes = tint.to_le_bytes();
        Ok(Self {
            color_type: 3,
            index,
            tint,
            argb: None,
            raw: [6, index, tint_bytes[0], tint_bytes[1], 0, 0, 0, 0],
        })
    }

    pub fn parse(data: &[u8]) -> Result<Self> {
        if data.len() != 8 {
            return Err(Error::InvalidLength {
                expected: 8,
                found: data.len(),
            });
        }
        let raw: [u8; 8] = data.try_into().map_err(|_| Error::InvalidLength {
            expected: 8,
            found: data.len(),
        })?;
        let color_type = raw[0] >> 1;
        if color_type > 3 {
            return Err(invalid("BrtColor", format!("color type {color_type}")));
        }
        let argb = if color_type == 2 {
            if raw[0] & 1 == 0 {
                return Err(invalid("BrtColor", "direct color is not marked valid"));
            }
            Some(
                (u32::from(raw[7]) << 24)
                    | (u32::from(raw[4]) << 16)
                    | (u32::from(raw[5]) << 8)
                    | u32::from(raw[6]),
            )
        } else {
            None
        };
        if color_type == 3 && raw[1] > 0x0b {
            return Err(invalid("BrtColor", format!("theme color index {}", raw[1])));
        }
        Ok(Self {
            color_type,
            index: raw[1],
            tint: i16::from_le_bytes([raw[2], raw[3]]),
            argb,
            raw,
        })
    }

    pub fn to_bytes(self) -> Result<[u8; 8]> {
        if self.color_type > 3 || (self.color_type == 3 && self.index > 0x0b) {
            return Err(invalid("BrtColor", "invalid color type or theme index"));
        }
        if self.color_type == 2 && self.argb.is_none() {
            return Err(invalid("BrtColor", "direct color has no ARGB value"));
        }
        if self.color_type != 2 && self.argb.is_some() {
            return Err(invalid("BrtColor", "non-direct color has an ARGB value"));
        }
        let parsed_raw = Self::parse(&self.raw).ok();
        if parsed_raw.as_ref().is_some_and(|raw| {
            raw.color_type == self.color_type
                && raw.index == self.index
                && raw.tint == self.tint
                && raw.argb == self.argb
        }) {
            return Ok(self.raw);
        }
        let tint = self.tint.to_le_bytes();
        let mut raw = [
            self.color_type << 1,
            self.index,
            tint[0],
            tint[1],
            0,
            0,
            0,
            0,
        ];
        if let Some(argb) = self.argb {
            raw[0] |= 1;
            raw[4] = ((argb >> 16) & 0xff) as u8;
            raw[5] = ((argb >> 8) & 0xff) as u8;
            raw[6] = (argb & 0xff) as u8;
            raw[7] = ((argb >> 24) & 0xff) as u8;
        }
        Ok(raw)
    }

    /// Parse an Office 2013 `BrtColor14` payload.
    pub fn parse_extension14(data: &[u8]) -> Result<Self> {
        if data.len() != 12 {
            return Err(Error::InvalidLength {
                expected: 12,
                found: data.len(),
            });
        }
        if data[..4] != [0; 4] {
            return Err(invalid("BrtColor14", "nonzero FRTBlank"));
        }
        Self::parse(&data[4..])
    }

    /// Serialize an Office 2013 `BrtColor14` payload.
    pub fn serialize_extension14(self) -> Result<[u8; 12]> {
        let mut data = [0; 12];
        data[4..].copy_from_slice(&self.to_bytes()?);
        Ok(data)
    }
}

impl Direction14 {
    fn from_u8(value: u8) -> Option<Self> {
        match value {
            0 => Some(Self::Context),
            1 => Some(Self::LeftToRight),
            2 => Some(Self::RightToLeft),
            _ => None,
        }
    }
}

impl AxisPosition14 {
    fn from_u8(value: u8) -> Option<Self> {
        match value {
            0 => Some(Self::Automatic),
            1 => Some(Self::Midpoint),
            2 => Some(Self::None),
            _ => None,
        }
    }
}

impl Bar14 {
    pub fn parse_header(data: &[u8]) -> Result<BarHeader14> {
        let mut cursor = CfCursor::new(data, "BrtBeginDatabar14");
        if cursor.read_u32()? != 0 {
            return Err(invalid("BrtBeginDatabar14", "nonzero FRTBlank"));
        }
        let min_length = cursor.read_u8()?;
        let max_length = cursor.read_u8()?;
        let show_value = cursor.read_bool8()?;
        let direction = Direction14::from_u8(cursor.read_u8()?)
            .ok_or_else(|| invalid("BrtBeginDatabar14", "invalid direction"))?;
        let axis_position = AxisPosition14::from_u8(cursor.read_u8()?)
            .ok_or_else(|| invalid("BrtBeginDatabar14", "invalid axis position"))?;
        let flags = cursor.read_u16()?;
        cursor.finish()?;
        if min_length > max_length || max_length > 100 {
            return Err(invalid(
                "BrtBeginDatabar14",
                "invalid minimum/maximum length",
            ));
        }
        Ok(BarHeader14 {
            min_length,
            max_length,
            show_value,
            direction,
            axis_position,
            border: flags & 0x01 != 0,
            gradient: flags & 0x02 != 0,
            custom_negative_fill: flags & 0x04 != 0,
            custom_negative_border: flags & 0x08 != 0,
            unused_flags: flags & 0xfff0,
        })
    }

    pub fn serialize_header(&self) -> Result<Vec<u8>> {
        if self.min_length > self.max_length
            || self.max_length > 100
            || self.unused_flags & 0x0f != 0
        {
            return Err(invalid("BrtBeginDatabar14", "invalid data-bar header"));
        }
        let mut flags = self.unused_flags;
        flags |= u16::from(self.border);
        flags |= u16::from(self.gradient) << 1;
        flags |= u16::from(self.custom_negative_fill) << 2;
        flags |= u16::from(self.custom_negative_border) << 3;
        let mut data = Vec::with_capacity(11);
        data.extend_from_slice(&0u32.to_le_bytes());
        data.extend_from_slice(&[
            self.min_length,
            self.max_length,
            u8::from(self.show_value),
            self.direction as u8,
            self.axis_position as u8,
        ]);
        data.extend_from_slice(&flags.to_le_bytes());
        Ok(data)
    }
}

impl Icon {
    pub fn parse(data: &[u8]) -> Result<Self> {
        let mut cursor = CfCursor::new(data, "BrtCFIcon");
        if cursor.read_u32()? != 0 {
            return Err(invalid("BrtCFIcon", "nonzero FRTBlank"));
        }
        let value = Self {
            icon_set: cursor.read_i32()?,
            index: cursor.read_i32()?,
        };
        cursor.finish()?;
        value.validate()?;
        Ok(value)
    }

    pub fn serialize(self) -> Result<[u8; 12]> {
        self.validate()?;
        let mut data = [0; 12];
        data[4..8].copy_from_slice(&self.icon_set.to_le_bytes());
        data[8..].copy_from_slice(&self.index.to_le_bytes());
        Ok(data)
    }

    fn validate(self) -> Result<()> {
        if self.icon_set == -1 {
            if self.index == -1 {
                return Ok(());
            }
        } else if let Ok(icon_set) = u8::try_from(self.icon_set)
            && icon_set <= 19
            && (0..icon_count14(icon_set) as i32).contains(&self.index)
        {
            return Ok(());
        }
        Err(invalid("BrtCFIcon", "invalid icon set or index"))
    }
}

impl IconSet14 {
    pub fn parse_header(data: &[u8]) -> Result<IconHeader14> {
        let mut cursor = CfCursor::new(data, "BrtBeginIconSet14");
        if cursor.read_u32()? != 0 {
            return Err(invalid("BrtBeginIconSet14", "nonzero FRTBlank"));
        }
        let icon_set_type = u8::try_from(cursor.read_u32()?)
            .map_err(|_| invalid("BrtBeginIconSet14", "icon-set type overflow"))?;
        if icon_set_type > 19 {
            return Err(invalid("BrtBeginIconSet14", "invalid icon-set type"));
        }
        let flags = cursor.read_u16()?;
        cursor.finish()?;
        if flags & 0xff80 != 0 {
            return Err(invalid("BrtBeginIconSet14", "reserved flags are nonzero"));
        }
        Ok(IconHeader14 {
            icon_set_type,
            custom: flags & 0x01 != 0,
            show_value: flags & 0x02 == 0,
            reverse: flags & 0x04 == 0,
            unused_flags: flags & 0x78,
        })
    }

    pub fn serialize_header(&self) -> Result<Vec<u8>> {
        if self.icon_set_type > 19 || self.unused_flags & !0x78 != 0 {
            return Err(invalid("BrtBeginIconSet14", "invalid icon-set header"));
        }
        let mut flags = self.unused_flags;
        flags |= u16::from(self.custom_icons.is_some());
        flags |= u16::from(!self.show_value) << 1;
        flags |= u16::from(!self.reverse) << 2;
        let mut data = Vec::with_capacity(10);
        data.extend_from_slice(&0u32.to_le_bytes());
        data.extend_from_slice(&u32::from(self.icon_set_type).to_le_bytes());
        data.extend_from_slice(&flags.to_le_bytes());
        Ok(data)
    }
}

impl Rule {
    pub fn parse(data: &[u8]) -> Result<Self> {
        let context = EmptyFormulaResolution;
        Self::parse_with_context(data, (0, 0), &context)
    }

    pub fn parse_with_context(
        data: &[u8],
        base: (u32, u32),
        context: &impl Resolution,
    ) -> Result<Self> {
        let mut cursor = CfCursor::new(data, "BrtBeginCFRule");
        let rule_type_raw = cursor.read_u32()?;
        let rule_type = RuleType::from_u32(rule_type_raw).ok_or_else(|| {
            invalid(
                "BrtBeginCFRule",
                format!("invalid rule type {rule_type_raw}"),
            )
        })?;
        let template = cursor.read_u32()?;
        validate_template(rule_type, template)?;
        let raw_dxf = cursor.read_u32()?;
        let dxf_id = (raw_dxf != u32::MAX).then_some(raw_dxf);
        if matches!(
            rule_type,
            RuleType::ColorScale | RuleType::DataBar | RuleType::IconSet
        ) && dxf_id.is_some()
        {
            return Err(invalid(
                "BrtBeginCFRule",
                "visual rule has a differential-format index",
            ));
        }
        let priority = cursor.read_u32()?;
        if priority == 0 || priority > i32::MAX as u32 {
            return Err(invalid(
                "BrtBeginCFRule",
                format!("invalid priority {priority}"),
            ));
        }
        let parameter = cursor.read_u32()?;
        let reserved1 = cursor.read_u32()?;
        let reserved2 = cursor.read_u32()?;
        let flags = cursor.read_u16()?;
        if reserved1 != 0 || reserved2 != 0 || flags & !0x1e != 0 {
            return Err(invalid("BrtBeginCFRule", "reserved fields are nonzero"));
        }
        let stop_if_true = flags & 0x02 != 0;
        let above_average = flags & 0x04 != 0;
        let bottom = flags & 0x08 != 0;
        let percent = flags & 0x10 != 0;
        if matches!(
            rule_type,
            RuleType::ColorScale | RuleType::DataBar | RuleType::IconSet
        ) && stop_if_true
        {
            return Err(invalid("BrtBeginCFRule", "visual rule sets stop-if-true"));
        }
        if rule_type != RuleType::TopN && (bottom || percent) {
            return Err(invalid(
                "BrtBeginCFRule",
                "non-filter rule sets bottom/percent flags",
            ));
        }
        validate_parameter_and_flags(
            rule_type,
            template,
            parameter,
            above_average,
            bottom,
            percent,
        )?;
        let declared = [cursor.read_u32()?, cursor.read_u32()?, cursor.read_u32()?];
        let text = cursor.read_nullable_string()?;
        if template == 8 {
            if text
                .as_ref()
                .is_none_or(|text| text.is_empty() || text.encode_utf16().count() > 255)
            {
                return Err(invalid(
                    "BrtBeginCFRule",
                    "contains-text template has an invalid text parameter",
                ));
            }
        } else if text.is_some() {
            return Err(invalid(
                "BrtBeginCFRule",
                "non-text template has a string parameter",
            ));
        }
        let mut formula_slots: [Option<ParsedFormula>; 3] = [None, None, None];
        for (index, size) in declared.into_iter().enumerate() {
            if size == 0 {
                continue;
            }
            let formula = cursor.read_formula()?;
            if formula.rgce.len() != size as usize {
                return Err(invalid(
                    "BrtBeginCFRule",
                    format!(
                        "formula {} declared {size} token bytes, found {}",
                        index + 1,
                        formula.rgce.len()
                    ),
                ));
            }
            formula_slots[index] = Some(formula);
        }
        cursor.finish()?;
        validate_formula_slots(rule_type, template, parameter, &formula_slots)?;

        let mut formulas = Vec::new();
        let mut formula_extras = Vec::new();
        let mut formula_texts = Vec::new();
        for formula in formula_slots.into_iter().flatten() {
            formulas.push(formula.rgce.clone());
            formula_extras.push(formula.rgcb.clone());
            formula_texts.push(render_formula(&formula, base, context)?);
        }
        let operator = (rule_type == RuleType::CellIs)
            .then(|| u8::try_from(parameter).ok())
            .flatten();
        if rule_type == RuleType::CellIs && !matches!(operator, Some(1..=8)) {
            return Err(invalid(
                "BrtBeginCFRule",
                format!("invalid cell comparison operator {parameter}"),
            ));
        }

        Ok(Rule {
            rule_type,
            dxf_id,
            priority,
            stop_if_true,
            formulas,
            formula_extras,
            formula_texts,
            color_scale: None,
            data_bar: None,
            icon_set: None,
            color_scale14: None,
            data_bar14: None,
            icon_set14: None,
            operator,
            parameter,
            template,
            text,
            above_average,
            bottom,
            percent,
            extension14: None,
            classic_extension_guid: None,
        })
    }

    /// Parse an Office 2013 `BrtBeginCFRule14` payload.
    pub fn parse_extension14(data: &[u8]) -> Result<Self> {
        let context = EmptyFormulaResolution;
        Self::parse_extension14_with_context(data, (0, 0), &context)
    }

    pub fn parse_extension14_with_context(
        data: &[u8],
        base: (u32, u32),
        context: &impl Resolution,
    ) -> Result<Self> {
        let (formulas, header_size) = parse_formula_header(data, "BrtBeginCFRule14", 2)?;
        let mut cursor = CfCursor::new(&data[header_size..], "BrtBeginCFRule14");
        let rule_type_raw = cursor.read_u32()?;
        let rule_type = RuleType::from_u32(rule_type_raw).ok_or_else(|| {
            invalid(
                "BrtBeginCFRule14",
                format!("invalid rule type {rule_type_raw}"),
            )
        })?;
        let template = cursor.read_u32()?;
        validate_extension14_template(rule_type, template)?;
        let raw_dxf = cursor.read_u32()?;
        let signed_priority = cursor.read_i32()?;
        if signed_priority != -1 && signed_priority <= 0 {
            return Err(invalid(
                "BrtBeginCFRule14",
                format!("invalid priority {signed_priority}"),
            ));
        }
        if signed_priority == -1 && (rule_type != RuleType::DataBar || raw_dxf != 0) {
            return Err(invalid(
                "BrtBeginCFRule14",
                "priority -1 requires a data-bar rule and zero DXF index",
            ));
        }
        let visual = matches!(
            rule_type,
            RuleType::ColorScale | RuleType::DataBar | RuleType::IconSet
        );
        if signed_priority > 0 && visual && raw_dxf != u32::MAX {
            return Err(invalid(
                "BrtBeginCFRule14",
                "visual rule has a differential-format index",
            ));
        }
        let dxf_id = if signed_priority == -1 || raw_dxf == u32::MAX {
            None
        } else {
            Some(raw_dxf)
        };
        let parameter = cursor.read_u32()?;
        let reserved1 = cursor.read_u32()?;
        let reserved2 = cursor.read_u32()?;
        let flags = cursor.read_u16()?;
        if reserved1 != 0 || reserved2 != 0 || flags & !0x1e != 0 {
            return Err(invalid("BrtBeginCFRule14", "reserved fields are nonzero"));
        }
        let stop_if_true = flags & 0x02 != 0;
        let above_average = flags & 0x04 != 0;
        let bottom = flags & 0x08 != 0;
        let percent = flags & 0x10 != 0;
        if visual && stop_if_true {
            return Err(invalid("BrtBeginCFRule14", "visual rule sets stop-if-true"));
        }
        validate_parameter_and_flags(
            rule_type,
            template,
            parameter,
            above_average,
            bottom,
            percent,
        )?;
        let declared = [cursor.read_u32()?, cursor.read_u32()?, cursor.read_u32()?];
        let unused = cursor.read_u32()?;
        let guid = cursor.read_array::<16>()?;
        let guid_present = cursor.read_bool32()?;
        let text = cursor.read_nullable_string()?;
        cursor.finish()?;

        if template == 8 {
            if text
                .as_ref()
                .is_none_or(|text| text.is_empty() || text.encode_utf16().count() > 255)
            {
                return Err(invalid(
                    "BrtBeginCFRule14",
                    "contains-text template has an invalid text parameter",
                ));
            }
        } else if text.is_some() {
            return Err(invalid(
                "BrtBeginCFRule14",
                "non-text template has a string parameter",
            ));
        }

        let mut formula_slots: [Option<ParsedFormula>; 3] = [None, None, None];
        let mut formula_iter = formulas.into_iter();
        for (index, declared_size) in declared.into_iter().enumerate() {
            if declared_size == 0 {
                continue;
            }
            let formula = formula_iter.next().ok_or_else(|| {
                invalid(
                    "BrtBeginCFRule14",
                    "declared formula is absent from FRTHeader",
                )
            })?;
            if formula.rgce.len() != declared_size as usize {
                return Err(invalid(
                    "BrtBeginCFRule14",
                    format!(
                        "formula {} declared {declared_size} token bytes, found {}",
                        index + 1,
                        formula.rgce.len()
                    ),
                ));
            }
            formula_slots[index] = Some(formula);
        }
        if formula_iter.next().is_some() {
            return Err(invalid(
                "BrtBeginCFRule14",
                "FRTHeader contains an undeclared formula",
            ));
        }
        validate_formula_slots(rule_type, template, parameter, &formula_slots)?;

        let mut binary_formulas = Vec::new();
        let mut formula_extras = Vec::new();
        let mut formula_texts = Vec::new();
        for formula in formula_slots.into_iter().flatten() {
            binary_formulas.push(formula.rgce.clone());
            formula_extras.push(formula.rgcb.clone());
            formula_texts.push(render_formula(&formula, base, context)?);
        }
        let operator = (rule_type == RuleType::CellIs)
            .then(|| u8::try_from(parameter).ok())
            .flatten();

        Ok(Self {
            rule_type,
            dxf_id,
            priority: u32::try_from(signed_priority).unwrap_or(0),
            stop_if_true,
            formulas: binary_formulas,
            formula_extras,
            formula_texts,
            color_scale: None,
            data_bar: None,
            icon_set: None,
            color_scale14: None,
            data_bar14: None,
            icon_set14: None,
            operator,
            parameter,
            template,
            text,
            above_average,
            bottom,
            percent,
            extension14: Some(RuleMetadata {
                priority: signed_priority,
                unused,
                guid,
                guid_present,
                linked_classic_priority: None,
            }),
            classic_extension_guid: None,
        })
    }

    /// Serialize an Office 2013 `BrtBeginCFRule14` payload.
    pub fn serialize_extension14(&self) -> Result<Vec<u8>> {
        let metadata = self.extension14.ok_or_else(|| {
            invalid(
                "BrtBeginCFRule14",
                "rule does not contain Office 2013 metadata",
            )
        })?;
        validate_extension14_template(self.rule_type, self.template)?;
        if metadata.priority != -1 && metadata.priority <= 0 {
            return Err(invalid(
                "BrtBeginCFRule14",
                format!("invalid priority {}", metadata.priority),
            ));
        }
        if metadata.priority > 0 && self.priority != metadata.priority as u32 {
            return Err(invalid(
                "BrtBeginCFRule14",
                "classic and extension priorities disagree",
            ));
        }
        if metadata.priority == -1 && self.rule_type != RuleType::DataBar {
            return Err(invalid(
                "BrtBeginCFRule14",
                "priority -1 is only valid for a data-bar extension",
            ));
        }
        let parameter = effective_rule_parameter(self)?;
        validate_parameter_and_flags(
            self.rule_type,
            self.template,
            parameter,
            self.above_average,
            self.bottom,
            self.percent,
        )?;
        let visual = matches!(
            self.rule_type,
            RuleType::ColorScale | RuleType::DataBar | RuleType::IconSet
        );
        if visual && (self.stop_if_true || (metadata.priority > 0 && self.dxf_id.is_some())) {
            return Err(invalid(
                "BrtBeginCFRule14",
                "visual rule has a DXF or stop-if-true flag",
            ));
        }
        if metadata.priority == -1 && self.dxf_id.is_some() {
            return Err(invalid(
                "BrtBeginCFRule14",
                "data-bar extension has a DXF index",
            ));
        }
        validate_rule_text(self.template, self.text.as_deref(), "BrtBeginCFRule14")?;

        let formulas = effective_rule_formulas(self)?;
        validate_formula_count(self.rule_type, self.template, parameter, formulas.len())?;
        let mut slots: [Option<&ParsedFormula>; 3] = [None, None, None];
        let start = if visual { 2 } else { 0 };
        for (index, formula) in formulas.iter().enumerate() {
            slots[start + index] = Some(formula);
        }
        let owned_slots = slots.each_ref().map(|formula| formula.cloned());
        validate_formula_slots(self.rule_type, self.template, parameter, &owned_slots)?;

        let mut payload = serialize_formula_header(&formulas, 2)?;
        payload.extend_from_slice(&(self.rule_type as u32).to_le_bytes());
        payload.extend_from_slice(&self.template.to_le_bytes());
        let raw_dxf = if metadata.priority == -1 {
            0
        } else {
            self.dxf_id.unwrap_or(u32::MAX)
        };
        payload.extend_from_slice(&raw_dxf.to_le_bytes());
        payload.extend_from_slice(&metadata.priority.to_le_bytes());
        payload.extend_from_slice(&parameter.to_le_bytes());
        payload.extend_from_slice(&0u32.to_le_bytes());
        payload.extend_from_slice(&0u32.to_le_bytes());
        let mut flags = 0u16;
        flags |= u16::from(self.stop_if_true) << 1;
        flags |= u16::from(self.above_average) << 2;
        flags |= u16::from(self.bottom) << 3;
        flags |= u16::from(self.percent) << 4;
        payload.extend_from_slice(&flags.to_le_bytes());
        for formula in &slots {
            payload.extend_from_slice(
                &u32::try_from(formula.map_or(0, |formula| formula.rgce.len()))
                    .map_err(|_| Error::InvalidFormula("formula is too large".to_string()))?
                    .to_le_bytes(),
            );
        }
        payload.extend_from_slice(&metadata.unused.to_le_bytes());
        payload.extend_from_slice(&metadata.guid);
        payload.extend_from_slice(&u32::from(metadata.guid_present).to_le_bytes());
        write_nullable_string(&mut payload, self.text.as_deref())?;
        Ok(payload)
    }
}

impl Formatting {
    /// Parse an Office 2013 `BrtBeginConditionalFormatting14` payload.
    pub fn parse_extension14_header(data: &[u8]) -> Result<(Self, u32)> {
        let (formatting, count, _) = Self::parse_extension14_header_with_base(data)?;
        Ok((formatting, count))
    }

    pub fn parse_extension14_header_with_base(data: &[u8]) -> Result<(Self, u32, (u32, u32))> {
        let (ranges, header_size) =
            parse_sqref_header(data, "BrtBeginConditionalFormatting14", i32::MAX as usize)?;
        let mut cursor = CfCursor::new(&data[header_size..], "BrtBeginConditionalFormatting14");
        let count = cursor.read_u32()?;
        let pivot_only = cursor.read_bool32()?;
        cursor.finish()?;
        let base = (ranges[0].0, ranges[0].2);
        let ranges = ranges
            .into_iter()
            .map(|(first_row, last_row, first_col, last_col)| {
                let first = cell_reference(first_row, first_col);
                let last = cell_reference(last_row, last_col);
                if first == last {
                    first
                } else {
                    format!("{first}:{last}")
                }
            })
            .collect();
        Ok((
            Self {
                ranges,
                rules: Vec::new(),
                pivot_only,
                record_kind: RecordKind::Extension14,
            },
            count,
            base,
        ))
    }

    /// Serialize an Office 2013 `BrtBeginConditionalFormatting14` payload.
    pub fn serialize_extension14_header(&self) -> Result<Vec<u8>> {
        let mut ranges = Vec::new();
        for range_list in &self.ranges {
            for range in range_list
                .split([',', ' '])
                .filter(|range| !range.is_empty())
            {
                let (first, last) = range.split_once(':').unwrap_or((range, range));
                let (first_row, first_col) = parse_cell_reference(first)?;
                let (last_row, last_col) = parse_cell_reference(last)?;
                ranges.push((first_row, last_row, first_col, last_col));
            }
        }
        let mut data = serialize_sqref_header(&ranges)?;
        data.extend_from_slice(
            &u32::try_from(self.rules.len())
                .map_err(|_| invalid("BrtBeginConditionalFormatting14", "rule count overflow"))?
                .to_le_bytes(),
        );
        data.extend_from_slice(&u32::from(self.pivot_only).to_le_bytes());
        Ok(data)
    }
}

pub fn parse_classic_header(data: &[u8]) -> Result<(Formatting, u32, (u32, u32))> {
    let mut cursor = CfCursor::new(data, "BrtBeginConditionalFormatting");
    let count = cursor.read_u32()?;
    let pivot_only = cursor.read_bool32()?;
    let ranges = cursor.read_ranges(1, 8_192)?;
    cursor.finish()?;
    let base = (ranges[0].0, ranges[0].2);
    let ranges = ranges
        .into_iter()
        .map(|(first_row, last_row, first_col, last_col)| {
            let first = cell_reference(first_row, first_col);
            let last = cell_reference(last_row, last_col);
            if first == last {
                first
            } else {
                format!("{first}:{last}")
            }
        })
        .collect();
    Ok((
        Formatting {
            ranges,
            rules: Vec::new(),
            pivot_only,
            record_kind: RecordKind::Classic,
        },
        count,
        base,
    ))
}

fn write_nullable_string(data: &mut Vec<u8>, value: Option<&str>) -> Result<()> {
    let Some(value) = value else {
        data.extend_from_slice(&u32::MAX.to_le_bytes());
        return Ok(());
    };
    let units = value.encode_utf16().collect::<Vec<_>>();
    data.extend_from_slice(
        &u32::try_from(units.len())
            .map_err(|_| invalid("XLNullableWideString", "string length overflow"))?
            .to_le_bytes(),
    );
    for unit in units {
        data.extend_from_slice(&unit.to_le_bytes());
    }
    Ok(())
}

pub(super) struct CfCursor<'a> {
    data: &'a [u8],
    offset: usize,
    record: &'static str,
}

impl<'a> CfCursor<'a> {
    pub(super) fn new(data: &'a [u8], record: &'static str) -> Self {
        Self {
            data,
            offset: 0,
            record,
        }
    }

    fn take(&mut self, size: usize) -> Result<&'a [u8]> {
        let end = self
            .offset
            .checked_add(size)
            .ok_or_else(|| invalid(self.record, "field size overflow"))?;
        let bytes = self
            .data
            .get(self.offset..end)
            .ok_or(Error::InvalidLength {
                expected: end,
                found: self.data.len(),
            })?;
        self.offset = end;
        Ok(bytes)
    }

    fn read_u16(&mut self) -> Result<u16> {
        let bytes = self.take(2)?;
        Ok(u16::from_le_bytes([bytes[0], bytes[1]]))
    }

    fn read_u8(&mut self) -> Result<u8> {
        Ok(self.take(1)?[0])
    }

    fn read_bool8(&mut self) -> Result<bool> {
        match self.read_u8()? {
            0 => Ok(false),
            1 => Ok(true),
            value => Err(invalid(self.record, format!("invalid Boolean {value}"))),
        }
    }

    fn read_u32(&mut self) -> Result<u32> {
        let bytes = self.take(4)?;
        Ok(u32::from_le_bytes(
            bytes.try_into().expect("four-byte field"),
        ))
    }

    fn read_i32(&mut self) -> Result<i32> {
        let bytes = self.take(4)?;
        Ok(i32::from_le_bytes(
            bytes.try_into().expect("four-byte field"),
        ))
    }

    fn read_array<const N: usize>(&mut self) -> Result<[u8; N]> {
        self.take(N)?.try_into().map_err(|_| Error::InvalidLength {
            expected: N,
            found: self.data.len().saturating_sub(self.offset),
        })
    }

    fn read_f64(&mut self) -> Result<f64> {
        let bytes = self.take(8)?;
        Ok(f64::from_le_bytes(
            bytes.try_into().expect("eight-byte field"),
        ))
    }

    fn read_bool32(&mut self) -> Result<bool> {
        match self.read_u32()? {
            0 => Ok(false),
            1 => Ok(true),
            value => Err(invalid(self.record, format!("invalid Boolean {value}"))),
        }
    }

    pub(super) fn read_nullable_string(&mut self) -> Result<Option<String>> {
        let count = self.read_u32()?;
        if count == u32::MAX {
            return Ok(None);
        }
        let count = count as usize;
        let bytes = self.take(
            count
                .checked_mul(2)
                .ok_or_else(|| invalid(self.record, "string size overflow"))?,
        )?;
        let units = bytes
            .chunks_exact(2)
            .map(|pair| u16::from_le_bytes([pair[0], pair[1]]))
            .collect::<Vec<_>>();
        String::from_utf16(&units)
            .map(Some)
            .map_err(|error| Error::Encoding(format!("invalid UTF-16: {error}")))
    }

    fn read_formula(&mut self) -> Result<ParsedFormula> {
        let cce = self.read_u32()? as usize;
        if cce == 0 || cce > MAX_CELL_FORMULA_BYTES {
            return Err(Error::InvalidFormula(format!(
                "conditional-format formula length {cce} is outside 1..={MAX_CELL_FORMULA_BYTES}"
            )));
        }
        let rgce = self.take(cce)?.to_vec();
        let cb = self.read_u32()? as usize;
        let rgcb = self.take(cb)?.to_vec();
        Ok(ParsedFormula { rgce, rgcb })
    }

    fn read_ranges(&mut self, minimum: usize, maximum: usize) -> Result<Vec<(u32, u32, u32, u32)>> {
        let raw_count = self.read_u32()? as i32;
        let count = usize::try_from(raw_count)
            .map_err(|_| invalid(self.record, "NULL range collection"))?;
        if !(minimum..=maximum).contains(&count)
            || count > self.data.len().saturating_sub(self.offset) / 16
        {
            return Err(invalid(self.record, format!("invalid range count {count}")));
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
                return Err(invalid(self.record, "invalid target range"));
            }
            ranges.push((first_row, last_row, first_col, last_col));
        }
        Ok(ranges)
    }

    pub(super) fn finish(self) -> Result<()> {
        if self.offset == self.data.len() {
            Ok(())
        } else {
            Err(Error::InvalidLength {
                expected: self.offset,
                found: self.data.len(),
            })
        }
    }
}

/// Write all classic and Office 2013 conditional-formatting collections for a worksheet.
pub fn write_conditional_formattings<W: Write>(
    writer: &mut Writer<W>,
    cond_fmts: &[Formatting],
) -> Result<()> {
    validate_extension_links(cond_fmts)?;
    let mut priorities = HashSet::new();
    for rule in cond_fmts.iter().flat_map(|formatting| &formatting.rules) {
        let priority = rule
            .extension14
            .map_or(i64::from(rule.priority), |metadata| {
                i64::from(metadata.priority)
            });
        if priority > 0 && !priorities.insert(priority) {
            return Err(invalid(
                "BrtBeginCFRule priority",
                format!("duplicate {priority}"),
            ));
        }
    }
    for formatting in cond_fmts {
        match formatting.record_kind {
            RecordKind::Classic => write_single_cond_formatting(writer, formatting)?,
            RecordKind::Extension14 => write_single_cond_formatting14(writer, formatting)?,
        }
    }
    Ok(())
}

fn write_single_cond_formatting<W: Write>(
    writer: &mut Writer<W>,
    formatting: &Formatting,
) -> Result<()> {
    writer.write_record(
        kind::BEGIN_COND_FORMATTING,
        &serialize_cond_formatting_header(formatting)?,
    )?;
    for rule in &formatting.rules {
        writer.write_record(kind::BEGIN_CF_RULE, &serialize_cf_rule(rule)?)?;
        write_rule_visualization(writer, rule)?;
        if let Some(guid) = rule.classic_extension_guid {
            writer.write_record(kind::CF_RULE_EXT, &serialize_rule_extension_guid(guid))?;
        }
        writer.write_record(kind::END_CF_RULE, &[])?;
    }
    writer.write_record(kind::END_COND_FORMATTING, &[])?;
    Ok(())
}

fn write_single_cond_formatting14<W: Write>(
    writer: &mut Writer<W>,
    formatting: &Formatting,
) -> Result<()> {
    writer.write_record(
        kind::BEGIN_COND_FORMATTING14,
        &formatting.serialize_extension14_header()?,
    )?;
    for rule in &formatting.rules {
        writer.write_record(kind::BEGIN_CF_RULE14, &rule.serialize_extension14()?)?;
        write_rule_visualization14(writer, rule)?;
        writer.write_record(kind::END_CF_RULE14, &[])?;
    }
    writer.write_record(kind::END_COND_FORMATTING14, &[])?;
    Ok(())
}

pub(super) fn serialize_cond_formatting_header(formatting: &Formatting) -> Result<Vec<u8>> {
    let rule_count = u32::try_from(formatting.rules.len())
        .map_err(|_| invalid("BrtBeginConditionalFormatting", "too many rules"))?;
    let mut ranges = Vec::new();
    for range in &formatting.ranges {
        ranges.extend(parse_range_list(range)?);
    }
    if ranges.is_empty() || ranges.len() > 8_192 {
        return Err(invalid(
            "BrtBeginConditionalFormatting",
            format!("classic range count {} is outside 1..=8192", ranges.len()),
        ));
    }
    let mut payload = Vec::with_capacity(12 + ranges.len() * 16);
    payload.extend_from_slice(&rule_count.to_le_bytes());
    payload.extend_from_slice(&u32::from(formatting.pivot_only).to_le_bytes());
    write_bin_range_list(&ranges, &mut payload)?;
    Ok(payload)
}

pub(super) fn serialize_cf_rule(rule: &Rule) -> Result<Vec<u8>> {
    validate_rule_metadata(rule)?;
    let parameter = effective_rule_parameter(rule)?;
    let formulas = effective_rule_formulas(rule)?;
    validate_formula_count(rule.rule_type, rule.template, parameter, formulas.len())?;

    let mut slots: [Option<&ParsedFormula>; 3] = [None, None, None];
    let start = if matches!(
        rule.rule_type,
        RuleType::ColorScale | RuleType::DataBar | RuleType::IconSet
    ) {
        2
    } else {
        0
    };
    for (index, formula) in formulas.iter().enumerate() {
        slots[start + index] = Some(formula);
    }

    let mut payload = Vec::with_capacity(64);
    payload.extend_from_slice(&(rule.rule_type as u32).to_le_bytes());
    payload.extend_from_slice(&rule.template.to_le_bytes());
    payload.extend_from_slice(&rule.dxf_id.unwrap_or(u32::MAX).to_le_bytes());
    payload.extend_from_slice(&rule.priority.to_le_bytes());
    payload.extend_from_slice(&parameter.to_le_bytes());
    payload.extend_from_slice(&0u32.to_le_bytes());
    payload.extend_from_slice(&0u32.to_le_bytes());
    let mut flags = 0u16;
    if rule.stop_if_true {
        flags |= 0x02;
    }
    if rule.above_average {
        flags |= 0x04;
    }
    if rule.bottom {
        flags |= 0x08;
    }
    if rule.percent {
        flags |= 0x10;
    }
    payload.extend_from_slice(&flags.to_le_bytes());
    for formula in &slots {
        let size = formula.map_or(0, |formula| formula.rgce.len());
        let size = u32::try_from(size)
            .map_err(|_| Error::InvalidFormula("formula is too large".to_string()))?;
        payload.extend_from_slice(&size.to_le_bytes());
    }
    write_nullable_wide_string(&mut payload, rule.text.as_deref())?;
    for formula in slots.into_iter().flatten() {
        payload.extend_from_slice(&formula.to_bytes()?);
    }
    Ok(payload)
}

fn write_rule_visualization<W: Write>(writer: &mut Writer<W>, rule: &Rule) -> Result<()> {
    match rule.rule_type {
        RuleType::ColorScale => {
            let scale = rule.color_scale.as_ref().expect("validated color scale");
            validate_scale_thresholds(scale)?;
            writer.write_record(kind::BEGIN_COLOR_SCALE, &[])?;
            write_cfvo(writer, &scale.min_cfvo, false)?;
            if let Some(midpoint) = &scale.mid_cfvo {
                write_cfvo(writer, midpoint, false)?;
            }
            write_cfvo(writer, &scale.max_cfvo, false)?;
            write_color(writer, scale.min_color_record, scale.min_color)?;
            if let (Some(record), Some(argb)) = (scale.mid_color_record, scale.mid_color) {
                write_color(writer, record, argb)?;
            }
            write_color(writer, scale.max_color_record, scale.max_color)?;
            writer.write_record(kind::END_COLOR_SCALE, &[])?;
        },
        RuleType::DataBar => {
            let bar = rule.data_bar.as_ref().expect("validated data bar");
            if bar.min_length > bar.max_length || bar.max_length > 100 {
                return Err(invalid("BrtBeginDatabar", "invalid minimum/maximum length"));
            }
            validate_boundary_thresholds(&bar.min_cfvo, &bar.max_cfvo, "BrtBeginDatabar")?;
            writer.write_record(
                kind::BEGIN_DATABAR,
                &[bar.min_length, bar.max_length, u8::from(bar.show_value)],
            )?;
            write_cfvo(writer, &bar.min_cfvo, false)?;
            write_cfvo(writer, &bar.max_cfvo, false)?;
            write_color(writer, bar.color_record, bar.color)?;
            writer.write_record(kind::END_DATABAR, &[])?;
        },
        RuleType::IconSet => {
            let set = rule.icon_set.as_ref().expect("validated icon set");
            let expected = icon_count(set.icon_set_type)?;
            if set.cfvos.len() != expected {
                return Err(invalid(
                    "BrtBeginIconSet",
                    format!("expected {expected} thresholds, found {}", set.cfvos.len()),
                ));
            }
            if set.cfvos.iter().any(|cfvo| matches!(cfvo.cfvo_type, 2 | 3)) {
                return Err(invalid(
                    "BrtBeginIconSet",
                    "min/max threshold is not allowed",
                ));
            }
            let mut flags = 0u16;
            if !set.show_value {
                flags |= 0x02;
            }
            if !set.reverse {
                flags |= 0x04;
            }
            let mut begin = Vec::with_capacity(6);
            begin.extend_from_slice(&u32::from(set.icon_set_type).to_le_bytes());
            begin.extend_from_slice(&flags.to_le_bytes());
            writer.write_record(kind::BEGIN_ICON_SET, &begin)?;
            for cfvo in &set.cfvos {
                write_cfvo(writer, cfvo, true)?;
            }
            writer.write_record(kind::END_ICON_SET, &[])?;
        },
        _ => {},
    }
    Ok(())
}

fn write_rule_visualization14<W: Write>(writer: &mut Writer<W>, rule: &Rule) -> Result<()> {
    if rule.color_scale.is_some() || rule.data_bar.is_some() || rule.icon_set.is_some() {
        return Err(invalid(
            "BrtBeginCFRule14",
            "classic visualization is set on an Office 2013 rule",
        ));
    }
    match rule.rule_type {
        RuleType::ColorScale => {
            let scale = rule
                .color_scale14
                .as_ref()
                .ok_or_else(|| invalid("BrtBeginCFRule14", "missing Office 2013 color scale"))?;
            if rule.data_bar14.is_some() || rule.icon_set14.is_some() {
                return Err(invalid(
                    "BrtBeginCFRule14",
                    "visualization does not match rule type",
                ));
            }
            validate_scale_thresholds14(scale)?;
            writer.write_record(kind::BEGIN_COLOR_SCALE14, &[])?;
            write_cfvo14(writer, &scale.min_cfvo, false)?;
            if let Some(midpoint) = &scale.mid_cfvo {
                write_cfvo14(writer, midpoint, false)?;
            }
            write_cfvo14(writer, &scale.max_cfvo, false)?;
            write_color14(writer, scale.min_color_record, scale.min_color)?;
            if let (Some(record), Some(argb)) = (scale.mid_color_record, scale.mid_color) {
                write_color14(writer, record, argb)?;
            }
            write_color14(writer, scale.max_color_record, scale.max_color)?;
            writer.write_record(kind::END_COLOR_SCALE14, &[])?;
        },
        RuleType::DataBar => {
            let bar = rule
                .data_bar14
                .as_ref()
                .ok_or_else(|| invalid("BrtBeginCFRule14", "missing Office 2013 data bar"))?;
            if rule.color_scale14.is_some() || rule.icon_set14.is_some() {
                return Err(invalid(
                    "BrtBeginCFRule14",
                    "visualization does not match rule type",
                ));
            }
            let priority = rule
                .extension14
                .ok_or_else(|| invalid("BrtBeginCFRule14", "missing extension metadata"))?
                .priority;
            validate_data_bar14(bar, priority)?;
            writer.write_record(kind::BEGIN_DATABAR14, &bar.serialize_header()?)?;
            write_cfvo14(writer, &bar.min_cfvo, false)?;
            write_cfvo14(writer, &bar.max_cfvo, false)?;
            for color in [
                bar.positive_color,
                bar.border_color,
                bar.negative_color,
                bar.negative_border_color,
                bar.axis_color,
            ]
            .into_iter()
            .flatten()
            {
                writer.write_record(kind::COLOR14, &color.serialize_extension14()?)?;
            }
            writer.write_record(kind::END_DATABAR14, &[])?;
        },
        RuleType::IconSet => {
            let set = rule
                .icon_set14
                .as_ref()
                .ok_or_else(|| invalid("BrtBeginCFRule14", "missing Office 2013 icon set"))?;
            if rule.color_scale14.is_some() || rule.data_bar14.is_some() {
                return Err(invalid(
                    "BrtBeginCFRule14",
                    "visualization does not match rule type",
                ));
            }
            validate_icon_set14(set)?;
            writer.write_record(kind::BEGIN_ICON_SET14, &set.serialize_header()?)?;
            for cfvo in &set.cfvos {
                write_cfvo14(writer, cfvo, true)?;
            }
            if let Some(icons) = &set.custom_icons {
                for icon in icons {
                    writer.write_record(kind::CF_ICON, &icon.serialize()?)?;
                }
            }
            writer.write_record(kind::END_ICON_SET14, &[])?;
        },
        _ => {
            if rule.color_scale14.is_some()
                || rule.data_bar14.is_some()
                || rule.icon_set14.is_some()
            {
                return Err(invalid(
                    "BrtBeginCFRule14",
                    "non-visual rule contains a visualization",
                ));
            }
        },
    }
    Ok(())
}

fn write_cfvo14<W: Write>(writer: &mut Writer<W>, cfvo: &Value, icon_set: bool) -> Result<()> {
    let formula = effective_cfvo_formula(cfvo)?;
    let numeric_value = if formula.is_none() && matches!(cfvo.cfvo_type, 1 | 4 | 5) {
        cfvo.value
            .as_deref()
            .and_then(|value| value.parse().ok())
            .unwrap_or(cfvo.numeric_value)
    } else {
        cfvo.numeric_value
    };
    writer.write_record(
        kind::CFVO14,
        &cfvo.serialize_extension14_with(
            formula.as_ref(),
            numeric_value,
            icon_set || cfvo.save_greater_than_or_equal,
        )?,
    )?;
    Ok(())
}

fn write_color14<W: Write>(writer: &mut Writer<W>, record: Color, legacy_argb: u32) -> Result<()> {
    let record = if record.argb == Some(legacy_argb) || (record.argb.is_none() && legacy_argb == 0)
    {
        record
    } else {
        Color::from_argb(legacy_argb)
    };
    writer.write_record(kind::COLOR14, &record.serialize_extension14()?)?;
    Ok(())
}

fn write_cfvo<W: Write>(writer: &mut Writer<W>, cfvo: &Value, icon_set: bool) -> Result<()> {
    if !matches!(cfvo.cfvo_type, 1 | 2 | 3 | 4 | 5 | 7) {
        return Err(invalid(
            "BrtCFVO",
            format!("invalid type {}", cfvo.cfvo_type),
        ));
    }
    let formula = effective_cfvo_formula(cfvo)?;
    if matches!(cfvo.cfvo_type, 2 | 3) && formula.is_some() {
        return Err(invalid("BrtCFVO", "min/max threshold contains a formula"));
    }
    if cfvo.cfvo_type == 7 && formula.is_none() {
        return Err(invalid("BrtCFVO", "formula threshold omits its formula"));
    }
    let numeric_value = if formula.is_none() && matches!(cfvo.cfvo_type, 1 | 4 | 5) {
        cfvo.value
            .as_deref()
            .and_then(|value| value.parse().ok())
            .unwrap_or(cfvo.numeric_value)
    } else {
        cfvo.numeric_value
    };
    if !numeric_value.is_finite()
        || (formula.is_none()
            && matches!(cfvo.cfvo_type, 4 | 5)
            && !(0.0..=100.0).contains(&numeric_value))
    {
        return Err(invalid("BrtCFVO", "invalid numeric parameter"));
    }
    let mut payload = Vec::with_capacity(32);
    payload.extend_from_slice(&u32::from(cfvo.cfvo_type).to_le_bytes());
    payload.extend_from_slice(&numeric_value.to_le_bytes());
    payload
        .extend_from_slice(&u32::from(icon_set || cfvo.save_greater_than_or_equal).to_le_bytes());
    payload.extend_from_slice(&u32::from(cfvo.greater_than_or_equal).to_le_bytes());
    let formula_size = formula.as_ref().map_or(0, |formula| formula.rgce.len());
    payload.extend_from_slice(
        &u32::try_from(formula_size)
            .map_err(|_| Error::InvalidFormula("formula is too large".to_string()))?
            .to_le_bytes(),
    );
    if let Some(formula) = formula {
        payload.extend_from_slice(&formula.to_bytes()?);
    }
    writer.write_record(kind::CFVO, &payload)?;
    Ok(())
}

fn write_color<W: Write>(writer: &mut Writer<W>, record: Color, legacy_argb: u32) -> Result<()> {
    let record = if record.argb == Some(legacy_argb) || (record.argb.is_none() && legacy_argb == 0)
    {
        record
    } else {
        Color::from_argb(legacy_argb)
    };
    writer.write_record(kind::COLOR, &record.to_bytes()?)?;
    Ok(())
}

fn write_nullable_wide_string(payload: &mut Vec<u8>, value: Option<&str>) -> Result<()> {
    let Some(value) = value else {
        payload.extend_from_slice(&u32::MAX.to_le_bytes());
        return Ok(());
    };
    let units = value.encode_utf16().count();
    payload.extend_from_slice(
        &u32::try_from(units)
            .map_err(|_| Error::Encoding("conditional-format text is too long".to_string()))?
            .to_le_bytes(),
    );
    payload.reserve(units.saturating_mul(2));
    for unit in value.encode_utf16() {
        payload.extend_from_slice(&unit.to_le_bytes());
    }
    Ok(())
}
