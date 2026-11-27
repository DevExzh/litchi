#![allow(
    clippy::cast_possible_truncation,
    clippy::cast_possible_wrap,
    clippy::expect_used,
    clippy::map_err_ignore,
    reason = "legacy module confines validated BIFF12 field narrowing or exact signed-bit reinterpretation, extraction after an immediately preceding structural invariant check, normalization into the module's stable typed public error to this codec boundary"
)]

//! Bounded XLSB wire helpers for conditional-formatting records.
//!
//! This owner contains the bounded cursor and FRT/BinRangeList primitives
//! shared by the typed conditional-formatting record codecs.

use crate::formula::{MAX_CELL_FORMULA_BYTES, ParsedFormula};
use std::io::Write;

use super::super::{Error, Result, invalid};

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
pub(super) struct CellRange {
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

pub(super) fn parse_range_list(value: &str) -> Result<Vec<CellRange>> {
    value
        .split([',', ' '])
        .filter(|part| !part.is_empty())
        .map(CellRange::parse)
        .collect()
}

pub(super) fn write_bin_range_list<W: Write>(ranges: &[CellRange], writer: &mut W) -> Result<()> {
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

pub(super) fn cell_reference(row: u32, column: u32) -> String {
    let Some(column) = column.checked_add(1) else {
        return format!("R{row}C{column}");
    };
    let Some(row) = row.checked_add(1) else {
        return format!("R{row}C{column}");
    };
    format!("{}{}", column_index_to_name(column), row)
}

pub(super) fn parse_cell_reference(value: &str) -> Result<(u32, u32)> {
    let normalized = value.trim().to_ascii_uppercase();
    let mut column = String::new();
    let mut row = String::new();
    let mut digit_seen = false;
    for character in normalized.chars() {
        if character.is_ascii_alphabetic() {
            if digit_seen {
                return Err(Error::InvalidCellReference(
                    "the cell reference is not valid",
                ));
            }
            column.push(character);
        } else if character.is_ascii_digit() {
            digit_seen = true;
            row.push(character);
        } else {
            return Err(invalid("cell reference", "unexpected value"));
        }
    }
    if column.is_empty() || row.is_empty() {
        return Err(Error::InvalidCellReference(
            "the cell reference is not valid",
        ));
    }
    let mut column_index = 0_u32;
    for character in column.bytes() {
        if !character.is_ascii_uppercase() {
            return Err(Error::InvalidCellReference(
                "the cell reference is not valid",
            ));
        }
        column_index = column_index
            .checked_mul(26)
            .and_then(|value| value.checked_add(u32::from(character - b'A' + 1)))
            .ok_or(Error::InvalidCellReference(
                "the cell reference is not valid",
            ))?;
    }
    let row_index = row
        .parse::<u32>()
        .map_err(|_| Error::InvalidCellReference("the cell reference is not valid"))?;
    if row_index == 0 || column_index == 0 {
        return Err(Error::InvalidCellReference(
            "the cell reference is not valid",
        ));
    }
    Ok((row_index - 1, column_index - 1))
}

pub(super) type FrtRange = (u32, u32, u32, u32);

pub(super) fn parse_sqref_header(
    data: &[u8],
    record: &'static str,
    maximum_ranges: usize,
    limits: crate::raw::Limits,
) -> Result<(Vec<FrtRange>, usize)> {
    let mut cursor = FrtCursor::new(data, record, limits)?;
    if cursor.read_u32()? != 0x02 {
        return Err(invalid(record, "FRTHeader is not sqref-only"));
    }
    if cursor.read_u32()? != 1 {
        return Err(invalid(record, "FRTSqrefs count is not 1"));
    }
    let flags = cursor.read_u32()?;
    if flags & 0x02 == 0 || flags & !0x0001_000f != 0 {
        return Err(invalid(record, "invalid FRTSqref flags"));
    }
    let count = usize::try_from(cursor.read_u32()? as i32)
        .map_err(|_| invalid(record, "NULL range collection"))?;
    if count == 0 || count > maximum_ranges || count > cursor.remaining() / 16 {
        return Err(invalid(record, "invalid range count"));
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

pub(super) fn serialize_sqref_header(
    ranges: &[FrtRange],
    limits: crate::raw::Limits,
) -> Result<Vec<u8>> {
    if ranges.is_empty() || ranges.len() > i32::MAX as usize {
        return Err(invalid("FRTHeader", "invalid range count"));
    }
    admit_payload(ranges.len().saturating_mul(16).saturating_add(16), limits)?;
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

pub(crate) fn parse_formula_header(
    data: &[u8],
    record: &'static str,
    maximum_formulas: usize,
    limits: crate::raw::Limits,
) -> Result<(Vec<ParsedFormula>, usize)> {
    let mut cursor = FrtCursor::new(data, record, limits)?;
    let flags = cursor.read_u32()?;
    if flags & !0x04 != 0 {
        return Err(invalid(record, "invalid FRTHeader flags"));
    }
    let mut formulas = Vec::new();
    if flags & 0x04 != 0 {
        let count = usize::try_from(cursor.read_u32()?)
            .map_err(|_| invalid(record, "FRT formula count overflow"))?;
        if count == 0 || count > maximum_formulas {
            return Err(invalid(record, "FRT formula count is outside 1..="));
        }
        formulas
            .try_reserve(count)
            .map_err(|_| Error::Unrecognized {
                typ: record,
                val: "formula allocation exceeds bounded capacity",
            })?;
        for _ in 0..count {
            formulas.push(cursor.read_formula()?);
        }
    }
    Ok((formulas, cursor.offset))
}

pub(super) fn serialize_formula_header(
    formulas: &[ParsedFormula],
    maximum_formulas: usize,
    limits: crate::raw::Limits,
) -> Result<Vec<u8>> {
    if formulas.len() > maximum_formulas {
        return Err(invalid("FRTHeader", "formula count exceeds"));
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
            return Err(Error::InvalidFormula(
                "FRT formula token length is outside its allowed range",
            ));
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
    admit_payload(data.len(), limits)?;
    Ok(data)
}

struct FrtCursor<'a> {
    data: &'a [u8],
    offset: usize,
    record: &'static str,
}

impl<'a> FrtCursor<'a> {
    pub(super) fn new(
        data: &'a [u8],
        record: &'static str,
        limits: crate::raw::Limits,
    ) -> Result<Self> {
        admit_payload(data.len(), limits)?;
        Ok(Self {
            data,
            offset: 0,
            record,
        })
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
            return Err(invalid(self.record, "invalid FRTFormula flags"));
        }
        let cce = usize::try_from(self.read_u32()?)
            .map_err(|_| invalid(self.record, "formula token length overflow"))?;
        let cb = usize::try_from(self.read_u32()?)
            .map_err(|_| invalid(self.record, "formula ancillary length overflow"))?;
        if cce == 0 || cce > MAX_CELL_FORMULA_BYTES {
            return Err(Error::InvalidFormula(
                "FRT formula token length is outside its allowed range",
            ));
        }
        Ok(ParsedFormula {
            rgce: self.take(cce)?.to_vec(),
            rgcb: self.take(cb)?.to_vec(),
        })
    }
}

/// Parse the fixed 20-byte `BrtCFRuleExt` payload under an explicit raw profile.
pub fn parse_rule_extension_guid(data: &[u8], limits: crate::raw::Limits) -> Result<[u8; 16]> {
    admit_payload(data.len(), limits)?;
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

pub(super) fn write_nullable_string(
    data: &mut Vec<u8>,
    value: Option<&str>,
    limits: crate::raw::Limits,
) -> Result<()> {
    let Some(value) = value else {
        data.extend_from_slice(&u32::MAX.to_le_bytes());
        return Ok(());
    };
    let units = value.encode_utf16().collect::<Vec<_>>();
    admit_units(units.len(), limits, 0)?;
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

pub(crate) struct CfCursor<'a> {
    data: &'a [u8],
    offset: usize,
    record: &'static str,
    limits: crate::raw::Limits,
}

impl<'a> CfCursor<'a> {
    pub(super) fn remaining(&self) -> usize {
        self.data.len().saturating_sub(self.offset)
    }

    /// Admit one record payload under the caller's raw profile.
    pub(crate) fn new(
        data: &'a [u8],
        record: &'static str,
        limits: crate::raw::Limits,
    ) -> Result<Self> {
        let limits = admit_payload(data.len(), limits)?;
        Ok(Self {
            data,
            offset: 0,
            record,
            limits,
        })
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

    pub(super) fn read_u16(&mut self) -> Result<u16> {
        let bytes = self.take(2)?;
        Ok(u16::from_le_bytes([bytes[0], bytes[1]]))
    }

    pub(super) fn read_u8(&mut self) -> Result<u8> {
        Ok(self.take(1)?[0])
    }

    pub(super) fn read_bool8(&mut self) -> Result<bool> {
        match self.read_u8()? {
            0 => Ok(false),
            1 => Ok(true),
            _value => Err(invalid(self.record, "invalid Boolean")),
        }
    }

    pub(super) fn read_u32(&mut self) -> Result<u32> {
        let bytes = self.take(4)?;
        Ok(u32::from_le_bytes(
            bytes.try_into().expect("four-byte field"),
        ))
    }

    pub(super) fn read_i32(&mut self) -> Result<i32> {
        let bytes = self.take(4)?;
        Ok(i32::from_le_bytes(
            bytes.try_into().expect("four-byte field"),
        ))
    }

    pub(super) fn read_array<const N: usize>(&mut self) -> Result<[u8; N]> {
        self.take(N)?.try_into().map_err(|_| Error::InvalidLength {
            expected: N,
            found: self.data.len().saturating_sub(self.offset),
        })
    }

    pub(super) fn read_f64(&mut self) -> Result<f64> {
        let bytes = self.take(8)?;
        Ok(f64::from_le_bytes(
            bytes.try_into().expect("eight-byte field"),
        ))
    }

    pub(super) fn read_bool32(&mut self) -> Result<bool> {
        match self.read_u32()? {
            0 => Ok(false),
            1 => Ok(true),
            _value => Err(invalid(self.record, "invalid Boolean")),
        }
    }

    pub(crate) fn read_nullable_string(&mut self) -> Result<Option<String>> {
        let count = self.read_u32()?;
        if count == u32::MAX {
            return Ok(None);
        }
        let count = count as usize;
        admit_units(count, self.limits, self.offset.saturating_sub(4))?;
        let bytes = self.take(
            count
                .checked_mul(2)
                .ok_or_else(|| invalid(self.record, "string size overflow"))?,
        )?;
        let units = bytes
            .as_chunks::<2>()
            .0
            .iter()
            .map(|pair| u16::from_le_bytes([pair[0], pair[1]]))
            .collect::<Vec<_>>();
        String::from_utf16(&units)
            .map(Some)
            .map_err(|_error| Error::Encoding("invalid UTF-16"))
    }

    pub(super) fn read_formula(&mut self) -> Result<ParsedFormula> {
        let cce = self.read_u32()? as usize;
        if cce == 0 || cce > MAX_CELL_FORMULA_BYTES {
            return Err(Error::InvalidFormula(
                "conditional-format formula length is outside its allowed range",
            ));
        }
        let rgce = self.take(cce)?.to_vec();
        let cb = self.read_u32()? as usize;
        let rgcb = self.take(cb)?.to_vec();
        Ok(ParsedFormula { rgce, rgcb })
    }

    pub(super) fn read_ranges(
        &mut self,
        minimum: usize,
        maximum: usize,
    ) -> Result<Vec<(u32, u32, u32, u32)>> {
        let raw_count = self.read_u32()? as i32;
        let count = usize::try_from(raw_count)
            .map_err(|_| invalid(self.record, "NULL range collection"))?;
        if !(minimum..=maximum).contains(&count)
            || count > self.data.len().saturating_sub(self.offset) / 16
        {
            return Err(invalid(self.record, "invalid range count"));
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

    pub(crate) fn finish(self) -> Result<()> {
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

/// Validate `limits` and admit one payload or output of `length` bytes.
pub(crate) fn admit_payload(
    length: usize,
    limits: crate::raw::Limits,
) -> Result<crate::raw::Limits> {
    Ok(limits.admit(length)?)
}

/// Refuse a declared or emitted string longer than the UTF-16-unit ceiling.
pub(crate) fn admit_units(units: usize, limits: crate::raw::Limits, offset: usize) -> Result<()> {
    Ok(limits.admit_units(units, offset)?)
}
