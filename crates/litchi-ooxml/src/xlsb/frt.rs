//! Bounded codecs for Future Record Type headers used by XLSB extensions.

use crate::xlsb::error::{XlsbError, XlsbResult};
use crate::xlsb::formula::{CellParsedFormula, MAX_CELL_FORMULA_BYTES};

pub(crate) type FrtRange = (u32, u32, u32, u32);

/// Parse the sqref-only FRTHeader used by `BrtBeginConditionalFormatting14`.
pub(crate) fn parse_sqref_header(
    data: &[u8],
    record: &'static str,
    maximum_ranges: usize,
) -> XlsbResult<(Vec<FrtRange>, usize)> {
    let mut cursor = FrtCursor::new(data, record);
    if cursor.read_u32()? != 0x02 {
        return Err(invalid(record, "FRTHeader is not sqref-only"));
    }
    if cursor.read_u32()? != 1 {
        return Err(invalid(record, "FRTSqrefs count is not 1"));
    }
    let sqref_flags = cursor.read_u32()?;
    if sqref_flags & 0x02 == 0 || sqref_flags & !0x0001_000f != 0 {
        return Err(invalid(
            record,
            format!("invalid FRTSqref flags 0x{sqref_flags:08X}"),
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

/// Serialize the sqref-only FRTHeader used by conditional formatting 14.
pub(crate) fn serialize_sqref_header(ranges: &[FrtRange]) -> XlsbResult<Vec<u8>> {
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

/// Parse an FRTHeader whose only optional field is `rgFormulas`.
///
/// Returns the formulas and the number of consumed bytes so the containing
/// record can continue parsing its fixed fields without copying the payload.
pub(crate) fn parse_formula_header(
    data: &[u8],
    record: &'static str,
    maximum_formulas: usize,
) -> XlsbResult<(Vec<CellParsedFormula>, usize)> {
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
        formulas.reserve(count);
        for _ in 0..count {
            formulas.push(cursor.read_formula()?);
        }
    }
    Ok((formulas, cursor.offset))
}

/// Serialize an FRTHeader whose only optional field is `rgFormulas`.
pub(crate) fn serialize_formula_header(
    formulas: &[CellParsedFormula],
    maximum_formulas: usize,
) -> XlsbResult<Vec<u8>> {
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
        validate_formula(formula)?;
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

fn validate_formula(formula: &CellParsedFormula) -> XlsbResult<()> {
    if formula.rgce.is_empty() || formula.rgce.len() > MAX_CELL_FORMULA_BYTES {
        return Err(XlsbError::InvalidFormula(format!(
            "FRT formula token length {} is outside 1..={MAX_CELL_FORMULA_BYTES}",
            formula.rgce.len()
        )));
    }
    Ok(())
}

struct FrtCursor<'a> {
    data: &'a [u8],
    offset: usize,
    record: &'static str,
}

impl<'a> FrtCursor<'a> {
    fn new(data: &'a [u8], record: &'static str) -> Self {
        Self {
            data,
            offset: 0,
            record,
        }
    }

    fn take(&mut self, count: usize) -> XlsbResult<&'a [u8]> {
        let end = self
            .offset
            .checked_add(count)
            .ok_or_else(|| invalid(self.record, "field size overflow"))?;
        let bytes = self
            .data
            .get(self.offset..end)
            .ok_or(XlsbError::InvalidLength {
                expected: end,
                found: self.data.len(),
            })?;
        self.offset = end;
        Ok(bytes)
    }

    fn read_u32(&mut self) -> XlsbResult<u32> {
        let bytes = self.take(4)?;
        Ok(u32::from_le_bytes(
            bytes.try_into().expect("four-byte field"),
        ))
    }

    fn remaining(&self) -> usize {
        self.data.len().saturating_sub(self.offset)
    }

    fn read_formula(&mut self) -> XlsbResult<CellParsedFormula> {
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
                "FRT formula token length {cce} is outside 1..={MAX_CELL_FORMULA_BYTES}"
            )));
        }
        Ok(CellParsedFormula {
            rgce: self.take(cce)?.to_vec(),
            rgcb: self.take(cb)?.to_vec(),
        })
    }
}

fn invalid(typ: impl Into<String>, val: impl Into<String>) -> XlsbError {
    XlsbError::Unrecognized {
        typ: typ.into(),
        val: val.into(),
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::xlsb::formula::FormulaCompiler;

    #[test]
    fn formula_header_roundtrips_token_and_ancillary_streams() {
        let formulas = [
            FormulaCompiler::compile("1+2").unwrap(),
            FormulaCompiler::compile("{1,2}").unwrap(),
        ];
        assert!(!formulas[1].rgcb.is_empty());
        let data = serialize_formula_header(&formulas, 2).unwrap();
        let (parsed, consumed) = parse_formula_header(&data, "test", 2).unwrap();
        assert_eq!(consumed, data.len());
        assert_eq!(parsed, formulas);
    }

    #[test]
    fn empty_formula_header_is_exact_frtblank_compatible_flags() {
        let data = serialize_formula_header(&[], 2).unwrap();
        assert_eq!(data, [0, 0, 0, 0]);
        assert_eq!(parse_formula_header(&data, "test", 2).unwrap(), (vec![], 4));
    }

    #[test]
    fn rejects_reserved_flags_counts_and_truncated_formulas() {
        assert!(parse_formula_header(&1u32.to_le_bytes(), "test", 2).is_err());
        let excessive = [4u32.to_le_bytes(), 3u32.to_le_bytes()].concat();
        assert!(parse_formula_header(&excessive, "test", 2).is_err());
        let truncated = [
            4u32.to_le_bytes(),
            1u32.to_le_bytes(),
            2u32.to_le_bytes(),
            5u32.to_le_bytes(),
            0u32.to_le_bytes(),
        ]
        .concat();
        assert!(parse_formula_header(&truncated, "test", 2).is_err());
    }

    #[test]
    fn sqref_header_roundtrips_more_than_classic_range_limit() {
        let ranges = vec![(0, 0, 0, 0); 8_193];
        let data = serialize_sqref_header(&ranges).unwrap();
        let (parsed, consumed) = parse_sqref_header(&data, "test", 10_000).unwrap();
        assert_eq!(consumed, data.len());
        assert_eq!(parsed, ranges);
    }

    #[test]
    fn sqref_header_rejects_reserved_flags_and_invalid_bounds() {
        let mut data = serialize_sqref_header(&[(0, 0, 0, 0)]).unwrap();
        data[8..12].copy_from_slice(&0x12u32.to_le_bytes());
        assert!(parse_sqref_header(&data, "test", 10).is_err());
        assert!(serialize_sqref_header(&[(0, 1_048_576, 0, 0)]).is_err());
    }
}
