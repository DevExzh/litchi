//! Bounded BIFF8 `RgbExtra` framing for cell formulas.
//!
//! `RgbExtra` is not padding: its structures correspond, in order, to
//! particular Ptgs in the preceding token stream. This module only retains a
//! suffix after that correspondence and every structure in the suffix have
//! been checked.

use crate::{Error, Result};

use super::array::validate_ser_ar;
use super::model::Ancillary;
use super::validation::{FORMULA_FIXED_SIZE, FORMULA_RECORD_TYPE, MAX_FORMULA_PAYLOAD, invalid};

/// A Ptg that owns one structure in `RgbExtra`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum FormulaExtraKind {
    Array,
    Memory,
}

/// Scan token wire lengths and collect the required `RgbExtra` structures.
///
/// This deliberately does not validate the formula expression or its stack.
/// Existing Formula readers accept opaque token streams when no suffix is
/// present; callers invoke this scanner only when they have nonempty `RgbExtra`
/// bytes and need an unambiguous correspondence before retaining them. The
/// ordinary CellParsedFormula path has no revision context, so Name/NameX/3-D
/// Ptgs are framed as tokens without claiming revision-owned extras; an
/// actually present revision tail consequently fails exact consumption.
/// ELF-owned tails remain an explicit unsupported case.
pub(crate) fn scan_formula_extra_kinds(
    tokens: &[u8],
    record_type: u16,
) -> Result<Vec<FormulaExtraKind>> {
    scan_extra_kinds(tokens, record_type, true, true, true)
}

/// Scan list-object formula tokens while retaining the historical permissive
/// high-bit, array-type, and memory-expression framing rules.
pub(crate) fn scan_list_extra_kinds(
    tokens: &[u8],
    record_type: u16,
) -> Result<Vec<FormulaExtraKind>> {
    scan_extra_kinds(tokens, record_type, false, false, false)
}

fn scan_extra_kinds(
    tokens: &[u8],
    record_type: u16,
    reject_high_bit: bool,
    check_memory_cce: bool,
    reject_elf: bool,
) -> Result<Vec<FormulaExtraKind>> {
    let mut extras = Vec::new();
    let mut boundaries: Option<Vec<usize>> = None;
    let mut memory_ends = Vec::new();

    let mut position = 0usize;
    while position < tokens.len() {
        let opcode = tokens[position];
        if reject_high_bit && opcode & 0x80 != 0 {
            return Err(invalid_for(
                record_type,
                "Formula RgbExtra token has a reserved high bit",
            ));
        }
        let base = if opcode < 0x20 {
            opcode
        } else {
            (opcode & 0x1f) | 0x20
        };
        let size = match base {
            0x03..=0x16 => 1,
            0x17 => {
                let count =
                    usize::from(*tokens.get(position + 1).ok_or_else(|| {
                        invalid_for(record_type, "truncated formula string token")
                    })?);
                let flags = *tokens
                    .get(position + 2)
                    .ok_or_else(|| invalid_for(record_type, "truncated formula string flags"))?;
                if flags & !1 != 0 {
                    return Err(invalid_for(record_type, "unsupported formula string flags"));
                }
                3usize
                    .checked_add(
                        count
                            .checked_mul(if flags == 0 { 1 } else { 2 })
                            .ok_or_else(|| {
                                invalid_for(record_type, "formula string length overflows")
                            })?,
                    )
                    .ok_or_else(|| invalid_for(record_type, "formula string length overflows"))?
            },
            0x19 => {
                let header = tokens
                    .get(position..position + 4)
                    .ok_or_else(|| invalid_for(record_type, "truncated Attr token"))?;
                if header[1] & 0x04 != 0 {
                    4usize
                        .checked_add(
                            (usize::from(u16::from_le_bytes([header[2], header[3]])) + 1)
                                .checked_mul(2)
                                .ok_or_else(|| {
                                    invalid_for(record_type, "Attr token length overflows")
                                })?,
                        )
                        .ok_or_else(|| invalid_for(record_type, "Attr token length overflows"))?
                } else {
                    4
                }
            },
            0x1c | 0x1d => 2,
            0x1e => 3,
            0x1f => 9,
            0x20 => {
                if reject_high_bit && !matches!((opcode >> 5) & 0x03, 2 | 3) {
                    return Err(invalid_for(
                        record_type,
                        "PtgArray must have a value or array operand type",
                    ));
                }
                extras
                    .try_reserve(1)
                    .map_err(|_error| Error::Allocation("scanning Formula RgbExtra kinds"))?;
                extras.push(FormulaExtraKind::Array);
                8
            },
            0x21 => 3,
            0x22 => 4,
            // Name/NameX/3-D Ptgs are framed here as ordinary tokens. Their
            // revision-owned RgbExtra structures require context this reader
            // does not have, so any such unowned tail is rejected below by
            // exact extra consumption.
            0x23 => 5,
            0x24 | 0x2a | 0x2c => 5,
            0x25 | 0x2b | 0x2d => 9,
            0x26 => {
                extras
                    .try_reserve(1)
                    .map_err(|_error| Error::Allocation("scanning Formula RgbExtra kinds"))?;
                if check_memory_cce {
                    let cce = usize::from(u16::from_le_bytes([
                        *tokens
                            .get(position + 5)
                            .ok_or_else(|| invalid_for(record_type, "truncated PtgMemArea cce"))?,
                        *tokens
                            .get(position + 6)
                            .ok_or_else(|| invalid_for(record_type, "truncated PtgMemArea cce"))?,
                    ]));
                    if cce == 0 {
                        return Err(invalid_for(
                            record_type,
                            "PtgMemArea cce must contain a binary reference expression",
                        ));
                    }
                    let expression_end = position
                        .checked_add(7)
                        .and_then(|value| value.checked_add(cce))
                        .ok_or_else(|| {
                            invalid_for(record_type, "PtgMemArea expression length overflows")
                        })?;
                    if expression_end > tokens.len() {
                        return Err(invalid_for(
                            record_type,
                            "PtgMemArea cce exceeds the remaining token stream",
                        ));
                    }
                    if boundaries.is_none() {
                        let mut initial = Vec::new();
                        initial.try_reserve(1).map_err(|_error| {
                            Error::Allocation("scanning Formula token boundaries")
                        })?;
                        boundaries = Some(initial);
                    }
                    memory_ends
                        .try_reserve(1)
                        .map_err(|_error| Error::Allocation("scanning PtgMemArea ranges"))?;
                    memory_ends.push(expression_end);
                }
                extras.push(FormulaExtraKind::Memory);
                7
            },
            0x27 => 7,
            0x29 => 3,
            0x39..=0x3d => {
                if matches!(base, 0x3b | 0x3d) {
                    11
                } else {
                    7
                }
            },
            0x18 => {
                if reject_elf {
                    return Err(Error::UnsupportedFeature(
                        "Formula RgbExtra containing an ELF token is unsupported".to_string(),
                    ));
                }
                return Err(invalid_for(
                    record_type,
                    "invalid or unframeable token in Formula RgbExtra stream",
                ));
            },
            _ => {
                return Err(invalid_for(
                    record_type,
                    "invalid or unframeable token in Formula RgbExtra stream",
                ));
            },
        };
        position = position
            .checked_add(size)
            .ok_or_else(|| invalid_for(record_type, "formula token length overflows"))?;
        if position > tokens.len() {
            return Err(invalid_for(record_type, "truncated formula token"));
        }
        if let Some(boundaries) = boundaries.as_mut() {
            boundaries
                .try_reserve(1)
                .map_err(|_error| Error::Allocation("scanning Formula token boundaries"))?;
            boundaries.push(position);
        }
    }
    if let Some(boundaries) = boundaries {
        for expression_end in memory_ends {
            if boundaries.binary_search(&expression_end).is_err() {
                return Err(invalid_for(
                    record_type,
                    "PtgMemArea cce does not end on a token boundary",
                ));
            }
        }
    }
    Ok(extras)
}

/// Parse the structures corresponding to `tokens` from `data`.
///
/// The returned offset permits list-object callers to continue parsing their
/// enclosing payload. Formula callers pass a suffix slice and require the
/// returned offset to equal the slice length.
pub(crate) fn parse_formula_extra_end(data: &[u8], tokens: &[u8], offset: usize) -> Result<usize> {
    let extras = scan_formula_extra_kinds(tokens, FORMULA_RECORD_TYPE)?;
    let mut offset = offset;
    for extra in extras {
        match extra {
            FormulaExtraKind::Memory => {
                let count = usize::from(read_u16(data, offset, "PtgExtraMem count")?);
                let bytes = count
                    .checked_mul(8)
                    .ok_or_else(|| invalid("PtgExtraMem length overflows"))?;
                let end = offset
                    .checked_add(2)
                    .and_then(|value| value.checked_add(bytes))
                    .ok_or_else(|| invalid("PtgExtraMem length overflows"))?;
                let ranges_start = offset
                    .checked_add(2)
                    .ok_or_else(|| invalid("PtgExtraMem length overflows"))?;
                let ranges = data
                    .get(ranges_start..end)
                    .ok_or_else(|| invalid("truncated PtgExtraMem"))?;
                for range in ranges.as_chunks::<8>().0.iter() {
                    validate_ref8u(range)?;
                }
                offset = end;
            },
            FormulaExtraKind::Array => {
                let dimensions_end = offset
                    .checked_add(3)
                    .ok_or_else(|| invalid("PtgExtraArray length overflows"))?;
                let dimensions = data
                    .get(offset..dimensions_end)
                    .ok_or_else(|| invalid("truncated PtgExtraArray dimensions"))?;
                let count = (usize::from(dimensions[0]) + 1)
                    .checked_mul(
                        usize::from(u16::from_le_bytes([dimensions[1], dimensions[2]])) + 1,
                    )
                    .ok_or_else(|| invalid("PtgExtraArray dimensions overflow"))?;
                offset = offset
                    .checked_add(3)
                    .ok_or_else(|| invalid("PtgExtraArray length overflows"))?;
                for _ in 0..count {
                    offset = validate_ser_ar(data, offset).map_err(|error| match error {
                        Error::InvalidRecord { message, .. } => {
                            invalid_for(FORMULA_RECORD_TYPE, message)
                        },
                        other => other,
                    })?;
                }
            },
        }
    }
    Ok(offset)
}

/// Validate a nonempty Formula `RgbExtra` suffix and require exact
/// consumption. Empty suffixes are deliberately not scanned so legacy
/// opaque/no-extra Formula token acceptance remains unchanged.
pub(crate) fn validate_formula_extra(tokens: &[u8], extra: &[u8]) -> Result<()> {
    if extra.is_empty() {
        return Ok(());
    }
    let end = parse_formula_extra_end(extra, tokens, 0)?;
    if end != extra.len() {
        return Err(invalid("Formula RgbExtra has trailing or unowned bytes"));
    }
    Ok(())
}

/// Return the Formula token and suffix slices after checking only BIFF record
/// framing. Callers that need to accept a nonempty suffix must additionally
/// call [`validate_formula_extra`].
pub(crate) fn formula_payload_parts(data: &[u8]) -> Result<(&[u8], &[u8])> {
    if data.len() < FORMULA_FIXED_SIZE {
        return Err(Error::InvalidLength {
            expected: FORMULA_FIXED_SIZE,
            found: data.len(),
        });
    }
    if data.len() > MAX_FORMULA_PAYLOAD {
        return Err(invalid(format!(
            "Formula payload exceeds the BIFF8 limit of {MAX_FORMULA_PAYLOAD} bytes"
        )));
    }
    let token_len = usize::from(u16::from_le_bytes([data[20], data[21]]));
    let token_end = FORMULA_FIXED_SIZE
        .checked_add(token_len)
        .ok_or_else(|| invalid("Formula token length overflows"))?;
    if token_end > data.len() {
        return Err(Error::InvalidLength {
            expected: token_end,
            found: data.len(),
        });
    }
    Ok((&data[FORMULA_FIXED_SIZE..token_end], &data[token_end..]))
}

/// Retain a validated suffix with the exact source cell and token stream that
/// make it safe to re-emit later.
pub(crate) fn retain_formula_extra(
    row: u16,
    col: u16,
    tokens: &[u8],
    extra: &[u8],
) -> Result<Option<Ancillary>> {
    if extra.is_empty() {
        return Ok(None);
    }
    validate_formula_extra(tokens, extra)?;

    Ancillary::new((row, col), tokens, extra).map(Some)
}

fn read_u16(data: &[u8], offset: usize, field: &'static str) -> Result<u16> {
    let end = offset
        .checked_add(2)
        .ok_or_else(|| invalid(format!("{field} offset overflows")))?;
    let bytes = data
        .get(offset..end)
        .ok_or_else(|| invalid(format!("truncated {field}")))?;
    Ok(u16::from_le_bytes([bytes[0], bytes[1]]))
}

fn validate_ref8u(range: &[u8; 8]) -> Result<()> {
    let first_row = u16::from_le_bytes([range[0], range[1]]);
    let last_row = u16::from_le_bytes([range[2], range[3]]);
    let first_col = u16::from_le_bytes([range[4], range[5]]);
    let last_col = u16::from_le_bytes([range[6], range[7]]);
    if first_row > last_row {
        return Err(invalid("PtgExtraMem Ref8U rows are reversed"));
    }
    if first_col > last_col || last_col > 0x00FF {
        return Err(invalid("PtgExtraMem Ref8U columns are invalid"));
    }
    Ok(())
}

fn invalid_for(record_type: u16, message: impl Into<String>) -> Error {
    Error::InvalidRecord {
        record_type,
        message: message.into(),
    }
}
