//! Row and column shifts of the records whose formulas are evaluated at
//! other cells: conditional formats (`CondFmt`/`CF`, `CondFmt12`/`CF12`,
//! `CFEx`), data validations (`DV`), array formulas (`Array`) and data
//! tables (`Table`). Record 0316.
//!
//! A conditional-format or validation formula is evaluated at every cell of
//! its range. Its `PtgRefN`/`PtgAreaN` relative components are offsets from
//! that cell (`[MS-XLS]` 2.5.198.104 `RgceLocRel`), so they keep naming the
//! same cells when every cell of the range and the cell it names lie on the
//! same side of the shift point, the rule shared formulas follow (record
//! 0308). Absolute components are coordinates and shift like a cell
//! formula's. A `PtgRef`/`PtgArea` with relative flags is ambiguous here:
//! `RgceLoc` stores coordinates, while LibreOffice reads the relative parts
//! of a conditional-format `PtgRef` as offsets; such a reference is admitted
//! only when no reading of it can reach the shift point. 3-D references
//! are closed with the workbook's other 3-D references
//! (`patch_rule_3d_references`).

use super::{
    AxisShift, CF, CONDFMT, CONTINUE, DV, Error, ExternSheets, RawRecord, Result, Target3d, binary,
    patch_condfmt, patch_ref8, patch_reference_span, patch_sqref, shift_axis_point, target_3d,
};
use crate::formula::FormulaReferenceSpan;
use litchi_core::edit::RefusalKind;

pub(super) const CONDFMT12: u16 = 0x0879;
pub(super) const CF12: u16 = 0x087a;
pub(super) const CFEX: u16 = 0x087b;

/// The record kinds this module shifts inside the shifted worksheet.
pub(super) const RULE_RECORDS: &[u16] = &[CONDFMT, CF, CONDFMT12, CF12, CFEX, DV];

/// A rectangle of cells, `(first_row, last_row, first_column, last_column)`.
#[derive(Clone, Copy)]
pub(super) struct Bound {
    rows: (u16, u16),
    columns: (u16, u16),
}

impl Bound {
    /// Reads a `Ref8U`.
    fn ref8(range: &[u8]) -> Result<Self> {
        Ok(Self {
            rows: (
                binary::read_u16_le_at(range, 0)?,
                binary::read_u16_le_at(range, 2)?,
            ),
            columns: (
                binary::read_u16_le_at(range, 4)?,
                binary::read_u16_le_at(range, 6)?,
            ),
        })
    }

    /// The smallest rectangle holding every range of a `SqRefU`.
    fn sqref(payload: &[u8]) -> Result<Self> {
        let count = usize::from(binary::read_u16_le_at(payload, 0)?);
        if count == 0 || payload.len() != 2 + count * 8 {
            return Err(Error::refused(
                RefusalKind::InvalidValue,
                "edit cells",
                "a range list does not fill its record and is not shifted",
            ));
        }
        let mut bound = Self::ref8(&payload[2..10])?;
        for range in payload[10..].as_chunks::<8>().0 {
            let next = Self::ref8(range)?;
            bound.rows = (bound.rows.0.min(next.rows.0), bound.rows.1.max(next.rows.1));
            bound.columns = (
                bound.columns.0.min(next.columns.0),
                bound.columns.1.max(next.columns.1),
            );
        }
        Ok(bound)
    }
}

/// The shifted axis of `shift`: `(start, count, insert, grid limit)`.
fn axis(shift: AxisShift) -> (u16, u16, bool, u32) {
    match shift {
        AxisShift::Rows {
            start,
            count,
            insert,
        } => (start, count, insert, 65_536),
        AxisShift::Columns {
            start,
            count,
            insert,
        } => (u16::from(start), u16::from(count), insert, 256),
    }
}

/// Whether a relative offset keeps naming the same cell from every cell of
/// `low..=high` on the shifted axis: each cell and the cell it names lie on
/// the same side of the shift point, so both move by the same amount or
/// neither moves (record 0308's shared-formula rule).
pub(super) fn offset_kept(low: u16, high: u16, offset: i32, shift: AxisShift) -> bool {
    let (start, count, insert, _) = axis(shift);
    let (start, end) = (i32::from(start), i32::from(start) + i32::from(count));
    let (low, high) = (i32::from(low), i32::from(high));
    let before = high < start && high + offset < start;
    let after = if insert {
        low >= start && low + offset >= start
    } else {
        low >= end && low + offset >= end
    };
    before || after
}

/// Shifts one relative-reference component (`PtgRefN`/`PtgAreaN` or a
/// shared formula's): a relative row or column must keep its offset
/// (`offset_kept` over `low..=high`), an absolute one shifts as a point.
pub(super) fn patch_relative_component(
    tokens: &mut [u8],
    row_offset: usize,
    flags_offset: usize,
    shift: AxisShift,
    low: u16,
    high: u16,
) -> Result<()> {
    let flags = binary::read_u16_le_at(tokens, flags_offset)?;
    let (start, count, insert, limit) = axis(shift);
    let crossing = || {
        Err(Error::refused(
            RefusalKind::UnsupportedOperation,
            "edit cells",
            "a relative reference would cross the shift point",
        ))
    };
    if matches!(shift, AxisShift::Rows { .. }) {
        if flags & 0x8000 != 0 {
            let offset = i32::from(binary::read_u16_le_at(tokens, row_offset)? as i16);
            return if offset_kept(low, high, offset, shift) {
                Ok(())
            } else {
                crossing()
            };
        }
        let row = binary::read_u16_le_at(tokens, row_offset)?;
        let row = shift_axis_point(row, start, count, insert, limit)?;
        write_u16(tokens, row_offset, row)
    } else {
        if flags & 0x4000 != 0 {
            let offset = i32::from((flags & 0x00ff) as u8 as i8);
            return if offset_kept(low, high, offset, shift) {
                Ok(())
            } else {
                crossing()
            };
        }
        let column = shift_axis_point(flags & 0x00ff, start, count, insert, limit)?;
        write_u16(tokens, flags_offset, (flags & 0xff00) | column)
    }
}

fn write_u16(bytes: &mut [u8], offset: usize, value: u16) -> Result<()> {
    bytes
        .get_mut(offset..offset + 2)
        .ok_or_else(|| Error::InvalidData("formula reference field is truncated".into()))?
        .copy_from_slice(&value.to_le_bytes());
    Ok(())
}

/// The column fields' relative flags of one reference, OR-ed.
fn relative_flags(tokens: &[u8], span: FormulaReferenceSpan) -> Result<u16> {
    Ok(match span {
        FormulaReferenceSpan::Cell { column, .. } => binary::read_u16_le_at(tokens, column)?,
        FormulaReferenceSpan::Area {
            first_column,
            last_column,
            ..
        } => {
            binary::read_u16_le_at(tokens, first_column)?
                | binary::read_u16_le_at(tokens, last_column)?
        },
    } & 0xc000)
}

/// The `(row, column)` field offsets of each endpoint of `span`: one for a
/// cell, two for an area.
fn endpoints(span: FormulaReferenceSpan) -> impl Iterator<Item = (usize, usize)> {
    let (first, second) = match span {
        FormulaReferenceSpan::Cell { row, column } => ((row, column), None),
        FormulaReferenceSpan::Area {
            first_row,
            last_row,
            first_column,
            last_column,
        } => ((first_row, first_column), Some((last_row, last_column))),
    };
    std::iter::once(first).chain(second)
}

/// Shifts one conditional-format or data-validation formula whose cells
/// are `bound` (module documentation).
pub(super) fn patch_rule_formula(
    tokens: &mut [u8],
    shift: AxisShift,
    bound: Bound,
    _what: &str,
) -> Result<()> {
    if tokens.is_empty() {
        return Ok(());
    }
    let (same_sheet, relative, _three_d) = crate::formula::shared_formula_reference_spans(tokens)
        .ok_or_else(|| {
        Error::refused(
            RefusalKind::UnsupportedOperation,
            "edit cells",
            "row/column movement cannot read this formula's references",
        )
    })?;
    let (low, high) = match shift {
        AxisShift::Rows { .. } => bound.rows,
        AxisShift::Columns { .. } => bound.columns,
    };
    let (start, ..) = axis(shift);
    for span in same_sheet {
        if relative_flags(tokens, span)? == 0 {
            patch_reference_span(tokens, span, shift)?;
            continue;
        }
        // Read as a coordinate the field must not reach the shift point;
        // read as an offset from every cell, neither may its target.
        for (row, column) in endpoints(span) {
            let value = match shift {
                AxisShift::Rows { .. } => binary::read_u16_le_at(tokens, row)?,
                AxisShift::Columns { .. } => binary::read_u16_le_at(tokens, column)? & 0x00ff,
            };
            let as_offset = match shift {
                AxisShift::Rows { .. } => i32::from(value as i16),
                AxisShift::Columns { .. } => i32::from(value as u8 as i8),
            };
            if value >= start || !(high < start && i32::from(high) + as_offset < i32::from(start)) {
                return Err(Error::refused(
                    RefusalKind::UnsupportedOperation,
                    "edit cells",
                    "row/column movement cannot close a formula's relative PtgRef or PtgArea",
                ));
            }
        }
    }
    for span in relative {
        for (row, column) in endpoints(span) {
            patch_relative_component(tokens, row, column, shift, low, high)?;
        }
    }
    Ok(())
}

/// The token ranges of a `CF` record's two formulas (`[MS-XLS]` 2.4.42):
/// `cce1` and `cce2` bytes at the end of the record.
pub(super) fn cf_formula_ranges(payload: &[u8]) -> Result<[std::ops::Range<usize>; 2]> {
    if payload.len() < 6 {
        return Err(Error::InvalidData("CF record is truncated".into()));
    }
    let first = usize::from(binary::read_u16_le_at(payload, 2)?);
    let second = usize::from(binary::read_u16_le_at(payload, 4)?);
    let start = payload
        .len()
        .checked_sub(first + second)
        .filter(|start| *start >= 6)
        .ok_or_else(|| Error::InvalidData("CF formulas exceed the record".into()))?;
    Ok([start..start + first, start + first..payload.len()])
}

/// The token ranges of a `DV` record's two formulas and the offset of its
/// `sqref` (`[MS-XLS]` 2.4.112).
fn dv_layout(payload: &[u8]) -> Result<([std::ops::Range<usize>; 2], usize)> {
    let mut offset = 4_usize;
    for _ in 0..4 {
        let characters = usize::from(binary::read_u16_le_at(payload, offset)?);
        let flags = *payload
            .get(offset + 2)
            .ok_or_else(|| Error::InvalidData("DV string is truncated".into()))?;
        offset += 3 + characters * if flags & 1 != 0 { 2 } else { 1 };
    }
    let mut formulas: [std::ops::Range<usize>; 2] = [0..0, 0..0];
    for formula in &mut formulas {
        let length = usize::from(binary::read_u16_le_at(payload, offset)?);
        let start = offset + 4;
        *formula = start..start + length;
        offset = start + length;
    }
    if formulas[1].end > payload.len() {
        return Err(Error::InvalidData("DV formula is truncated".into()));
    }
    Ok((formulas, offset))
}

/// One `CFVO` (`[MS-XLS]` 2.5.39) at `offset`: its formula's token range and
/// the offset after it.
fn cfvo(payload: &[u8], offset: usize) -> Result<(std::ops::Range<usize>, usize)> {
    let kind = *payload
        .get(offset)
        .ok_or_else(|| Error::InvalidData("CFVO is truncated".into()))?;
    let length = usize::from(binary::read_u16_le_at(payload, offset + 1)?);
    let start = offset + 3;
    let mut end = start + length;
    if length == 0 && !matches!(kind, 2 | 3) {
        end += 8;
    }
    if end > payload.len() {
        return Err(Error::InvalidData("CFVO is truncated".into()));
    }
    Ok((start..start + length, end))
}

/// The token ranges of every formula of a `CF12` record (`[MS-XLS]`
/// 2.4.43): `rgce1`, `rgce2`, `fmlaActive` and the `CFVO` formulas of its
/// color scale, data bar or icon set. A layout this walk does not account
/// for byte for byte is refused.
pub(super) fn cf12_formula_ranges(payload: &[u8]) -> Result<Vec<std::ops::Range<usize>>> {
    let truncated = || {
        Error::refused(
            RefusalKind::UnsupportedTarget,
            "edit cells",
            "CF12 record layout is not shiftable",
        )
    };
    let byte = |offset: usize| payload.get(offset).copied().ok_or_else(truncated);
    let word =
        |offset: usize| binary::read_u16_le_at(payload, offset).map_err(|_error| truncated());
    let condition = byte(12)?;
    let first = usize::from(word(14)?);
    let second = usize::from(word(16)?);
    let dxf = payload
        .get(18..22)
        .map(|bytes| u32::from_le_bytes([bytes[0], bytes[1], bytes[2], bytes[3]]))
        .ok_or_else(truncated)?;
    let mut offset = 18
        + if dxf == 0 {
            6
        } else {
            4 + usize::try_from(dxf).map_err(|_error| truncated())?
        };
    let mut ranges = vec![
        offset..offset + first,
        offset + first..offset + first + second,
    ];
    offset += first + second;
    let active = usize::from(word(offset)?);
    ranges.push(offset + 2..offset + 2 + active);
    // fmlaActive, flags, ipriority, icfTemplate, cbTemplateParm and the
    // 16-byte rgbTemplateParms.
    offset += 2 + active + 1 + 2 + 2 + 1 + 16;
    match condition {
        1 | 2 => {},
        3 => {
            let interpolation = byte(offset + 3)?;
            let gradient = byte(offset + 4)?;
            offset += 6;
            for _ in 0..interpolation {
                let (formula, end) = cfvo(payload, offset)?;
                ranges.push(formula);
                offset = end + 8;
            }
            offset += 24 * usize::from(gradient);
        },
        4 => {
            offset += 6 + 16;
            for _ in 0..2 {
                let (formula, end) = cfvo(payload, offset)?;
                ranges.push(formula);
                offset = end;
            }
        },
        5 => offset += 2 + usize::from(word(offset)?),
        6 => {
            let states = byte(offset + 3)?;
            offset += 6;
            for _ in 0..states {
                let (formula, end) = cfvo(payload, offset)?;
                ranges.push(formula);
                offset = end + 5;
            }
        },
        _ => return Err(truncated()),
    }
    if offset != payload.len() || ranges.iter().any(|range| range.end > payload.len()) {
        return Err(truncated());
    }
    Ok(ranges)
}

/// The formula token ranges of a rule record of any worksheet, for the
/// workbook-wide 3-D pass. `CondFmt`, `CondFmt12` and `CFEx` hold none.
fn rule_formula_ranges(kind: u16, payload: &[u8]) -> Result<Vec<std::ops::Range<usize>>> {
    match kind {
        CF => Ok(cf_formula_ranges(payload)?.to_vec()),
        CF12 => cf12_formula_ranges(payload),
        DV => Ok(dv_layout(payload)?.0.to_vec()),
        _ => Ok(Vec::new()),
    }
}

fn payload_mut<'a>(worksheet: &'a mut [u8], record: &RawRecord) -> Result<&'a mut [u8]> {
    worksheet
        .get_mut(record.start + 4..record.end)
        .ok_or_else(|| Error::InvalidData("rule record payload is truncated".into()))
}

/// Shifts every conditional-format and data-validation record of the
/// shifted worksheet in place. `records` are the worksheet's records; each
/// rule is checked against the ranges its formulas are evaluated at before
/// those ranges shift.
pub(super) fn patch_rule_records(
    worksheet: &mut [u8],
    records: &[RawRecord],
    shift: AxisShift,
) -> Result<()> {
    let mut bound: Option<Bound> = None;
    for (position, record) in records.iter().enumerate() {
        if !RULE_RECORDS.contains(&record.kind) {
            continue;
        }
        if records
            .get(position + 1)
            .is_some_and(|next| next.kind == CONTINUE)
        {
            return Err(Error::refused(
                RefusalKind::UnsupportedOperation,
                "edit cells",
                "row/column movement cannot close a continued record",
            ));
        }
        let payload = payload_mut(worksheet, record)?;
        match record.kind {
            CONDFMT => {
                bound = Some(Bound::ref8(payload.get(4..12).ok_or_else(|| {
                    Error::InvalidData("CondFmt record is truncated".into())
                })?)?);
                patch_condfmt(payload, shift)?;
            },
            CONDFMT12 => {
                if payload.len() < 26 {
                    return Err(Error::InvalidData("CondFmt12 record is truncated".into()));
                }
                bound = Some(Bound::ref8(&payload[16..24])?);
                patch_ref8(&mut payload[4..12], shift)?;
                patch_condfmt(&mut payload[12..], shift)?;
            },
            CFEX => {
                if payload.len() < 18 {
                    return Err(Error::InvalidData("CFEx record is truncated".into()));
                }
                // A CFEx that announces a CF12 rule evaluates it over its
                // own range; one that extends a legacy rule carries no
                // formula.
                bound = Some(Bound::ref8(&payload[4..12])?);
                patch_ref8(&mut payload[4..12], shift)?;
            },
            CF | CF12 => {
                let bound = bound.ok_or_else(|| {
                    Error::refused(
                        RefusalKind::UnsupportedTarget,
                        "edit cells",
                        "a conditional-format rule has no range",
                    )
                })?;
                let ranges = if record.kind == CF {
                    cf_formula_ranges(payload)?.to_vec()
                } else {
                    cf12_formula_ranges(payload)?
                };
                for range in ranges {
                    let tokens = payload
                        .get_mut(range)
                        .ok_or_else(|| Error::InvalidData("CF formula is truncated".into()))?;
                    patch_rule_formula(tokens, shift, bound, "conditional format")?;
                }
            },
            DV => {
                let (formulas, sqref) = dv_layout(payload)?;
                let cells = Bound::sqref(
                    payload
                        .get(sqref..)
                        .ok_or_else(|| Error::InvalidData("DV range list is truncated".into()))?,
                )?;
                for range in formulas {
                    let tokens = payload
                        .get_mut(range)
                        .ok_or_else(|| Error::InvalidData("DV formula is truncated".into()))?;
                    patch_rule_formula(tokens, shift, cells, "data validation")?;
                }
                patch_sqref(&mut payload[sqref..], shift)?;
            },
            _ => {},
        }
    }
    Ok(())
}

/// Shifts the 3-D references of a rule, shared-formula or array formula
/// that name the shifted worksheet, in any worksheet. A relative component
/// of such a reference is refused.
pub(super) fn patch_rule_3d_references(
    tokens: &mut [u8],
    sheets: &ExternSheets,
    tab: usize,
    shift: AxisShift,
) -> Result<()> {
    if tokens.is_empty() {
        return Ok(());
    }
    let (_, _, three_d) =
        crate::formula::shared_formula_reference_spans(tokens).ok_or_else(|| {
            Error::refused(RefusalKind::UnsupportedOperation, "edit cells", "row/column movement cannot read the references of a shared, array, conditional-format or validation formula")
        })?;
    for reference in three_d {
        if matches!(
            target_3d(sheets, reference.extern_sheet, tab)?,
            Target3d::Elsewhere
        ) {
            continue;
        }
        let Some(span) = reference.reference else {
            continue;
        };
        if relative_flags(tokens, span)? != 0 {
            return Err(Error::refused(
                RefusalKind::UnsupportedTarget,
                "edit cells",
                "row/column movement cannot close a relative 3-D reference in a rule or shared formula",
            ));
        }
        patch_reference_span(tokens, span, shift)?;
    }
    Ok(())
}

/// The 3-D pass over one rule record of any worksheet.
pub(super) fn patch_rule_record_3d(
    kind: u16,
    payload: &mut [u8],
    sheets: &ExternSheets,
    tab: usize,
    shift: AxisShift,
) -> Result<()> {
    for range in rule_formula_ranges(kind, payload)? {
        let tokens = payload
            .get_mut(range)
            .ok_or_else(|| Error::InvalidData("rule formula is truncated".into()))?;
        patch_rule_3d_references(tokens, sheets, tab, shift)?;
    }
    Ok(())
}

/// The token range of an `Array` record's formula (`[MS-XLS]` 2.4.4):
/// `ref` (6 bytes), flags (2), unused (4), then an `ArrayParsedFormula`
/// whose `rgcb` must be empty here (its `PtgMemArea` and `PtgArray` data are
/// not shifted).
pub(super) fn array_tokens(payload: &[u8]) -> Result<std::ops::Range<usize>> {
    let length = usize::from(binary::read_u16_le_at(payload, 12)?);
    let end = 14 + length;
    if end > payload.len() {
        return Err(Error::InvalidData("Array formula is truncated".into()));
    }
    if end != payload.len() {
        return Err(Error::refused(
            RefusalKind::UnsupportedTarget,
            "edit cells",
            "row/column movement refuses an array formula with ancillary data",
        ));
    }
    Ok(14..end)
}

/// Whether `first..=last` moves whole under `shift` (`Some(true)`), stays
/// (`Some(false)`), or is split by the shift point (`None`).
fn range_moves(first: u16, last: u16, shift: AxisShift) -> Option<bool> {
    let (start, count, insert, _) = axis(shift);
    if last < start {
        Some(false)
    } else if (insert && first >= start) || (!insert && first >= start.saturating_add(count)) {
        Some(true)
    } else {
        None
    }
}

/// Moves a `Ref` (`rwFirst`, `rwLast`, `colFirst`, `colLast` as bytes) whole,
/// with `margin` cells before it that must move with it (a data table's
/// formula row and column). A shift point inside is refused.
fn patch_ref_whole(range: &mut [u8], shift: AxisShift, margin: u16, _what: &str) -> Result<()> {
    let (start, count, insert, limit) = axis(shift);
    let (first, last) = match shift {
        AxisShift::Rows { .. } => (
            binary::read_u16_le_at(range, 0)?,
            binary::read_u16_le_at(range, 2)?,
        ),
        AxisShift::Columns { .. } => (u16::from(range[4]), u16::from(range[5])),
    };
    match range_moves(first.saturating_sub(margin), last, shift) {
        Some(false) => Ok(()),
        Some(true) => {
            let first = shift_axis_point(first, start, count, insert, limit)?;
            let last = shift_axis_point(last, start, count, insert, limit)?;
            match shift {
                AxisShift::Rows { .. } => {
                    range[0..2].copy_from_slice(&first.to_le_bytes());
                    range[2..4].copy_from_slice(&last.to_le_bytes());
                },
                AxisShift::Columns { .. } => {
                    range[4] = u8::try_from(first).map_err(|_error| {
                        Error::refused(
                            RefusalKind::InvalidValue,
                            "edit cells",
                            "column outside BIFF8",
                        )
                    })?;
                    range[5] = u8::try_from(last).map_err(|_error| {
                        Error::refused(
                            RefusalKind::InvalidValue,
                            "edit cells",
                            "column outside BIFF8",
                        )
                    })?;
                },
            }
            Ok(())
        },
        None => Err(Error::refused(
            RefusalKind::UnsupportedOperation,
            "edit cells",
            "a row/column shift through this range would split it",
        )),
    }
}

/// Shifts an `Array` record of the shifted worksheet: its range moves whole
/// and its formula's same-sheet references shift like a cell formula's
/// (array formulas store coordinates); 3-D references are closed by the
/// workbook pass.
pub(super) fn patch_array(payload: &mut [u8], shift: AxisShift) -> Result<()> {
    if payload.len() < 14 {
        return Err(Error::InvalidData("Array record is truncated".into()));
    }
    patch_ref_whole(&mut payload[0..6], shift, 0, "an array formula's range")?;
    patch_array_formula(payload, shift)
}

/// The formula part of [`patch_array`], for an `Array` record whose range
/// moved with its row (`move_owner_range`).
pub(super) fn patch_array_formula(payload: &mut [u8], shift: AxisShift) -> Result<()> {
    let tokens = array_tokens(payload)?;
    super::patch_formula_tokens(&mut payload[tokens], shift)
}

/// Moves the range of an `Array` or `Table` record (`record`, header
/// included) that lives in a row the shift moves from `from` to `to` rows,
/// or shifts its columns: certification has already proved the range moves
/// whole.
pub(super) fn move_owner_range(
    record: &mut [u8],
    shift: AxisShift,
    from: u16,
    to: u16,
) -> Result<()> {
    let range = record
        .get_mut(4..10)
        .ok_or_else(|| Error::InvalidData("formula owner record is truncated".into()))?;
    match shift {
        AxisShift::Rows { .. } => {
            // The record's row and its range move by the same amount.
            let delta = i32::from(to) - i32::from(from);
            for at in [0, 2] {
                let row = i32::from(binary::read_u16_le_at(range, at)?) + delta;
                let row = u16::try_from(row).map_err(|_error| {
                    Error::refused(
                        RefusalKind::InvalidValue,
                        "edit cells",
                        "row movement moves a formula owner outside BIFF8",
                    )
                })?;
                range[at..at + 2].copy_from_slice(&row.to_le_bytes());
            }
        },
        AxisShift::Columns { .. } => {
            let (start, count, insert, limit) = axis(shift);
            for at in [4, 5] {
                let column = shift_axis_point(u16::from(range[at]), start, count, insert, limit)?;
                range[at] = u8::try_from(column).map_err(|_error| {
                    Error::refused(
                        RefusalKind::InvalidValue,
                        "edit cells",
                        "column outside BIFF8",
                    )
                })?;
            }
        },
    }
    Ok(())
}

/// Shifts a `Table` record (`[MS-XLS]` 2.4.319): its range moves whole with
/// the formula row and column before it, and its input cells shift as
/// points unless marked deleted.
pub(super) fn patch_table(payload: &mut [u8], shift: AxisShift) -> Result<()> {
    if payload.len() != 16 {
        return Err(Error::InvalidData("Table record is not 16 bytes".into()));
    }
    patch_ref_whole(&mut payload[0..6], shift, 1, "a data table")?;
    patch_table_inputs(payload, shift)
}

/// The input-cell part of [`patch_table`], for a `Table` record whose range
/// moved with its row (`move_owner_range`).
pub(super) fn patch_table_inputs(payload: &mut [u8], shift: AxisShift) -> Result<()> {
    if payload.len() != 16 {
        return Err(Error::InvalidData("Table record is not 16 bytes".into()));
    }
    let flags = binary::read_u16_le_at(payload, 6)?;
    let (start, count, insert, limit) = axis(shift);
    // The second input cell exists only in a two-variable table (fTbl2);
    // a deleted input cell keeps its sentinel.
    let two_variable = flags & 0x0008 != 0;
    for (deleted, row_offset, column_offset, used) in
        [(0x0010, 8, 10, true), (0x0020, 12, 14, two_variable)]
    {
        let row = binary::read_u16_le_at(payload, row_offset)?;
        let column = binary::read_u16_le_at(payload, column_offset)?;
        if !used || flags & deleted != 0 || row == 0xffff || column == 0xffff {
            continue;
        }
        let offset = match shift {
            AxisShift::Rows { .. } => row_offset,
            AxisShift::Columns { .. } => column_offset,
        };
        let value = binary::read_u16_le_at(payload, offset)?;
        let value = shift_axis_point(value, start, count, insert, limit)?;
        payload[offset..offset + 2].copy_from_slice(&value.to_le_bytes());
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::super::raw_records;
    use super::*;

    const INSERT_ROWS: AxisShift = AxisShift::Rows {
        start: 2,
        count: 3,
        insert: true,
    };
    const DELETE_ROWS: AxisShift = AxisShift::Rows {
        start: 2,
        count: 3,
        insert: false,
    };

    fn record(kind: u16, payload: &[u8]) -> Vec<u8> {
        let mut bytes = kind.to_le_bytes().to_vec();
        bytes.extend_from_slice(&u16::try_from(payload.len()).unwrap().to_le_bytes());
        bytes.extend_from_slice(payload);
        bytes
    }

    fn ref8(first_row: u16, last_row: u16, first_column: u16, last_column: u16) -> Vec<u8> {
        [first_row, last_row, first_column, last_column]
            .iter()
            .flat_map(|value| value.to_le_bytes())
            .collect()
    }

    /// `PtgRefN` (value class) with a row offset and an absolute column.
    fn ref_n(row_offset: i16, column: u8) -> Vec<u8> {
        let mut tokens = vec![0x4c];
        tokens.extend_from_slice(&row_offset.to_le_bytes());
        tokens.extend_from_slice(&(0x8000 | u16::from(column)).to_le_bytes());
        tokens
    }

    fn bound(first_row: u16, last_row: u16) -> Bound {
        Bound {
            rows: (first_row, last_row),
            columns: (0, 0),
        }
    }

    #[test]
    fn relative_offsets_are_kept_only_on_one_side_of_the_shift_point() {
        // Every cell after the point, and so is every cell it names.
        assert!(offset_kept(4, 6, -1, INSERT_ROWS));
        // Every cell before the point, and so is every cell it names.
        assert!(offset_kept(0, 1, 0, INSERT_ROWS));
        // The cells are before the point; the cells they name are not.
        assert!(!offset_kept(0, 1, 1, INSERT_ROWS));
        // The cells are after the point; the cells they name are not.
        assert!(!offset_kept(3, 6, -2, INSERT_ROWS));
        // A deletion keeps only cells and targets after the deleted rows.
        assert!(offset_kept(5, 6, 0, DELETE_ROWS));
        assert!(!offset_kept(4, 6, 0, DELETE_ROWS));
    }

    #[test]
    fn rule_formulas_keep_offsets_and_shift_coordinates() {
        // `PtgRefN` one row up from cells 5..=7: unchanged by the insertion.
        let mut tokens = ref_n(-1, 0);
        patch_rule_formula(&mut tokens, INSERT_ROWS, bound(5, 7), "test").unwrap();
        assert_eq!(tokens, ref_n(-1, 0));
        // From cells 2..=3 the row above crosses the point: refused.
        let mut crossing = ref_n(-1, 0);
        assert!(patch_rule_formula(&mut crossing, INSERT_ROWS, bound(2, 3), "test").is_err());
        // An absolute `PtgRef` ($A$5) shifts like a cell formula's.
        let mut absolute = vec![0x44, 4, 0, 0, 0];
        patch_rule_formula(&mut absolute, INSERT_ROWS, bound(0, 0), "test").unwrap();
        assert_eq!(absolute, vec![0x44, 7, 0, 0, 0]);
        patch_rule_formula(&mut absolute, DELETE_ROWS, bound(0, 0), "test").unwrap();
        assert_eq!(absolute, vec![0x44, 4, 0, 0, 0]);
        // A relative `PtgRef` is admitted only out of the shift's reach.
        let mut relative = vec![0x44, 0, 0, 0, 0xc0];
        patch_rule_formula(&mut relative, INSERT_ROWS, bound(0, 0), "test").unwrap();
        assert_eq!(relative, vec![0x44, 0, 0, 0, 0xc0]);
        let mut reaching = vec![0x44, 4, 0, 0, 0xc0];
        assert!(patch_rule_formula(&mut reaching, INSERT_ROWS, bound(0, 0), "test").is_err());
    }

    fn cf(formula: &[u8]) -> Vec<u8> {
        // ct = formula, cp = 0, cce1, cce2 = 0, rgbdxf flags with no
        // blocks, then the formula.
        let mut payload = vec![2, 0];
        payload.extend_from_slice(&u16::try_from(formula.len()).unwrap().to_le_bytes());
        payload.extend_from_slice(&0u16.to_le_bytes());
        payload.extend_from_slice(&[0xff, 0xff, 0x3f, 0x00, 0x00, 0x00]);
        payload.extend_from_slice(formula);
        record(CF, &payload)
    }

    fn condfmt(first_row: u16, last_row: u16) -> Vec<u8> {
        let mut payload = vec![1, 0, 1, 0];
        payload.extend_from_slice(&ref8(first_row, last_row, 0, 0));
        payload.extend_from_slice(&1u16.to_le_bytes());
        payload.extend_from_slice(&ref8(first_row, last_row, 0, 0));
        record(CONDFMT, &payload)
    }

    fn cfex(first_row: u16, last_row: u16) -> Vec<u8> {
        let mut payload = CFEX.to_le_bytes().to_vec();
        payload.extend_from_slice(&1u16.to_le_bytes());
        payload.extend_from_slice(&ref8(first_row, last_row, 0, 0));
        payload.extend_from_slice(&1u32.to_le_bytes());
        payload.extend_from_slice(&1u16.to_le_bytes());
        record(CFEX, &payload)
    }

    fn patched(bytes: &[u8], shift: AxisShift) -> Result<Vec<u8>> {
        let mut bytes = bytes.to_vec();
        let records = raw_records(&bytes)?;
        patch_rule_records(&mut bytes, &records, shift)?;
        Ok(bytes)
    }

    #[test]
    fn conditional_formats_and_extensions_shift_and_invert_exactly() {
        let mut source = condfmt(5, 9);
        source.extend(cf(&ref_n(-1, 0)));
        source.extend(cfex(5, 9));
        let inserted = patched(&source, INSERT_ROWS).unwrap();
        let mut expected = condfmt(8, 12);
        expected.extend(cf(&ref_n(-1, 0)));
        expected.extend(cfex(8, 12));
        assert_eq!(inserted, expected);
        assert_eq!(patched(&inserted, DELETE_ROWS).unwrap(), source);
        // The same rule over rows 2..=4 names a row across the point.
        let mut crossing = condfmt(2, 4);
        crossing.extend(cf(&ref_n(-1, 0)));
        assert!(patched(&crossing, INSERT_ROWS).is_err());
    }

    fn cfvo(kind: u8, formula: &[u8]) -> Vec<u8> {
        let mut bytes = vec![kind];
        bytes.extend_from_slice(&u16::try_from(formula.len()).unwrap().to_le_bytes());
        bytes.extend_from_slice(formula);
        if formula.is_empty() && !matches!(kind, 2 | 3) {
            bytes.extend_from_slice(&0f64.to_le_bytes());
        }
        bytes
    }

    /// A data-bar `CF12` whose maximum is the formula `formula`.
    fn data_bar(formula: &[u8]) -> Vec<u8> {
        let mut payload = CF12.to_le_bytes().to_vec();
        payload.extend_from_slice(&[0; 10]);
        payload.extend_from_slice(&[4, 0, 0, 0, 0, 0]);
        payload.extend_from_slice(&[0, 0, 0, 0, 0, 0]);
        payload.extend_from_slice(&0u16.to_le_bytes());
        payload.extend_from_slice(&[0, 1, 0, 0, 0, 16]);
        payload.extend_from_slice(&[0; 16]);
        payload.extend_from_slice(&[0, 0, 0, 2, 0, 100]);
        payload.extend_from_slice(&[0; 16]);
        payload.extend(cfvo(2, &[]));
        payload.extend(cfvo(7, formula));
        payload
    }

    #[test]
    fn cf12_formulas_are_found_byte_for_byte() {
        let absolute = [0x44, 4, 0, 0, 0];
        let payload = data_bar(&absolute);
        let ranges = cf12_formula_ranges(&payload).unwrap();
        assert_eq!(ranges.len(), 5);
        assert_eq!(&payload[ranges[4].clone()], &absolute);
        // A trailing byte is not accounted for: refused.
        let mut trailing = payload.clone();
        trailing.push(0);
        assert!(cf12_formula_ranges(&trailing).is_err());

        let mut condfmt12 = CONDFMT12.to_le_bytes().to_vec();
        condfmt12.extend_from_slice(&1u16.to_le_bytes());
        condfmt12.extend_from_slice(&ref8(0, 9, 0, 0));
        condfmt12.extend_from_slice(&[1, 0, 1, 0]);
        condfmt12.extend_from_slice(&ref8(0, 9, 0, 0));
        condfmt12.extend_from_slice(&1u16.to_le_bytes());
        condfmt12.extend_from_slice(&ref8(0, 9, 0, 0));
        let mut source = record(CONDFMT12, &condfmt12);
        source.extend(record(CF12, &payload));
        let inserted = patched(&source, INSERT_ROWS).unwrap();
        // The range grows over the point and the absolute maximum moves.
        assert_eq!(&inserted[4 + 4..4 + 12], &ref8(0, 12, 0, 0)[..]);
        let cf12 = &inserted[4 + condfmt12.len() + 4..];
        assert_eq!(&cf12[ranges[4].clone()], &[0x44, 7, 0, 0, 0]);
        assert_eq!(patched(&inserted, DELETE_ROWS).unwrap(), source);
    }

    fn table(flags: u16, first_input: (u16, u16), second_input: (u16, u16)) -> Vec<u8> {
        let mut payload = vec![3, 0, 6, 0, 1, 2];
        payload.extend_from_slice(&flags.to_le_bytes());
        for value in [first_input.0, first_input.1, second_input.0, second_input.1] {
            payload.extend_from_slice(&value.to_le_bytes());
        }
        payload
    }

    #[test]
    fn data_tables_move_whole_with_their_input_cells() {
        let shift = AxisShift::Rows {
            start: 0,
            count: 2,
            insert: true,
        };
        // A two-variable table over B4:C7 with inputs A1 and A2.
        let mut payload = table(0x0008, (0, 0), (1, 0));
        patch_table(&mut payload, shift).unwrap();
        assert_eq!(payload, {
            let mut expected = table(0x0008, (2, 0), (3, 0));
            expected[0..2].copy_from_slice(&5u16.to_le_bytes());
            expected[2..4].copy_from_slice(&8u16.to_le_bytes());
            expected
        });
        // A one-variable table keeps its unused second input as written.
        let mut single = table(0x0000, (0, 0), (0, 0));
        patch_table(&mut single, shift).unwrap();
        assert_eq!(&single[12..16], &[0, 0, 0, 0]);
        // Inserting at the table's first row splits it from its formula row.
        let inside = AxisShift::Rows {
            start: 3,
            count: 1,
            insert: true,
        };
        assert!(patch_table(&mut table(0x0008, (0, 0), (1, 0)), inside).is_err());
        // Before the formula row the whole table moves.
        let before = AxisShift::Rows {
            start: 2,
            count: 1,
            insert: true,
        };
        patch_table(&mut table(0x0008, (0, 0), (1, 0)), before).unwrap();
    }

    #[test]
    fn array_formulas_move_whole_and_refuse_ancillary_data() {
        let mut payload = vec![4, 0, 5, 0, 1, 1, 0, 0, 0, 0, 0, 0];
        payload.extend_from_slice(&5u16.to_le_bytes());
        payload.extend_from_slice(&[0x44, 0, 0, 0, 0]);
        let mut shifted = payload.clone();
        patch_array(&mut shifted, INSERT_ROWS).unwrap();
        assert_eq!(&shifted[0..4], &[7, 0, 8, 0]);
        // A1 is before the point and stays.
        assert_eq!(&shifted[14..], &[0x44, 0, 0, 0, 0]);
        patch_array(&mut shifted, DELETE_ROWS).unwrap();
        assert_eq!(shifted, payload);
        let mut split = payload.clone();
        let through = AxisShift::Rows {
            start: 5,
            count: 1,
            insert: true,
        };
        assert!(patch_array(&mut split, through).is_err());
        let mut extra = payload;
        extra.push(0);
        assert!(patch_array(&mut extra, INSERT_ROWS).is_err());
    }
}
