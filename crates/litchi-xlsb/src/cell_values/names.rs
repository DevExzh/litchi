//! Defined-name replay for the primary workbook transaction (record 0261).
//!
//! A defined name that refers to a worksheet range is one `BrtName` record
//! whose formula is a `PtgRef3d` or `PtgArea3d` over an `Xti` entry of the workbook's
//! `BrtExternSheet` table. Setting a name splices `workbook.bin`: an existing
//! name of the same scope keeps its flags, comment and every other byte and
//! receives the new formula; a new name is appended after the last
//! `BrtName`, so no `PtgName` index in a formula is renumbered. The `Xti` is
//! reused when the table already has one for the sheet, otherwise appended;
//! a workbook without an externals block gains `BrtBeginExternals`,
//! `BrtSupSelf`, `BrtExternSheet` and `BrtEndExternals`. Removing a name is
//! allowed only for the last `BrtName` record for the same reason; the
//! complete workbook readback then refuses any formula that still names it.

#![deny(
    clippy::cast_lossless,
    clippy::cast_possible_truncation,
    clippy::cast_possible_wrap,
    clippy::cast_precision_loss,
    clippy::cast_sign_loss,
    clippy::checked_conversions,
    clippy::expect_used,
    clippy::let_underscore_must_use,
    clippy::unnecessary_unwrap,
    reason = "workbook-stream splices use checked BIFF12 offsets and widths"
)]

use std::ops::Range as Span;

use crate::Workbook;
use crate::formula::{ParsedFormula, Range};
use crate::package::error::{Error, Result};
use crate::package::formula::excel_name_eq;
use crate::raw::{Encoder, Header, Kind, Limits as RawLimits, Records, kind};

/// `PtgRef3d` and `PtgArea3d` with absolute rows and columns.
const PTG_REF_3D: u8 = 0x3A;
const PTG_AREA_3D: u8 = 0x3B;

struct Record {
    kind: Kind,
    span: Span<usize>,
    payload: Span<usize>,
}

fn records(stream: &[u8], limits: RawLimits) -> Result<Vec<Record>> {
    let mut records = Vec::new();
    for record in Records::new(stream, limits) {
        let record = record?;
        let source = stream.get(record.offset()..).ok_or(Error::InvalidFormat(
            "workbook record offset is outside the stream",
        ))?;
        let (_, header) = Header::parse(source, limits)?;
        let payload_start = record
            .offset()
            .checked_add(header)
            .ok_or(Error::CapacityOverflow {
                resource: "workbook record range",
            })?;
        let end = payload_start
            .checked_add(record.len())
            .ok_or(Error::CapacityOverflow {
                resource: "workbook record range",
            })?;
        records.try_reserve(1).map_err(|source| Error::Allocation {
            resource: "workbook record index",
            source,
        })?;
        records.push(Record {
            kind: record.kind(),
            span: record.offset()..end,
            payload: payload_start..end,
        });
    }
    Ok(records)
}

fn encode(kind: Kind, payload: &[u8], limits: RawLimits) -> Result<Vec<u8>> {
    let mut bytes = Vec::new();
    Encoder::new(&mut bytes, limits).write_record(kind, payload)?;
    Ok(bytes)
}

fn read_u32(payload: &[u8], offset: usize) -> Result<u32> {
    let bytes = payload
        .get(offset..offset + 4)
        .ok_or(Error::InvalidFormat("BrtExternSheet is truncated"))?;
    Ok(u32::from_le_bytes([bytes[0], bytes[1], bytes[2], bytes[3]]))
}

/// One splice of the workbook stream: replace `span` with `bytes`.
struct Splice {
    span: Span<usize>,
    bytes: Vec<u8>,
}

fn apply_splices(stream: &[u8], mut splices: Vec<Splice>) -> Result<Vec<u8>> {
    splices.sort_by_key(|splice| (splice.span.start, splice.span.end));
    let mut output = Vec::new();
    output
        .try_reserve(stream.len().saturating_add(256))
        .map_err(|source| Error::Allocation {
            resource: "workbook defined-name rewrite",
            source,
        })?;
    let mut cursor = 0;
    for splice in splices {
        if splice.span.start < cursor {
            return Err(Error::InvalidFormat("overlapping workbook stream splices"));
        }
        output.extend_from_slice(&stream[cursor..splice.span.start]);
        output.extend_from_slice(&splice.bytes);
        cursor = splice.span.end;
    }
    output.extend_from_slice(&stream[cursor..]);
    Ok(output)
}

/// The `Xti` index of `sheet`, with the splices that create it when the
/// table does not have one yet.
fn sheet_xti(
    stream: &[u8],
    records: &[Record],
    sheet: u32,
    splices: &mut Vec<Splice>,
    limits: RawLimits,
) -> Result<u16> {
    let sheet = i32::try_from(sheet)
        .map_err(|_error| Error::InvalidFormat("sheet position exceeds the Xti field"))?;
    let begin = records
        .iter()
        .position(|record| record.kind == kind::BEGIN_EXTERNALS);
    let Some(begin) = begin else {
        // No externals block: author one before the first record that
        // follows EXTERNALS in the workbook grammar.
        let anchor = records
            .iter()
            .position(|record| record.kind == kind::END_BUNDLE_SHS)
            .ok_or(Error::InvalidFormat(
                "workbook stream has no BrtEndBundleShs",
            ))?;
        let at = records[anchor + 1..]
            .iter()
            .find(|record| matches!(record.kind, kind::NAME | kind::CALC_PROP | kind::END_BOOK))
            .map_or(records[anchor].span.end, |record| record.span.start);
        let mut table = Vec::new();
        table.extend_from_slice(&1_u32.to_le_bytes());
        table.extend_from_slice(&0_u32.to_le_bytes());
        table.extend_from_slice(&sheet.to_le_bytes());
        table.extend_from_slice(&sheet.to_le_bytes());
        let mut bytes = encode(kind::BEGIN_EXTERNALS, &[], limits)?;
        bytes.extend(encode(kind::SUP_SELF, &[], limits)?);
        bytes.extend(encode(kind::EXTERN_SHEET, &table, limits)?);
        bytes.extend(encode(kind::END_EXTERNALS, &[], limits)?);
        splices.push(Splice {
            span: at..at,
            bytes,
        });
        return Ok(0);
    };
    let block = &records[begin..];
    let table = block
        .iter()
        .position(|record| record.kind == kind::EXTERN_SHEET)
        .map(|index| &block[index])
        .ok_or_else(|| {
            Error::refused(
                litchi_core::edit::RefusalKind::UnsupportedContent,
                "edit",
                "the externals block has no BrtExternSheet",
            )
        })?;
    let supbooks = block
        .iter()
        .take_while(|record| record.kind != kind::EXTERN_SHEET)
        .filter(|record| {
            matches!(
                record.kind,
                kind::SUP_SELF | kind::SUP_SAME | kind::SUP_BOOK_SRC | kind::SUP_ADDIN
            )
        })
        .map(|record| record.kind)
        .collect::<Vec<_>>();
    let self_book = supbooks.iter().position(|kind| *kind == kind::SUP_SELF);
    let payload = &stream[table.payload.clone()];
    let count = read_u32(payload, 0)?;
    let count_usize = usize::try_from(count)
        .map_err(|_error| Error::InvalidFormat("BrtExternSheet count exceeds this platform"))?;
    if let Some(book) = self_book {
        let book = u32::try_from(book)
            .map_err(|_error| Error::InvalidFormat("supporting link index overflow"))?;
        for index in 0..count_usize {
            let base = 4 + index * 12;
            if read_u32(payload, base)? == book
                && read_u32(payload, base + 4)? == sheet.cast_unsigned()
                && read_u32(payload, base + 8)? == sheet.cast_unsigned()
            {
                return u16::try_from(index).map_err(|_error| {
                    Error::refused(
                        litchi_core::edit::RefusalKind::UnsupportedContent,
                        "edit",
                        "the Xti index exceeds 65,535",
                    )
                });
            }
        }
    }
    let book = match self_book {
        Some(book) => book,
        None => {
            // Append BrtSupSelf as the last supporting link so existing
            // supporting-link indexes keep their meaning.
            splices.push(Splice {
                span: table.span.start..table.span.start,
                bytes: encode(kind::SUP_SELF, &[], limits)?,
            });
            supbooks.len()
        },
    };
    let book = u32::try_from(book)
        .map_err(|_error| Error::InvalidFormat("supporting link index overflow"))?;
    let next = count.checked_add(1).ok_or_else(|| {
        Error::refused(
            litchi_core::edit::RefusalKind::UnsupportedContent,
            "edit",
            "the Xti table is full",
        )
    })?;
    let ixti = u16::try_from(count).map_err(|_error| {
        Error::refused(
            litchi_core::edit::RefusalKind::UnsupportedContent,
            "edit",
            "the Xti index exceeds 65,535",
        )
    })?;
    let mut rewritten = Vec::new();
    rewritten.extend_from_slice(&next.to_le_bytes());
    rewritten.extend_from_slice(payload.get(4..).unwrap_or_default());
    rewritten.extend_from_slice(&book.to_le_bytes());
    rewritten.extend_from_slice(&sheet.to_le_bytes());
    rewritten.extend_from_slice(&sheet.to_le_bytes());
    splices.push(Splice {
        span: table.span.clone(),
        bytes: encode(kind::EXTERN_SHEET, &rewritten, limits)?,
    });
    Ok(ixti)
}

fn area_formula(ixti: u16, range: Range, raw_limits: crate::raw::Limits) -> Result<Vec<u8>> {
    let first_column = u16::try_from(range.col_first)
        .map_err(|_error| Error::InvalidCellReference("column exceeds the PtgArea3d field"))?;
    let last_column = u16::try_from(range.col_last)
        .map_err(|_error| Error::InvalidCellReference("column exceeds the PtgArea3d field"))?;
    let mut rgce = Vec::with_capacity(15);
    if range.row_first == range.row_last && first_column == last_column {
        // One cell is a `PtgRef3d`, as Excel stores it.
        rgce.push(PTG_REF_3D);
        rgce.extend_from_slice(&ixti.to_le_bytes());
        rgce.extend_from_slice(&range.row_first.to_le_bytes());
        rgce.extend_from_slice(&first_column.to_le_bytes());
    } else {
        rgce.push(PTG_AREA_3D);
        rgce.extend_from_slice(&ixti.to_le_bytes());
        rgce.extend_from_slice(&range.row_first.to_le_bytes());
        rgce.extend_from_slice(&range.row_last.to_le_bytes());
        rgce.extend_from_slice(&first_column.to_le_bytes());
        rgce.extend_from_slice(&last_column.to_le_bytes());
    }
    Ok(ParsedFormula {
        rgce,
        rgcb: Vec::new(),
    }
    .to_bytes(raw_limits)?)
}

/// The offset of the formula and the end of the name in a `BrtName`
/// payload: flags (4), chKey (1), itab (4), then the name string.
fn formula_offset(payload: &[u8]) -> Result<usize> {
    let units = read_u32(payload, 9)?;
    let units = usize::try_from(units)
        .map_err(|_error| Error::InvalidFormat("BrtName name length exceeds this platform"))?;
    units
        .checked_mul(2)
        .and_then(|bytes| bytes.checked_add(13))
        .filter(|end| *end <= payload.len())
        .ok_or(Error::InvalidFormat("BrtName name is truncated"))
}

/// Set (`Some`) or remove (`None`) the defined name `name` of `scope`
/// (a complete-catalog position, or the workbook). Returns whether the
/// workbook changed.
pub(super) fn apply(
    workbook: &mut Workbook,
    name: &str,
    scope: Option<usize>,
    target: Option<(usize, Range)>,
) -> Result<bool> {
    let raw_limits = workbook.raw_limits_value();
    let part = workbook.package.main_document_part()?;
    let stream = part.blob().to_vec();
    let records = records(&stream, raw_limits)?;
    let scope_wire = match scope {
        Some(position) => u32::try_from(position)
            .map_err(|_error| Error::InvalidFormat("scope position exceeds the BrtName field"))?,
        None => u32::MAX,
    };
    let names = records
        .iter()
        .enumerate()
        .filter(|(_, record)| record.kind == kind::NAME)
        .map(|(index, _)| index)
        .collect::<Vec<_>>();
    let mut existing = None;
    for &index in &names {
        let definition = crate::named_ranges::Definition::parse(
            &stream[records[index].payload.clone()],
            raw_limits,
        )?;
        if excel_name_eq(&definition.name, name)
            && definition.sheet_id.unwrap_or(u32::MAX) == scope_wire
        {
            existing = Some(index);
        }
    }
    let mut splices = Vec::new();
    match target {
        None => {
            let Some(index) = existing else {
                return Ok(false);
            };
            if names.last() != Some(&index) {
                return Err(Error::refused(
                    litchi_core::edit::RefusalKind::UnsupportedContent,
                    "edit",
                    "only the last defined name can be removed without renumbering name references",
                ));
            }
            splices.push(Splice {
                span: records[index].span.clone(),
                bytes: Vec::new(),
            });
        },
        Some((sheet, range)) => {
            let sheet = u32::try_from(sheet)
                .map_err(|_error| Error::InvalidFormat("sheet position exceeds the Xti field"))?;
            let ixti = sheet_xti(&stream, &records, sheet, &mut splices, raw_limits)?;
            let formula = area_formula(ixti, range, raw_limits)?;
            match existing {
                Some(index) => {
                    let payload = &stream[records[index].payload.clone()];
                    let start = formula_offset(payload)?;
                    let (_, consumed) = ParsedFormula::parse(&payload[start..], raw_limits)?;
                    let old = payload
                        .get(start..start + consumed)
                        .ok_or(Error::InvalidFormat("BrtName formula is truncated"))?;
                    if old == formula.as_slice() {
                        return Ok(false);
                    }
                    let mut rewritten = payload[..start].to_vec();
                    rewritten.extend_from_slice(&formula);
                    rewritten.extend_from_slice(&payload[start + consumed..]);
                    splices.push(Splice {
                        span: records[index].span.clone(),
                        bytes: encode(kind::NAME, &rewritten, raw_limits)?,
                    });
                },
                None => {
                    crate::named_ranges::validate_name(name)?;
                    let mut payload = Vec::new();
                    {
                        let mut writer = Encoder::new(&mut payload, raw_limits);
                        writer.write_u32(0)?;
                        writer.write_u8(0)?;
                        writer.write_u32(scope_wire)?;
                        writer.write_wide_string(name)?;
                    }
                    payload.extend_from_slice(&formula);
                    payload.extend_from_slice(&u32::MAX.to_le_bytes());
                    let bytes = encode(kind::NAME, &payload, raw_limits)?;
                    let at = match names.last() {
                        Some(&last) => records[last].span.end,
                        None => match records
                            .iter()
                            .position(|record| record.kind == kind::END_EXTERNALS)
                        {
                            Some(end) => records[end].span.end,
                            // The externals block authored by `sheet_xti`
                            // ends where it is inserted.
                            None => splices
                                .iter()
                                .find(|splice| splice.span.is_empty())
                                .map(|splice| splice.span.end)
                                .ok_or(Error::InvalidFormat(
                                    "no position for a new defined name",
                                ))?,
                        },
                    };
                    splices.push(Splice {
                        span: at..at,
                        bytes,
                    });
                },
            }
        },
    }
    let rewritten = apply_splices(&stream, splices)?;
    let uri = part.partname().clone();
    workbook.edit_opc(|package| {
        package.get_part_mut(&uri)?.set_blob(rewritten);
        Ok(())
    })?;
    Ok(true)
}
