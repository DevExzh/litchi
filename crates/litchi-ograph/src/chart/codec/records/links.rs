//! BRAI link and formula record codecs.

use super::super::super::Kind as ChartKind;
use super::super::super::model::{CellRef, Context, Link, Role, RowCol, Source};
use super::wire::{byte_at, copy, exact, invalid, invalid_model, limit, u16_at, vec_with_capacity};
use crate::{Error, Limits, Result};
use litchi_biff::RecordRef;

pub(super) fn parse_link(record: RecordRef<'_>, context: Context, limits: Limits) -> Result<Link> {
    let data = record.payload();
    match context.kind() {
        ChartKind::Graph => {
            exact(record, 8)?;
            let flags = u16_at(data, 2, record)?;
            if flags & 0x0002 == 0 || flags & !0x0003 != 0 {
                return invalid(record, "Graph BRAI reserved bits are invalid");
            }
            let row_col = RowCol::new(u16_at(data, 6, record)?).ok_or(Error::InvalidChart {
                offset: record.offset(),
                reason: "Graph BRAI row or column exceeds 3,999",
            })?;
            Ok(Link::Graph {
                role: parse_role(byte_at(data, 0, record)?, record)?,
                source: parse_source(byte_at(data, 1, record)?, record)?,
                unlinked_format: flags & 1 != 0,
                number_format: u16_at(data, 4, record)?,
                row_col,
            })
        },
        ChartKind::Excel => {
            if data.len() < 8 {
                return invalid(record, "Excel BRAI is shorter than eight bytes");
            }
            let flags = u16_at(data, 2, record)?;
            if flags & !1 != 0 {
                return invalid(record, "Excel BRAI reserved bits are nonzero");
            }
            let formula_len = usize::from(u16_at(data, 6, record)?);
            let expected = 8usize.checked_add(formula_len).ok_or(Error::SizeOverflow {
                resource: "BRAI formula",
            })?;
            if data.len() != expected {
                return invalid(
                    record,
                    "Excel BRAI formula length does not match its payload",
                );
            }
            if formula_len > limits.max_formula_bytes {
                return limit("formula bytes", formula_len, limits.max_formula_bytes);
            }
            let tokens = data.get(8..).ok_or(Error::InvalidChart {
                offset: record.offset(),
                reason: "Excel BRAI formula is truncated",
            })?;
            let formula = copy(tokens, "formula tokens", limits.max_formula_bytes)?;
            let refs = parse_refs(tokens, record)?;
            let link = Link::Excel {
                role: parse_role(byte_at(data, 0, record)?, record)?,
                source: parse_source(byte_at(data, 1, record)?, record)?,
                unlinked_format: flags & 1 != 0,
                number_format: u16_at(data, 4, record)?,
                formula,
                refs,
            };
            if matches!(
                &link,
                Link::Excel {
                    source: Source::Automatic,
                    formula,
                    ..
                } if !formula.is_empty()
            ) {
                return invalid(record, "automatic Excel BRAI has a nonempty formula");
            }
            Ok(link)
        },
    }
}

fn parse_refs(tokens: &[u8], record: RecordRef<'_>) -> Result<Vec<CellRef>> {
    if tokens.is_empty() {
        return Ok(Vec::new());
    }
    let opcode = byte_at(tokens, 0, record)? & 0x1F;
    let value = match (opcode, tokens.len()) {
        (0x1A, 7) => {
            let col = u16_at(tokens, 5, record)? & 0x3FFF;
            let col = u8::try_from(col).ok().ok_or(Error::InvalidChart {
                offset: record.offset(),
                reason: "chart formula column exceeds the BIFF8 grid",
            })?;
            Some(CellRef {
                external_sheet: u16_at(tokens, 1, record)?,
                first_row: u16_at(tokens, 3, record)?,
                last_row: u16_at(tokens, 3, record)?,
                first_col: col,
                last_col: col,
            })
        },
        (0x1B, 11) => {
            let first_col = u16_at(tokens, 7, record)? & 0x3FFF;
            let last_col = u16_at(tokens, 9, record)? & 0x3FFF;
            Some(CellRef {
                external_sheet: u16_at(tokens, 1, record)?,
                first_row: u16_at(tokens, 3, record)?,
                last_row: u16_at(tokens, 5, record)?,
                first_col: u8::try_from(first_col).ok().ok_or(Error::InvalidChart {
                    offset: record.offset(),
                    reason: "chart formula column exceeds the BIFF8 grid",
                })?,
                last_col: u8::try_from(last_col).ok().ok_or(Error::InvalidChart {
                    offset: record.offset(),
                    reason: "chart formula column exceeds the BIFF8 grid",
                })?,
            })
        },
        _ => None,
    };
    let Some(value) = value else {
        return Ok(Vec::new());
    };
    let mut refs = Vec::new();
    refs.try_reserve_exact(1).ok().ok_or(Error::Allocation {
        resource: "chart references",
    })?;
    refs.push(value);
    Ok(refs)
}

fn parse_role(value: u8, record: RecordRef<'_>) -> Result<Role> {
    match value {
        0 => Ok(Role::Name),
        1 => Ok(Role::Values),
        2 => Ok(Role::Categories),
        3 => Ok(Role::Bubbles),
        _ => invalid(record, "BRAI role is outside the defined range"),
    }
}

fn parse_source(value: u8, record: RecordRef<'_>) -> Result<Source> {
    match value {
        0 => Ok(Source::Automatic),
        1 => Ok(Source::Literal),
        2 => Ok(Source::Cells),
        _ => invalid(record, "BRAI source is outside the defined range"),
    }
}

pub(super) fn validate_link(link: &Link, context: Context, limits: Limits) -> Result<()> {
    match (context.kind(), link) {
        (ChartKind::Graph, Link::Graph { .. }) => {},
        (
            ChartKind::Excel,
            Link::Excel {
                source,
                formula,
                refs,
                ..
            },
        ) => {
            if formula.len() > limits.max_formula_bytes {
                return limit("formula bytes", formula.len(), limits.max_formula_bytes);
            }
            let maximum = limits.biff.max_record_bytes.saturating_sub(8);
            if formula.len() > maximum {
                return limit("formula bytes", formula.len(), maximum);
            }
            if *source == Source::Automatic && !formula.is_empty() {
                return invalid_model("link", "automatic Excel BRAI has a nonempty formula");
            }
            for value in refs {
                if value.first_row > value.last_row || value.first_col > value.last_col {
                    return invalid_model("link", "cell range is reversed");
                }
                if let Some(count) = context.external_sheet_count()
                    && usize::from(value.external_sheet) >= count
                {
                    return invalid_model("link", "external-sheet index is out of range");
                }
            }
        },
        (ChartKind::Graph, Link::Excel { .. }) => {
            return invalid_model("link", "Excel BRAI cannot be encoded in a Graph chart");
        },
        (ChartKind::Excel, Link::Graph { .. }) => {
            return invalid_model("link", "Graph BRAI cannot be encoded in an Excel chart");
        },
    }
    Ok(())
}

pub(super) fn encode_link(link: &Link, context: Context, limits: Limits) -> Result<Vec<u8>> {
    validate_link(link, context, limits)?;
    match link {
        Link::Graph {
            role,
            source,
            unlinked_format,
            number_format,
            row_col,
        } => {
            let mut data = vec_with_capacity(8, "Graph BRAI")?;
            data.push(*role as u8);
            data.push(*source as u8);
            data.extend_from_slice(&(u16::from(*unlinked_format) | 2).to_le_bytes());
            data.extend_from_slice(&number_format.to_le_bytes());
            data.extend_from_slice(&row_col.get().to_le_bytes());
            Ok(data)
        },
        Link::Excel {
            role,
            source,
            unlinked_format,
            number_format,
            formula,
            ..
        } => {
            let capacity = 8usize
                .checked_add(formula.len())
                .ok_or(Error::SizeOverflow {
                    resource: "Excel BRAI",
                })?;
            let mut data = vec_with_capacity(capacity, "Excel BRAI")?;
            data.push(*role as u8);
            data.push(*source as u8);
            data.extend_from_slice(&u16::from(*unlinked_format).to_le_bytes());
            data.extend_from_slice(&number_format.to_le_bytes());
            let length = u16::try_from(formula.len())
                .ok()
                .ok_or(Error::InvalidModel {
                    field: "link",
                    reason: "formula length exceeds u16",
                })?;
            data.extend_from_slice(&length.to_le_bytes());
            data.extend_from_slice(formula);
            Ok(data)
        },
    }
}
