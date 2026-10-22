#![allow(
    clippy::expect_used,
    clippy::map_err_ignore,
    reason = "legacy module confines extraction after an immediately preceding structural invariant check, normalization into the module's stable typed public error to this codec boundary"
)]

//! Fixed and variable payload codecs for the scenario record family.

use super::super::model::{CellRange, ChangedCell};
use crate::package::error::{Error, Result};
use crate::raw::{Cursor, Encoder, Kind};
use std::io::Write;

const BEGIN_MANAGER: u16 = 500;
const END_MANAGER: u16 = 501;
const BEGIN_SCENARIO: u16 = 502;
const END_SCENARIO: u16 = 503;
const SCENARIO_CELL: u16 = 504;
const NO_RESULT: u32 = u32::MAX;

fn kind(value: u16) -> Kind {
    Kind::new(value).expect("scenario record kind is within BIFF12 range")
}

pub(crate) fn begin_manager() -> Kind {
    kind(BEGIN_MANAGER)
}

pub(crate) fn end_manager() -> Kind {
    kind(END_MANAGER)
}

pub(crate) fn begin_scenario() -> Kind {
    kind(BEGIN_SCENARIO)
}

pub(crate) fn end_scenario() -> Kind {
    kind(END_SCENARIO)
}

pub(crate) fn scenario_cell() -> Kind {
    kind(SCENARIO_CELL)
}

pub(crate) struct Header {
    pub current: u16,
    pub shown: u16,
    pub result_ranges: Vec<CellRange>,
}

pub(crate) struct ChildHeader {
    pub count: u16,
    pub locked: bool,
    pub hidden: bool,
    pub name: String,
    pub comment: String,
    pub user_name: String,
}

pub(crate) fn parse_manager_header(data: &[u8], limits: crate::raw::Limits) -> Result<Header> {
    let mut cursor = Cursor::new(data, "BrtBeginScenMan", limits);
    let current = cursor.read_u16()?;
    let shown = cursor.read_u16()?;
    let count = cursor.read_u32()?;
    let result_ranges = if count == NO_RESULT {
        Vec::new()
    } else {
        let count = usize::try_from(count).map_err(|_| Error::Unrecognized {
            typ: "BrtBeginScenMan sqrfxResult",
            val: "range count overflows usize",
        })?;
        if !(1..=32).contains(&count) {
            return Err(Error::Unrecognized {
                typ: "BrtBeginScenMan sqrfxResult",
                val: "range count is outside 132",
            });
        }
        let mut ranges = Vec::with_capacity(count);
        for _ in 0..count {
            ranges.push(CellRange::new(
                cursor.read_u32()?,
                cursor.read_u32()?,
                cursor.read_u32()?,
                cursor.read_u32()?,
            )?);
        }
        ranges
    };
    cursor.finish()?;
    Ok(Header {
        current,
        shown,
        result_ranges,
    })
}

pub(crate) fn parse_scenario_header(
    data: &[u8],
    limits: crate::raw::Limits,
) -> Result<ChildHeader> {
    let mut cursor = Cursor::new(data, "BrtBeginSct", limits);
    let count = cursor.read_u16()?;
    let locked = cursor.read_bool32()?;
    let hidden = cursor.read_bool32()?;
    let name = cursor.read_wide_string()?;
    let comment = cursor.read_wide_string()?;
    let user_name = cursor.read_wide_string()?;
    cursor.finish()?;
    Ok(ChildHeader {
        count,
        locked,
        hidden,
        name,
        comment,
        user_name,
    })
}

pub(crate) fn parse_changed_cell(data: &[u8], limits: crate::raw::Limits) -> Result<ChangedCell> {
    let mut cursor = Cursor::new(data, "BrtSlc", limits);
    let row = cursor.read_u32()?;
    let column = cursor.read_u32()?;
    let reserved = cursor.read_u32()?;
    if reserved != 0 {
        return Err(Error::Unrecognized {
            typ: "BrtSlc fReserved",
            val: "0x",
        });
    }
    let unused = cursor.read_u32()?;
    let number_format = cursor.read_u16()?;
    let value = cursor.read_wide_string()?;
    cursor.finish()?;
    ChangedCell::from_wire(row, column, number_format, value, unused)
}

pub(crate) fn write_manager_header(
    current: Option<usize>,
    shown: Option<usize>,
    result_ranges: &[CellRange],
    limits: crate::raw::Limits,
) -> Result<Vec<u8>> {
    let mut payload = Vec::new();
    let mut writer = Encoder::new(&mut payload, limits);
    writer.write_u16(match current {
        Some(index) => u16::try_from(index).map_err(|_| Error::Unrecognized {
            typ: "BrtBeginScenMan isctCur",
            val: "scenario index exceeds u16",
        })?,
        None => u16::MAX,
    })?;
    writer.write_u16(match shown {
        Some(index) => u16::try_from(index).map_err(|_| Error::Unrecognized {
            typ: "BrtBeginScenMan isctShown",
            val: "scenario index exceeds u16",
        })?,
        None => u16::MAX,
    })?;
    if result_ranges.is_empty() {
        writer.write_u32(NO_RESULT)?;
    } else {
        writer.write_u32(u32::try_from(result_ranges.len()).map_err(|_| {
            Error::Unrecognized {
                typ: "BrtBeginScenMan sqrfxResult",
                val: "range count exceeds u32",
            }
        })?)?;
        for range in result_ranges {
            writer.write_u32(range.row_first())?;
            writer.write_u32(range.row_last())?;
            writer.write_u32(range.column_first())?;
            writer.write_u32(range.column_last())?;
        }
    }
    Ok(payload)
}

pub(crate) fn write_scenario_header(
    count: usize,
    locked: bool,
    hidden: bool,
    name: &str,
    comment: &str,
    user_name: &str,
    limits: crate::raw::Limits,
) -> Result<Vec<u8>> {
    let mut payload = Vec::new();
    let mut writer = Encoder::new(&mut payload, limits);
    writer.write_u16(u16::try_from(count).map_err(|_| Error::Unrecognized {
        typ: "BrtBeginSct cref",
        val: "changed-cell count exceeds u16",
    })?)?;
    writer.write_u32(u32::from(locked))?;
    writer.write_u32(u32::from(hidden))?;
    writer.write_wide_string(name)?;
    writer.write_wide_string(comment)?;
    writer.write_wide_string(user_name)?;
    Ok(payload)
}

pub(crate) fn write_changed_cell(
    cell: &ChangedCell,
    limits: crate::raw::Limits,
) -> Result<Vec<u8>> {
    let mut payload = Vec::new();
    let mut writer = Encoder::new(&mut payload, limits);
    writer.write_u32(cell.row())?;
    writer.write_u32(cell.column())?;
    writer.write_u32(0)?;
    writer.write_u32(cell.unused())?;
    writer.write_u16(cell.number_format())?;
    writer.write_wide_string(cell.value())?;
    Ok(payload)
}

pub(crate) fn write_record<W: Write>(
    writer: &mut Encoder<W>,
    kind: Kind,
    payload: &[u8],
) -> Result<()> {
    writer.write_record(kind, payload).map_err(Into::into)
}
