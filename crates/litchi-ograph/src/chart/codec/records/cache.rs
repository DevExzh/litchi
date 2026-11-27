//! Cached cell and Dimensions-related BIFF record codecs.

use super::super::super::Kind as ChartKind;
use super::super::super::cache as chart_cache;
use super::super::super::model::{Cache, RowCol, Value, XlValue};
use super::text::{biff_string, xl_unicode_string};
use super::wire::{
    CELL_LABEL, EXCEL_BLANK, EXCEL_BOOL_ERR, EXCEL_NUMBER, GRAPH_BLANK, GRAPH_NUMBER, push_record,
    put_byte, put_f64, put_u16, u16_at, vec_with_capacity,
};
use crate::{Error, Result};
use litchi_biff::{Encoder, Kind as RecordKind, RecordRef};

pub(super) fn encode_cache(out: &mut Encoder, cache: &Cache) -> Result<()> {
    match cache {
        Cache::Excel {
            row,
            col,
            xf,
            value,
            ..
        } => encode_excel_cache(out, *row, *col, *xf, value),
        Cache::Graph {
            row,
            col,
            ifmt,
            value,
        } => encode_graph_cache(out, *row, *col, *ifmt, value),
    }
}

fn encode_excel_cache(
    out: &mut Encoder,
    row: u16,
    col: u8,
    xf: chart_cache::Xf,
    value: &XlValue,
) -> Result<()> {
    match value {
        XlValue::Number(number) => {
            let mut data = [0u8; 14];
            put_u16(&mut data, 0, row)?;
            put_u16(&mut data, 2, u16::from(col))?;
            put_u16(&mut data, 4, xf.get())?;
            put_f64(&mut data, 6, *number)?;
            push_record(out, EXCEL_NUMBER, &data)
        },
        XlValue::Text(text) => {
            let string = xl_unicode_string(text)?;
            let capacity = 6usize
                .checked_add(string.len())
                .ok_or(Error::SizeOverflow {
                    resource: "cached chart string",
                })?;
            let mut data = vec_with_capacity(capacity, "cached chart string")?;
            data.extend_from_slice(&row.to_le_bytes());
            data.extend_from_slice(&u16::from(col).to_le_bytes());
            data.extend_from_slice(&xf.get().to_le_bytes());
            data.extend_from_slice(&string);
            push_record(out, CELL_LABEL, &data)
        },
        XlValue::Bool(value) => {
            let mut data = [0u8; 8];
            put_u16(&mut data, 0, row)?;
            put_u16(&mut data, 2, u16::from(col))?;
            put_u16(&mut data, 4, xf.get())?;
            put_byte(&mut data, 6, u8::from(*value))?;
            push_record(out, EXCEL_BOOL_ERR, &data)
        },
        XlValue::Error(value) => {
            let mut data = [0u8; 8];
            put_u16(&mut data, 0, row)?;
            put_u16(&mut data, 2, u16::from(col))?;
            put_u16(&mut data, 4, xf.get())?;
            put_byte(&mut data, 6, *value as u8)?;
            put_byte(&mut data, 7, 1)?;
            push_record(out, EXCEL_BOOL_ERR, &data)
        },
        XlValue::Blank => {
            let mut data = [0u8; 6];
            put_u16(&mut data, 0, row)?;
            put_u16(&mut data, 2, u16::from(col))?;
            put_u16(&mut data, 4, xf.get())?;
            push_record(out, EXCEL_BLANK, &data)
        },
    }
}

fn encode_graph_cache(
    out: &mut Encoder,
    row: RowCol,
    col: RowCol,
    ifmt: chart_cache::Ifmt,
    value: &Value,
) -> Result<()> {
    match value {
        Value::Number(number) => {
            let mut data = [0u8; 15];
            put_u16(&mut data, 0, row.get())?;
            put_u16(&mut data, 2, col.get())?;
            put_u16(&mut data, 5, ifmt.get())?;
            put_f64(&mut data, 7, *number)?;
            push_record(out, GRAPH_NUMBER, &data)
        },
        Value::Text(text) => {
            let string = biff_string(text)?;
            let capacity = 7usize
                .checked_add(string.len())
                .ok_or(Error::SizeOverflow {
                    resource: "cached chart string",
                })?;
            let mut data = vec_with_capacity(capacity, "cached chart string")?;
            data.extend_from_slice(&row.get().to_le_bytes());
            data.extend_from_slice(&col.get().to_le_bytes());
            data.push(0);
            data.extend_from_slice(&ifmt.get().to_le_bytes());
            data.extend_from_slice(&string);
            push_record(out, CELL_LABEL, &data)
        },
        Value::Blank => {
            let mut data = [0u8; 7];
            put_u16(&mut data, 0, row.get())?;
            put_u16(&mut data, 2, col.get())?;
            put_u16(&mut data, 5, ifmt.get())?;
            push_record(out, GRAPH_BLANK, &data)
        },
    }
}

pub(super) fn blank_kind(kind: ChartKind) -> RecordKind {
    match kind {
        ChartKind::Graph => GRAPH_BLANK,
        ChartKind::Excel => EXCEL_BLANK,
    }
}

pub(super) fn number_kind(kind: ChartKind) -> RecordKind {
    match kind {
        ChartKind::Graph => GRAPH_NUMBER,
        ChartKind::Excel => EXCEL_NUMBER,
    }
}

pub(super) fn excel_cache(
    data: &[u8],
    section: chart_cache::Index,
    xf: chart_cache::Xf,
    value: XlValue,
    record: RecordRef<'_>,
) -> Result<Cache> {
    let row = u16_at(data, 0, record)?;
    let col = u16_at(data, 2, record)?;
    Ok(Cache::excel(
        section,
        row,
        u8::try_from(col).ok().ok_or(Error::InvalidChart {
            offset: record.offset(),
            reason: "Excel cache column exceeds the BIFF8 grid",
        })?,
        xf,
        value,
    ))
}

pub(super) fn graph_cache(
    data: &[u8],
    ifmt: chart_cache::Ifmt,
    value: Value,
    record: RecordRef<'_>,
) -> Result<Cache> {
    Ok(Cache::graph(
        RowCol::new(u16_at(data, 0, record)?).ok_or(Error::InvalidChart {
            offset: record.offset(),
            reason: "Graph cache row exceeds 3,999",
        })?,
        RowCol::new(u16_at(data, 2, record)?).ok_or(Error::InvalidChart {
            offset: record.offset(),
            reason: "Graph cache column exceeds 3,999",
        })?,
        ifmt,
        value,
    ))
}
