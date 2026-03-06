//! Cell record BIFF8 writers.

use crate::ole::xls::{XlsError, XlsResult};
use std::io::Write;

use super::write_record_header;

fn encode_rk(value: f64) -> Option<u32> {
    let int_value = value as i32;
    if f64::from(int_value) == value && (-(1 << 29)..(1 << 29)).contains(&int_value) {
        return Some(((int_value as u32) << 2) | 0x02);
    }

    let scaled = value * 100.0;
    let scaled_int = scaled as i32;
    if f64::from(scaled_int) == scaled && (-(1 << 29)..(1 << 29)).contains(&scaled_int) {
        return Some(((scaled_int as u32) << 2) | 0x03);
    }

    None
}

/// Write NUMBER record (floating point cell)
///
/// Record type: 0x0203
///
/// # Arguments
///
/// * `writer` - Output writer
/// * `row` - Row index (0-based)
/// * `col` - Column index (0-based)
/// * `value` - Cell value (f64)
pub fn write_number<W: Write>(
    writer: &mut W,
    row: u32,
    col: u16,
    xf_index: u16,
    value: f64,
) -> XlsResult<()> {
    // BIFF8 stores row as a 16-bit index (0..65535)
    let row_u16 = u16::try_from(row).map_err(|_| {
        XlsError::InvalidData(format!(
            "Row index {} exceeds BIFF8 limit 65535 for NUMBER record",
            row
        ))
    })?;

    write_record_header(writer, 0x0203, 14)?;

    writer.write_all(&row_u16.to_le_bytes())?;
    writer.write_all(&col.to_le_bytes())?;

    // XF record index
    writer.write_all(&xf_index.to_le_bytes())?;

    // IEEE 754 floating point value
    writer.write_all(&value.to_le_bytes())?;

    Ok(())
}

pub fn write_mulrk<W: Write>(
    writer: &mut W,
    row: u32,
    first_col: u16,
    values: &[(u16, f64)],
) -> XlsResult<()> {
    let row_u16 = u16::try_from(row).map_err(|_| {
        XlsError::InvalidData(format!(
            "Row index {} exceeds BIFF8 limit 65535 for MULRK record",
            row
        ))
    })?;
    if values.len() < 2 {
        return Err(XlsError::InvalidData(
            "MULRK record requires at least two contiguous numeric cells".to_string(),
        ));
    }
    let last_col = first_col
        .checked_add(u16::try_from(values.len() - 1).map_err(|_| {
            XlsError::InvalidData("MULRK column span exceeds BIFF8 limit".to_string())
        })?)
        .ok_or_else(|| {
            XlsError::InvalidData("MULRK last column exceeds BIFF8 limit".to_string())
        })?;
    let data_len = 6u16
        .checked_add(u16::try_from(values.len() * 6).map_err(|_| {
            XlsError::InvalidData("MULRK payload exceeds BIFF8 size limit".to_string())
        })?)
        .ok_or_else(|| {
            XlsError::InvalidData("MULRK payload exceeds BIFF8 size limit".to_string())
        })?;

    write_record_header(writer, 0x00BD, data_len)?;
    writer.write_all(&row_u16.to_le_bytes())?;
    writer.write_all(&first_col.to_le_bytes())?;
    for (xf_index, value) in values {
        let rk = encode_rk(*value).ok_or_else(|| {
            XlsError::InvalidData(format!("Value {value} cannot be encoded as RK for MULRK"))
        })?;
        writer.write_all(&xf_index.to_le_bytes())?;
        writer.write_all(&rk.to_le_bytes())?;
    }
    writer.write_all(&last_col.to_le_bytes())?;
    Ok(())
}

/// Write LABELSST record (string cell with reference to SST)
///
/// Record type: 0x00FD
///
/// # Arguments
///
/// * `writer` - Output writer
/// * `row` - Row index (0-based)
/// * `col` - Column index (0-based)
/// * `sst_index` - Index into shared string table
pub fn write_labelsst<W: Write>(
    writer: &mut W,
    row: u32,
    col: u16,
    xf_index: u16,
    sst_index: u32,
) -> XlsResult<()> {
    // BIFF8 stores row as a 16-bit index (0..65535)
    let row_u16 = u16::try_from(row).map_err(|_| {
        XlsError::InvalidData(format!(
            "Row index {} exceeds BIFF8 limit 65535 for LABELSST record",
            row
        ))
    })?;

    // 2 (row) + 2 (col) + 2 (xf) + 4 (sst index) = 10 bytes
    write_record_header(writer, 0x00FD, 10)?;

    writer.write_all(&row_u16.to_le_bytes())?;
    writer.write_all(&col.to_le_bytes())?;

    // XF record index
    writer.write_all(&xf_index.to_le_bytes())?;

    // SST index
    writer.write_all(&sst_index.to_le_bytes())?;

    Ok(())
}

/// Write BOOLERR record (boolean or error cell)
///
/// Record type: 0x0205
///
/// # Arguments
///
/// * `writer` - Output writer
/// * `row` - Row index (0-based)
/// * `col` - Column index (0-based)
/// * `value` - Boolean value
pub fn write_boolerr<W: Write>(
    writer: &mut W,
    row: u32,
    col: u16,
    xf_index: u16,
    value: bool,
) -> XlsResult<()> {
    // BIFF8 stores row as a 16-bit index (0..65535)
    let row_u16 = u16::try_from(row).map_err(|_| {
        XlsError::InvalidData(format!(
            "Row index {} exceeds BIFF8 limit 65535 for BOOLERR record",
            row
        ))
    })?;

    // 2 (row) + 2 (col) + 2 (xf) + 1 (value) + 1 (is-error flag) = 8 bytes
    write_record_header(writer, 0x0205, 8)?;

    writer.write_all(&row_u16.to_le_bytes())?;
    writer.write_all(&col.to_le_bytes())?;

    // XF record index
    writer.write_all(&xf_index.to_le_bytes())?;

    // Boolean value (0 = false, 1 = true) + error flag (0 = boolean)
    writer.write_all(&[if value { 1 } else { 0 }, 0])?;

    Ok(())
}
