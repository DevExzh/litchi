//! Cell record BIFF8 writers.

use crate::{XlsError, XlsResult};
use std::io::Write;

use super::{write_record, write_record_header};

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
pub(crate) fn write_number<W: Write>(
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

pub(crate) fn write_mulrk<W: Write>(
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
pub(crate) fn write_labelsst<W: Write>(
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
pub(crate) fn write_boolerr<W: Write>(
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

/// Write a BIFF8 FORMULA record with an empty cached result.
///
/// The `fAlwaysCalc` flag requests recalculation when Excel opens the file.
pub(crate) fn write_formula<W: Write>(
    writer: &mut W,
    row: u32,
    col: u16,
    xf_index: u16,
    tokens: &[u8],
) -> XlsResult<()> {
    let row_u16 = u16::try_from(row).map_err(|_| {
        XlsError::InvalidData(format!(
            "Row index {row} exceeds BIFF8 limit 65535 for FORMULA record"
        ))
    })?;
    if tokens.is_empty() {
        return Err(XlsError::InvalidFormula(
            "Formula token stream cannot be empty".to_string(),
        ));
    }
    // A BIFF record payload is limited to 8,224 bytes. FORMULA contributes
    // 22 fixed bytes before the token stream.
    if tokens.len() > 8_202 {
        return Err(XlsError::InvalidFormula(
            "Formula token stream exceeds BIFF8 record limit".to_string(),
        ));
    }
    let token_len = u16::try_from(tokens.len())
        .map_err(|_| XlsError::InvalidFormula("Formula token length exceeds u16".to_string()))?;
    let data_len = 22u16
        .checked_add(token_len)
        .ok_or_else(|| XlsError::InvalidFormula("Formula record length overflow".to_string()))?;

    write_record_header(writer, 0x0006, data_len)?;
    writer.write_all(&row_u16.to_le_bytes())?;
    writer.write_all(&col.to_le_bytes())?;
    writer.write_all(&xf_index.to_le_bytes())?;
    // FormulaValue special cached EMPTY: type, reserved, data, reserved[3], marker.
    writer.write_all(&[3, 0, 0, 0, 0, 0, 0xff, 0xff])?;
    writer.write_all(&0x0001u16.to_le_bytes())?; // fAlwaysCalc
    writer.write_all(&0u32.to_le_bytes())?; // chn
    writer.write_all(&token_len.to_le_bytes())?;
    writer.write_all(tokens)?;
    Ok(())
}

/// Write a TABLE record (MS-XLS 2.4.319) for a what-if data table.
///
/// Record type: 0x0236
pub(crate) fn write_table<W: Write>(writer: &mut W, table: &crate::XlsDataTable) -> XlsResult<()> {
    let payload = table.to_payload();
    write_record(writer, 0x0236, &payload)
}

#[cfg(test)]
mod tests {
    use super::write_formula;
    use crate::records::{CellRecord, FormulaValue, XlsEncoding};

    #[test]
    fn writes_formula_record_with_recalculation_and_empty_cache() {
        let tokens = [0x1e, 2, 0, 0x1e, 3, 0, 0x03];
        let mut bytes = Vec::new();
        write_formula(&mut bytes, 4, 5, 15, &tokens).unwrap();
        assert_eq!(&bytes[..4], &[0x06, 0x00, 29, 0]);
        assert_eq!(u16::from_le_bytes([bytes[18], bytes[19]]), 1);

        let record = CellRecord::parse(0x0006, &bytes[4..], &XlsEncoding::Utf16Le).unwrap();
        assert!(
            matches!(
                record,
                CellRecord::Formula {
                    row: 4,
                    col: 5,
                    xf_index: 15,
                    value: FormulaValue::Empty,
                    ref formula,
                } if formula == &tokens
            ),
            "{record:?}"
        );
    }
}
