use crate::ole::xls::{XlsError, XlsResult};
use std::io::Write;

use super::write_record_header;

pub fn write_cfheader<W: Write>(
    writer: &mut W,
    ranges: &[(u32, u32, u16, u16)],
    num_rules: u16,
) -> XlsResult<()> {
    if ranges.is_empty() {
        return Err(XlsError::InvalidData(
            "CFHEADER must have at least one target range".to_string(),
        ));
    }

    let numcf = num_rules;
    if numcf == 0 {
        return Err(XlsError::InvalidData(
            "CFHEADER must describe at least one rule".to_string(),
        ));
    }

    let range_count_u16 = u16::try_from(ranges.len())
        .map_err(|_| XlsError::InvalidData("Too many conditional formatting ranges".to_string()))?;

    // Compute enclosing cell range over all regions
    let mut enc_first_row = u32::MAX;
    let mut enc_last_row = 0u32;
    let mut enc_first_col = u16::MAX;
    let mut enc_last_col = 0u16;

    for (first_row, last_row, first_col, last_col) in ranges {
        if *first_row > *last_row || *first_col > *last_col {
            return Err(XlsError::InvalidData(
                "Conditional format range must have first <= last for rows and columns".to_string(),
            ));
        }
        enc_first_row = enc_first_row.min(*first_row);
        enc_last_row = enc_last_row.max(*last_row);
        enc_first_col = enc_first_col.min(*first_col);
        enc_last_col = enc_last_col.max(*last_col);
    }

    let enc_first_row_u16 = u16::try_from(enc_first_row).map_err(|_| {
        XlsError::InvalidData(format!(
            "Row index {} exceeds BIFF8 limit 65535 for CFHEADER record",
            enc_first_row
        ))
    })?;
    let enc_last_row_u16 = u16::try_from(enc_last_row).map_err(|_| {
        XlsError::InvalidData(format!(
            "Row index {} exceeds BIFF8 limit 65535 for CFHEADER record",
            enc_last_row
        ))
    })?;

    // Columns are stored as u16 in BIFF, but Excel only uses the low 8 bits.
    // We still enforce the u16 limit here for consistency with other writers.
    let enc_first_col_u16 = enc_first_col;
    let enc_last_col_u16 = enc_last_col;

    // Data size: 2 (numcf) + 2 (need_recalculation_and_id) + 8 (enclosing range)
    //          + 2 (count) + 8 * count (CellRangeAddress list)
    let data_len: u16 = 14u16.saturating_add(range_count_u16.saturating_mul(8));

    if data_len > 8224 {
        return Err(XlsError::InvalidData(
            "CFHEADER record exceeds maximum BIFF8 record size".to_string(),
        ));
    }

    write_record_header(writer, 0x01B0, data_len)?;

    // Number of conditional formats
    writer.write_all(&numcf.to_le_bytes())?;

    // needRecalculation (bit0) + ID (bits 1..15). We set both to zero.
    writer.write_all(&0u16.to_le_bytes())?;

    // Enclosing cell range
    writer.write_all(&enc_first_row_u16.to_le_bytes())?;
    writer.write_all(&enc_last_row_u16.to_le_bytes())?;
    writer.write_all(&enc_first_col_u16.to_le_bytes())?;
    writer.write_all(&enc_last_col_u16.to_le_bytes())?;

    // CellRangeAddressList: count + individual ranges
    writer.write_all(&range_count_u16.to_le_bytes())?;

    for (first_row_u32, last_row_u32, first_col, last_col) in ranges {
        let first_row_u16 = u16::try_from(*first_row_u32).map_err(|_| {
            XlsError::InvalidData(format!(
                "Row index {} exceeds BIFF8 limit 65535 for CFHEADER range",
                first_row_u32
            ))
        })?;
        let last_row_u16 = u16::try_from(*last_row_u32).map_err(|_| {
            XlsError::InvalidData(format!(
                "Row index {} exceeds BIFF8 limit 65535 for CFHEADER range",
                last_row_u32
            ))
        })?;

        writer.write_all(&first_row_u16.to_le_bytes())?;
        writer.write_all(&last_row_u16.to_le_bytes())?;
        writer.write_all(&first_col.to_le_bytes())?;
        writer.write_all(&last_col.to_le_bytes())?;
    }

    Ok(())
}

pub fn write_cfrule<W: Write>(
    writer: &mut W,
    condition_type: u8,
    comparison_op: u8,
    formula1: &[u8],
    formula2: &[u8],
    pattern: Option<(u16, u16, u16)>,
) -> XlsResult<()> {
    let f1_len = u16::try_from(formula1.len()).map_err(|_| {
        XlsError::InvalidData("Conditional format formula1 exceeds BIFF8 size limit".to_string())
    })?;
    let f2_len = u16::try_from(formula2.len()).map_err(|_| {
        XlsError::InvalidData("Conditional format formula2 exceeds BIFF8 size limit".to_string())
    })?;

    // Base size: 1 (condition_type) + 1 (comparison_op)
    //          + 2 (formula1_len) + 2 (formula2_len)
    //          + 4 (formatting_options) + 2 (formatting_not_used)
    //          + optional 4-byte PatternFormatting block
    //          + formula1 bytes + formula2 bytes
    let pattern_block_len: u16 = if pattern.is_some() { 4 } else { 0 };
    let data_len = 12u16
        .saturating_add(pattern_block_len)
        .saturating_add(f1_len)
        .saturating_add(f2_len);

    if data_len > 8224 {
        return Err(XlsError::InvalidData(
            "CFRULE record exceeds maximum BIFF8 record size".to_string(),
        ));
    }

    write_record_header(writer, 0x01B1, data_len)?;

    writer.write_all(&[condition_type])?;
    writer.write_all(&[comparison_op])?;
    writer.write_all(&f1_len.to_le_bytes())?;
    writer.write_all(&f2_len.to_le_bytes())?;

    // formatting_options: start with all modification bits set (no attrs changed).
    let mut formatting_options: u32 = 0x003F_FFFF;
    if pattern.is_some() {
        // Indicate presence of a PatternFormatting block (CFRuleBase.patt).
        formatting_options |= 0x2000_0000;
    }
    writer.write_all(&formatting_options.to_le_bytes())?;

    // formatting_not_used: Excel writes 0x8002 and does not appear to care.
    writer.write_all(&0x8002u16.to_le_bytes())?;

    if let Some((pattern_code, fg, bg)) = pattern {
        let pattern_style: u16 = (pattern_code & 0x003F) << 10;
        let fg_idx = fg & 0x007F;
        let bg_idx = bg & 0x007F;
        let color_indexes: u16 = (fg_idx & 0x007F) | ((bg_idx & 0x007F) << 7);

        writer.write_all(&pattern_style.to_le_bytes())?;
        writer.write_all(&color_indexes.to_le_bytes())?;
    }

    writer.write_all(formula1)?;
    writer.write_all(formula2)?;

    Ok(())
}
