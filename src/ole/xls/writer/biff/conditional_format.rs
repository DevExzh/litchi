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

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_write_cfheader_single_range() {
        let mut buffer = Vec::new();
        let ranges = vec![(0u32, 9u32, 0u16, 1u16)];
        let result = write_cfheader(&mut buffer, &ranges, 1);

        assert!(result.is_ok());
        // Check record header
        assert_eq!(buffer[0..2], [0xB0, 0x01]); // Record type CFHEADER
        // Check numcf
        assert_eq!(u16::from_le_bytes([buffer[4], buffer[5]]), 1);
    }

    #[test]
    fn test_write_cfheader_multiple_ranges() {
        let mut buffer = Vec::new();
        let ranges = vec![(0u32, 9u32, 0u16, 1u16), (10u32, 19u32, 0u16, 1u16)];
        let result = write_cfheader(&mut buffer, &ranges, 2);

        assert!(result.is_ok());
        // Record header is 4 bytes, then data starts
        // Data layout: 2 (numcf) + 2 (flags) + 8 (enclosing range) + 2 (range count) + ranges...
        // So range count is at offset 4 + 12 = 16
        let range_count = u16::from_le_bytes([buffer[16], buffer[17]]);
        assert_eq!(range_count, 2);
    }

    #[test]
    fn test_write_cfheader_empty_ranges() {
        let mut buffer = Vec::new();
        let ranges: Vec<(u32, u32, u16, u16)> = vec![];
        let result = write_cfheader(&mut buffer, &ranges, 1);

        assert!(result.is_err());
    }

    #[test]
    fn test_write_cfheader_zero_rules() {
        let mut buffer = Vec::new();
        let ranges = vec![(0u32, 9u32, 0u16, 1u16)];
        let result = write_cfheader(&mut buffer, &ranges, 0);

        assert!(result.is_err());
    }

    #[test]
    fn test_write_cfheader_invalid_range() {
        let mut buffer = Vec::new();
        // First row > last row
        let ranges = vec![(10u32, 0u32, 0u16, 1u16)];
        let result = write_cfheader(&mut buffer, &ranges, 1);

        assert!(result.is_err());
    }

    #[test]
    fn test_write_cfheader_row_limit_exceeded() {
        let mut buffer = Vec::new();
        // Row index exceeds BIFF8 limit (65535)
        let ranges = vec![(0u32, 70000u32, 0u16, 1u16)];
        let result = write_cfheader(&mut buffer, &ranges, 1);

        assert!(result.is_err());
    }

    #[test]
    fn test_write_cfrule_formula_only() {
        let mut buffer = Vec::new();
        let formula1 = vec![0x01, 0x02, 0x03]; // dummy formula bytes
        let formula2 = vec![];

        let result = write_cfrule(
            &mut buffer,
            0x02, // CONDITION_TYPE_FORMULA
            0x00, // NO_COMPARISON
            &formula1,
            &formula2,
            None, // no pattern
        );

        assert!(result.is_ok());
        // Check record type
        assert_eq!(buffer[0..2], [0xB1, 0x01]); // CFRULE
    }

    #[test]
    fn test_write_cfrule_with_pattern() {
        let mut buffer = Vec::new();
        let formula1 = vec![0x01, 0x02, 0x03];
        let formula2 = vec![];
        let pattern = Some((0x0001u16, 0x0040u16, 0x0041u16)); // pattern, fg, bg

        let result = write_cfrule(&mut buffer, 0x02, 0x00, &formula1, &formula2, pattern);

        assert!(result.is_ok());
        // Pattern should be included in data
        assert!(buffer.len() > 12); // Base size + pattern block
    }

    #[test]
    fn test_write_cfrule_formula_sizes() {
        let mut buffer = Vec::new();
        let formula1 = vec![0x01; 100];
        let formula2 = vec![0x02; 50];

        let result = write_cfrule(&mut buffer, 0x02, 0x00, &formula1, &formula2, None);

        assert!(result.is_ok());
        // Check formula lengths in header
        // Record header: 4 bytes (2 type + 2 length)
        // condition_type: 1 byte at offset 4
        // comparison_op: 1 byte at offset 5
        // f1_len: 2 bytes at offset 6
        // f2_len: 2 bytes at offset 8
        let f1_len = u16::from_le_bytes([buffer[6], buffer[7]]);
        let f2_len = u16::from_le_bytes([buffer[8], buffer[9]]);
        assert_eq!(f1_len, 100);
        assert_eq!(f2_len, 50);
    }

    #[test]
    fn test_cfheader_computes_enclosing_range() {
        let mut buffer = Vec::new();
        // Multiple disjoint ranges
        let ranges = vec![
            (5u32, 10u32, 2u16, 4u16),  // smaller range
            (0u32, 20u32, 0u16, 10u16), // larger range that encompasses
        ];
        let result = write_cfheader(&mut buffer, &ranges, 1);

        assert!(result.is_ok());
        // Enclosing range should be (0, 20, 0, 10)
        let enc_first_row = u16::from_le_bytes([buffer[8], buffer[9]]);
        let enc_last_row = u16::from_le_bytes([buffer[10], buffer[11]]);
        let enc_first_col = u16::from_le_bytes([buffer[12], buffer[13]]);
        let enc_last_col = u16::from_le_bytes([buffer[14], buffer[15]]);

        assert_eq!(enc_first_row, 0);
        assert_eq!(enc_last_row, 20);
        assert_eq!(enc_first_col, 0);
        assert_eq!(enc_last_col, 10);
    }
}
