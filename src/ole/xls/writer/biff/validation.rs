use crate::ole::xls::{XlsError, XlsResult};
use std::io::Write;

use super::{unicode_string_size, write_record_header, write_unicode_string_biff8};

/// Configuration for a single DV (data validation) record.
#[derive(Debug, Clone)]
pub(crate) struct DvConfig<'a> {
    pub data_type: u8,
    pub operator: u8,
    pub error_style: u8,
    pub empty_cell_allowed: bool,
    pub suppress_dropdown_arrow: bool,
    pub is_explicit_list_formula: bool,
    pub show_prompt_on_cell_selected: bool,
    pub prompt_title: Option<&'a str>,
    pub prompt_text: Option<&'a str>,
    pub show_error_on_invalid_value: bool,
    pub error_title: Option<&'a str>,
    pub error_text: Option<&'a str>,
    pub formula1: Option<&'a [u8]>,
    pub formula2: Option<&'a [u8]>,
}

pub fn write_dval<W: Write>(writer: &mut W, dv_count: u32) -> XlsResult<()> {
    // DVAL: options(2) + horiz_pos(4) + vert_pos(4) + cbo_id(4) + dv_no(4) = 18 bytes
    write_record_header(writer, 0x01B2, 18)?;

    // Use conservative defaults matching POI's DVALRecord:
    //  - options = 0
    //  - horiz/vert positions = 0
    //  - cbo_id = 0xFFFFFFFF (no associated drop-down object)
    //  - dv_no = number of following DV records
    writer.write_all(&0u16.to_le_bytes())?; // options
    writer.write_all(&0u32.to_le_bytes())?; // horiz_pos
    writer.write_all(&0u32.to_le_bytes())?; // vert_pos
    writer.write_all(&0xFFFFFFFFu32.to_le_bytes())?; // cbo_id
    writer.write_all(&dv_count.to_le_bytes())?; // dv_no

    Ok(())
}

pub fn write_dv<W: Write>(
    writer: &mut W,
    cfg: &DvConfig<'_>,
    ranges: &[(u32, u32, u16, u16)],
) -> XlsResult<()> {
    if ranges.is_empty() {
        return Err(XlsError::InvalidData(
            "DV record must have at least one target range".to_string(),
        ));
    }

    let dv_count_u16 = u16::try_from(ranges.len()).map_err(|_| {
        XlsError::InvalidData("Too many ranges in data validation rule".to_string())
    })?;

    // Normalise titles/text: Excel encodes "not present" as a single NUL character.
    let prompt_title_val: &str = match cfg.prompt_title {
        Some(s) if !s.is_empty() => s,
        _ => "\u{0000}",
    };
    let error_title_val: &str = match cfg.error_title {
        Some(s) if !s.is_empty() => s,
        _ => "\u{0000}",
    };
    let prompt_text_val: &str = match cfg.prompt_text {
        Some(s) if !s.is_empty() => s,
        _ => "\u{0000}",
    };
    let error_text_val: &str = match cfg.error_text {
        Some(s) if !s.is_empty() => s,
        _ => "\u{0000}",
    };

    let f1_len: u16 = cfg
        .formula1
        .map(|f| {
            u16::try_from(f.len()).map_err(|_| {
                XlsError::InvalidData(
                    "Data validation formula1 exceeds BIFF8 size limit".to_string(),
                )
            })
        })
        .transpose()?
        .unwrap_or(0);

    let f2_len: u16 = cfg
        .formula2
        .map(|f| {
            u16::try_from(f.len()).map_err(|_| {
                XlsError::InvalidData(
                    "Data validation formula2 exceeds BIFF8 size limit".to_string(),
                )
            })
        })
        .transpose()?
        .unwrap_or(0);

    // Compute data size to populate the record header.
    let mut data_len: u16 = 4; // option_flags (u32)
    data_len = data_len.saturating_add(unicode_string_size(prompt_title_val));
    data_len = data_len.saturating_add(unicode_string_size(error_title_val));
    data_len = data_len.saturating_add(unicode_string_size(prompt_text_val));
    data_len = data_len.saturating_add(unicode_string_size(error_text_val));
    data_len = data_len.saturating_add(2 + 2 + f1_len); // field_size_first_formula + not_used_1 + formula1
    data_len = data_len.saturating_add(2 + 2 + f2_len); // field_size_sec_formula + not_used_2 + formula2
    data_len = data_len.saturating_add(2 + dv_count_u16.saturating_mul(8)); // CellRangeAddressList

    if data_len > 8224 {
        return Err(XlsError::InvalidData(
            "DV record exceeds maximum BIFF8 record size".to_string(),
        ));
    }

    write_record_header(writer, 0x01BE, data_len)?;

    // Option flags bitfield
    let mut option_flags: u32 = 0;
    option_flags |= (u32::from(cfg.data_type)) & 0x0000_000F;
    option_flags |= ((u32::from(cfg.error_style)) & 0x0000_0007) << 4;
    if cfg.is_explicit_list_formula {
        option_flags |= 0x0000_0080;
    }
    if cfg.empty_cell_allowed {
        option_flags |= 0x0000_0100;
    }
    if cfg.suppress_dropdown_arrow {
        option_flags |= 0x0000_0200;
    }
    if cfg.show_prompt_on_cell_selected {
        option_flags |= 0x0004_0000;
    }
    if cfg.show_error_on_invalid_value {
        option_flags |= 0x0008_0000;
    }
    option_flags |= ((u32::from(cfg.operator)) & 0x0000_0007) << 20;

    writer.write_all(&option_flags.to_le_bytes())?;

    // Prompt / error strings
    write_unicode_string_biff8(writer, prompt_title_val)?;
    write_unicode_string_biff8(writer, error_title_val)?;
    write_unicode_string_biff8(writer, prompt_text_val)?;
    write_unicode_string_biff8(writer, error_text_val)?;

    // First formula
    writer.write_all(&f1_len.to_le_bytes())?;
    writer.write_all(&0x3FE0u16.to_le_bytes())?; // not_used_1
    if let Some(bytes) = cfg.formula1 {
        writer.write_all(bytes)?;
    }

    // Second formula
    writer.write_all(&f2_len.to_le_bytes())?;
    writer.write_all(&0u16.to_le_bytes())?; // not_used_2
    if let Some(bytes) = cfg.formula2 {
        writer.write_all(bytes)?;
    }

    // CellRangeAddressList with all affected ranges
    writer.write_all(&dv_count_u16.to_le_bytes())?;
    for (first_row_u32, last_row_u32, first_col, last_col) in ranges {
        let first_row = u16::try_from(*first_row_u32).map_err(|_| {
            XlsError::InvalidData(format!(
                "Row index {} exceeds BIFF8 limit 65535 for DV record",
                first_row_u32
            ))
        })?;
        let last_row = u16::try_from(*last_row_u32).map_err(|_| {
            XlsError::InvalidData(format!(
                "Row index {} exceeds BIFF8 limit 65535 for DV record",
                last_row_u32
            ))
        })?;

        writer.write_all(&first_row.to_le_bytes())?;
        writer.write_all(&last_row.to_le_bytes())?;
        writer.write_all(&first_col.to_le_bytes())?;
        writer.write_all(&last_col.to_le_bytes())?;
    }

    Ok(())
}
