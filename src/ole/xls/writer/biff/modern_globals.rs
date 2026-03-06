use crate::ole::xls::XlsResult;
use std::io::Write;

use super::write_record_header;

const SXDBEX_CREATION_TIMESTAMP: [u8; 8] = [0xFC, 0xE5, 0x58, 0x44, 0xBC, 0x7D, 0xE6, 0x40];

const TABLE_STYLES_RECORD_ID: u16 = 0x088E;
const TABLE_STYLES_RESERVED: [u8; 8] = [0; 8];
const TABLE_STYLES_COUNT_HINT: u32 = 0x0000_0090;
const DEFAULT_TABLE_STYLE_NAME: &str = "TableStyleMedium2";
const DEFAULT_PIVOT_STYLE_NAME: &str = "PivotStyleLight16";

const SXADDL_RECORD_ID: u16 = 0x0864;
const SXADDL_FRT_RESERVED: u16 = 0;
const SXADDL_VERSION_SENTINEL: u32 = 0xFFFF_FFFF;
const SXADDL_VERSION_BUILD: u16 = 0x0304;

const COMPRESS_PICTURES_RECORD_ID: u16 = 0x089A;
const COMPRESS_PICTURES_RESERVED: [u8; 8] = [0; 8];
const COMPRESS_PICTURES_ENABLED: u32 = 1;
const COMPRESS_PICTURES_TARGET: u32 = 0;
const COMPRESS_PICTURES_DEFAULT_DPI: u32 = 8;

const COMPAT12_RECORD_ID: u16 = 0x08A3;
const COMPAT12_RESERVED: [u8; 12] = [0; 12];

#[derive(Clone, Copy)]
#[repr(u8)]
enum PivotCacheSxAddlClass {
    Cache = 0x03,
}

#[derive(Clone, Copy)]
#[repr(u8)]
enum PivotCacheSxAddlType {
    CacheId = 0x00,
    Ver10Info = 0x02,
    FeatureFlags = 0x18,
    FlagA = 0x01,
    FlagB = 0x41,
    FlagC = 0x34,
    EndInfo = 0xFF,
}

struct SxAddlHeader {
    class: PivotCacheSxAddlClass,
    record_type: PivotCacheSxAddlType,
    id: u32,
}

fn write_frt_header<W: Write>(writer: &mut W, record_id: u16) -> XlsResult<()> {
    writer.write_all(&record_id.to_le_bytes())?;
    writer.write_all(&0u16.to_le_bytes())?;
    Ok(())
}

fn write_wide_string<W: Write>(writer: &mut W, value: &str) -> XlsResult<()> {
    for ch in value.encode_utf16() {
        writer.write_all(&ch.to_le_bytes())?;
    }
    Ok(())
}

fn write_sxaddl_header<W: Write>(writer: &mut W, header: SxAddlHeader) -> XlsResult<()> {
    write_frt_header(writer, SXADDL_RECORD_ID)?;
    writer.write_all(&[header.class as u8])?;
    writer.write_all(&[header.record_type as u8])?;
    writer.write_all(&header.id.to_le_bytes())?;
    writer.write_all(&SXADDL_FRT_RESERVED.to_le_bytes())?;
    Ok(())
}

pub fn write_table_styles<W: Write>(writer: &mut W) -> XlsResult<()> {
    let table_style_name_len = DEFAULT_TABLE_STYLE_NAME.encode_utf16().count() as u16;
    let pivot_style_name_len = DEFAULT_PIVOT_STYLE_NAME.encode_utf16().count() as u16;
    let data_len =
        20u16 + table_style_name_len.saturating_mul(2) + pivot_style_name_len.saturating_mul(2);

    write_record_header(writer, TABLE_STYLES_RECORD_ID, data_len)?;
    write_frt_header(writer, TABLE_STYLES_RECORD_ID)?;
    writer.write_all(&TABLE_STYLES_RESERVED)?;
    writer.write_all(&TABLE_STYLES_COUNT_HINT.to_le_bytes())?;
    writer.write_all(&table_style_name_len.to_le_bytes())?;
    writer.write_all(&pivot_style_name_len.to_le_bytes())?;
    write_wide_string(writer, DEFAULT_TABLE_STYLE_NAME)?;
    write_wide_string(writer, DEFAULT_PIVOT_STYLE_NAME)?;
    Ok(())
}

pub fn write_pivot_cache_sxaddl_block<W: Write>(writer: &mut W) -> XlsResult<()> {
    write_record_header(writer, SXADDL_RECORD_ID, 12)?;
    write_sxaddl_header(
        writer,
        SxAddlHeader {
            class: PivotCacheSxAddlClass::Cache,
            record_type: PivotCacheSxAddlType::CacheId,
            id: 1,
        },
    )?;

    write_record_header(writer, SXADDL_RECORD_ID, 28)?;
    write_sxaddl_header(
        writer,
        SxAddlHeader {
            class: PivotCacheSxAddlClass::Cache,
            record_type: PivotCacheSxAddlType::Ver10Info,
            id: 0,
        },
    )?;
    writer.write_all(&SXADDL_VERSION_SENTINEL.to_le_bytes())?;
    writer.write_all(&SXADDL_VERSION_BUILD.to_le_bytes())?;
    writer.write_all(&SXDBEX_CREATION_TIMESTAMP)?;
    writer.write_all(&SXADDL_FRT_RESERVED.to_le_bytes())?;

    for (record_type, id) in [
        (PivotCacheSxAddlType::FeatureFlags, 4u32),
        (PivotCacheSxAddlType::FlagA, 2u32),
        (PivotCacheSxAddlType::FlagB, 0u32),
        (PivotCacheSxAddlType::FlagC, 1u32),
        (PivotCacheSxAddlType::FlagA, 0x00FFu32),
        (PivotCacheSxAddlType::EndInfo, 0u32),
    ] {
        write_record_header(writer, SXADDL_RECORD_ID, 12)?;
        write_sxaddl_header(
            writer,
            SxAddlHeader {
                class: PivotCacheSxAddlClass::Cache,
                record_type,
                id,
            },
        )?;
    }
    Ok(())
}

pub fn write_compress_pictures<W: Write>(writer: &mut W) -> XlsResult<()> {
    write_record_header(writer, COMPRESS_PICTURES_RECORD_ID, 24)?;
    write_frt_header(writer, COMPRESS_PICTURES_RECORD_ID)?;
    writer.write_all(&COMPRESS_PICTURES_RESERVED)?;
    writer.write_all(&COMPRESS_PICTURES_ENABLED.to_le_bytes())?;
    writer.write_all(&COMPRESS_PICTURES_TARGET.to_le_bytes())?;
    writer.write_all(&COMPRESS_PICTURES_DEFAULT_DPI.to_le_bytes())?;
    Ok(())
}

pub fn write_compat12<W: Write>(writer: &mut W) -> XlsResult<()> {
    write_record_header(writer, COMPAT12_RECORD_ID, 16)?;
    write_frt_header(writer, COMPAT12_RECORD_ID)?;
    writer.write_all(&COMPAT12_RESERVED)?;
    Ok(())
}

pub fn sxdbex_creation_timestamp_bytes() -> [u8; 8] {
    SXDBEX_CREATION_TIMESTAMP
}
