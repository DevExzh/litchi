use crate::ole::xls::XlsResult;
use std::io::Write;

use super::write_record_header;

const XFCRC_RECORD_ID: u16 = 0x087C;
const XFEXT_RECORD_ID: u16 = 0x087D;
const FRT_RESERVED: u16 = 0;
const XFEXT_RESERVED: [u8; 10] = [0; 10];
const XFCRC_EXTENSION_COUNT: u16 = 0x0043;
const XFCRC_CHECKSUM: u32 = 0xB463_87D8;
const XFEXT_PROPERTY_COUNT: u16 = 0x0002;
const XFEXT_PROPERTY_KIND: u16 = 0x000D;
const XFEXT_PROPERTY_SIZE: u16 = 0x0014;
const XFEXT_PROPERTY_FLAGS: u32 = 0x0000_0003;
const XFEXT_FORMAT_KIND: u16 = 0x000E;
const XFEXT_FORMAT_SIZE: u16 = 0x0005;
const XFEXT_FORMAT_GRBIT: u8 = 0x02;
const PIVOT_XF_FORMAT_CODE: &[u8; 8] = b"00_ ;_ *";

fn write_frt_header<W: Write>(writer: &mut W, record_id: u16) -> XlsResult<()> {
    writer.write_all(&record_id.to_le_bytes())?;
    writer.write_all(&FRT_RESERVED.to_le_bytes())?;
    Ok(())
}

fn write_xfcrc<W: Write>(writer: &mut W) -> XlsResult<()> {
    write_record_header(writer, XFCRC_RECORD_ID, 20)?;
    write_frt_header(writer, XFCRC_RECORD_ID)?;
    writer.write_all(&XFEXT_RESERVED)?;
    writer.write_all(&XFCRC_EXTENSION_COUNT.to_le_bytes())?;
    writer.write_all(&XFCRC_CHECKSUM.to_le_bytes())?;
    Ok(())
}

fn write_xfext<W: Write>(writer: &mut W, xf_index: u16) -> XlsResult<()> {
    write_record_header(writer, XFEXT_RECORD_ID, 45)?;
    write_frt_header(writer, XFEXT_RECORD_ID)?;
    writer.write_all(&XFEXT_RESERVED)?;
    writer.write_all(&xf_index.to_le_bytes())?;
    writer.write_all(&FRT_RESERVED.to_le_bytes())?;
    writer.write_all(&XFEXT_PROPERTY_COUNT.to_le_bytes())?;
    writer.write_all(&XFEXT_PROPERTY_KIND.to_le_bytes())?;
    writer.write_all(&XFEXT_PROPERTY_SIZE.to_le_bytes())?;
    writer.write_all(&XFEXT_PROPERTY_FLAGS.to_le_bytes())?;
    writer.write_all(&1u32.to_le_bytes())?;
    writer.write_all(PIVOT_XF_FORMAT_CODE)?;
    writer.write_all(&XFEXT_FORMAT_KIND.to_le_bytes())?;
    writer.write_all(&XFEXT_FORMAT_SIZE.to_le_bytes())?;
    writer.write_all(&[XFEXT_FORMAT_GRBIT])?;
    Ok(())
}

pub fn write_pivot_xfext_block<W: Write>(writer: &mut W) -> XlsResult<()> {
    write_xfcrc(writer)?;

    for xf_index in [64u16, 65u16, 66u16] {
        write_xfext(writer, xf_index)?;
    }

    Ok(())
}
