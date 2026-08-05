use crate::Result;
use crate::vba::{
    CODE_NAME_RECORD_TYPE, OB_NO_MACROS_RECORD_TYPE, OB_PROJ_RECORD_TYPE, validate_code_name,
};
use crate::writer::biff::write_record_header;
use std::io::Write;

/// Payload length of the two zero-length VBA marker records.
const MARKER_RECORD_LEN: u16 = 0;
/// `cch` plus the option byte that precede a `CodeName` string's characters.
const CODE_NAME_HEADER_LEN: u16 = 3;

pub(crate) fn write_ob_proj<W: Write>(writer: &mut W) -> Result<()> {
    write_record_header(writer, OB_PROJ_RECORD_TYPE, MARKER_RECORD_LEN)
}
pub(crate) fn write_ob_no_macros<W: Write>(writer: &mut W) -> Result<()> {
    write_record_header(writer, OB_NO_MACROS_RECORD_TYPE, MARKER_RECORD_LEN)
}
pub(crate) fn write_code_name<W: Write>(writer: &mut W, value: &str) -> Result<()> {
    validate_code_name(value)?;
    let count = value.encode_utf16().count() as u16;
    let wide = !value.chars().all(|character| u32::from(character) <= 0xFF);
    let data_len = CODE_NAME_HEADER_LEN + count * if wide { 2 } else { 1 };
    write_record_header(writer, CODE_NAME_RECORD_TYPE, data_len)?;
    writer.write_all(&count.to_le_bytes())?;
    writer.write_all(&[u8::from(wide)])?;
    if wide {
        for unit in value.encode_utf16() {
            writer.write_all(&unit.to_le_bytes())?;
        }
    } else {
        for character in value.chars() {
            writer.write_all(&[character as u8])?;
        }
    }
    Ok(())
}
