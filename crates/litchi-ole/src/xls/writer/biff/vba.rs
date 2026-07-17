use crate::xls::XlsResult;
use crate::xls::vba::validate_code_name;
use crate::xls::writer::biff::write_record_header;
use std::io::Write;

pub(crate) fn write_ob_proj<W: Write>(writer: &mut W) -> XlsResult<()> {
    write_record_header(writer, 0x00D3, 0)
}
pub(crate) fn write_ob_no_macros<W: Write>(writer: &mut W) -> XlsResult<()> {
    write_record_header(writer, 0x01BF, 0)
}
pub(crate) fn write_code_name<W: Write>(writer: &mut W, value: &str) -> XlsResult<()> {
    validate_code_name(value)?;
    let count = value.encode_utf16().count() as u16;
    let wide = !value.chars().all(|character| u32::from(character) <= 0xFF);
    let data_len = 3 + count * if wide { 2 } else { 1 };
    write_record_header(writer, 0x01BA, data_len)?;
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
