//! Bounded MS-DOC `Hplxsdr` and `fcCustomXForm` codecs.

use super::model::{Collection, Reference};
use crate::package::{Error as PackageError, Result};
use crate::parts::fib::FileInformationBlock;

/// Table-pointer index of `fcHplxsdr`/`lcbHplxsdr`.
pub(super) const HPLXSDR: usize = 136;
/// Table-pointer index of `fcCustomXForm`/`lcbCustomXForm`.
pub(super) const CUSTOM_XFORM: usize = 140;

/// `fExtend` value of an extended STTB (MS-DOC 2.2.4).
pub(super) const STTB_F_EXTEND: u16 = 0xFFFF;
/// Fixed header of an extended STTB with a 4-byte `cData`: `fExtend`,
/// `cData`, and `cbExtra` (MS-DOC 2.2.4).
const STTB_HEADER_LEN: usize = 8;
/// `cXSDR` and the 4-byte STTB `cData` fields are signed integers whose
/// minimum value is zero, so the sign bit must be clear.
const MAX_SIGNED_COUNT: u32 = 0x7FFF_FFFF;
/// Minimum size of one `XSDR`: two empty length-prefixed strings and two
/// empty STTB headers (MS-DOC 2.9.352).
const MIN_XSDR_LEN: usize = 2 + 2 + STTB_HEADER_LEN + STTB_HEADER_LEN;
/// Maximum byte length of the `fcCustomXForm` path array (MS-DOC
/// `FibRgFcLcb2007`).
pub(super) const MAX_CUSTOM_XFORM_BYTES: u32 = 4168;

impl Collection {
    /// Parse the `Hplxsdr` addressed by the FIB, or `None` when the document
    /// carries none.
    pub fn parse(fib: &FileInformationBlock, table_stream: &[u8]) -> Result<Option<Collection>> {
        let Some(data) = optional_slice(fib, table_stream, HPLXSDR, "Hplxsdr")? else {
            return Ok(None);
        };
        if data.len() < 4 {
            return Err(corrupted("Hplxsdr is truncated"));
        }
        let count = read_u32(data, 0, "Hplxsdr cXSDR")?;
        if count > MAX_SIGNED_COUNT {
            return Err(corrupted("Hplxsdr cXSDR is negative"));
        }
        let count = usize::try_from(count).map_err(|_| corrupted("Hplxsdr count exceeds usize"))?;
        if count > (data.len() - 4) / MIN_XSDR_LEN {
            return Err(corrupted("Hplxsdr byte length does not match its count"));
        }

        let mut schemas = Vec::with_capacity(count);
        let mut offset = 4usize;
        for _ in 0..count {
            let (schema, size) = parse_xsdr(&data[offset..])?;
            schemas.push(schema);
            offset += size;
        }
        if offset != data.len() {
            return Err(corrupted("Hplxsdr contains trailing bytes"));
        }
        Ok(Some(Self::from_schemas(schemas)))
    }
}

/// Parse the custom XML save transform path (`fcCustomXForm`): the full path
/// and file name of the XML stylesheet Word applies when saving the document
/// in XML format, or `None` when the document carries none.
///
/// The path is inert: it is exposed verbatim and never opened, resolved, or
/// applied.
pub fn parse_custom_xml_transform(
    fib: &FileInformationBlock,
    table_stream: &[u8],
) -> Result<Option<String>> {
    let Some((offset, length)) = fib.get_table_pointer(CUSTOM_XFORM) else {
        return Ok(None);
    };
    if length == 0 {
        return Ok(None);
    }
    if length > MAX_CUSTOM_XFORM_BYTES || length % 2 != 0 {
        return Err(corrupted(
            "fcCustomXForm length exceeds 4168 bytes or is not even",
        ));
    }
    let start =
        usize::try_from(offset).map_err(|_| corrupted("fcCustomXForm offset exceeds usize"))?;
    let end = start
        .checked_add(
            usize::try_from(length).map_err(|_| corrupted("fcCustomXForm length exceeds usize"))?,
        )
        .ok_or_else(|| corrupted("fcCustomXForm range overflows"))?;
    let data = table_stream
        .get(start..end)
        .ok_or_else(|| corrupted("fcCustomXForm extends beyond the table stream"))?;
    let mut units = data
        .chunks_exact(2)
        .map(|chunk| u16::from_le_bytes([chunk[0], chunk[1]]))
        .collect::<Vec<_>>();
    // Producers commonly terminate the path array with a null code unit;
    // the spec defines the array by its byte length alone.
    if units.last() == Some(&0) {
        units.pop();
    }
    String::from_utf16(&units)
        .map(Some)
        .map_err(|_| corrupted("fcCustomXForm is invalid UTF-16"))
}

/// Parse one `XSDR`, returning the schema reference and the consumed byte
/// count.
fn parse_xsdr(data: &[u8]) -> Result<(Reference, usize)> {
    let mut offset = 0usize;
    let uri = parse_length_prefixed_string(data, &mut offset, "XSDR wzURI")?;
    let manifest_location =
        parse_length_prefixed_string(data, &mut offset, "XSDR wzManifestLocation")?;
    let elements = parse_string_table(data, &mut offset, "XSDR sttbElements")?;
    let attributes = parse_string_table(data, &mut offset, "XSDR sttbAttributes")?;
    Ok((
        Reference {
            uri,
            manifest_location,
            elements,
            attributes,
        },
        offset,
    ))
}

/// Parse a 16-bit length-prefixed UTF-16 string that is not null-terminated
/// (MS-DOC 2.9.352), advancing `offset` past it.
fn parse_length_prefixed_string(data: &[u8], offset: &mut usize, field: &str) -> Result<String> {
    let chars = usize::from(read_u16(data, *offset, field)?);
    let start = *offset + 2;
    let end = start
        .checked_add(
            chars
                .checked_mul(2)
                .ok_or_else(|| corrupted(format!("{field} byte length overflows")))?,
        )
        .ok_or_else(|| corrupted(format!("{field} range overflows")))?;
    let bytes = data
        .get(start..end)
        .ok_or_else(|| corrupted(format!("{field} is truncated")))?;
    *offset = end;
    decode_utf16(bytes, field)
}

/// Parse an extended STTB with a 4-byte `cData` (MS-DOC 2.2.4), advancing
/// `offset` past it. Per-entry extra data (`cbExtra`) is skipped verbatim.
fn parse_string_table(data: &[u8], offset: &mut usize, name: &str) -> Result<Vec<String>> {
    if data.len() < *offset + STTB_HEADER_LEN {
        return Err(corrupted(format!("{name} is truncated")));
    }
    if read_u16(data, *offset, "STTB fExtend")? != STTB_F_EXTEND {
        return Err(corrupted(format!("{name} is not an extended STTB")));
    }
    let count = read_u32(data, *offset + 2, "STTB cData")?;
    if count > MAX_SIGNED_COUNT {
        return Err(corrupted(format!("{name} cData is negative")));
    }
    let count =
        usize::try_from(count).map_err(|_| corrupted(format!("{name} count exceeds usize")))?;
    let extra = usize::from(read_u16(data, *offset + 6, "STTB cbExtra")?);
    let minimum_entry = 2usize
        .checked_add(extra)
        .ok_or_else(|| corrupted(format!("{name} entry size overflows")))?;
    if count > (data.len() - *offset - STTB_HEADER_LEN) / minimum_entry {
        return Err(corrupted(format!(
            "{name} byte length does not match its count"
        )));
    }

    let mut cursor = *offset + STTB_HEADER_LEN;
    let mut strings = Vec::with_capacity(count);
    for _ in 0..count {
        let chars = usize::from(read_u16(data, cursor, "STTB cchData")?);
        let start = cursor + 2;
        let end = start
            .checked_add(
                chars
                    .checked_mul(2)
                    .ok_or_else(|| corrupted(format!("{name} string byte length overflows")))?,
            )
            .ok_or_else(|| corrupted(format!("{name} string range overflows")))?;
        let bytes = data
            .get(start..end)
            .ok_or_else(|| corrupted(format!("{name} string is truncated")))?;
        strings.push(decode_utf16(bytes, name)?);
        cursor = end
            .checked_add(extra)
            .ok_or_else(|| corrupted(format!("{name} extra data range overflows")))?;
        if cursor > data.len() {
            return Err(corrupted(format!("{name} extra data is truncated")));
        }
    }
    *offset = cursor;
    Ok(strings)
}

fn decode_utf16(bytes: &[u8], field: &str) -> Result<String> {
    let units = bytes
        .chunks_exact(2)
        .map(|chunk| u16::from_le_bytes([chunk[0], chunk[1]]))
        .collect::<Vec<_>>();
    String::from_utf16(&units).map_err(|_| corrupted(format!("{field} is invalid UTF-16")))
}

fn optional_slice<'a>(
    fib: &FileInformationBlock,
    table_stream: &'a [u8],
    index: usize,
    name: &str,
) -> Result<Option<&'a [u8]>> {
    let Some((offset, length)) = fib.get_table_pointer(index) else {
        return Ok(None);
    };
    if length == 0 {
        return Ok(None);
    }
    let start =
        usize::try_from(offset).map_err(|_| corrupted(format!("{name} offset exceeds usize")))?;
    let length =
        usize::try_from(length).map_err(|_| corrupted(format!("{name} length exceeds usize")))?;
    let end = start
        .checked_add(length)
        .ok_or_else(|| corrupted(format!("{name} range overflows")))?;
    table_stream
        .get(start..end)
        .map(Some)
        .ok_or_else(|| corrupted(format!("{name} extends beyond the table stream")))
}

fn read_u16(data: &[u8], offset: usize, field: &str) -> Result<u16> {
    litchi_core::binary::read_u16_le(data, offset)
        .map_err(|error| corrupted(format!("invalid {field}: {error}")))
}

fn read_u32(data: &[u8], offset: usize, field: &str) -> Result<u32> {
    litchi_core::binary::read_u32_le(data, offset)
        .map_err(|error| corrupted(format!("invalid {field}: {error}")))
}

fn corrupted(message: impl Into<String>) -> PackageError {
    PackageError::Corrupted(message.into())
}
