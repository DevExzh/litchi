//! Binary codecs for the contextual [MS-OSHARED] property composites.

use super::super::super::model::*;
use super::wire::{
    ValueReader, append_bytes, append_u16, append_u32, decode_ansi, decode_utf16, encode_ansi,
    reserve_bytes,
};
use litchi_cfb::OleError;
use litchi_cfb::consts::{VT_I4, VT_LPSTR, VT_LPWSTR, VT_VARIANT, VT_VECTOR};

pub(super) const HEADING_PAIRS_TYPE: u16 = VT_VECTOR | VT_VARIANT;
pub(super) const DOC_PARTS_ANSI_TYPE: u16 = VT_VECTOR | VT_LPSTR;
pub(super) const DOC_PARTS_UNICODE_TYPE: u16 = VT_VECTOR | VT_LPWSTR;

pub(super) fn append_heading_pairs(
    out: &mut Vec<u8>,
    value: &HeadingPairs,
    codepage: u16,
) -> Result<(), OleError> {
    value.validate()?;
    let element_count = checked_u32(
        checked_mul(value.len(), 2, "Heading pair element count")?,
        "Heading pair element count",
    )?;
    append_u32(out, element_count, "serialized heading pairs")?;
    for pair in value.pairs() {
        append_heading_string(out, pair.heading(), codepage)?;
        append_u16(out, VT_I4, "serialized heading pair")?;
        append_u16(out, 0, "serialized heading pair")?;
        append_bytes(
            out,
            &(pair.part_count() as i32).to_le_bytes(),
            "serialized heading pair",
        )?;
    }
    Ok(())
}

pub(super) fn append_doc_parts(
    out: &mut Vec<u8>,
    value: &DocParts,
    codepage: u16,
) -> Result<(), OleError> {
    value.validate_for_codepage(codepage)?;
    append_u32(
        out,
        checked_u32(value.len(), "Document-part count")?,
        "serialized document parts",
    )?;
    for part in value.values() {
        match value.encoding() {
            TextEncoding::Ansi => append_unaligned_ansi(out, part, codepage)?,
            TextEncoding::Unicode => append_lpwstr(out, part)?,
        }
    }
    Ok(())
}

pub(super) fn parse_heading_pairs(
    reader: &mut ValueReader<'_>,
    codepage: u16,
) -> Result<Value, OleError> {
    let element_count = usize::try_from(reader.read_u32("heading pair element count")?)
        .map_err(|_| invalid("Heading pair element count is too large"))?;
    if element_count % 2 != 0 {
        return Err(invalid("Heading pair element count must be even"));
    }
    let pair_count = element_count / 2;
    if pair_count > MAX_COMPOSITE_ELEMENTS {
        return Err(invalid("Heading pair count exceeds the safety limit"));
    }
    if pair_count > reader.remaining_len() / 16 {
        return Err(invalid("Heading pair count exceeds its property range"));
    }
    let mut pairs = try_vec_with_capacity(pair_count, "heading pairs")?;
    for _ in 0..pair_count {
        let heading = read_heading_string(reader, codepage)?;
        let variant_type = reader.read_u16("heading pair part-count type")?;
        if variant_type != VT_I4 {
            return Err(invalid("Heading pair part count must be VT_I4"));
        }
        if reader.read_u16("heading pair part-count reserved field")? != 0 {
            return Err(invalid(
                "Heading pair part-count reserved field must be zero",
            ));
        }
        let part_count = reader.read_i32("heading pair part count")?;
        if part_count < 0 {
            return Err(invalid("Heading pair part count must be nonnegative"));
        }
        pairs.push(HeadingPair::new(heading, part_count as u32)?);
    }
    Ok(Value::HeadingPairs(HeadingPairs::new(pairs)?))
}

pub(super) fn parse_doc_parts(
    reader: &mut ValueReader<'_>,
    encoding: TextEncoding,
    codepage: u16,
) -> Result<Value, OleError> {
    validate_encoding(encoding, codepage)?;
    let count = usize::try_from(reader.read_u32("document-part count")?)
        .map_err(|_| invalid("Document-part count is too large"))?;
    if count > MAX_COMPOSITE_ELEMENTS {
        return Err(invalid("Document-part count exceeds the safety limit"));
    }
    if count > reader.remaining_len() / 4 {
        return Err(invalid("Document-part count exceeds its property range"));
    }
    let mut values = try_vec_with_capacity(count, "document parts")?;
    for _ in 0..count {
        let value = match encoding {
            TextEncoding::Ansi => read_unaligned_ansi(reader, codepage, "document-part name")?,
            TextEncoding::Unicode => read_lpwstr(reader, "document-part name")?,
        };
        values.push(value);
    }
    Ok(Value::DocParts(DocParts::new(encoding, values)?))
}

fn append_heading_string(out: &mut Vec<u8>, value: &str, codepage: u16) -> Result<(), OleError> {
    if codepage == UNICODE_CODEPAGE {
        append_u16(out, VT_LPWSTR, "serialized heading string")?;
        append_u16(out, 0, "serialized heading string")?;
        append_lpwstr(out, value)
    } else {
        append_u16(out, VT_LPSTR, "serialized heading string")?;
        append_u16(out, 0, "serialized heading string")?;
        append_unaligned_ansi(out, value, codepage)
    }
}

fn append_unaligned_ansi(out: &mut Vec<u8>, value: &str, codepage: u16) -> Result<(), OleError> {
    let bytes = encode_ansi(value, codepage)?;
    let length = checked_add(bytes.len(), 1, "UnalignedLpstr length")?;
    append_u32(
        out,
        checked_u32(length, "UnalignedLpstr length")?,
        "serialized UnalignedLpstr",
    )?;
    reserve_bytes(out, length, "serialized UnalignedLpstr")?;
    out.extend_from_slice(&bytes);
    out.push(0);
    Ok(())
}

fn append_lpwstr(out: &mut Vec<u8>, value: &str) -> Result<(), OleError> {
    let units = checked_add(value.encode_utf16().count(), 1, "Lpwstr length")?;
    let byte_len = checked_mul(units, 2, "Lpwstr length")?;
    append_u32(
        out,
        checked_u32(units, "Lpwstr length")?,
        "serialized Lpwstr",
    )?;
    reserve_bytes(out, byte_len, "serialized Lpwstr")?;
    for unit in value.encode_utf16() {
        out.extend_from_slice(&unit.to_le_bytes());
    }
    out.extend_from_slice(&0u16.to_le_bytes());
    append_relative_padding(out, byte_len)
}

fn append_relative_padding(out: &mut Vec<u8>, value_len: usize) -> Result<(), OleError> {
    let padding = (4 - (value_len & 3)) & 3;
    reserve_bytes(out, padding, "serialized composite string padding")?;
    out.resize(out.len() + padding, 0);
    Ok(())
}

fn read_heading_string(reader: &mut ValueReader<'_>, codepage: u16) -> Result<String, OleError> {
    let string_type = reader.read_u16("heading string type")?;
    if reader.read_u16("heading string reserved field")? != 0 {
        return Err(invalid("Heading string reserved field must be zero"));
    }
    match string_type {
        VT_LPSTR if codepage != UNICODE_CODEPAGE => {
            read_unaligned_ansi(reader, codepage, "heading string")
        },
        VT_LPWSTR if codepage == UNICODE_CODEPAGE => read_lpwstr(reader, "heading string"),
        VT_LPSTR | VT_LPWSTR => Err(invalid(
            "Heading string type does not match the section code page",
        )),
        _ => Err(invalid("Heading string type must be VT_LPSTR or VT_LPWSTR")),
    }
}

fn read_unaligned_ansi(
    reader: &mut ValueReader<'_>,
    codepage: u16,
    description: &str,
) -> Result<String, OleError> {
    let length = usize::try_from(reader.read_u32(&format!("{description} length"))?)
        .map_err(|_| invalid(format!("{description} length is too large")))?;
    if length == 0 {
        return Ok(String::new());
    }
    let raw = reader.take(length, description)?;
    let Some(end) = raw.iter().position(|byte| *byte == 0) else {
        return Err(invalid(format!("{description} is not NUL-terminated")));
    };
    decode_ansi(&raw[..end], codepage, description)
}

fn read_lpwstr(reader: &mut ValueReader<'_>, description: &str) -> Result<String, OleError> {
    let units = usize::try_from(reader.read_u32(&format!("{description} length"))?)
        .map_err(|_| invalid(format!("{description} length is too large")))?;
    if units == 0 {
        return Ok(String::new());
    }
    let byte_len = checked_mul(units, 2, description)?;
    let raw = reader.take(byte_len, description)?;
    let end = raw
        .chunks_exact(2)
        .position(|pair| pair == [0, 0])
        .map(|units| units * 2)
        .ok_or_else(|| invalid(format!("{description} is not UTF-16LE terminated")))?;
    let value = decode_utf16(&raw[..end], description)?;
    let padding = (4 - (byte_len & 3)) & 3;
    reader.take(padding, &format!("{description} padding"))?;
    Ok(value)
}

fn validate_encoding(encoding: TextEncoding, codepage: u16) -> Result<(), OleError> {
    let expected = if codepage == UNICODE_CODEPAGE {
        TextEncoding::Unicode
    } else {
        TextEncoding::Ansi
    };
    if encoding != expected {
        return Err(invalid(
            "Document-part string type does not match the section code page",
        ));
    }
    Ok(())
}
