//! Small exact wire helpers for the table-topology rewrites.
//!
//! Every rewrite here keeps the bytes of each field it does not change and
//! re-encodes only the fields it does, so unknown fields survive.

use litchi_iwa_common::wire::parse_wire_view;

use super::{Error, Result, reserve};

/// One field of a parsed message.
#[derive(Clone, Copy)]
pub(super) struct Field<'a> {
    pub(super) number: u32,
    pub(super) wire_type: u8,
    pub(super) raw: &'a [u8],
    pub(super) payload: &'a [u8],
}

/// The fields of `data`, in order.
pub(super) fn fields(data: &[u8]) -> Result<Vec<Field<'_>>> {
    let view = parse_wire_view(data, litchi_iwa_common::WireLimits::DEFAULT)
        .map_err(|error| Error::from_wire(&error))?;
    let mut output = Vec::new();
    reserve(&mut output, view.len())?;
    for field in view.fields() {
        output.push(Field {
            number: field.number(),
            wire_type: field.wire_type(),
            raw: field.raw(),
            payload: field.payload(),
        });
    }
    Ok(output)
}

/// Decode one canonical varint payload.
pub(super) fn varint(payload: &[u8]) -> Result<u64> {
    let (value, length) = litchi_iwa_common::varint::decode_varint_from_bytes(payload)
        .map_err(|_error| Error::Invalid)?;
    if length != payload.len() {
        return Err(Error::Invalid);
    }
    Ok(value)
}

/// Append a canonical varint.
pub(super) fn push_varint(output: &mut Vec<u8>, value: u64) -> Result<()> {
    reserve(output, 10)?;
    litchi_iwa_common::varint::encode_varint_into(output, value);
    Ok(())
}

/// Append a varint field.
pub(super) fn push_varint_field(output: &mut Vec<u8>, number: u32, value: u64) -> Result<()> {
    push_varint(output, u64::from(number) << 3)?;
    push_varint(output, value)
}

/// Append a length-delimited field.
pub(super) fn push_bytes_field(output: &mut Vec<u8>, number: u32, payload: &[u8]) -> Result<()> {
    push_varint(output, (u64::from(number) << 3) | 2)?;
    push_varint(
        output,
        u64::try_from(payload.len()).map_err(|_error| Error::Invalid)?,
    )?;
    reserve(output, payload.len())?;
    output.extend_from_slice(payload);
    Ok(())
}

/// The payload of the unique length-delimited field `number`, if present.
pub(super) fn single_bytes(data: &[u8], number: u32) -> Result<Option<&[u8]>> {
    let mut found = None;
    for field in fields(data)? {
        if field.number != number {
            continue;
        }
        if field.wire_type != 2 || found.replace(field.payload).is_some() {
            return Err(Error::Invalid);
        }
    }
    Ok(found)
}

/// The value of the unique varint field `number`, if present.
pub(super) fn single_varint(data: &[u8], number: u32) -> Result<Option<u64>> {
    let mut found = None;
    for field in fields(data)? {
        if field.number != number {
            continue;
        }
        if field.wire_type != 0 || found.replace(varint(field.payload)?).is_some() {
            return Err(Error::Invalid);
        }
    }
    Ok(found)
}

/// The payloads of every length-delimited field `number`.
pub(super) fn repeated_bytes(data: &[u8], number: u32) -> Result<Vec<&[u8]>> {
    let mut output = Vec::new();
    for field in fields(data)? {
        if field.number != number {
            continue;
        }
        if field.wire_type != 2 {
            return Err(Error::Invalid);
        }
        reserve(&mut output, 1)?;
        output.push(field.payload);
    }
    Ok(output)
}

/// The values of every unpacked varint field `number`; a packed occurrence
/// is decoded too.
pub(super) fn repeated_varints(data: &[u8], number: u32) -> Result<Vec<u64>> {
    let mut output = Vec::new();
    for field in fields(data)? {
        if field.number != number {
            continue;
        }
        match field.wire_type {
            0 => {
                reserve(&mut output, 1)?;
                output.push(varint(field.payload)?);
            },
            2 => {
                let mut rest = field.payload;
                while !rest.is_empty() {
                    let (value, length) = litchi_iwa_common::varint::decode_varint_from_bytes(rest)
                        .map_err(|_error| Error::Invalid)?;
                    reserve(&mut output, 1)?;
                    output.push(value);
                    rest = rest.get(length..).ok_or(Error::Invalid)?;
                }
            },
            _ => return Err(Error::Invalid),
        }
    }
    Ok(output)
}

/// Rebuild `data`, replacing each field for which `map` returns `Some`
/// (an empty replacement drops the field) and copying the rest.
pub(super) fn rebuild(
    data: &[u8],
    mut map: impl FnMut(Field<'_>) -> Result<Option<Vec<u8>>>,
) -> Result<Vec<u8>> {
    let mut output = Vec::new();
    reserve(&mut output, data.len())?;
    for field in fields(data)? {
        match map(field)? {
            Some(replacement) => {
                reserve(&mut output, replacement.len())?;
                output.extend_from_slice(&replacement);
            },
            None => {
                output.extend_from_slice(field.raw);
            },
        }
    }
    Ok(output)
}

/// Encode a complete varint field.
pub(super) fn varint_field(number: u32, value: u64) -> Result<Vec<u8>> {
    let mut output = Vec::new();
    push_varint_field(&mut output, number, value)?;
    Ok(output)
}

/// Encode a complete length-delimited field.
pub(super) fn bytes_field(number: u32, payload: &[u8]) -> Result<Vec<u8>> {
    let mut output = Vec::new();
    push_bytes_field(&mut output, number, payload)?;
    Ok(output)
}
