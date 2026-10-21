//! Small exact wire helpers for the table-topology rewrites.
//!
//! Every rewrite here keeps the bytes of each field it does not change and
//! re-encodes only the fields it does, so unknown fields survive.

use litchi_iwa_common::wire::parse_wire_view;

use super::super::structure::{Fault, FaultResult, map_wire, references, reserve};

/// One field of a parsed message.
#[derive(Clone, Copy)]
pub(in crate::package) struct Field<'a> {
    pub(in crate::package) number: u32,
    pub(in crate::package) wire_type: u8,
    pub(in crate::package) raw: &'a [u8],
    pub(in crate::package) payload: &'a [u8],
}

/// The fields of `data`, in order.
pub(in crate::package) fn fields(data: &[u8]) -> FaultResult<Vec<Field<'_>>> {
    let view = parse_wire_view(data, litchi_iwa_common::WireLimits::DEFAULT)
        .map_err(|error| map_wire(&error))?;
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
pub(in crate::package) fn varint(payload: &[u8]) -> FaultResult<u64> {
    let (value, length) = litchi_iwa_common::varint::decode_varint_from_bytes(payload)
        .map_err(|_error| Fault::Invalid)?;
    if length != payload.len() {
        return Err(Fault::Invalid);
    }
    Ok(value)
}

/// Append a canonical varint.
pub(in crate::package) fn push_varint(output: &mut Vec<u8>, value: u64) -> FaultResult<()> {
    reserve(output, 10)?;
    litchi_iwa_common::varint::encode_varint_into(output, value);
    Ok(())
}

/// Append a varint field.
pub(in crate::package) fn push_varint_field(
    output: &mut Vec<u8>,
    number: u32,
    value: u64,
) -> FaultResult<()> {
    push_varint(output, u64::from(number) << 3)?;
    push_varint(output, value)
}

/// Append a length-delimited field.
pub(in crate::package) fn push_bytes_field(
    output: &mut Vec<u8>,
    number: u32,
    payload: &[u8],
) -> FaultResult<()> {
    push_varint(output, (u64::from(number) << 3) | 2)?;
    push_varint(
        output,
        u64::try_from(payload.len()).map_err(|_error| Fault::Invalid)?,
    )?;
    reserve(output, payload.len())?;
    output.extend_from_slice(payload);
    Ok(())
}

/// The payload of the unique length-delimited field `number`, if present.
pub(in crate::package) fn single_bytes(data: &[u8], number: u32) -> FaultResult<Option<&[u8]>> {
    let mut found = None;
    for field in fields(data)? {
        if field.number != number {
            continue;
        }
        if field.wire_type != 2 || found.replace(field.payload).is_some() {
            return Err(Fault::Invalid);
        }
    }
    Ok(found)
}

/// The value of the unique varint field `number`, if present.
pub(in crate::package) fn single_varint(data: &[u8], number: u32) -> FaultResult<Option<u64>> {
    let mut found = None;
    for field in fields(data)? {
        if field.number != number {
            continue;
        }
        if field.wire_type != 0 || found.replace(varint(field.payload)?).is_some() {
            return Err(Fault::Invalid);
        }
    }
    Ok(found)
}

/// The identifier of the unique `TSP.Reference` field `number`, if present.
pub(in crate::package) fn single_reference(data: &[u8], number: u32) -> FaultResult<Option<u64>> {
    single_bytes(data, number)?
        .map(references::reference_identifier)
        .transpose()
}

/// The payloads of every length-delimited field `number`.
pub(in crate::package) fn repeated_bytes(data: &[u8], number: u32) -> FaultResult<Vec<&[u8]>> {
    let mut output = Vec::new();
    for field in fields(data)? {
        if field.number != number {
            continue;
        }
        if field.wire_type != 2 {
            return Err(Fault::Invalid);
        }
        reserve(&mut output, 1)?;
        output.push(field.payload);
    }
    Ok(output)
}

/// The values of every unpacked varint field `number`; a packed occurrence
/// is decoded too.
pub(in crate::package) fn repeated_varints(data: &[u8], number: u32) -> FaultResult<Vec<u64>> {
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
                        .map_err(|_error| Fault::Invalid)?;
                    reserve(&mut output, 1)?;
                    output.push(value);
                    rest = rest.get(length..).ok_or(Fault::Invalid)?;
                }
            },
            _ => return Err(Fault::Invalid),
        }
    }
    Ok(output)
}

/// Rebuild `data`, replacing each field for which `map` returns `Some`
/// (an empty replacement drops the field) and copying the rest.
pub(in crate::package) fn rebuild(
    data: &[u8],
    mut map: impl FnMut(Field<'_>) -> FaultResult<Option<Vec<u8>>>,
) -> FaultResult<Vec<u8>> {
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
pub(in crate::package) fn varint_field(number: u32, value: u64) -> FaultResult<Vec<u8>> {
    let mut output = Vec::new();
    push_varint_field(&mut output, number, value)?;
    Ok(output)
}

/// Encode a complete length-delimited field.
pub(in crate::package) fn bytes_field(number: u32, payload: &[u8]) -> FaultResult<Vec<u8>> {
    let mut output = Vec::new();
    push_bytes_field(&mut output, number, payload)?;
    Ok(output)
}
