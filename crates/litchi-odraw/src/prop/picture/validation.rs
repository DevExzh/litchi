use super::model::{Flags, MAX_NAME_BYTES, MAX_NAME_UNITS, Name};
use crate::prop::{Id, Prop, Props, Value};
use crate::{Error, Result};

pub(super) fn decode<'data>(
    properties: &Props<'data>,
) -> Result<Option<super::model::Metadata<'data>>> {
    let name = picture_name(properties)?;
    let flags = picture_flags(properties)?;
    if name.is_none() && !properties.has(Id::BlipFlags) {
        return Ok(None);
    }
    Ok(Some(super::model::Metadata::new(
        name,
        flags.unwrap_or(Flags::from_raw(0)?),
    )))
}

pub(super) fn validate_name(raw: &[u8]) -> Result<()> {
    if raw.is_empty() || raw.len() > MAX_NAME_BYTES || raw.len() % 2 != 0 {
        return Err(Error::MalformedProperties {
            reason: "picture name must be a bounded even-length UTF-16 string",
        });
    }
    let units = raw
        .chunks_exact(2)
        .map(|pair| u16::from_le_bytes([pair[0], pair[1]]))
        .collect::<Vec<_>>();
    if units.len() > MAX_NAME_UNITS + 1 || units.last().copied() != Some(0) {
        return Err(Error::MalformedProperties {
            reason: "picture name must have one terminating NUL",
        });
    }
    if units[..units.len() - 1].contains(&0) {
        return Err(Error::MalformedProperties {
            reason: "picture name contains an interior NUL",
        });
    }
    String::from_utf16(&units[..units.len() - 1]).map_err(|_| Error::MalformedProperties {
        reason: "picture name is not valid UTF-16",
    })?;
    Ok(())
}

pub(super) fn encode_name(text: &str) -> Result<Vec<u8>> {
    let units = text.encode_utf16().collect::<Vec<_>>();
    if units.len() > MAX_NAME_UNITS || units.contains(&0) {
        return Err(Error::MalformedProperties {
            reason: "picture name is empty-safe but contains an interior NUL or exceeds the bound",
        });
    }
    let byte_len = units
        .len()
        .checked_add(1)
        .and_then(|count| count.checked_mul(2))
        .ok_or(Error::ArithmeticOverflow {
            context: "picture name byte length",
        })?;
    let mut raw = Vec::with_capacity(byte_len);
    for unit in units {
        raw.extend_from_slice(&unit.to_le_bytes());
    }
    raw.extend_from_slice(&0_u16.to_le_bytes());
    validate_name(&raw)?;
    Ok(raw)
}

fn picture_name<'data>(properties: &Props<'data>) -> Result<Option<Name<'data>>> {
    let Some(property) = properties.prop(Id::PictureFileName) else {
        return Ok(None);
    };
    if !property.is_complex() {
        if property.raw_value() != 0 {
            return Err(Error::MalformedProperties {
                reason: "picture name length is nonzero without its complex value",
            });
        }
        return Ok(None);
    }
    let raw = match property.value() {
        Value::Complex(raw) => *raw,
        _ => {
            return Err(Error::MalformedProperties {
                reason: "picture name is not a scalar Unicode payload",
            });
        },
    };
    validate_name(raw)?;
    Ok(Some(Name::from_raw(raw)))
}

fn picture_flags(properties: &Props<'_>) -> Result<Option<Flags>> {
    let Some(property) = properties.prop(Id::BlipFlags) else {
        return Ok(None);
    };
    if property.is_complex() {
        return Err(Error::MalformedProperties {
            reason: "picture flags must be a simple property",
        });
    }
    Ok(Some(Flags::from_raw(property.raw_value() as u32)?))
}

pub(super) fn name_bytes<'data>(property: &Prop<'data>) -> Result<Option<&'data [u8]>> {
    if !property.is_complex() {
        if property.raw_value() != 0 {
            return Err(Error::MalformedProperties {
                reason: "picture name length is nonzero without its complex value",
            });
        }
        return Ok(None);
    }
    match property.value() {
        Value::Complex(raw) => {
            validate_name(raw)?;
            Ok(Some(raw))
        },
        _ => Err(Error::MalformedProperties {
            reason: "picture name is not a scalar Unicode payload",
        }),
    }
}
