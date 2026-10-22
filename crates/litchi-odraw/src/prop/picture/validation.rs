use super::error::check;
use super::{Error, Flags, LimitKind, Limits, Metadata, Name, Refusal, RefusalKind, Result};
use crate::prop::{Id, Props, Value};

pub(super) fn decode<'data>(
    properties: &Props<'data>,
    limits: Limits,
) -> Result<Option<Metadata<'data>>> {
    check(
        LimitKind::Properties,
        properties.len(),
        limits.properties_value(),
    )?;
    let name = picture_name(properties, limits)?;
    let flags = picture_flags(properties)?;
    if name.is_none() && !properties.has(Id::BlipFlags) {
        return Ok(None);
    }
    Ok(Some(Metadata::new(
        name,
        flags.unwrap_or(Flags::from_raw(0)?),
    )))
}

fn malformed(reason: &'static str) -> Error {
    crate::Error::MalformedProperties { reason }.into()
}

fn validate_name(raw: &[u8], limits: Limits) -> Result<()> {
    if raw.is_empty() || !raw.len().is_multiple_of(2) {
        return Err(malformed(
            "picture name must be a bounded even-length UTF-16 string",
        ));
    }
    let count = raw.len() / 2 - 1;
    check(LimitKind::NameUnits, count, limits.name_units_value())?;
    if raw[raw.len() - 2..] != [0, 0] {
        return Err(malformed("picture name must have one terminating NUL"));
    }
    let units = raw[..raw.len() - 2]
        .as_chunks::<2>()
        .0
        .iter()
        .map(|bytes| u16::from_le_bytes(*bytes));
    for character in char::decode_utf16(units) {
        match character {
            Ok('\0') => return Err(malformed("picture name contains an interior NUL")),
            Err(_) => return Err(malformed("picture name is not valid UTF-16")),
            Ok(_) => {},
        }
    }
    Ok(())
}

pub(super) fn encode_name(text: &str, limits: Limits) -> Result<Vec<u8>> {
    let capacity = text
        .len()
        .saturating_add(1)
        .saturating_mul(2)
        .min((limits.name_units_value() + 1) * 2)
        .min(limits.property_bytes_value());
    let mut bytes = Vec::with_capacity(capacity);
    for (index, unit) in text.encode_utf16().enumerate() {
        check(LimitKind::NameUnits, index + 1, limits.name_units_value())?;
        if unit == 0 {
            return Err(Error::Refused(Refusal::new(
                RefusalKind::InvalidValue,
                "set_name",
                "picture name contains an interior NUL",
            )));
        }
        check(
            LimitKind::PropertyBytes,
            bytes.len() + 2,
            limits.property_bytes_value(),
        )?;
        bytes.extend_from_slice(&unit.to_le_bytes());
    }
    check(
        LimitKind::PropertyBytes,
        bytes.len() + 2,
        limits.property_bytes_value(),
    )?;
    bytes.extend_from_slice(&0_u16.to_le_bytes());
    Ok(bytes)
}

fn picture_name<'data>(properties: &Props<'data>, limits: Limits) -> Result<Option<Name<'data>>> {
    let Some(property) = properties.prop(Id::PictureFileName) else {
        return Ok(None);
    };
    if !property.is_complex() {
        if property.raw_value() != 0 {
            return Err(malformed(
                "picture name length is nonzero without its complex value",
            ));
        }
        return Ok(None);
    }
    let raw = match property.value() {
        Value::Complex(raw) => *raw,
        Value::Simple(_) | Value::Array(_) => {
            return Err(malformed("picture name is not a scalar Unicode payload"));
        },
    };
    validate_name(raw, limits)?;
    Ok(Some(Name::from_raw(raw)))
}

fn picture_flags(properties: &Props<'_>) -> Result<Option<Flags>> {
    let Some(property) = properties.prop(Id::BlipFlags) else {
        return Ok(None);
    };
    if property.is_complex() {
        return Err(malformed("picture flags must be a simple property"));
    }
    Ok(Some(Flags::from_raw(property.raw_value().cast_unsigned())?))
}
