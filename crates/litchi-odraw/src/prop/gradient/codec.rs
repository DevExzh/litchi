use super::model::{MAX_STOPS, Stop, Stops};
use super::validation::validate;
use crate::prop::{Array, Id, Props, Value};
use crate::{Error, Result};

const ELEMENT_SIZE: u16 = 8;
const ARRAY_HEADER_SIZE: usize = 6;

/// Decodes the typed `fillShadeColors_complex` property when present.
///
/// # Errors
///
/// Returns `Error::MalformedProperties` if the property is not a complex
/// `IMsoArray` value or the decoded array fails shade-stop validation.
pub fn parse<'data>(properties: &Props<'data>) -> Result<Option<Stops<'data>>> {
    let Some(property) = properties.prop(Id::FillShadeColors) else {
        return Ok(None);
    };
    let array = match property.value() {
        Value::Array(array) => *array,
        Value::Simple(_) | Value::Complex(_) => {
            return Err(Error::MalformedProperties {
                reason: "fillShadeColors must be a complex IMsoArray property",
            });
        },
    };
    let stops = from_array(array)?;
    Ok(Some(stops))
}

/// Decodes one standalone `IMsoArray` payload of `MSOSHADECOLOR` elements.
///
/// # Errors
///
/// Returns `Error::MalformedProperties` if the payload is not an exact
/// `IMsoArray` of eight-byte elements or fails shade-stop validation, and
/// `Error::ArithmeticOverflow` if the declared array extent overflows.
pub fn parse_payload(payload: &[u8]) -> Result<Stops<'_>> {
    from_array(Array::new(payload)?)
}

/// Encodes checked shade stops into the exact `IMsoArray` representation.
///
/// # Errors
///
/// Returns `Error::MalformedProperties` if the stop count exceeds the safe
/// bound or the stops fail shade-stop validation, and
/// `Error::ArithmeticOverflow` if the encoded array extent overflows.
pub fn encode(stops: &[Stop]) -> Result<Vec<u8>> {
    if stops.len() > MAX_STOPS || stops.len() > usize::from(u16::MAX) {
        return Err(Error::MalformedProperties {
            reason: "gradient stop count exceeds the safe bound",
        });
    }
    let payload_len = stops
        .len()
        .checked_mul(usize::from(ELEMENT_SIZE))
        .and_then(|len| ARRAY_HEADER_SIZE.checked_add(len))
        .ok_or(Error::ArithmeticOverflow {
            context: "gradient stop array extent",
        })?;
    let mut payload = Vec::with_capacity(payload_len);
    let count = u16::try_from(stops.len()).map_err(|_err| Error::MalformedProperties {
        reason: "gradient stop count exceeds the safe bound",
    })?;
    payload.extend_from_slice(&count.to_le_bytes());
    payload.extend_from_slice(&count.to_le_bytes());
    payload.extend_from_slice(&ELEMENT_SIZE.to_le_bytes());
    for stop in stops {
        payload.extend_from_slice(&stop.color().raw().to_le_bytes());
        payload.extend_from_slice(&stop.position().raw().to_le_bytes());
    }

    // Validate the canonical result as well as the caller's typed inputs.  The
    // borrowed return is intentionally dropped; callers receive owned bytes.
    parse_payload(&payload)?;
    Ok(payload)
}

fn from_array(array: Array<'_>) -> Result<Stops<'_>> {
    if array.raw_element_size() != ELEMENT_SIZE {
        return Err(Error::MalformedProperties {
            reason: "gradient stop array element size is not MSOSHADECOLOR",
        });
    }
    validate(&array)?;
    let stops = Stops::from_array(array);
    Ok(stops)
}
