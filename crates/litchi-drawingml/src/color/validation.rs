//! Structural and scalar validation for `DrawingML` color fragments.

use quick_xml::Reader;
use quick_xml::events::Event;

use crate::{Error, Result};

/// Maximum accepted color-fragment size.
pub const MAX_XML_BYTES: usize = 64 * 1024;
/// Maximum accepted nesting depth for an opaque fragment.
pub const MAX_DEPTH: usize = 32;
/// Maximum accepted element count for an opaque fragment.
pub const MAX_NODES: usize = 256;
/// Maximum number of typed color transforms in one color choice.
pub const MAX_TRANSFORMS: usize = 64;

const MIN_PERCENTAGE: i32 = -100_000;
const MAX_PERCENTAGE: i32 = 100_000;
const MAX_POSITIVE_ANGLE: u32 = 21_600_000;

pub(super) fn percentage(value: i32, kind: &'static str) -> Result<i32> {
    if !(MIN_PERCENTAGE..=MAX_PERCENTAGE).contains(&value) {
        return Err(Error::Invalid(format!(
            "DrawingML {kind} must be between -100000 and 100000"
        )));
    }
    Ok(value)
}

pub(super) fn positive_percentage(value: u32, kind: &'static str) -> Result<u32> {
    if value > MAX_PERCENTAGE as u32 {
        return Err(Error::Invalid(format!(
            "DrawingML {kind} must be at most 100000"
        )));
    }
    Ok(value)
}

pub(super) fn positive_angle(value: u32) -> Result<u32> {
    if value > MAX_POSITIVE_ANGLE {
        return Err(Error::Invalid(
            "DrawingML positive fixed angles must be at most 21600000".into(),
        ));
    }
    Ok(value)
}

pub(super) fn parse_percentage(value: &str, kind: &'static str) -> Result<i32> {
    let value = parse_scaled(value, kind)?;
    percentage(value, kind)
}

pub(super) fn parse_positive_percentage(value: &str, kind: &'static str) -> Result<u32> {
    let value = parse_scaled(value, kind)?;
    if value < 0 {
        return Err(Error::Invalid(format!(
            "DrawingML {kind} cannot be negative"
        )));
    }
    let value = u32::try_from(value).map_err(|error| {
        Error::Invalid(format!(
            "DrawingML {kind} cannot be represented as a positive percentage: {error}"
        ))
    })?;
    positive_percentage(value, kind)
}

pub(super) fn parse_angle(value: &str) -> Result<i32> {
    value.trim().parse::<i32>().map_err(|_error| {
        Error::Invalid(format!(
            "DrawingML angles must be signed 60000ths of a degree: {value:?}"
        ))
    })
}

pub(super) fn parse_positive_angle(value: &str) -> Result<u32> {
    let value = value.trim().parse::<u32>().map_err(|_error| {
        Error::Invalid(format!(
            "DrawingML positive angles must be unsigned 60000ths of a degree: {value:?}"
        ))
    })?;
    positive_angle(value)
}

/// Parse the `DrawingML` percentage lexical union.
///
/// The standard stores values as thousandths of a percent. Office also reads
/// a human-readable percent sign form; accepting both keeps the shared owner
/// useful for real-world files while always writing the canonical integer form.
fn parse_scaled(value: &str, kind: &'static str) -> Result<i32> {
    let value = value.trim();
    if let Some(value) = value.strip_suffix('%') {
        let value = value.trim();
        if value.is_empty() {
            return Err(Error::Invalid(format!(
                "DrawingML {kind} has an empty percent value"
            )));
        }

        let (negative, value) = match value.as_bytes().first() {
            Some(b'-') => (true, &value[1..]),
            Some(b'+') => (false, &value[1..]),
            _ => (false, value),
        };
        if value.is_empty() {
            return Err(Error::Invalid(format!(
                "DrawingML {kind} has an empty percent value"
            )));
        }

        let (whole, fraction) = value.split_once('.').unwrap_or((value, ""));
        if fraction.len() > 3 || !fraction.bytes().all(|byte| byte.is_ascii_digit()) {
            return Err(Error::Invalid(format!(
                "DrawingML {kind} percent values have at most three fractional digits"
            )));
        }
        let whole = if whole.is_empty() {
            0
        } else {
            whole.parse::<u64>().map_err(|_error| {
                Error::Invalid(format!("invalid DrawingML {kind} percent: {value:?}%"))
            })?
        };
        let fraction = if fraction.is_empty() {
            0
        } else {
            let parsed = fraction.parse::<u32>().map_err(|_error| {
                Error::Invalid(format!("invalid DrawingML {kind} percent: {value:?}%"))
            })?;
            let digits = u32::try_from(fraction.len()).map_err(|error| {
                Error::Invalid(format!(
                    "DrawingML {kind} percent precision is invalid: {error}"
                ))
            })?;
            parsed * 10_u32.pow(3 - digits)
        };
        let scaled = whole
            .checked_mul(1000)
            .and_then(|whole| whole.checked_add(u64::from(fraction)))
            .ok_or_else(|| Error::Invalid(format!("DrawingML {kind} percent overflows")))?;
        let signed = i64::try_from(scaled)
            .ok()
            .and_then(|scaled| {
                if negative {
                    scaled.checked_neg()
                } else {
                    Some(scaled)
                }
            })
            .and_then(|scaled| i32::try_from(scaled).ok())
            .ok_or_else(|| Error::Invalid(format!("DrawingML {kind} percent overflows")))?;
        return Ok(signed);
    }

    value.parse::<i32>().map_err(|_error| {
        Error::Invalid(format!(
            "DrawingML {kind} must be an integer thousandth of a percent: {value:?}"
        ))
    })
}

/// Validate and return the original fragment without allocating a copy.
pub(crate) fn validated_fragment(xml: &[u8]) -> Result<&[u8]> {
    if xml.len() > MAX_XML_BYTES {
        return Err(Error::Limit {
            resource: "DrawingML color XML",
            limit: MAX_XML_BYTES,
        });
    }

    let mut reader = Reader::from_reader(xml);
    let mut stack: Vec<Vec<u8>> = Vec::new();
    let mut root_seen = false;
    let mut root_closed = false;
    let mut nodes = 0usize;

    loop {
        let event = reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        match event {
            Event::Start(element) => {
                nodes = nodes
                    .checked_add(1)
                    .ok_or_else(|| Error::Invalid("DrawingML color node count overflow".into()))?;
                if nodes > MAX_NODES {
                    return Err(Error::Limit {
                        resource: "DrawingML color nodes",
                        limit: MAX_NODES,
                    });
                }
                if stack.is_empty() {
                    if root_seen {
                        return Err(Error::Invalid(
                            "DrawingML color fragment contains multiple roots".into(),
                        ));
                    }
                    root_seen = true;
                }
                if stack.len() >= MAX_DEPTH {
                    return Err(Error::Limit {
                        resource: "DrawingML color depth",
                        limit: MAX_DEPTH,
                    });
                }
                stack.push(element.name().as_ref().to_vec());
            },
            Event::Empty(_) => {
                nodes = nodes
                    .checked_add(1)
                    .ok_or_else(|| Error::Invalid("DrawingML color node count overflow".into()))?;
                if nodes > MAX_NODES {
                    return Err(Error::Limit {
                        resource: "DrawingML color nodes",
                        limit: MAX_NODES,
                    });
                }
                if stack.is_empty() {
                    if root_seen {
                        return Err(Error::Invalid(
                            "DrawingML color fragment contains multiple roots".into(),
                        ));
                    }
                    root_seen = true;
                    root_closed = true;
                }
            },
            Event::End(element) => {
                let Some(expected) = stack.pop() else {
                    return Err(Error::Invalid(
                        "DrawingML color fragment has an unmatched closing element".into(),
                    ));
                };
                if expected.as_slice() != element.name().as_ref() {
                    return Err(Error::Invalid(
                        "DrawingML color fragment has mismatched closing elements".into(),
                    ));
                }
                if stack.is_empty() {
                    root_closed = true;
                }
            },
            Event::Text(text) if stack.is_empty() => {
                if !text
                    .decode()
                    .map_err(|error| Error::Xml(error.to_string()))?
                    .trim()
                    .is_empty()
                {
                    return Err(Error::Invalid(
                        "DrawingML color fragment contains text outside its root".into(),
                    ));
                }
            },
            Event::CData(_) | Event::GeneralRef(_) if stack.is_empty() => {
                return Err(Error::Invalid(
                    "DrawingML color fragment contains data outside its root".into(),
                ));
            },
            Event::Decl(_) | Event::DocType(_) => {
                return Err(Error::Invalid(
                    "DrawingML color fragment cannot contain an XML declaration or doctype".into(),
                ));
            },
            Event::Eof => break,
            Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::PI(_)
            | Event::GeneralRef(_) => {},
        }
    }

    if !root_seen || !root_closed || !stack.is_empty() {
        return Err(Error::Invalid(
            "DrawingML color fragment must contain one complete root".into(),
        ));
    }
    Ok(xml)
}
