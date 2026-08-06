//! Bounded XML codec for DrawingML color fragments.

use quick_xml::Reader;
use quick_xml::events::{BytesStart, Event};

use crate::{Error, Result};

use super::model::{Rgb, Scheme, Unknown, Value, write_hex};

/// Maximum accepted color-fragment size.
pub const MAX_XML_BYTES: usize = 64 * 1024;
/// Maximum accepted nesting depth for an opaque fragment.
pub const MAX_DEPTH: usize = 32;
/// Maximum accepted element count for an opaque fragment.
pub const MAX_NODES: usize = 256;

/// Read one DrawingML color-choice fragment.
pub fn read(xml: &[u8]) -> Result<Value> {
    let validated = validated_fragment(xml)?;
    let mut reader = Reader::from_reader(validated);

    loop {
        match reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?
        {
            Event::Text(text) if text.decode().map_err(xml_error)?.trim().is_empty() => {},
            Event::Start(element) => {
                let Some(value) = typed_value(&element, reader.decoder())? else {
                    return Ok(Value::Unknown(Unknown::from_validated(validated)));
                };
                let root_name = element.name().as_ref().to_vec();
                match reader
                    .read_event()
                    .map_err(|error| Error::Xml(error.to_string()))?
                {
                    Event::End(end) if end.name().as_ref() == root_name.as_slice() => {
                        return if tail_is_empty(&mut reader)? {
                            Ok(value)
                        } else {
                            Ok(Value::Unknown(Unknown::from_validated(validated)))
                        };
                    },
                    _ => return Ok(Value::Unknown(Unknown::from_validated(validated))),
                }
            },
            Event::Empty(element) => {
                let Some(value) = typed_value(&element, reader.decoder())? else {
                    return Ok(Value::Unknown(Unknown::from_validated(validated)));
                };
                return if tail_is_empty(&mut reader)? {
                    Ok(value)
                } else {
                    Ok(Value::Unknown(Unknown::from_validated(validated)))
                };
            },
            _ => return Ok(Value::Unknown(Unknown::from_validated(validated))),
        }
    }
}

/// Write one DrawingML color-choice fragment using the conventional `a`
/// prefix. The result is a fragment and intentionally has no namespace
/// declaration so the host can retain its own namespace spelling.
pub fn write(value: &Value) -> Result<Vec<u8>> {
    match value {
        Value::Rgb(rgb) => {
            let mut output = String::with_capacity(29);
            output.push_str("<a:srgbClr val=\"");
            write_hex(&mut output, *rgb);
            output.push_str("\"/>");
            Ok(output.into_bytes())
        },
        Value::Scheme(scheme) => {
            let mut output = String::with_capacity(31);
            output.push_str("<a:schemeClr val=\"");
            output.push_str(scheme.token());
            output.push_str("\"/>");
            Ok(output.into_bytes())
        },
        Value::Unknown(value) => Ok(value.as_xml().to_vec()),
    }
}

fn typed_value(
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
) -> Result<Option<Value>> {
    let local = element.local_name();
    let kind = match local.as_ref() {
        b"srgbClr" => 0_u8,
        b"schemeClr" => 1_u8,
        _ => return Ok(None),
    };

    let Some(value) = attribute(element, b"val", decoder)? else {
        return Err(Error::Invalid(format!(
            "DrawingML {} color is missing val",
            String::from_utf8_lossy(local.as_ref())
        )));
    };

    match kind {
        0 => Ok(Some(Value::Rgb(Rgb::parse(&value)?))),
        1 => Ok(Scheme::from_token(&value).map(Value::Scheme)),
        _ => unreachable!("typed DrawingML color kind is closed"),
    }
}

fn attribute(
    element: &BytesStart<'_>,
    expected: &[u8],
    decoder: quick_xml::encoding::Decoder,
) -> Result<Option<String>> {
    let mut value = None;
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        let key = attribute.key.as_ref();
        if key == b"xmlns" || key.starts_with(b"xmlns:") {
            continue;
        }
        if attribute.key.prefix().is_some() || attribute.key.local_name().as_ref() != expected {
            return Ok(None);
        }
        if value.is_some() {
            return Err(Error::Invalid(format!(
                "duplicate DrawingML color attribute '{}'",
                String::from_utf8_lossy(expected)
            )));
        }
        value = Some(
            attribute
                .decoded_and_normalized_value(quick_xml::XmlVersion::Explicit1_0, decoder)
                .map_err(|error| Error::Xml(error.to_string()))?
                .into_owned(),
        );
    }
    Ok(value)
}

fn tail_is_empty(reader: &mut Reader<&[u8]>) -> Result<bool> {
    loop {
        match reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?
        {
            Event::Text(text) if text.decode().map_err(xml_error)?.trim().is_empty() => {},
            Event::Eof => return Ok(true),
            _ => return Ok(false),
        }
    }
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
                if !text.decode().map_err(xml_error)?.trim().is_empty() {
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
            _ => {},
        }
    }

    if !root_seen || !root_closed || !stack.is_empty() {
        return Err(Error::Invalid(
            "DrawingML color fragment must contain one complete root".into(),
        ));
    }
    Ok(xml)
}

fn xml_error(error: quick_xml::encoding::EncodingError) -> Error {
    Error::Xml(error.to_string())
}
