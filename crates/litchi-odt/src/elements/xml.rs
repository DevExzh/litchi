use super::element::{Element, ElementBase};
use litchi_core::{Error, Result};
use quick_xml::XmlVersion;
use quick_xml::events::{BytesRef, BytesStart};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;

pub(crate) const OFFICE_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
pub(crate) const TEXT_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:text:1.0";
pub(crate) const DRAW_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
pub(crate) const STYLE_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:style:1.0";
pub(crate) const NUMBER_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0";
pub(crate) const SCRIPT_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:script:1.0";
pub(crate) const XLINK_NAMESPACE: &[u8] = b"http://www.w3.org/1999/xlink";
pub(crate) const XML_NAMESPACE: &[u8] = b"http://www.w3.org/XML/1998/namespace";
pub(crate) const DC_NAMESPACE: &[u8] = b"http://purl.org/dc/elements/1.1/";
pub(crate) const META_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:meta:1.0";

const MAX_INLINE_TEXT_BYTES: usize = 64 * 1024 * 1024;
const MAX_SPACE_COUNT: usize = 1_000_000;

pub(crate) fn is_bound(namespace: &ResolveResult<'_>, expected: &[u8]) -> bool {
    matches!(namespace, ResolveResult::Bound(Namespace(uri)) if *uri == expected)
}

pub(crate) fn namespaced_attribute(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    expected_namespace: &[u8],
    expected_local_name: &[u8],
    context: &str,
) -> Result<Option<String>> {
    let mut value = None;
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| {
            Error::InvalidFormat(format!("invalid {context} attribute: {error}"))
        })?;
        let (namespace, local_name) = reader.resolver().resolve_attribute(attribute.key);
        if is_bound(&namespace, expected_namespace) && local_name.as_ref() == expected_local_name {
            if value.is_some() {
                return Err(Error::InvalidFormat(format!(
                    "duplicate expanded {context} attribute '{}'",
                    String::from_utf8_lossy(expected_local_name)
                )));
            }
            value = Some(
                attribute
                    .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                    .map_err(|error| {
                        Error::InvalidFormat(format!("invalid {context} attribute value: {error}"))
                    })?
                    .into_owned(),
            );
        }
    }
    Ok(value)
}

pub(crate) fn copy_canonical_attributes(
    reader: &NsReader<&[u8]>,
    source: &BytesStart<'_>,
    element: &mut Element,
    context: &str,
) -> Result<()> {
    for attribute in source.attributes() {
        let attribute = attribute.map_err(|error| {
            Error::InvalidFormat(format!("invalid {context} attribute: {error}"))
        })?;
        if attribute.key.as_ref() == b"xmlns" || attribute.key.as_ref().starts_with(b"xmlns:") {
            continue;
        }
        let (namespace, local_name) = reader.resolver().resolve_attribute(attribute.key);
        let local_name = std::str::from_utf8(local_name.as_ref()).map_err(|_error| {
            Error::InvalidFormat(format!("non-UTF-8 {context} attribute name"))
        })?;
        let name = match namespace {
            ResolveResult::Bound(Namespace(uri)) if uri == OFFICE_NAMESPACE => {
                format!("office:{local_name}")
            },
            ResolveResult::Bound(Namespace(uri)) if uri == TEXT_NAMESPACE => {
                format!("text:{local_name}")
            },
            ResolveResult::Bound(Namespace(uri)) if uri == DRAW_NAMESPACE => {
                format!("draw:{local_name}")
            },
            ResolveResult::Bound(Namespace(uri)) if uri == STYLE_NAMESPACE => {
                format!("style:{local_name}")
            },
            ResolveResult::Bound(Namespace(uri)) if uri == NUMBER_NAMESPACE => {
                format!("number:{local_name}")
            },
            ResolveResult::Bound(Namespace(uri)) if uri == SCRIPT_NAMESPACE => {
                format!("script:{local_name}")
            },
            ResolveResult::Bound(Namespace(uri)) if uri == XLINK_NAMESPACE => {
                format!("xlink:{local_name}")
            },
            ResolveResult::Bound(Namespace(uri)) if uri == XML_NAMESPACE => {
                format!("xml:{local_name}")
            },
            ResolveResult::Bound(_) | ResolveResult::Unbound => {
                std::str::from_utf8(attribute.key.as_ref())
                    .map(str::to_string)
                    .map_err(|_error| {
                        Error::InvalidFormat(format!("non-UTF-8 {context} attribute name"))
                    })?
            },
            ResolveResult::Unknown(prefix) => {
                return Err(Error::InvalidFormat(format!(
                    "unknown {context} attribute namespace prefix '{}'",
                    String::from_utf8_lossy(&prefix)
                )));
            },
        };
        if element.has_attribute(&name) {
            return Err(Error::InvalidFormat(format!(
                "duplicate expanded {context} attribute '{name}'"
            )));
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
            .map_err(|error| {
                Error::InvalidFormat(format!("invalid {context} attribute value: {error}"))
            })?;
        element.set_attribute(&name, &value);
    }
    Ok(())
}

pub(crate) fn append_text_control(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    output: &mut String,
) -> Result<()> {
    match element.local_name().as_ref() {
        b"s" => {
            let count = namespaced_attribute(reader, element, TEXT_NAMESPACE, b"c", "text:s")?
                .map(|value| {
                    value.parse::<usize>().map_err(|_error| {
                        Error::InvalidFormat("text:c must be a non-negative integer".to_string())
                    })
                })
                .transpose()?
                .unwrap_or(1);
            if count > MAX_SPACE_COUNT {
                return Err(Error::InvalidFormat(format!(
                    "text:s count exceeds {MAX_SPACE_COUNT}"
                )));
            }
            ensure_text_capacity(output, count)?;
            output.extend(std::iter::repeat_n(' ', count));
        },
        b"tab" => append_checked(output, "\t")?,
        b"line-break" => append_checked(output, "\n")?,
        _ => {},
    }
    Ok(())
}

pub(crate) fn append_checked(output: &mut String, value: &str) -> Result<()> {
    ensure_text_capacity(output, value.len())?;
    output.push_str(value);
    Ok(())
}

fn ensure_text_capacity(output: &str, additional: usize) -> Result<()> {
    let length = output
        .len()
        .checked_add(additional)
        .ok_or_else(|| Error::InvalidFormat("ODF inline text size overflow".to_string()))?;
    if length > MAX_INLINE_TEXT_BYTES {
        return Err(Error::InvalidFormat(format!(
            "ODF inline text exceeds {MAX_INLINE_TEXT_BYTES} bytes"
        )));
    }
    Ok(())
}

pub(crate) fn decode_reference(reference: &BytesRef<'_>, context: &str) -> Result<String> {
    if let Some(character) = reference.resolve_char_ref().map_err(|error| {
        Error::InvalidFormat(format!("invalid {context} character reference: {error}"))
    })? {
        return Ok(character.to_string());
    }
    let name = reference
        .decode()
        .map_err(|error| Error::InvalidFormat(format!("invalid {context} entity: {error}")))?;
    match name.as_ref() {
        "amp" => Ok("&".to_string()),
        "lt" => Ok("<".to_string()),
        "gt" => Ok(">".to_string()),
        "quot" => Ok("\"".to_string()),
        "apos" => Ok("'".to_string()),
        _ => Err(Error::InvalidFormat(format!(
            "unsupported {context} entity '&{name};'"
        ))),
    }
}
