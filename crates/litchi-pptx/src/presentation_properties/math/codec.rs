//! Strict XML codec for the presentation-level a14:m extension.

use super::model::{BinaryBreak, BinarySubtractionBreak, Properties};
use crate::{Error, Result};
use litchi_ooxml_common::xml::{OMML_NAMESPACE_URI, xsd_token_atom};
use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, QName, ResolveResult};
use quick_xml::reader::NsReader;

const DRAWING_MATH_NAMESPACE: &[u8] = b"http://schemas.microsoft.com/office/drawing/2010/main";
const OMML_NAMESPACE: &[u8] = OMML_NAMESPACE_URI.as_bytes();
const STRICT_OMML_NAMESPACE: &[u8] = b"http://purl.oclc.org/ooxml/officeDocument/math";
const MAX_BYTES: usize = 8 * 1024 * 1024;
const MAX_NODES: usize = 8;
const MAX_DEPTH: usize = 4;

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Kind {
    Root,
    MathProperties,
    BinaryBreak,
    BinarySubtractionBreak,
}

#[derive(Clone, Debug, PartialEq, Eq)]
enum ResolvedNamespace {
    Unbound,
    Bound(Vec<u8>),
    Unknown(Vec<u8>),
}

/// Parse one complete a14:m payload.
pub(crate) fn parse(xml: &[u8]) -> Result<Properties> {
    if xml.len() > MAX_BYTES {
        return Err(Error::Limit {
            resource: "presentation math XML",
            limit: MAX_BYTES,
        });
    }

    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut event_buffer = Vec::new();
    let mut stack = Vec::new();
    let mut root_seen = false;
    let mut root_closed = false;
    let mut nodes = 0usize;
    let mut child_stage = 0u8;
    let mut value = Properties::new();

    loop {
        event_buffer.clear();
        let (namespace, event) = reader
            .read_resolved_event_into(&mut event_buffer)
            .map_err(xml_error)?;
        let namespace = own_namespace(namespace);
        match event {
            Event::Start(element) => {
                nodes = nodes.checked_add(1).ok_or(Error::Limit {
                    resource: "presentation math XML nodes",
                    limit: MAX_NODES,
                })?;
                if nodes > MAX_NODES {
                    return Err(Error::Limit {
                        resource: "presentation math XML nodes",
                        limit: MAX_NODES,
                    });
                }
                if stack.len() >= MAX_DEPTH {
                    return Err(Error::Limit {
                        resource: "presentation math XML depth",
                        limit: MAX_DEPTH,
                    });
                }
                if root_closed {
                    return Err(invalid("math extension contains multiple roots"));
                }
                let kind = start_kind(
                    &namespace,
                    &element,
                    &reader,
                    &mut root_seen,
                    &mut child_stage,
                    &mut value,
                    stack.len(),
                )?;
                stack.push(kind);
            },
            Event::Empty(element) => {
                nodes = nodes.checked_add(1).ok_or(Error::Limit {
                    resource: "presentation math XML nodes",
                    limit: MAX_NODES,
                })?;
                if nodes > MAX_NODES {
                    return Err(Error::Limit {
                        resource: "presentation math XML nodes",
                        limit: MAX_NODES,
                    });
                }
                if root_closed {
                    return Err(invalid("math extension contains multiple roots"));
                }
                let kind = start_kind(
                    &namespace,
                    &element,
                    &reader,
                    &mut root_seen,
                    &mut child_stage,
                    &mut value,
                    stack.len(),
                )?;
                if kind == Kind::Root {
                    return Err(invalid("a14:m must contain mathPr"));
                }
                // An empty mathPr is the valid schema-default snapshot.
            },
            Event::End(element) => {
                let kind = stack
                    .pop()
                    .ok_or_else(|| invalid("math extension has an unexpected closing tag"))?;
                expect_end(&namespace, element.name(), kind)?;
                if stack.is_empty() {
                    root_closed = true;
                }
            },
            Event::Text(text) => {
                let text = text.decode().map_err(xml_error)?;
                if !text.trim().is_empty() {
                    return Err(invalid("math extension contains unexpected text"));
                }
            },
            Event::CData(_) | Event::DocType(_) | Event::GeneralRef(_) => {
                return Err(invalid("math extension contains unsupported XML content"));
            },
            Event::Comment(_) => {},
            Event::Decl(_) | Event::PI(_) => {
                if root_seen {
                    return Err(invalid("math extension has XML content after its root"));
                }
            },
            Event::Eof => break,
        }
    }

    if !root_seen || !root_closed || !stack.is_empty() {
        return Err(invalid("math extension has an incomplete root"));
    }
    value.validate()?;
    Ok(value)
}

/// Write one complete a14:m payload in the requested package conformance.
pub(crate) fn write(out: &mut String, value: &Properties, strict: bool) -> Result<()> {
    value.validate()?;
    let omml = if strict {
        "http://purl.oclc.org/ooxml/officeDocument/math"
    } else {
        OMML_NAMESPACE_URI
    };
    out.push_str(
        "<a14:m xmlns:a14=\"http://schemas.microsoft.com/office/drawing/2010/main\" \
         xmlns:m=\"",
    );
    out.push_str(omml);
    out.push_str("\"><m:mathPr>");
    if let Some(value) = value.binary_break {
        out.push_str("<m:brkBin m:val=\"");
        out.push_str(value.wire_value());
        out.push_str("\"/>");
    }
    if let Some(value) = value.binary_subtraction_break {
        out.push_str("<m:brkBinSub m:val=\"");
        out.push_str(value.wire_value());
        out.push_str("\"/>");
    }
    out.push_str("</m:mathPr></a14:m>");
    Ok(())
}

fn start_kind(
    namespace: &ResolvedNamespace,
    element: &BytesStart<'_>,
    reader: &NsReader<&[u8]>,
    root_seen: &mut bool,
    child_stage: &mut u8,
    value: &mut Properties,
    depth: usize,
) -> Result<Kind> {
    let local = element.name().local_name();
    match depth {
        0 => {
            if *root_seen || !is_name(namespace, local.as_ref(), DRAWING_MATH_NAMESPACE, b"m") {
                return Err(invalid("expected one a14:m root"));
            }
            *root_seen = true;
            no_attributes(element, reader, "a14:m")?;
            Ok(Kind::Root)
        },
        1 => {
            if !is_omml_name(namespace, local.as_ref(), b"mathPr") {
                return Err(invalid("a14:m must contain m:mathPr"));
            }
            if *child_stage != 0 {
                return Err(invalid("a14:m contains duplicate mathPr"));
            }
            no_attributes(element, reader, "m:mathPr")?;
            *child_stage = 1;
            Ok(Kind::MathProperties)
        },
        2 => {
            if is_omml_name(namespace, local.as_ref(), b"brkBin") {
                if *child_stage > 2 || value.binary_break.is_some() {
                    return Err(invalid("invalid brkBin order or duplicate"));
                }
                value.binary_break = Some(BinaryBreak::from_wire(&required_value(
                    element, reader, "m:brkBin",
                )?)?);
                *child_stage = 2;
                Ok(Kind::BinaryBreak)
            } else if is_omml_name(namespace, local.as_ref(), b"brkBinSub") {
                if *child_stage > 3 || value.binary_subtraction_break.is_some() {
                    return Err(invalid("invalid brkBinSub order or duplicate"));
                }
                value.binary_subtraction_break = Some(BinarySubtractionBreak::from_wire(
                    &required_value(element, reader, "m:brkBinSub")?,
                )?);
                *child_stage = 3;
                Ok(Kind::BinarySubtractionBreak)
            } else {
                Err(invalid("mathPr contains an unsupported child"))
            }
        },
        _ => Err(invalid(
            "math extension is deeper than the presentation schema",
        )),
    }
}

fn expect_end(namespace: &ResolvedNamespace, name: QName<'_>, kind: Kind) -> Result<()> {
    let local = name.local_name();
    let valid = match kind {
        Kind::Root => is_name(namespace, local.as_ref(), DRAWING_MATH_NAMESPACE, b"m"),
        Kind::MathProperties => is_omml_name(namespace, local.as_ref(), b"mathPr"),
        Kind::BinaryBreak => is_omml_name(namespace, local.as_ref(), b"brkBin"),
        Kind::BinarySubtractionBreak => is_omml_name(namespace, local.as_ref(), b"brkBinSub"),
    };
    if valid {
        Ok(())
    } else {
        Err(invalid("math extension has mismatched closing tag"))
    }
}

fn required_value(
    element: &BytesStart<'_>,
    reader: &NsReader<&[u8]>,
    label: &str,
) -> Result<String> {
    let mut value = None;
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(xml_error)?;
        if attribute.key.as_namespace_binding().is_some() {
            continue;
        }
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        let namespace = resolved_namespace(namespace, "attribute")?;
        let local = local.as_ref();
        if (namespace != OMML_NAMESPACE && namespace != STRICT_OMML_NAMESPACE) || local != b"val" {
            return Err(invalid(format!("unexpected attribute on {label}")));
        }
        if value.is_some() {
            return Err(invalid(format!("duplicate attribute on {label}")));
        }
        value = Some(
            attribute
                .decoded_and_normalized_value(XmlVersion::Explicit1_0, reader.decoder())
                .map_err(xml_error)?
                .into_owned(),
        );
    }
    let value = value.ok_or_else(|| invalid(format!("{label} requires m:val")))?;
    let value = xsd_token_atom(&value)
        .ok_or_else(|| invalid(format!("{label} has an invalid m:val token")))?;
    Ok(value.to_owned())
}

fn no_attributes(element: &BytesStart<'_>, reader: &NsReader<&[u8]>, label: &str) -> Result<()> {
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(xml_error)?;
        if attribute.key.as_namespace_binding().is_some() {
            continue;
        }
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        let namespace = resolved_namespace(namespace, "attribute")?;
        return Err(invalid(format!(
            "unexpected {{{}}}{} attribute on {label}",
            String::from_utf8_lossy(namespace),
            String::from_utf8_lossy(local.as_ref())
        )));
    }
    Ok(())
}

fn is_name(namespace: &ResolvedNamespace, local: &[u8], expected: &[u8], name: &[u8]) -> bool {
    local == name
        && matches!(namespace, ResolvedNamespace::Bound(value) if value.as_slice() == expected)
}

fn is_omml_name(namespace: &ResolvedNamespace, local: &[u8], name: &[u8]) -> bool {
    local == name
        && matches!(namespace, ResolvedNamespace::Bound(value) if value.as_slice() == OMML_NAMESPACE || value.as_slice() == STRICT_OMML_NAMESPACE)
}

fn own_namespace(namespace: ResolveResult<'_>) -> ResolvedNamespace {
    match namespace {
        ResolveResult::Unbound => ResolvedNamespace::Unbound,
        ResolveResult::Bound(Namespace(value)) => ResolvedNamespace::Bound(value.to_vec()),
        ResolveResult::Unknown(prefix) => ResolvedNamespace::Unknown(prefix),
    }
}

fn resolved_namespace<'a>(namespace: ResolveResult<'a>, kind: &str) -> Result<&'a [u8]> {
    match namespace {
        ResolveResult::Bound(Namespace(value)) => Ok(value),
        ResolveResult::Unbound => Err(invalid(format!("unbound XML {kind} namespace"))),
        ResolveResult::Unknown(prefix) => Err(invalid(format!(
            "unknown XML {kind} namespace prefix '{}'",
            String::from_utf8_lossy(prefix.as_ref())
        ))),
    }
}

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}

fn xml_error(error: impl std::fmt::Display) -> Error {
    Error::Xml(error.to_string())
}
