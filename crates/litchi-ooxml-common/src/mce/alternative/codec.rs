//! Lossless scanner for one `mc:AlternateContent` element.

use std::{io::BufRead, sync::Arc};

use quick_xml::{
    XmlVersion,
    events::{BytesStart, Event},
    name::{Namespace, QName, ResolveResult},
    reader::NsReader,
};

use super::super::{Error, NAMESPACE};
use super::model::{Alternatives, Kind, Limits, Span, Stored};
use crate::{xml::decode_xml_reference, xml_name};

/// Read and validate one self-contained `mc:AlternateContent` element.
///
/// The returned snapshot retains the input bytes exactly. Namespace
/// declarations must be present in the supplied fragment; a caller that has
/// only an inherited namespace scope should first materialize that scope on
/// the fragment root.
pub fn read(xml: &[u8], limits: &Limits) -> Result<Alternatives, Error> {
    if xml.len() > limits.bytes {
        return Err(limit("alternate-content XML bytes"));
    }

    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut frames = Vec::new();
    let mut branches = Vec::new();
    let mut root_seen = false;
    let mut root_closed = false;
    let mut nodes = 0usize;

    loop {
        let start = position(&reader)?;
        let (namespace, event) = reader
            .read_resolved_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        let namespace = NamespaceKind::from(namespace);
        let end = position(&reader)?;

        if matches!(event, Event::Eof) {
            break;
        }
        charge(&mut nodes, limits.nodes, "alternate-content XML events")?;

        match event {
            Event::Start(element) => {
                validate_element(&reader, &element, &namespace)?;
                if frames.len() >= limits.depth {
                    return Err(limit("alternate-content XML depth"));
                }

                if frames.is_empty() {
                    if root_seen || root_closed {
                        return Err(invalid("AlternateContent must have one root element"));
                    }
                    require_name(&namespace, element.name(), b"AlternateContent")?;
                    validate_attributes(&reader, &element, AttributePolicy::Root)?;
                    root_seen = true;
                    frames.push(Frame::Root);
                } else if matches!(frames.last(), Some(Frame::Root)) {
                    let (kind, requirements) = parse_branch(&reader, &element, &namespace)?;
                    frames.push(Frame::Branch(OpenBranch {
                        kind,
                        span: Span {
                            start,
                            content_start: end,
                            content_end: end,
                            end,
                        },
                        requirements,
                    }));
                } else {
                    validate_attributes(&reader, &element, AttributePolicy::Any)?;
                    frames.push(Frame::Nested);
                }
            },
            Event::Empty(element) => {
                validate_element(&reader, &element, &namespace)?;
                if frames.len() >= limits.depth {
                    return Err(limit("alternate-content XML depth"));
                }

                if frames.is_empty() {
                    if root_seen || root_closed {
                        return Err(invalid("AlternateContent must have one root element"));
                    }
                    require_name(&namespace, element.name(), b"AlternateContent")?;
                    validate_attributes(&reader, &element, AttributePolicy::Root)?;
                    return Err(invalid("AlternateContent cannot be empty"));
                }

                if matches!(frames.last(), Some(Frame::Root)) {
                    let (kind, requirements) = parse_branch(&reader, &element, &namespace)?;
                    let span = Span {
                        start,
                        content_start: end,
                        content_end: end,
                        end,
                    };
                    add_branch_with_span(&mut branches, kind, span, requirements, limits)?;
                } else {
                    validate_attributes(&reader, &element, AttributePolicy::Any)?;
                }
            },
            Event::End(_) => {
                let frame = frames
                    .pop()
                    .ok_or_else(|| invalid("unexpected AlternateContent closing tag"))?;
                match frame {
                    Frame::Root => {
                        if branches.is_empty() {
                            return Err(invalid("AlternateContent requires a Choice"));
                        }
                        root_closed = true;
                    },
                    Frame::Branch(mut open) => {
                        open.span.content_end = start;
                        open.span.end = end;
                        let stored = Stored {
                            kind: open.kind,
                            span: open.span,
                            requirements: open.requirements,
                        };
                        add_branch_with_span(
                            &mut branches,
                            stored.kind,
                            stored.span,
                            stored.requirements,
                            limits,
                        )?;
                    },
                    Frame::Nested => {},
                }
            },
            Event::Text(text) => {
                if outside_branch(&frames) && !is_xml_whitespace(&text)? {
                    return Err(invalid(
                        "AlternateContent may contain only whitespace outside branches",
                    ));
                }
            },
            Event::CData(_) => {
                if outside_branch(&frames) {
                    return Err(invalid("CDATA is not allowed directly in AlternateContent"));
                }
            },
            Event::GeneralRef(reference) => {
                decode_xml_reference(&reference).map_err(|error| Error::Xml(error.to_string()))?;
                if outside_branch(&frames) {
                    return Err(invalid(
                        "entity references are not allowed directly in AlternateContent",
                    ));
                }
            },
            Event::Comment(_) => {},
            Event::PI(_) | Event::Decl(_) => {
                return Err(invalid(
                    "processing instructions and declarations are not valid in an AlternateContent fragment",
                ));
            },
            Event::DocType(_) => {
                return Err(invalid(
                    "DTD declarations are forbidden in AlternateContent",
                ));
            },
            Event::Eof => unreachable!("EOF handled before event dispatch"),
        }
    }

    if !root_seen || !root_closed || !frames.is_empty() {
        return Err(invalid("unterminated AlternateContent fragment"));
    }

    Ok(Alternatives {
        source: Arc::from(xml.to_owned()),
        branches: branches.into_boxed_slice(),
    })
}

#[derive(Debug)]
enum Frame {
    Root,
    Branch(OpenBranch),
    Nested,
}

#[derive(Debug, Clone)]
enum NamespaceKind {
    Bound(Vec<u8>),
    Unbound,
    Unknown,
}

impl From<ResolveResult<'_>> for NamespaceKind {
    fn from(result: ResolveResult<'_>) -> Self {
        match result {
            ResolveResult::Bound(Namespace(namespace)) => Self::Bound(namespace.to_vec()),
            ResolveResult::Unbound => Self::Unbound,
            ResolveResult::Unknown(_) => Self::Unknown,
        }
    }
}

#[derive(Debug)]
struct OpenBranch {
    kind: Kind,
    span: Span,
    requirements: Box<[Box<str>]>,
}

#[derive(Debug, Clone, Copy)]
enum AttributePolicy {
    Root,
    Choice,
    Fallback,
    Any,
}

fn parse_branch<R: BufRead>(
    reader: &NsReader<R>,
    element: &BytesStart<'_>,
    namespace: &NamespaceKind,
) -> Result<(Kind, Box<[Box<str>]>), Error> {
    if !is_mce_namespace(namespace) {
        return Err(invalid(
            "AlternateContent children must use the MCE namespace",
        ));
    }

    match element.name().local_name().as_ref() {
        b"Choice" => {
            validate_attributes(reader, element, AttributePolicy::Choice)?;
            let value = choice_requires(reader, element)?;
            let mut requirements = Vec::new();
            for prefix in value.split_whitespace() {
                if !xml_name::is_ncname(prefix) {
                    return Err(invalid("Choice Requires contains an invalid prefix"));
                }
                let qualified = format!("{prefix}:x");
                let resolved = reader
                    .resolver()
                    .resolve_element(QName(qualified.as_bytes()))
                    .0;
                let namespace = match resolved {
                    ResolveResult::Bound(Namespace(namespace)) => std::str::from_utf8(namespace)
                        .map_err(|error| Error::Xml(error.to_string()))?,
                    ResolveResult::Unknown(_) => {
                        return Err(invalid(format!(
                            "Choice Requires prefix '{prefix}' is unbound"
                        )));
                    },
                    ResolveResult::Unbound => {
                        return Err(invalid(format!(
                            "Choice Requires prefix '{prefix}' is unbound"
                        )));
                    },
                };
                if requirements
                    .iter()
                    .any(|known: &Box<str>| known.as_ref() == namespace)
                {
                    return Err(invalid("Choice Requires contains a duplicate namespace"));
                }
                requirements.push(namespace.into());
            }
            if requirements.is_empty() {
                return Err(invalid("Choice Requires must contain a prefix"));
            }
            Ok((Kind::Choice, requirements.into_boxed_slice()))
        },
        b"Fallback" => {
            validate_attributes(reader, element, AttributePolicy::Fallback)?;
            Ok((Kind::Fallback, Box::new([])))
        },
        _ => Err(invalid(
            "AlternateContent may contain only Choice and Fallback",
        )),
    }
}

fn choice_requires<R: BufRead>(
    reader: &NsReader<R>,
    element: &BytesStart<'_>,
) -> Result<String, Error> {
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        if attribute.key.as_ref() == b"Requires" {
            return attribute
                .decoded_and_normalized_value(XmlVersion::Explicit1_0, reader.decoder())
                .map(std::borrow::Cow::into_owned)
                .map_err(|error| Error::Xml(error.to_string()));
        }
    }
    Err(invalid("Choice requires a Requires attribute"))
}

fn add_branch_with_span(
    branches: &mut Vec<Stored>,
    kind: Kind,
    span: Span,
    requirements: Box<[Box<str>]>,
    limits: &Limits,
) -> Result<(), Error> {
    if branches.len() >= limits.branches {
        return Err(limit("alternate-content branch count"));
    }
    if kind == Kind::Choice && branches.iter().any(|branch| branch.kind == Kind::Fallback) {
        return Err(invalid("Choice cannot follow Fallback"));
    }
    if kind == Kind::Fallback && branches.iter().any(|branch| branch.kind == Kind::Fallback) {
        return Err(invalid(
            "AlternateContent cannot contain two Fallback branches",
        ));
    }
    branches.push(Stored {
        kind,
        span,
        requirements,
    });
    Ok(())
}

fn validate_element<R: BufRead>(
    reader: &NsReader<R>,
    element: &BytesStart<'_>,
    namespace: &NamespaceKind,
) -> Result<(), Error> {
    let element_name = element.name();
    let name = std::str::from_utf8(element_name.as_ref())
        .map_err(|error| Error::Xml(error.to_string()))?;
    if !xml_name::is_qualified_name(name) {
        return Err(invalid("element name is not a valid XML QName"));
    }
    if matches!(namespace, NamespaceKind::Unknown) {
        return Err(invalid("element name uses an unbound namespace prefix"));
    }
    validate_attributes(reader, element, AttributePolicy::Any)
}

fn validate_attributes<R: BufRead>(
    reader: &NsReader<R>,
    element: &BytesStart<'_>,
    policy: AttributePolicy,
) -> Result<(), Error> {
    let mut requires = false;
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        let raw_name = std::str::from_utf8(attribute.key.as_ref())
            .map_err(|error| Error::Xml(error.to_string()))?;
        if raw_name == "xmlns" || raw_name.starts_with("xmlns:") {
            continue;
        }
        if !xml_name::is_qualified_name(raw_name) {
            return Err(invalid("attribute name is not a valid XML QName"));
        }
        if matches!(
            reader.resolver().resolve_attribute(attribute.key).0,
            ResolveResult::Unknown(_)
        ) {
            return Err(invalid("attribute name uses an unbound namespace prefix"));
        }
        match policy {
            AttributePolicy::Any => {},
            AttributePolicy::Choice if raw_name == "Requires" => {
                if requires {
                    return Err(invalid("Choice has duplicate Requires attributes"));
                }
                requires = true;
            },
            AttributePolicy::Root | AttributePolicy::Fallback | AttributePolicy::Choice => {
                return Err(invalid("MCE wrapper has an unsupported attribute"));
            },
        }
    }
    if matches!(policy, AttributePolicy::Choice) && !requires {
        return Err(invalid("Choice requires a Requires attribute"));
    }
    Ok(())
}

fn require_name(
    namespace: &NamespaceKind,
    name: QName<'_>,
    expected_local: &[u8],
) -> Result<(), Error> {
    if is_mce_namespace(namespace) && name.local_name().as_ref() == expected_local {
        Ok(())
    } else {
        Err(invalid("root is not an MCE AlternateContent element"))
    }
}

fn is_mce_namespace(namespace: &NamespaceKind) -> bool {
    matches!(namespace, NamespaceKind::Bound(value) if value == NAMESPACE.as_bytes())
}

fn outside_branch(frames: &[Frame]) -> bool {
    matches!(frames.last(), None | Some(Frame::Root))
}

fn is_xml_whitespace(text: &quick_xml::events::BytesText<'_>) -> Result<bool, Error> {
    let text = text
        .decode()
        .map_err(|error| Error::Xml(error.to_string()))?;
    Ok(text
        .chars()
        .all(|character| matches!(character, ' ' | '\t' | '\r' | '\n')))
}

fn position<R: BufRead>(reader: &NsReader<R>) -> Result<usize, Error> {
    usize::try_from(reader.buffer_position())
        .map_err(|_| invalid("AlternateContent byte offset exceeds platform size"))
}

fn charge(current: &mut usize, max: usize, resource: &'static str) -> Result<(), Error> {
    *current = current.checked_add(1).ok_or_else(|| limit(resource))?;
    if *current > max {
        return Err(limit(resource));
    }
    Ok(())
}

fn invalid(message: impl Into<String>) -> Error {
    Error::NonConformant(message.into())
}

fn limit(resource: &'static str) -> Error {
    Error::LimitExceeded(resource.into())
}
