//! Strict XML codec for the presentation-level a14:m extension.

use super::model::{BinaryBreak, BinarySubtractionBreak, Properties};
use crate::presentation_properties::Limits;
use crate::{Error, Result};
use litchi_ooxml_common::xml::attributes::BytesStartExt as _;
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

/// The last element the `a14:m` payload accepted, in schema order:
///
/// ```text
/// Start --mathPr--> MathProperties --brkBin--> BinaryBreak --brkBinSub--> BinarySubtractionBreak
/// MathProperties --brkBinSub--> BinarySubtractionBreak
/// ```
///
/// `mathPr` is accepted once, in `Start` only; a second one is a duplicate
/// in every later stage. `brkBin` and `brkBinSub` are both optional and
/// accepted at most once: each is refused once the stage has reached or
/// passed it, which covers both a duplicate and `brkBin` after `brkBinSub`.
#[derive(Clone, Copy, Debug, PartialEq, Eq, PartialOrd, Ord)]
enum MathStage {
    Start,
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

/// Where the math-properties reader is relative to the `a14:m` root.
///
/// ```text
/// Prolog --declaration, processing instruction--> Prolog
/// Prolog --a14:m start tag----------------------> Open
/// Open   --a14:m end tag------------------------> Closed
/// any    --comment, whitespace------------------> unchanged
/// ```
///
/// A declaration or a processing instruction is accepted any number of times,
/// but only in `Prolog`. Refused: either of them in `Open` or `Closed`; an
/// empty-element root; a top-level element other than `a14:m` in `Prolog`; any
/// top-level element in `Closed` ("multiple roots"); text other than
/// whitespace, CDATA, a reference or a DOCTYPE anywhere; and the end of input
/// before `Closed`.
#[derive(Clone, Copy, PartialEq, Eq)]
enum RootPhase {
    /// No root start tag has been read.
    Prolog,
    /// Inside the root.
    Open,
    /// The root has ended.
    Closed,
}

/// Parse one complete a14:m payload.
pub(crate) fn parse(xml: &[u8], limits: Limits) -> Result<Properties> {
    let max_bytes = limits.max_input_bytes().min(MAX_BYTES);
    let max_nodes = limits.max_nodes().min(MAX_NODES);
    let max_depth = limits.max_depth().min(MAX_DEPTH);
    if xml.len() > max_bytes {
        return Err(Error::Limit {
            kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::XmlBytes),
            observed: xml.len(),
            maximum: max_bytes,
        });
    }

    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut event_buffer = Vec::new();
    let mut stack = Vec::new();
    let mut phase = RootPhase::Prolog;
    let mut nodes = 0usize;
    let mut stage = MathStage::Start;
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
                    kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::XmlNodes),
                    observed: usize::MAX,
                    maximum: max_nodes,
                })?;
                if nodes > max_nodes {
                    return Err(Error::Limit {
                        kind: crate::opened::LimitKind::Codec(
                            crate::opened::CodecLimitKind::XmlNodes,
                        ),
                        observed: nodes,
                        maximum: max_nodes,
                    });
                }
                if stack.len() >= max_depth {
                    return Err(Error::Limit {
                        kind: crate::opened::LimitKind::Codec(
                            crate::opened::CodecLimitKind::XmlDepth,
                        ),
                        observed: stack.len().checked_add(1).unwrap_or(usize::MAX),
                        maximum: max_depth,
                    });
                }
                if phase == RootPhase::Closed {
                    return Err(invalid("math extension contains multiple roots"));
                }
                let kind = start_kind(
                    &namespace,
                    &element,
                    &reader,
                    &mut phase,
                    &mut stage,
                    &mut value,
                    stack.len(),
                )?;
                stack.push(kind);
            },
            Event::Empty(element) => {
                nodes = nodes.checked_add(1).ok_or(Error::Limit {
                    kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::XmlNodes),
                    observed: usize::MAX,
                    maximum: max_nodes,
                })?;
                if nodes > max_nodes {
                    return Err(Error::Limit {
                        kind: crate::opened::LimitKind::Codec(
                            crate::opened::CodecLimitKind::XmlNodes,
                        ),
                        observed: nodes,
                        maximum: max_nodes,
                    });
                }
                if phase == RootPhase::Closed {
                    return Err(invalid("math extension contains multiple roots"));
                }
                let kind = start_kind(
                    &namespace,
                    &element,
                    &reader,
                    &mut phase,
                    &mut stage,
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
                    phase = RootPhase::Closed;
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
            Event::Decl(_) | Event::PI(_) => match phase {
                RootPhase::Prolog => {},
                RootPhase::Open | RootPhase::Closed => {
                    return Err(invalid("math extension has XML content after its root"));
                },
            },
            Event::Eof => break,
        }
    }

    match phase {
        RootPhase::Closed if stack.is_empty() => {},
        RootPhase::Prolog | RootPhase::Open | RootPhase::Closed => {
            return Err(invalid("math extension has an incomplete root"));
        },
    }
    value.validate()?;
    Ok(value)
}

/// Write one complete a14:m payload in the requested package conformance.
pub(crate) fn write(
    out: &mut super::super::codec::XmlSink,
    value: &Properties,
    strict: bool,
    _limits: Limits,
) -> Result<()> {
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
    out.check()
}

fn start_kind(
    namespace: &ResolvedNamespace,
    element: &BytesStart<'_>,
    reader: &NsReader<&[u8]>,
    phase: &mut RootPhase,
    stage: &mut MathStage,
    value: &mut Properties,
    depth: usize,
) -> Result<Kind> {
    let local = element.name().local_name();
    match depth {
        0 => {
            if *phase != RootPhase::Prolog
                || !is_name(namespace, local.as_ref(), DRAWING_MATH_NAMESPACE, b"m")
            {
                return Err(invalid("expected one a14:m root"));
            }
            *phase = RootPhase::Open;
            no_attributes(element, reader, "a14:m")?;
            Ok(Kind::Root)
        },
        1 => {
            if !is_omml_name(namespace, local.as_ref(), b"mathPr") {
                return Err(invalid("a14:m must contain m:mathPr"));
            }
            if *stage != MathStage::Start {
                return Err(invalid("a14:m contains duplicate mathPr"));
            }
            no_attributes(element, reader, "m:mathPr")?;
            *stage = MathStage::MathProperties;
            Ok(Kind::MathProperties)
        },
        2 => {
            if is_omml_name(namespace, local.as_ref(), b"brkBin") {
                if *stage >= MathStage::BinaryBreak {
                    return Err(invalid("invalid brkBin order or duplicate"));
                }
                value.binary_break = Some(BinaryBreak::from_wire(&required_value(
                    element, reader, "m:brkBin",
                )?)?);
                *stage = MathStage::BinaryBreak;
                Ok(Kind::BinaryBreak)
            } else if is_omml_name(namespace, local.as_ref(), b"brkBinSub") {
                if *stage >= MathStage::BinarySubtractionBreak {
                    return Err(invalid("invalid brkBinSub order or duplicate"));
                }
                value.binary_subtraction_break = Some(BinarySubtractionBreak::from_wire(
                    &required_value(element, reader, "m:brkBinSub")?,
                )?);
                *stage = MathStage::BinarySubtractionBreak;
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
    for attribute in element.checked_attributes() {
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
    for attribute in element.checked_attributes() {
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

#[cfg(test)]
mod envelope_tests {
    //! What the math-properties reader accepts before, around and after `a14:m`, and the exact error reported otherwise. Each expected
    //! line is the row label followed by the result.

    const DECLARATION: &str = r#"<?xml version="1.0" encoding="UTF-8"?>"#;

    /// One reader's root element in the forms the rows are built from.
    struct Root<'a> {
        /// The start tag.
        open: &'a str,
        /// Content that makes the element valid for its reader.
        body: &'a str,
        /// The end tag.
        close: &'a str,
        /// The element as one empty-element tag.
        empty: &'a str,
    }

    fn rows(root: &Root<'_>) -> Vec<(&'static str, String)> {
        let Root {
            open,
            body,
            close,
            empty,
        } = *root;
        let document = format!("{open}{body}{close}");
        let d = document.as_str();
        vec![
            ("empty input", String::new()),
            ("whitespace only", " \n".to_owned()),
            ("declaration only", DECLARATION.to_owned()),
            ("empty-element root", empty.to_owned()),
            ("start and end root", format!("{open}{close}")),
            ("root then whitespace", format!("{d}\n ")),
            ("root then comment", format!("{d}<!--c-->")),
            ("root then PI", format!("{d}<?pi x?>")),
            ("root then text", format!("{d}text")),
            ("root then second root", format!("{d}{d}")),
            ("root then foreign start", format!("{d}<foreign></foreign>")),
            ("root then foreign empty", format!("{d}<foreign/>")),
            ("declaration after root", format!("{d}{DECLARATION}")),
            ("two declarations", format!("{DECLARATION}{DECLARATION}{d}")),
            (
                "comment before declaration",
                format!("<!--c-->{DECLARATION}{d}"),
            ),
            ("text before root", format!("text{d}")),
            ("unterminated root", format!("{open}{body}")),
            ("stray end tag", format!("{d}</foreign>")),
            ("DOCTYPE", format!("<!DOCTYPE x>{d}")),
            ("top-level CDATA", format!("<![CDATA[x]]>{d}")),
            ("declaration then root", format!("{DECLARATION}{d}")),
            (
                "whitespace before declaration",
                format!(" {DECLARATION}{d}"),
            ),
            ("PI before declaration", format!("<?pi x?>{DECLARATION}{d}")),
            ("comment then root", format!("<!--c-->{d}")),
            ("PI then root", format!("<?pi x?>{d}")),
            ("root then text then second root", format!("{d}text{d}")),
            ("empty-element root then root", format!("{empty}{d}")),
            (
                "declaration after empty-element root",
                format!("{empty}{DECLARATION}"),
            ),
        ]
    }

    fn assert_matrix(root: &Root<'_>, expected: &[&str], read: impl Fn(&[u8]) -> String) {
        let actual = rows(root)
            .into_iter()
            .map(|(label, xml)| format!("{label}: {}", read(xml.as_bytes())))
            .collect::<Vec<_>>();
        assert_eq!(actual, expected, "actual rows:\n{actual:#?}");
    }

    fn outcome<T, E: std::fmt::Debug>(
        result: Result<T, E>,
        summary: impl Fn(T) -> String,
    ) -> String {
        match result {
            Ok(value) => format!("Ok({})", summary(value)),
            Err(error) => format!("{error:?}"),
        }
    }

    /// The reader of the math-properties extension payload.
    #[test]
    fn math_properties_envelope() {
        let root = Root {
            open: r#"<a14:m xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">"#,
            body: r"<m:mathPr/>",
            close: r"</a14:m>",
            empty: r#"<a14:m xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"/>"#,
        };
        assert_matrix(
            &root,
            &[
                "empty input: Invalid(\"math extension has an incomplete root\")",
                "whitespace only: Invalid(\"math extension has an incomplete root\")",
                "declaration only: Invalid(\"math extension has an incomplete root\")",
                "empty-element root: Invalid(\"a14:m must contain mathPr\")",
                "start and end root: Ok()",
                "root then whitespace: Ok()",
                "root then comment: Ok()",
                "root then PI: Invalid(\"math extension has XML content after its root\")",
                "root then text: Invalid(\"math extension contains unexpected text\")",
                "root then second root: Invalid(\"math extension contains multiple roots\")",
                "root then foreign start: Invalid(\"math extension contains multiple roots\")",
                "root then foreign empty: Invalid(\"math extension contains multiple roots\")",
                "declaration after root: Invalid(\"math extension has XML content after its root\")",
                "two declarations: Ok()",
                "comment before declaration: Ok()",
                "text before root: Invalid(\"math extension contains unexpected text\")",
                "unterminated root: Invalid(\"math extension has an incomplete root\")",
                "stray end tag: Xml(\"ill-formed document: close tag `</foreign>` does not match any open tag\")",
                "DOCTYPE: Invalid(\"math extension contains unsupported XML content\")",
                "top-level CDATA: Invalid(\"math extension contains unsupported XML content\")",
                "declaration then root: Ok()",
                "whitespace before declaration: Ok()",
                "PI before declaration: Ok()",
                "comment then root: Ok()",
                "PI then root: Ok()",
                "root then text then second root: Invalid(\"math extension contains unexpected text\")",
                "empty-element root then root: Invalid(\"a14:m must contain mathPr\")",
                "declaration after empty-element root: Invalid(\"a14:m must contain mathPr\")",
            ],
            |xml| {
                outcome(
                    super::parse(xml, crate::presentation_properties::Limits::default()),
                    |_| String::new(),
                )
            },
        );
    }
}
