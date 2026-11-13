//! Bounded `PresentationML` XML scanning for action settings.

use super::model::Trigger;
use super::{Limits, invalid, limit};
use crate::{Error, Result};
use litchi_ooxml_common::mce::{Capabilities, process_markup_compatibility};
use litchi_ooxml_common::relationships::attribute_value;
use litchi_ooxml_common::xml::{is_drawingml_name, unqualified_attribute_value};
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, NamespaceResolver, QName, ResolveResult};
use quick_xml::reader::NsReader;

pub(super) const MAX_SLIDE_XML_BYTES: usize = 32 * 1024 * 1024;
const MAX_TOTAL_SLIDE_XML_BYTES: usize = 256 * 1024 * 1024;
const MAX_ACTION_SETTINGS: usize = 4_096;
const MAX_XML_NODES: usize = 250_000;
const MAX_XML_DEPTH: usize = 128;
pub(super) const MAX_ATTRIBUTE_BYTES: usize = 4_096;

const PRESENTATIONML_NAMESPACE: &[u8] =
    b"http://schemas.openxmlformats.org/presentationml/2006/main";
const STRICT_PRESENTATIONML_NAMESPACE: &[u8] = b"http://purl.oclc.org/ooxml/presentationml/main";

pub(super) struct Parsed {
    pub(super) trigger: Trigger,
    pub(super) action: Option<String>,
    pub(super) relationship_id: Option<String>,
    pub(super) tooltip: Option<String>,
    pub(super) target_frame: Option<String>,
}

impl Limits {
    pub(super) fn add_slide_xml(&mut self, bytes: usize) -> Result<()> {
        if bytes > MAX_SLIDE_XML_BYTES {
            return Err(limit(
                crate::error::CodecLimitKind::XmlBytes,
                bytes,
                MAX_SLIDE_XML_BYTES,
            ));
        }
        self.total_slide_xml_bytes =
            self.total_slide_xml_bytes
                .checked_add(bytes)
                .ok_or_else(|| {
                    limit(
                        crate::error::CodecLimitKind::XmlBytes,
                        usize::MAX,
                        MAX_TOTAL_SLIDE_XML_BYTES,
                    )
                })?;
        if self.total_slide_xml_bytes > MAX_TOTAL_SLIDE_XML_BYTES {
            return Err(limit(
                crate::error::CodecLimitKind::XmlBytes,
                self.total_slide_xml_bytes,
                MAX_TOTAL_SLIDE_XML_BYTES,
            ));
        }
        Ok(())
    }

    pub(super) fn add_action(&mut self) -> Result<()> {
        self.action_count = self.action_count.checked_add(1).ok_or_else(|| {
            limit(
                crate::error::CodecLimitKind::ActionSettings,
                usize::MAX,
                MAX_ACTION_SETTINGS,
            )
        })?;
        if self.action_count > MAX_ACTION_SETTINGS {
            return Err(limit(
                crate::error::CodecLimitKind::ActionSettings,
                self.action_count,
                MAX_ACTION_SETTINGS,
            ));
        }
        Ok(())
    }
}

/// Where the action scan is relative to the `p:sld` root.
///
/// ```text
/// Prolog --p:sld start tag-----> Open
/// Prolog --empty p:sld tag-----> Closed
/// Open   --p:sld end tag-------> Closed
/// any    --declaration, comment, processing instruction, text, CDATA--> unchanged
/// ```
///
/// The last line holds in every phase: a declaration is accepted anywhere and
/// content outside the root is ignored. Refused: any other top-level element,
/// before or after the root; a DOCTYPE; and the end of input before `Closed`.
#[derive(Clone, Copy, PartialEq, Eq)]
enum RootPhase {
    /// No root start tag has been read.
    Prolog,
    /// Inside the root.
    Open,
    /// The root has ended.
    Closed,
}

pub(super) fn scan(xml_bytes: &[u8], limits: &mut Limits) -> Result<Vec<Parsed>> {
    if xml_bytes.len() > MAX_SLIDE_XML_BYTES {
        return Err(limit(
            crate::error::CodecLimitKind::XmlBytes,
            xml_bytes.len(),
            MAX_SLIDE_XML_BYTES,
        ));
    }

    let capabilities = Capabilities::ooxml_baseline();
    let mce_limits = litchi_ooxml_common::mce::Limits {
        max_input_bytes: MAX_SLIDE_XML_BYTES,
        max_output_bytes: MAX_SLIDE_XML_BYTES,
        max_depth: MAX_XML_DEPTH,
        max_namespace_bindings: 4_096,
        max_directive_tokens: 4_096,
        max_choices_per_alternate: 1_024,
        max_attributes_per_element: litchi_ooxml_common::mce::DEFAULT_MAX_ATTRIBUTES_PER_ELEMENT,
    };
    let xml = process_markup_compatibility(xml_bytes, &capabilities, &mce_limits)?.xml;
    let mut reader = NsReader::from_reader(xml.as_ref());
    let mut actions = Vec::new();
    let mut nodes = 0usize;
    let mut depth = 0usize;
    let mut phase = RootPhase::Prolog;

    loop {
        let decoder = reader.decoder();
        let event = reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?
            .into_owned();
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);

        match event {
            Event::Start(element) => {
                increment_nodes(&mut nodes)?;
                depth = depth.checked_add(1).ok_or_else(|| {
                    limit(
                        crate::error::CodecLimitKind::XmlDepth,
                        usize::MAX,
                        MAX_XML_DEPTH,
                    )
                })?;
                if depth > MAX_XML_DEPTH {
                    return Err(limit(
                        crate::error::CodecLimitKind::XmlDepth,
                        depth,
                        MAX_XML_DEPTH,
                    ));
                }
                if depth == 1 {
                    validate_slide_root(&namespace, element.name(), phase)?;
                    phase = RootPhase::Open;
                }
                maybe_push(
                    &mut actions,
                    &namespace,
                    &element,
                    decoder,
                    &resolver,
                    limits,
                )?;
            },
            Event::Empty(element) => {
                increment_nodes(&mut nodes)?;
                let child_depth = depth.checked_add(1).ok_or_else(|| {
                    limit(
                        crate::error::CodecLimitKind::XmlDepth,
                        usize::MAX,
                        MAX_XML_DEPTH,
                    )
                })?;
                if child_depth > MAX_XML_DEPTH {
                    return Err(limit(
                        crate::error::CodecLimitKind::XmlDepth,
                        child_depth,
                        MAX_XML_DEPTH,
                    ));
                }
                if child_depth == 1 {
                    validate_slide_root(&namespace, element.name(), phase)?;
                    phase = RootPhase::Closed;
                }
                maybe_push(
                    &mut actions,
                    &namespace,
                    &element,
                    decoder,
                    &resolver,
                    limits,
                )?;
            },
            Event::End(element) => {
                if depth == 0 {
                    return Err(invalid("invalid slide XML nesting"));
                }
                if depth == 1 {
                    if !is_presentationml_name(&namespace, element.name(), b"sld") {
                        return Err(invalid(
                            "slide XML must close with a PresentationML sld element",
                        ));
                    }
                    phase = RootPhase::Closed;
                }
                depth -= 1;
            },
            Event::DocType(_) => return Err(invalid("slide XML must not contain a DTD")),
            Event::Eof => match phase {
                RootPhase::Closed if depth == 0 => break,
                RootPhase::Prolog | RootPhase::Open | RootPhase::Closed => {
                    return Err(invalid("unterminated or missing PresentationML slide root"));
                },
            },
            // Accepted in every phase and ignored.
            Event::Decl(_)
            | Event::Comment(_)
            | Event::PI(_)
            | Event::Text(_)
            | Event::CData(_)
            | Event::GeneralRef(_) => {},
        }
    }

    Ok(actions)
}

fn maybe_push(
    actions: &mut Vec<Parsed>,
    namespace: &ResolveResult<'_>,
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
    limits: &mut Limits,
) -> Result<()> {
    let trigger = if is_drawingml_name(namespace, element.name(), b"hlinkClick") {
        Trigger::Click
    } else if is_drawingml_name(namespace, element.name(), b"hlinkHover")
        || is_drawingml_name(namespace, element.name(), b"hlinkMouseOver")
    {
        Trigger::Hover
    } else {
        return Ok(());
    };

    limits.add_action()?;
    let relationship_id = bounded_optional(attribute_value(element, b"id", decoder, resolver)?)?
        .filter(|value| !value.is_empty());
    let action = bounded_optional(unqualified_attribute_value(element, b"action", decoder)?)?;
    let tooltip = bounded_optional(unqualified_attribute_value(element, b"tooltip", decoder)?)?;
    let target_frame =
        bounded_optional(unqualified_attribute_value(element, b"tgtFrame", decoder)?)?;
    actions.push(Parsed {
        trigger,
        action,
        relationship_id,
        tooltip,
        target_frame,
    });
    Ok(())
}

fn validate_slide_root(
    namespace: &ResolveResult<'_>,
    name: QName<'_>,
    phase: RootPhase,
) -> Result<()> {
    match phase {
        RootPhase::Prolog if is_presentationml_name(namespace, name, b"sld") => Ok(()),
        RootPhase::Prolog | RootPhase::Open | RootPhase::Closed => Err(invalid(
            "slide XML must have one PresentationML sld root element",
        )),
    }
}

fn is_presentationml_name(
    namespace: &ResolveResult<'_>,
    name: QName<'_>,
    local_name: &[u8],
) -> bool {
    if name.local_name().as_ref() != local_name {
        return false;
    }
    match namespace {
        ResolveResult::Bound(Namespace(value)) => {
            *value == PRESENTATIONML_NAMESPACE || *value == STRICT_PRESENTATIONML_NAMESPACE
        },
        ResolveResult::Unknown(prefix) => prefix.as_slice() == b"p",
        ResolveResult::Unbound => false,
    }
}

fn bounded_optional(value: Option<String>) -> Result<Option<String>> {
    if let Some(value) = &value
        && value.len() > MAX_ATTRIBUTE_BYTES
    {
        return Err(limit(
            crate::error::CodecLimitKind::XmlAttributeBytes,
            value.len(),
            MAX_ATTRIBUTE_BYTES,
        ));
    }
    Ok(value)
}

fn increment_nodes(nodes: &mut usize) -> Result<()> {
    *nodes = nodes.checked_add(1).ok_or_else(|| {
        limit(
            crate::error::CodecLimitKind::XmlNodes,
            usize::MAX,
            MAX_XML_NODES,
        )
    })?;
    if *nodes > MAX_XML_NODES {
        return Err(limit(
            crate::error::CodecLimitKind::XmlNodes,
            *nodes,
            MAX_XML_NODES,
        ));
    }
    Ok(())
}

#[cfg(test)]
mod envelope_tests {
    //! What the action scan accepts before, around and after `p:sld`, and the exact error reported otherwise. Each expected
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

    /// The action scan of one slide.
    #[test]
    fn slide_action_scan_envelope() {
        let root = Root {
            open: r#"<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">"#,
            body: r"<p:cSld><p:spTree/></p:cSld>",
            close: r"</p:sld>",
            empty: r#"<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>"#,
        };
        assert_matrix(
            &root,
            &[
                "empty input: Invalid(\"unterminated or missing PresentationML slide root\")",
                "whitespace only: Invalid(\"unterminated or missing PresentationML slide root\")",
                "declaration only: Invalid(\"unterminated or missing PresentationML slide root\")",
                "empty-element root: Ok(0)",
                "start and end root: Ok(0)",
                "root then whitespace: Ok(0)",
                "root then comment: Ok(0)",
                "root then PI: Ok(0)",
                "root then text: Ok(0)",
                "root then second root: Invalid(\"slide XML must have one PresentationML sld root element\")",
                "root then foreign start: Invalid(\"slide XML must have one PresentationML sld root element\")",
                "root then foreign empty: Invalid(\"slide XML must have one PresentationML sld root element\")",
                "declaration after root: Ok(0)",
                "two declarations: Ok(0)",
                "comment before declaration: Ok(0)",
                "text before root: Ok(0)",
                "unterminated root: Invalid(\"unterminated or missing PresentationML slide root\")",
                "stray end tag: Xml(\"ill-formed document: close tag `</foreign>` does not match any open tag\")",
                "DOCTYPE: Invalid(\"slide XML must not contain a DTD\")",
                "top-level CDATA: Ok(0)",
                "declaration then root: Ok(0)",
                "whitespace before declaration: Ok(0)",
                "PI before declaration: Ok(0)",
                "comment then root: Ok(0)",
                "PI then root: Ok(0)",
                "root then text then second root: Invalid(\"slide XML must have one PresentationML sld root element\")",
                "empty-element root then root: Invalid(\"slide XML must have one PresentationML sld root element\")",
                "declaration after empty-element root: Ok(0)",
            ],
            |xml| {
                outcome(super::scan(xml, &mut super::Limits::default()), |values| {
                    values.len().to_string()
                })
            },
        );
    }
}
