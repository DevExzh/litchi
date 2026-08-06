//! ODS host navigation and bounded XML replacement.
//!
//! The calculation-settings element itself is parsed and written exclusively
//! by `litchi-odf-common::calculation`.  This codec only identifies its legal
//! ODS host and performs checked byte-range replacement around that element.

use litchi_core::{Error, Result};
use litchi_odf_common::calculation::{Settings, write};
use quick_xml::events::Event;
use quick_xml::name::{Namespace, NamespaceResolver, ResolveResult};
use quick_xml::reader::NsReader;

const OFFICE_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const TABLE_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:table:1.0";
const MAX_XML_BYTES: usize = 64 * 1024 * 1024;
const MAX_EVENTS: usize = 1_000_000;
const MAX_DEPTH: usize = 256;

/// A byte span for one XML element.
#[derive(Debug, Clone)]
pub(crate) struct Span {
    pub(crate) start: usize,
    pub(crate) tag_end: usize,
    pub(crate) end: usize,
    pub(crate) empty: bool,
    pub(crate) qname: String,
}

/// The legal spreadsheet host and its optional owned child.
#[derive(Debug, Clone)]
pub(crate) struct Location {
    pub(crate) spreadsheet: Span,
    pub(crate) calculation: Option<Span>,
    pub(crate) opaque: bool,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Kind {
    Document,
    Body,
    Spreadsheet,
    Calculation,
    Other,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum NamespaceKind {
    Office,
    Table,
    Other,
}

#[derive(Debug)]
struct OpenElement {
    kind: Kind,
    start: usize,
    tag_end: usize,
    qname: String,
}

/// Locate the direct ODS spreadsheet host and calculation-settings child.
pub(crate) fn locate(xml: &str) -> Result<Location> {
    if xml.len() > MAX_XML_BYTES {
        return Err(Error::InvalidFormat(
            "ODS calculation-settings source exceeds the size limit".to_string(),
        ));
    }

    let mut reader = NsReader::from_str(xml);
    reader.config_mut().check_end_names = true;
    let mut buffer = Vec::new();
    let mut stack = Vec::<OpenElement>::new();
    let mut spreadsheet = None;
    let mut calculation = None;
    let mut root_seen = false;
    let mut root_closed = false;
    let mut events = 0usize;
    let mut opaque = false;

    loop {
        events = events
            .checked_add(1)
            .ok_or_else(|| Error::InvalidFormat("XML event count overflow".to_string()))?;
        if events > MAX_EVENTS {
            return Err(Error::InvalidFormat(
                "ODS calculation-settings source exceeds the event limit".to_string(),
            ));
        }

        let event_start = position(&reader)?;
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::InvalidFormat(format!("invalid ODS content.xml: {error}")))?;
        let namespace = namespace_kind(&namespace);
        let event = event.into_owned();
        let event_end = position(&reader)?;

        match event {
            Event::Start(element) => {
                if stack.is_empty() {
                    if root_seen || root_closed {
                        return Err(Error::InvalidFormat(
                            "ODS content.xml has more than one root element".to_string(),
                        ));
                    }
                    root_seen = true;
                }
                if stack.len() >= MAX_DEPTH {
                    return Err(Error::InvalidFormat(
                        "ODS calculation-settings source exceeds the nesting limit".to_string(),
                    ));
                }
                let parent = stack.last().map(|open| open.kind);
                let kind = classify(parent, namespace, element.local_name().as_ref());
                if kind == Kind::Calculation
                    || stack.iter().any(|open| open.kind == Kind::Calculation)
                {
                    validate_attributes(&element, reader.resolver())?;
                }
                let qname = String::from_utf8(element.name().as_ref().to_vec()).map_err(|_| {
                    Error::InvalidFormat("ODS XML element name is not valid UTF-8".to_string())
                })?;
                stack.push(OpenElement {
                    kind,
                    start: event_start,
                    tag_end: event_end,
                    qname,
                });
            },
            Event::Empty(element) => {
                if stack.is_empty() {
                    if root_seen || root_closed {
                        return Err(Error::InvalidFormat(
                            "ODS content.xml has more than one root element".to_string(),
                        ));
                    }
                    root_seen = true;
                    root_closed = true;
                }
                if stack.len() >= MAX_DEPTH {
                    return Err(Error::InvalidFormat(
                        "ODS calculation-settings source exceeds the nesting limit".to_string(),
                    ));
                }
                let parent = stack.last().map(|open| open.kind);
                let kind = classify(parent, namespace, element.local_name().as_ref());
                if kind == Kind::Calculation
                    || stack.iter().any(|open| open.kind == Kind::Calculation)
                {
                    validate_attributes(&element, reader.resolver())?;
                }
                let qname = String::from_utf8(element.name().as_ref().to_vec()).map_err(|_| {
                    Error::InvalidFormat("ODS XML element name is not valid UTF-8".to_string())
                })?;
                let span = Span {
                    start: event_start,
                    tag_end: event_end,
                    end: event_end,
                    empty: true,
                    qname,
                };
                record(kind, span, &mut spreadsheet, &mut calculation)?;
            },
            Event::End(_) => {
                let open = stack.pop().ok_or_else(|| {
                    Error::InvalidFormat("unbalanced ODS content.xml elements".to_string())
                })?;
                let span = Span {
                    start: open.start,
                    tag_end: open.tag_end,
                    end: event_end,
                    empty: false,
                    qname: open.qname,
                };
                record(open.kind, span, &mut spreadsheet, &mut calculation)?;
                if stack.is_empty() {
                    root_closed = true;
                }
            },
            Event::Comment(_) | Event::PI(_)
                if stack.iter().any(|open| open.kind == Kind::Calculation) =>
            {
                opaque = true;
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }

    if !root_seen || !root_closed || !stack.is_empty() {
        return Err(Error::InvalidFormat(
            "incomplete ODS content.xml document".to_string(),
        ));
    }
    let spreadsheet = spreadsheet.ok_or_else(|| {
        Error::InvalidFormat("ODS content.xml has no office:spreadsheet host".to_string())
    })?;
    Ok(Location {
        spreadsheet,
        calculation,
        opaque,
    })
}

fn classify(parent: Option<Kind>, namespace: NamespaceKind, local: &[u8]) -> Kind {
    if parent.is_none() && is_element(namespace, NamespaceKind::Office, local, b"document-content")
    {
        Kind::Document
    } else if parent == Some(Kind::Document)
        && is_element(namespace, NamespaceKind::Office, local, b"body")
    {
        Kind::Body
    } else if parent == Some(Kind::Body)
        && is_element(namespace, NamespaceKind::Office, local, b"spreadsheet")
    {
        Kind::Spreadsheet
    } else if parent == Some(Kind::Spreadsheet)
        && is_element(
            namespace,
            NamespaceKind::Table,
            local,
            b"calculation-settings",
        )
    {
        Kind::Calculation
    } else {
        Kind::Other
    }
}

fn record(
    kind: Kind,
    span: Span,
    spreadsheet: &mut Option<Span>,
    calculation: &mut Option<Span>,
) -> Result<()> {
    match kind {
        Kind::Spreadsheet => {
            if spreadsheet.is_some() {
                return Err(Error::InvalidFormat(
                    "ODS content.xml has more than one office:spreadsheet host".to_string(),
                ));
            }
            *spreadsheet = Some(span);
        },
        Kind::Calculation => {
            if calculation.is_some() {
                return Err(Error::InvalidFormat(
                    "ODS content.xml has duplicate table:calculation-settings".to_string(),
                ));
            }
            *calculation = Some(span);
        },
        _ => {},
    }
    Ok(())
}

/// Replace only the owned calculation-settings element, or insert/remove it
/// as a direct child of the spreadsheet host.
pub(crate) fn replace(
    source: &str,
    location: &Location,
    settings: Option<&Settings>,
) -> Result<String> {
    if location.opaque && settings.is_some() {
        return Err(Error::InvalidFormat(
            "ODS calculation-settings contains opaque markup that cannot be safely rewritten"
                .to_string(),
        ));
    }
    let fragment = settings.map(render).transpose()?;
    if let Some(calculation) = &location.calculation {
        return splice(
            source,
            calculation.start,
            calculation.end,
            fragment.as_deref().unwrap_or_default(),
        );
    }
    let Some(fragment) = fragment else {
        return Ok(source.to_owned());
    };
    let spreadsheet = &location.spreadsheet;
    if spreadsheet.empty {
        let opening = source
            .get(spreadsheet.start..spreadsheet.tag_end)
            .ok_or_else(|| Error::InvalidFormat("invalid spreadsheet XML span".to_string()))?;
        let opening = opening.strip_suffix("/>").ok_or_else(|| {
            Error::InvalidFormat("empty spreadsheet XML span has no close token".to_string())
        })?;
        let mut expanded = String::with_capacity(opening.len() + fragment.len() + 32);
        expanded.push_str(opening);
        expanded.push('>');
        expanded.push_str(&fragment);
        expanded.push_str("</");
        expanded.push_str(&spreadsheet.qname);
        expanded.push('>');
        splice(source, spreadsheet.start, spreadsheet.end, &expanded)
    } else {
        splice(source, spreadsheet.tag_end, spreadsheet.tag_end, &fragment)
    }
}

fn render(settings: &Settings) -> Result<String> {
    let mut output = String::new();
    write(&mut output, Some(settings))?;
    // The shared writer intentionally emits a compact semantic fragment.  A
    // local declaration makes the fragment valid even when the source uses a
    // different table prefix or has no inherited `table` binding.
    const START: &str = "<table:calculation-settings";
    let declaration = format!(
        " xmlns:table=\"{}\"",
        String::from_utf8_lossy(TABLE_NAMESPACE)
    );
    let position = START.len();
    if !output.starts_with(START) || output.len() < position {
        return Err(Error::InvalidFormat(
            "shared calculation writer emitted an unexpected root".to_string(),
        ));
    }
    output.insert_str(position, &declaration);
    Ok(output)
}

fn splice(source: &str, start: usize, end: usize, replacement: &str) -> Result<String> {
    if start > end
        || end > source.len()
        || !source.is_char_boundary(start)
        || !source.is_char_boundary(end)
    {
        return Err(Error::InvalidFormat(
            "invalid ODS calculation-settings XML span".to_string(),
        ));
    }
    let mut output = String::with_capacity(source.len() - (end - start) + replacement.len());
    output.push_str(&source[..start]);
    output.push_str(replacement);
    output.push_str(&source[end..]);
    Ok(output)
}

fn position(reader: &NsReader<&[u8]>) -> Result<usize> {
    usize::try_from(reader.buffer_position())
        .map_err(|_| Error::InvalidFormat("ODS XML position overflows usize".to_string()))
}

fn is_element(
    namespace: NamespaceKind,
    expected_namespace: NamespaceKind,
    local: &[u8],
    expected_local: &[u8],
) -> bool {
    namespace == expected_namespace && local == expected_local
}

fn validate_attributes(
    element: &quick_xml::events::BytesStart<'_>,
    resolver: &NamespaceResolver,
) -> Result<()> {
    let local = element.local_name();
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| {
            Error::InvalidFormat(format!(
                "invalid ODS calculation-settings attribute: {error}"
            ))
        })?;
        if attribute.value.len() > 64 * 1024 {
            return Err(Error::InvalidFormat(
                "ODS calculation-settings attribute exceeds the size limit".to_string(),
            ));
        }
        let key = attribute.key.as_ref();
        if key == b"xmlns"
            || attribute
                .key
                .prefix()
                .is_some_and(|prefix| prefix.as_ref() == b"xmlns")
        {
            continue;
        }
        let (namespace, name) = resolver.resolve_attribute(attribute.key);
        let allowed = matches!(namespace, ResolveResult::Bound(Namespace(uri)) if *uri == *TABLE_NAMESPACE)
            && allowed_attribute(local.as_ref(), name.as_ref());
        if !allowed {
            return Err(Error::InvalidFormat(
                "ODS calculation-settings contains an unknown attribute; refusing lossy edit"
                    .to_string(),
            ));
        }
    }
    Ok(())
}

fn allowed_attribute(element: &[u8], attribute: &[u8]) -> bool {
    match element {
        b"calculation-settings" => matches!(
            attribute,
            b"case-sensitive"
                | b"precision-as-shown"
                | b"search-criteria-must-apply-to-whole-cell"
                | b"automatic-find-labels"
                | b"use-regular-expressions"
                | b"use-wildcards"
                | b"null-year"
        ),
        b"null-date" => matches!(attribute, b"value-type" | b"date-value"),
        b"iteration" => matches!(attribute, b"status" | b"steps" | b"maximum-difference"),
        _ => false,
    }
}

fn namespace_kind(namespace: &ResolveResult<'_>) -> NamespaceKind {
    match namespace {
        ResolveResult::Bound(Namespace(uri)) if *uri == OFFICE_NAMESPACE => NamespaceKind::Office,
        ResolveResult::Bound(Namespace(uri)) if *uri == TABLE_NAMESPACE => NamespaceKind::Table,
        _ => NamespaceKind::Other,
    }
}
