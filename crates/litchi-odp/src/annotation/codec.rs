//! Bounded presentation annotation scanning and span-preserving XML edits.

use super::model::{Anchor, Info, Position};
use super::package::rebuild;
use super::validation::{
    MAX_ANNOTATIONS, MAX_EVENTS, MAX_FRAGMENT_BYTES, MAX_XML_BYTES, bounds, invalid,
    validate_anchor, validate_annotation,
};
use crate::core::OwnedPackage;
use litchi_core::{Error, Result};
use litchi_odf_common::annotation::{Annotation, Builder};
use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use std::collections::{BTreeMap, HashMap};

const OFFICE_NS: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const DRAW_NS: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
const OFFICE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const TEXT: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const TABLE: &str = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
const DRAW: &str = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
const SVG: &str = "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0";
const DC: &str = "http://purl.org/dc/elements/1.1/";
const META: &str = "urn:oasis:names:tc:opendocument:xmlns:meta:1.0";
const XLINK: &str = "http://www.w3.org/1999/xlink";
const LOEXT: &str = "urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0";

#[derive(Clone)]
struct Span {
    start: usize,
    end: usize,
    close_start: Option<usize>,
    qname: String,
}

#[derive(Clone)]
struct Site {
    position: Position,
    span: Span,
}

struct Record {
    span: Span,
    annotation: Annotation,
    position: Position,
}

struct Scan {
    records: Vec<Record>,
    sites: Vec<Site>,
}

enum FrameKind {
    Page { site: usize, index: usize },
    Shape { site: usize },
    Annotation { record: usize },
    Other,
}

struct Frame {
    kind: FrameKind,
    namespace_changes: Vec<(String, Option<String>)>,
}

struct ActiveBuilder {
    record: usize,
    builder: Builder,
}

/// Read annotations in document order.
pub(crate) fn annotations(content: &str) -> Result<Vec<Info>> {
    let scan = scan(content)?;
    scan.records
        .into_iter()
        .enumerate()
        .map(|(index, record)| {
            Ok(Info {
                index,
                annotation: record.annotation,
                anchor: anchor(record.position),
            })
        })
        .collect()
}

/// Find a named annotation without exposing physical XML locations.
pub(crate) fn find(content: &str, name: &str) -> Result<Option<Info>> {
    if name.is_empty() {
        return invalid("annotation name cannot be empty");
    }
    Ok(annotations(content)?
        .into_iter()
        .find(|item| item.annotation.name() == Some(name)))
}

/// Add an annotation at a validated page or shape anchor.
pub(crate) fn add(
    package: &OwnedPackage,
    content: &str,
    anchor: &Anchor,
    annotation: &Annotation,
) -> Result<(Vec<u8>, usize)> {
    validate_anchor(anchor)?;
    validate_annotation(annotation)?;
    let parsed = scan(content)?;
    validate_new_name(&parsed, annotation.name())?;
    let site = site_for(&parsed, anchor.position())?;
    let fragment = serialize(annotation)?;
    let updated = insert_child(content, site, &fragment)?;
    let index = parsed
        .records
        .iter()
        .filter(|record| record.span.start < site.span.start)
        .count();
    scan(&updated)?;
    rebuild(package, &updated).map(|bytes| (bytes, index))
}

/// Replace one annotation body while retaining its page/shape anchor.
pub(crate) fn replace(
    package: &OwnedPackage,
    content: &str,
    index: usize,
    annotation: &Annotation,
) -> Result<Vec<u8>> {
    validate_annotation(annotation)?;
    let parsed = scan(content)?;
    let record = parsed
        .records
        .get(index)
        .ok_or_else(|| bounds(index, parsed.records.len()))?;
    if record.annotation == *annotation {
        return Ok(package.as_bytes().to_vec());
    }
    validate_new_name_except(&parsed, annotation.name(), index)?;
    let updated = apply_edits(
        content,
        vec![Edit {
            start: record.span.start,
            end: record.span.end,
            replacement: serialize(annotation)?,
        }],
    )?;
    scan(&updated)?;
    rebuild(package, &updated)
}

/// Remove one annotation atomically.
pub(crate) fn remove(package: &OwnedPackage, content: &str, index: usize) -> Result<Vec<u8>> {
    let parsed = scan(content)?;
    let record = parsed
        .records
        .get(index)
        .ok_or_else(|| bounds(index, parsed.records.len()))?;
    let updated = apply_edits(
        content,
        vec![Edit {
            start: record.span.start,
            end: record.span.end,
            replacement: String::new(),
        }],
    )?;
    scan(&updated)?;
    rebuild(package, &updated)
}

fn anchor(position: Position) -> Anchor {
    Anchor::from_position(position)
}

fn scan(xml: &str) -> Result<Scan> {
    if xml.len() > MAX_XML_BYTES {
        return invalid("presentation content.xml exceeds the annotation size limit");
    }

    let mut reader = NsReader::from_str(xml);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = true;
    let mut buffer = Vec::new();
    let mut frames = Vec::<Frame>::new();
    let mut namespaces = BTreeMap::new();
    let mut builders = Vec::<ActiveBuilder>::new();
    let mut records = Vec::<Record>::new();
    let mut sites = Vec::<Site>::new();
    let mut next_page = 0usize;
    let mut events = 0usize;

    loop {
        let start = position(&reader)?;
        let (namespace, event) = {
            let (resolved, event) = reader
                .read_resolved_event_into(&mut buffer)
                .map_err(|error| invalid_error(format!("invalid presentation XML: {error}")))?;
            (namespace(&resolved), event.into_owned())
        };
        let end = position(&reader)?;
        events = events
            .checked_add(1)
            .ok_or_else(|| invalid_error("presentation XML event count overflow"))?;
        if events > MAX_EVENTS {
            return invalid("presentation XML exceeds the event limit");
        }
        match event {
            Event::Start(element) => {
                let changes =
                    apply_namespace_declarations(&element, reader.decoder(), &mut namespaces)?;
                let local = element.local_name();
                let top_level_annotation = namespace == NamespaceKind::Office
                    && local.as_ref() == b"annotation"
                    && builders.is_empty();

                if top_level_annotation {
                    if records.len() >= MAX_ANNOTATIONS {
                        return invalid("presentation exceeds the annotation limit");
                    }
                    let position = current_position(&frames, &sites)?;
                    let record = records.len();
                    records.push(Record {
                        span: Span {
                            start,
                            end: 0,
                            close_start: None,
                            qname: qname(element.name().as_ref())?,
                        },
                        annotation: Annotation::default(),
                        position,
                    });
                    builders.push(ActiveBuilder {
                        record,
                        builder: Builder::new(&element, reader.decoder(), namespaces.clone())?,
                    });
                    frames.push(Frame {
                        kind: FrameKind::Annotation { record },
                        namespace_changes: changes,
                    });
                } else {
                    for active in &mut builders {
                        active.builder.start(&element, reader.decoder())?;
                    }
                    let kind = if builders.is_empty() {
                        structural_kind(
                            &reader,
                            &element,
                            namespace,
                            start,
                            end,
                            false,
                            &mut sites,
                            &mut next_page,
                            &frames,
                        )?
                    } else {
                        FrameKind::Other
                    };
                    frames.push(Frame {
                        kind,
                        namespace_changes: changes,
                    });
                }
                if frames.len() > 256 {
                    return invalid("presentation XML nesting exceeds the safety limit");
                }
            },
            Event::Empty(element) => {
                let changes =
                    apply_namespace_declarations(&element, reader.decoder(), &mut namespaces)?;
                let local = element.local_name();
                let top_level_annotation = namespace == NamespaceKind::Office
                    && local.as_ref() == b"annotation"
                    && builders.is_empty();
                if top_level_annotation {
                    if records.len() >= MAX_ANNOTATIONS {
                        return invalid("presentation exceeds the annotation limit");
                    }
                    let position = current_position(&frames, &sites)?;
                    let builder = Builder::new(&element, reader.decoder(), namespaces.clone())?;
                    records.push(Record {
                        span: Span {
                            start,
                            end,
                            close_start: None,
                            qname: qname(element.name().as_ref())?,
                        },
                        annotation: builder.finish()?,
                        position,
                    });
                } else {
                    for active in &mut builders {
                        active.builder.empty(&element, reader.decoder())?;
                    }
                    if builders.is_empty() {
                        structural_kind(
                            &reader,
                            &element,
                            namespace,
                            start,
                            end,
                            true,
                            &mut sites,
                            &mut next_page,
                            &frames,
                        )?;
                    }
                }
                restore_namespaces(changes, &mut namespaces);
            },
            Event::End(_element) => {
                let frame = frames
                    .pop()
                    .ok_or_else(|| invalid_error("presentation XML depth underflow"))?;
                match frame.kind {
                    FrameKind::Annotation { record } => {
                        let active = builders
                            .pop()
                            .ok_or_else(|| invalid_error("missing annotation builder"))?;
                        if active.record != record {
                            return invalid("mismatched presentation annotation builder");
                        }
                        records[record].span.end = end;
                        records[record].span.close_start = Some(start);
                        records[record].annotation = active.builder.finish()?;
                    },
                    kind => {
                        for active in &mut builders {
                            active.builder.end_element()?;
                        }
                        finish_site(&kind, start, end, &mut sites);
                    },
                }
                restore_namespaces(frame.namespace_changes, &mut namespaces);
            },
            Event::Text(value) => {
                for active in &mut builders {
                    active.builder.text(&value)?;
                }
            },
            Event::CData(value) => {
                for active in &mut builders {
                    active.builder.cdata(&value)?;
                }
            },
            Event::GeneralRef(value) => {
                for active in &mut builders {
                    active.builder.reference(&value)?;
                }
            },
            Event::DocType(_) => return invalid("DOCTYPE is not allowed in presentation XML"),
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }

    if !frames.is_empty() || !builders.is_empty() {
        return invalid("unterminated presentation XML");
    }
    validate_record_names(&records)?;
    sites.sort_by_key(|site| site.span.start);
    Ok(Scan { records, sites })
}

#[allow(clippy::too_many_arguments)]
fn structural_kind(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    namespace: NamespaceKind,
    start: usize,
    end: usize,
    empty: bool,
    sites: &mut Vec<Site>,
    next_page: &mut usize,
    frames: &[Frame],
) -> Result<FrameKind> {
    let local = element.local_name();
    let local = local.as_ref();
    if namespace == NamespaceKind::Draw && local == b"page" {
        let page = *next_page;
        *next_page = next_page
            .checked_add(1)
            .ok_or_else(|| invalid_error("presentation page index overflow"))?;
        let site = push_site(
            sites,
            Position::Page { index: page },
            start,
            end,
            empty,
            element,
        )?;
        return Ok(FrameKind::Page { site, index: page });
    }
    if namespace == NamespaceKind::Draw
        && local != b"page"
        && current_page(frames).is_some()
        && let Some(name) = attribute(reader, element, DRAW_NS, b"name")?
    {
        let page_index = current_page(frames)
            .ok_or_else(|| invalid_error("presentation shape is outside a draw:page"))?;
        let site = push_site(
            sites,
            Position::Shape { page_index, name },
            start,
            end,
            empty,
            element,
        )?;
        return Ok(FrameKind::Shape { site });
    }
    Ok(FrameKind::Other)
}

fn current_page(frames: &[Frame]) -> Option<usize> {
    frames.iter().rev().find_map(|frame| {
        if let FrameKind::Page { index, .. } = frame.kind {
            Some(index)
        } else {
            None
        }
    })
}

fn current_position(frames: &[Frame], sites: &[Site]) -> Result<Position> {
    for frame in frames.iter().rev() {
        match frame.kind {
            FrameKind::Shape { site } | FrameKind::Page { site, .. } => {
                return sites
                    .get(site)
                    .map(|site| site.position.clone())
                    .ok_or_else(|| invalid_error("presentation annotation site disappeared"));
            },
            FrameKind::Annotation { .. } | FrameKind::Other => {},
        }
    }
    invalid("presentation annotation is outside a draw:page")
}

fn push_site(
    sites: &mut Vec<Site>,
    position: Position,
    start: usize,
    end: usize,
    empty: bool,
    element: &BytesStart<'_>,
) -> Result<usize> {
    let index = sites.len();
    sites.push(Site {
        position,
        span: Span {
            start,
            end: if empty { end } else { 0 },
            close_start: None,
            qname: qname(element.name().as_ref())?,
        },
    });
    Ok(index)
}

fn finish_site(kind: &FrameKind, close_start: usize, end: usize, sites: &mut [Site]) {
    let site = match kind {
        FrameKind::Page { site, .. } | FrameKind::Shape { site } => Some(*site),
        FrameKind::Annotation { .. } | FrameKind::Other => None,
    };
    if let Some(site) = site
        && let Some(target) = sites.get_mut(site)
    {
        target.span.end = end;
        target.span.close_start = Some(close_start);
    }
}

fn site_for<'a>(scan: &'a Scan, position: &Position) -> Result<&'a Site> {
    let mut matches = scan.sites.iter().filter(|site| &site.position == position);
    let site = matches.next().ok_or_else(|| {
        invalid_error(format!(
            "presentation annotation anchor {position:?} was not found"
        ))
    })?;
    if matches.next().is_some() {
        return invalid("presentation annotation anchor is ambiguous");
    }
    Ok(site)
}

fn validate_record_names(records: &[Record]) -> Result<()> {
    let mut names = HashMap::new();
    for record in records {
        if let Some(name) = record.annotation.name()
            && names.insert(name.to_string(), ()).is_some()
        {
            return invalid(format!("duplicate annotation name '{name}'"));
        }
    }
    Ok(())
}

fn validate_new_name(scan: &Scan, name: Option<&str>) -> Result<()> {
    validate_new_name_except(scan, name, usize::MAX)
}

fn validate_new_name_except(scan: &Scan, name: Option<&str>, except: usize) -> Result<()> {
    let Some(name) = name else { return Ok(()) };
    if name.is_empty() {
        return invalid("annotation office:name cannot be empty");
    }
    if scan
        .records
        .iter()
        .enumerate()
        .any(|(index, record)| index != except && record.annotation.name() == Some(name))
    {
        return invalid(format!("duplicate annotation name '{name}'"));
    }
    Ok(())
}

pub(super) fn serialize(annotation: &Annotation) -> Result<String> {
    let mut annotation = annotation.clone();
    for (prefix, uri) in [
        ("office", OFFICE),
        ("text", TEXT),
        ("table", TABLE),
        ("draw", DRAW),
        ("svg", SVG),
        ("dc", DC),
        ("meta", META),
        ("xlink", XLINK),
        ("loext", LOEXT),
    ] {
        annotation.set_namespace(prefix, uri)?;
    }
    validate_annotation(&annotation)?;
    let mut output = String::new();
    annotation.write_xml(&mut output);
    if output.len() > MAX_FRAGMENT_BYTES {
        return invalid("annotation XML exceeds the fragment size limit");
    }
    Ok(output)
}

fn insert_child(xml: &str, site: &Site, fragment: &str) -> Result<String> {
    let edit = if let Some(close) = site.span.close_start {
        Edit {
            start: close,
            end: close,
            replacement: fragment.to_string(),
        }
    } else {
        let raw = xml
            .get(site.span.start..site.span.end)
            .ok_or_else(|| invalid_error("invalid empty presentation anchor span"))?;
        let slash = raw
            .rfind("/>")
            .ok_or_else(|| invalid_error("invalid empty presentation anchor"))?;
        Edit {
            start: site.span.start,
            end: site.span.end,
            replacement: format!("{}>{}</{}>", &raw[..slash], fragment, site.span.qname),
        }
    };
    apply_edits(xml, vec![edit])
}

pub(super) struct Edit {
    pub(super) start: usize,
    pub(super) end: usize,
    pub(super) replacement: String,
}

pub(super) fn apply_edits(xml: &str, mut edits: Vec<Edit>) -> Result<String> {
    edits.sort_by(|left, right| right.start.cmp(&left.start).then(right.end.cmp(&left.end)));
    let mut output = xml.to_string();
    let mut previous_start = xml.len();
    for edit in edits {
        if edit.start > edit.end
            || edit.end > xml.len()
            || edit.end > previous_start
            || !xml.is_char_boundary(edit.start)
            || !xml.is_char_boundary(edit.end)
        {
            return invalid("overlapping or invalid presentation XML edit");
        }
        output.replace_range(edit.start..edit.end, &edit.replacement);
        previous_start = edit.start;
    }
    Ok(output)
}

fn apply_namespace_declarations(
    element: &BytesStart<'_>,
    decoder: quick_xml::Decoder,
    namespaces: &mut BTreeMap<String, String>,
) -> Result<Vec<(String, Option<String>)>> {
    let mut changes = Vec::new();
    for attribute in element.attributes() {
        let attribute = attribute
            .map_err(|error| invalid_error(format!("invalid presentation namespace: {error}")))?;
        let raw = qname(attribute.key.as_ref())?;
        let prefix = if raw == "xmlns" {
            Some(String::new())
        } else {
            raw.strip_prefix("xmlns:").map(str::to_string)
        };
        let Some(prefix) = prefix else { continue };
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
            .map_err(|error| {
                invalid_error(format!("invalid presentation namespace value: {error}"))
            })?
            .into_owned();
        let previous = namespaces.insert(prefix.clone(), value);
        changes.push((prefix, previous));
    }
    Ok(changes)
}

fn restore_namespaces(
    changes: Vec<(String, Option<String>)>,
    namespaces: &mut BTreeMap<String, String>,
) {
    for (prefix, previous) in changes.into_iter().rev() {
        if let Some(previous) = previous {
            namespaces.insert(prefix, previous);
        } else {
            namespaces.remove(&prefix);
        }
    }
}

fn attribute(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    expected_namespace: &[u8],
    expected_local: &[u8],
) -> Result<Option<String>> {
    let mut result = None;
    for raw in element.attributes() {
        let raw =
            raw.map_err(|error| invalid_error(format!("invalid presentation attribute: {error}")))?;
        let (namespace, local) = reader.resolver().resolve_attribute(raw.key);
        if matches!(namespace, ResolveResult::Bound(Namespace(value)) if *value == *expected_namespace)
            && local.as_ref() == expected_local
        {
            if result.is_some() {
                return invalid("duplicate expanded presentation anchor attribute");
            }
            result = Some(
                raw.decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                    .map_err(|error| {
                        invalid_error(format!("invalid presentation attribute value: {error}"))
                    })?
                    .into_owned(),
            );
        }
    }
    Ok(result)
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum NamespaceKind {
    Office,
    Draw,
    Other,
}

fn namespace(value: &ResolveResult<'_>) -> NamespaceKind {
    match value {
        ResolveResult::Bound(Namespace(value)) if *value == OFFICE_NS => NamespaceKind::Office,
        ResolveResult::Bound(Namespace(value)) if *value == DRAW_NS => NamespaceKind::Draw,
        _ => NamespaceKind::Other,
    }
}

fn position(reader: &NsReader<&[u8]>) -> Result<usize> {
    usize::try_from(reader.buffer_position())
        .map_err(|_| invalid_error("presentation XML position overflow"))
}

fn qname(value: &[u8]) -> Result<String> {
    std::str::from_utf8(value)
        .map(str::to_string)
        .map_err(|_| invalid_error("invalid presentation qualified name"))
}

fn invalid_error(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}
