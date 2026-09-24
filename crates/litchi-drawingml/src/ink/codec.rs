//! Bounded namespace-aware InkML XML reader and source-preserving writer.

use std::{fmt, io::Write, sync::Arc};

use litchi_ooxml_common::xml_name::is_qualified_name;
use quick_xml::{
    XmlVersion,
    events::{BytesDecl, BytesRef, BytesStart, Event},
    name::{Namespace, QName, ResolveResult},
    reader::NsReader,
};

use crate::{Error, Result};

use super::model::{
    BrushProperty, BrushPropertyName, Context, ContextKind, Document, Guid, Metadata, SemanticType,
    SourceSpan, Trace, ValueError,
};
use super::{
    INKML_NAMESPACE, MAX_ATTRIBUTE_VALUE_BYTES, MAX_BRUSH_PROPERTIES, MAX_CONTEXTS, MAX_DEPTH,
    MAX_NODES, MAX_SOURCE_BYTES, MAX_TOKEN_BYTES, MAX_TRACES, NAMESPACE,
};

const MAX_NAMESPACE_DECLARATIONS: usize = 256;
const MAX_ATTRIBUTES_PER_ELEMENT: usize = 256;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum NamespaceId {
    Inkml,
    Msink,
    Other,
    Unbound,
    Unknown,
}

/// Read one complete InkML content part.
///
/// The bounded scanner also accepts the documented legacy fragment form: an
/// unresolved expected `inkml` root prefix with no declaration, with the same
/// expected aliases allowed in its descendants. Complete package hosts must
/// require a namespace-bound InkML root before calling this reader.
///
/// # Errors
///
/// Returns an error for malformed XML, a wrong root, invalid typed metadata,
/// or an exhausted resource budget.
pub fn read(xml: &[u8]) -> Result<Document> {
    enforce_source_limit(xml.len())?;
    let metadata = scan(xml, Projection::all())?.into_metadata();
    let source = copy_source(xml)?;
    Document::new(source, metadata).map_err(value_error)
}

/// Read one complete InkML content part from an already shared immutable
/// payload.
///
/// The bounded scanner also accepts the documented legacy fragment form: an
/// unresolved expected `inkml` root prefix with no declaration, with the same
/// expected aliases allowed in its descendants. Complete package hosts must
/// require a namespace-bound InkML root before calling this reader.
///
/// The caller's `Arc<Vec<u8>>` allocation is retained by the returned
/// document after bounded validation succeeds. The input allocation is not
/// copied.
///
/// # Errors
///
/// Returns an error for malformed XML, a wrong root, invalid typed metadata,
/// or an exhausted resource budget. The source allocation is retained only
/// on success.
pub fn read_shared(xml: Arc<Vec<u8>>) -> Result<Document> {
    enforce_source_limit(xml.len())?;
    let metadata = scan(xml.as_slice(), Projection::all())?.into_metadata();
    Document::new(xml, metadata).map_err(value_error)
}

/// Read bounded typed InkML metadata without retaining the input XML bytes.
///
/// The bounded scanner also accepts the documented legacy fragment form: an
/// unresolved expected `inkml` root prefix with no declaration, with the same
/// expected aliases allowed in its descendants. Complete package hosts must
/// require a namespace-bound InkML root before calling this reader.
///
/// # Errors
///
/// Returns an error for malformed XML, a wrong root, invalid typed metadata,
/// or an exhausted resource budget. The input slice is borrowed only during
/// validation and is never retained by the returned projection.
pub fn read_metadata(xml: &[u8]) -> Result<Metadata> {
    enforce_source_limit(xml.len())?;
    Ok(scan(xml, Projection::all())?.into_metadata())
}

/// Read a source-backed document while retaining only the semantic elements
/// selected by the owning package profile.
///
/// Every selected span must be a source span produced by the same XML source,
/// listed in strictly increasing order for its semantic kind, and must match
/// exactly one generic InkML semantic element. The scanner still validates all
/// XML, namespace declarations, attributes, nodes, depth, text, and resource
/// limits; selection only suppresses typed decoding for unselected semantic
/// elements. An unmatched or out-of-order selection is rejected.
pub fn read_shared_with_source_spans(
    xml: Arc<Vec<u8>>,
    contexts: &[SourceSpan],
    traces: &[SourceSpan],
    brush_properties: &[SourceSpan],
    links: &[SourceSpan],
) -> Result<Document> {
    enforce_source_limit(xml.len())?;
    let projection = Projection::filtered(xml.len(), contexts, traces, brush_properties, links)?;
    let metadata = scan(xml.as_slice(), projection)?.into_metadata();
    Document::new(xml, metadata).map_err(value_error)
}

/// Read bounded typed InkML metadata while retaining only the semantic
/// elements selected by the owning package profile.
///
/// The selected spans must be source ordered, in bounds, and must each match
/// exactly one semantic element. All XML remains subject to the generic
/// scanner's structural and resource limits; unselected typed elements are
/// skipped before their typed attributes are decoded.
pub fn read_metadata_with_source_spans(
    xml: &[u8],
    contexts: &[SourceSpan],
    traces: &[SourceSpan],
    brush_properties: &[SourceSpan],
    links: &[SourceSpan],
) -> Result<Metadata> {
    enforce_source_limit(xml.len())?;
    let projection = Projection::filtered(xml.len(), contexts, traces, brush_properties, links)?;
    Ok(scan(xml, projection)?.into_metadata())
}

#[derive(Debug)]
struct Parsed {
    contexts: Vec<Context>,
    traces: Vec<Trace>,
    brush_properties: Vec<BrushProperty>,
}

impl Parsed {
    fn into_metadata(self) -> Metadata {
        Metadata::from_parts(self.contexts, self.traces, self.brush_properties)
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum SemanticKind {
    Context,
    Trace,
    BrushProperty,
    Link,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct SelectedSpan {
    kind: SemanticKind,
    expected: SourceSpan,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum ProjectionToken {
    All,
    Selected(SelectedSpan),
}

struct Projection<'a> {
    all: bool,
    contexts: &'a [SourceSpan],
    traces: &'a [SourceSpan],
    brush_properties: &'a [SourceSpan],
    links: &'a [SourceSpan],
    matched_contexts: usize,
    matched_traces: usize,
    matched_brush_properties: usize,
    matched_links: usize,
}

impl<'a> Projection<'a> {
    fn all() -> Self {
        Self {
            all: true,
            contexts: &[],
            traces: &[],
            brush_properties: &[],
            links: &[],
            matched_contexts: 0,
            matched_traces: 0,
            matched_brush_properties: 0,
            matched_links: 0,
        }
    }

    fn filtered(
        source_len: usize,
        contexts: &'a [SourceSpan],
        traces: &'a [SourceSpan],
        brush_properties: &'a [SourceSpan],
        links: &'a [SourceSpan],
    ) -> Result<Self> {
        validate_selected_spans(source_len, "InkML context projection", contexts)?;
        validate_selected_spans(source_len, "InkML trace projection", traces)?;
        validate_selected_spans(
            source_len,
            "InkML brush-property projection",
            brush_properties,
        )?;
        validate_selected_spans(source_len, "InkML link projection", links)?;
        Ok(Self {
            all: false,
            contexts,
            traces,
            brush_properties,
            links,
            matched_contexts: 0,
            matched_traces: 0,
            matched_brush_properties: 0,
            matched_links: 0,
        })
    }

    fn select(&self, kind: SemanticKind, start: usize) -> Option<ProjectionToken> {
        if self.all {
            return Some(ProjectionToken::All);
        }
        let spans = match kind {
            SemanticKind::Context => self.contexts,
            SemanticKind::Trace => self.traces,
            SemanticKind::BrushProperty => self.brush_properties,
            SemanticKind::Link => self.links,
        };
        spans
            .binary_search_by_key(&start, |span| span.start())
            .ok()
            .map(|index| {
                ProjectionToken::Selected(SelectedSpan {
                    kind,
                    expected: spans[index],
                })
            })
    }

    fn complete(&mut self, token: ProjectionToken, actual: SourceSpan) -> Result<()> {
        let ProjectionToken::Selected(selected) = token else {
            return Ok(());
        };
        if selected.expected != actual {
            return Err(invalid(
                "InkML semantic projection span does not match its source element",
            ));
        }
        let matched = match selected.kind {
            SemanticKind::Context => &mut self.matched_contexts,
            SemanticKind::Trace => &mut self.matched_traces,
            SemanticKind::BrushProperty => &mut self.matched_brush_properties,
            SemanticKind::Link => &mut self.matched_links,
        };
        *matched = matched
            .checked_add(1)
            .ok_or_else(|| invalid("InkML semantic projection count overflowed"))?;
        Ok(())
    }

    fn finish(self) -> Result<()> {
        if self.all
            || (self.matched_contexts == self.contexts.len()
                && self.matched_traces == self.traces.len()
                && self.matched_brush_properties == self.brush_properties.len()
                && self.matched_links == self.links.len())
        {
            Ok(())
        } else {
            Err(invalid(
                "InkML semantic projection contains an unmatched source span",
            ))
        }
    }
}

fn validate_selected_spans(
    source_len: usize,
    resource: &'static str,
    spans: &[SourceSpan],
) -> Result<()> {
    let mut previous = None;
    for &span in spans {
        if span.start() > span.end() || span.end() > source_len {
            return Err(invalid(format!("{resource} span is outside the source")));
        }
        if previous.is_some_and(|previous: SourceSpan| previous.start() >= span.start()) {
            return Err(invalid(format!(
                "{resource} spans are not strictly source ordered"
            )));
        }
        previous = Some(span);
    }
    Ok(())
}

fn scan(xml: &[u8], mut projection: Projection<'_>) -> Result<Parsed> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = true;
    reader.config_mut().check_comments = true;
    reader
        .resolver_mut()
        .set_max_declarations_per_element(MAX_NAMESPACE_DECLARATIONS);
    let mut stack = Vec::new();
    let mut contexts = Vec::new();
    let mut traces = Vec::new();
    let mut brush_properties = Vec::new();
    let mut root_seen = false;
    let mut root_closed = false;
    let mut nodes = 0usize;
    let mut declaration_seen = false;
    let mut preamble_content_seen = false;
    let mut legacy_fragment = false;

    loop {
        let start = position(&reader, "InkML")?;
        let event = reader.read_event().map_err(xml_error)?;
        let (resolved, event) = reader.resolver().resolve_event(event);
        let end = position(&reader, "InkML")?;
        let is_declaration = matches!(&event, Event::Decl(_));
        match event {
            Event::Decl(declaration)
                if !root_seen && !declaration_seen && !preamble_content_seen =>
            {
                validate_declaration(&declaration)?;
                declaration_seen = true;
            },
            Event::Start(element) if !root_seen => {
                validate_element(&element, &reader)?;
                legacy_fragment = require_root(&element, &resolved)?;
                root_seen = true;
                increment_nodes(&mut nodes)?;
                enforce_depth(1)?;
                let frame = frame_for_start(
                    &element,
                    &resolved,
                    legacy_fragment,
                    start,
                    end,
                    &reader,
                    &mut contexts,
                    &mut traces,
                    &mut brush_properties,
                    &mut projection,
                )?;
                reserve_one(&mut stack, "InkML XML stack")?;
                stack.push(frame);
            },
            Event::Empty(element) if !root_seen => {
                validate_element(&element, &reader)?;
                legacy_fragment = require_root(&element, &resolved)?;
                root_seen = true;
                root_closed = true;
                increment_nodes(&mut nodes)?;
                enforce_depth(1)?;
                parse_empty(
                    &element,
                    &resolved,
                    legacy_fragment,
                    start,
                    end,
                    &reader,
                    &mut contexts,
                    &mut traces,
                    &mut brush_properties,
                    &mut projection,
                )?;
            },
            Event::Start(element) if root_seen && !root_closed => {
                validate_element(&element, &reader)?;
                validate_unknown_prefix(&resolved, &element, legacy_fragment)?;
                increment_nodes(&mut nodes)?;
                enforce_depth(stack.len().saturating_add(1))?;
                let frame = frame_for_start(
                    &element,
                    &resolved,
                    legacy_fragment,
                    start,
                    end,
                    &reader,
                    &mut contexts,
                    &mut traces,
                    &mut brush_properties,
                    &mut projection,
                )?;
                reserve_one(&mut stack, "InkML XML stack")?;
                stack.push(frame);
            },
            Event::Empty(element) if root_seen && !root_closed => {
                validate_element(&element, &reader)?;
                validate_unknown_prefix(&resolved, &element, legacy_fragment)?;
                increment_nodes(&mut nodes)?;
                enforce_depth(stack.len().saturating_add(1))?;
                parse_empty(
                    &element,
                    &resolved,
                    legacy_fragment,
                    start,
                    end,
                    &reader,
                    &mut contexts,
                    &mut traces,
                    &mut brush_properties,
                    &mut projection,
                )?;
            },
            Event::Start(_) | Event::Empty(_) if root_closed => {
                return Err(invalid("InkML document has more than one root element"));
            },
            Event::End(element) if root_seen && !root_closed => {
                let frame = stack
                    .pop()
                    .ok_or_else(|| invalid("InkML XML has an unexpected closing element"))?;
                validate_end_name(&element, &resolved, legacy_fragment)?;
                if frame.namespace != namespace_id(&resolved)? {
                    return Err(invalid("InkML XML has mismatched closing elements"));
                }
                if let Some(index) = frame.context {
                    contexts[index].source = SourceSpan::new(frame.start, end);
                }
                if let Some(index) = frame.trace {
                    traces[index].source = SourceSpan::new(frame.start, end);
                    traces[index].data = SourceSpan::new(frame.data_start, start);
                }
                if let Some(index) = frame.brush_property {
                    brush_properties[index].source = SourceSpan::new(frame.start, end);
                }
                if let Some(token) = frame.projection {
                    projection.complete(token, SourceSpan::new(frame.start, end))?;
                }
                if stack.is_empty() {
                    root_closed = true;
                }
            },
            Event::End(_) if root_closed => {
                return Err(invalid("InkML XML has content after its root"));
            },
            Event::Start(_) | Event::Empty(_) | Event::End(_) => {
                return Err(invalid("InkML XML has an invalid root transition"));
            },
            Event::Text(text)
                if (!root_seen || root_closed)
                    && !validate_text(&text, "InkML text")?
                        .bytes()
                        .all(|byte| byte.is_ascii_whitespace()) =>
            {
                return Err(invalid("InkML has text outside its root"));
            },
            Event::Text(text) => {
                validate_text(&text, "InkML text")?;
            },
            Event::CData(data) => {
                validate_text(&data, "InkML CDATA")?;
                if !root_seen || root_closed {
                    return Err(invalid("InkML CDATA is not allowed outside its root"));
                }
            },
            Event::GeneralRef(reference) => {
                validate_reference(&reference)?;
                if !root_seen || root_closed {
                    return Err(invalid("InkML has a reference outside its root"));
                }
            },
            Event::Comment(comment) => {
                validate_text(&comment, "InkML comment")?;
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid("InkML rejects DTDs and processing instructions"));
            },
            Event::Decl(_) => {
                return Err(invalid("InkML has a duplicate or late XML declaration"));
            },
            Event::Eof => break,
        }
        if !is_declaration {
            preamble_content_seen = true;
        }
    }

    if !root_seen || !root_closed || !stack.is_empty() {
        return Err(invalid("InkML document root is absent or unterminated"));
    }
    if contexts.len() > MAX_CONTEXTS {
        return Err(limit("InkML context nodes", MAX_CONTEXTS));
    }
    if traces.len() > MAX_TRACES {
        return Err(limit("InkML traces", MAX_TRACES));
    }
    if brush_properties.len() > MAX_BRUSH_PROPERTIES {
        return Err(limit("InkML brush properties", MAX_BRUSH_PROPERTIES));
    }
    projection.finish()?;
    Ok(Parsed {
        contexts,
        traces,
        brush_properties,
    })
}

/// Write an InkML document using its exact retained source bytes.
///
/// # Errors
///
/// Returns an error when source validation fails or the sink write fails.
pub fn write(document: &Document) -> Result<Vec<u8>> {
    super::validation::validate(document).map_err(value_error)?;
    Ok(document.source().to_vec())
}

/// Write an InkML document to a caller-provided sink.
///
/// # Errors
///
/// Returns an error when source validation fails or the sink write fails.
pub fn write_to<W: Write>(writer: &mut W, document: &Document) -> Result<()> {
    writer.write_all(&write(document)?)?;
    Ok(())
}

#[derive(Debug)]
struct Frame {
    namespace: NamespaceId,
    start: usize,
    data_start: usize,
    context: Option<usize>,
    trace: Option<usize>,
    brush_property: Option<usize>,
    projection: Option<ProjectionToken>,
}

fn frame_for_start<R: std::io::BufRead>(
    element: &BytesStart<'_>,
    resolved: &ResolveResult<'_>,
    legacy_fragment: bool,
    start: usize,
    data_start: usize,
    reader: &NsReader<R>,
    contexts: &mut Vec<Context>,
    traces: &mut Vec<Trace>,
    brush_properties: &mut Vec<BrushProperty>,
    projection: &mut Projection<'_>,
) -> Result<Frame> {
    let local = element.name().local_name();
    let is_link = is_namespace(
        resolved,
        element,
        b"msink",
        NAMESPACE.as_bytes(),
        legacy_fragment,
    ) && matches!(local.as_ref(), b"sourceLink" | b"destinationLink");
    let mut projection_token = None;
    if is_link {
        if let Some(token) = projection.select(SemanticKind::Link, start) {
            validate_link(element, reader)?;
            projection_token = Some(token);
        }
    }
    let context = if is_namespace(
        resolved,
        element,
        b"msink",
        NAMESPACE.as_bytes(),
        legacy_fragment,
    ) && local.as_ref() == b"context"
    {
        let token = projection.select(SemanticKind::Context, start);
        if let Some(token) = token {
            if contexts.len() >= MAX_CONTEXTS {
                return Err(limit("InkML context nodes", MAX_CONTEXTS));
            }
            let index = contexts.len();
            reserve_one(contexts, "InkML context records")?;
            contexts.push(parse_context(element, start, reader)?);
            projection_token = Some(token);
            Some(index)
        } else {
            None
        }
    } else {
        None
    };
    let trace = if is_namespace(
        resolved,
        element,
        b"inkml",
        INKML_NAMESPACE.as_bytes(),
        legacy_fragment,
    ) && local.as_ref() == b"trace"
    {
        let token = projection.select(SemanticKind::Trace, start);
        if let Some(token) = token {
            if traces.len() >= MAX_TRACES {
                return Err(limit("InkML traces", MAX_TRACES));
            }
            let index = traces.len();
            reserve_one(traces, "InkML trace records")?;
            traces.push(parse_trace(element, start, data_start, false, reader)?);
            projection_token = Some(token);
            Some(index)
        } else {
            None
        }
    } else {
        None
    };
    let brush_property = if is_namespace(
        resolved,
        element,
        b"inkml",
        INKML_NAMESPACE.as_bytes(),
        legacy_fragment,
    ) && local.as_ref() == b"brushProperty"
    {
        let token = projection.select(SemanticKind::BrushProperty, start);
        if let Some(token) = token {
            if brush_properties.len() >= MAX_BRUSH_PROPERTIES {
                return Err(limit("InkML brush properties", MAX_BRUSH_PROPERTIES));
            }
            let index = brush_properties.len();
            reserve_one(brush_properties, "InkML brush-property records")?;
            brush_properties.push(parse_brush_property(
                element, start, data_start, false, reader,
            )?);
            projection_token = Some(token);
            Some(index)
        } else {
            None
        }
    } else {
        None
    };
    Ok(Frame {
        namespace: namespace_id(resolved)?,
        start,
        data_start,
        context,
        trace,
        brush_property,
        projection: projection_token,
    })
}

fn parse_empty<R: std::io::BufRead>(
    element: &BytesStart<'_>,
    resolved: &ResolveResult<'_>,
    legacy_fragment: bool,
    start: usize,
    end: usize,
    reader: &NsReader<R>,
    contexts: &mut Vec<Context>,
    traces: &mut Vec<Trace>,
    brush_properties: &mut Vec<BrushProperty>,
    projection: &mut Projection<'_>,
) -> Result<()> {
    let local = element.name().local_name();
    let is_link = is_namespace(
        resolved,
        element,
        b"msink",
        NAMESPACE.as_bytes(),
        legacy_fragment,
    ) && matches!(local.as_ref(), b"sourceLink" | b"destinationLink");
    if is_namespace(
        resolved,
        element,
        b"msink",
        NAMESPACE.as_bytes(),
        legacy_fragment,
    ) && local.as_ref() == b"context"
    {
        if let Some(token) = projection.select(SemanticKind::Context, start) {
            if contexts.len() >= MAX_CONTEXTS {
                return Err(limit("InkML context nodes", MAX_CONTEXTS));
            }
            let mut context = parse_context(element, start, reader)?;
            context.source = SourceSpan::new(start, end);
            reserve_one(contexts, "InkML context records")?;
            contexts.push(context);
            projection.complete(token, SourceSpan::new(start, end))?;
        }
    } else if is_link {
        if let Some(token) = projection.select(SemanticKind::Link, start) {
            validate_link(element, reader)?;
            projection.complete(token, SourceSpan::new(start, end))?;
        }
    } else if is_namespace(
        resolved,
        element,
        b"inkml",
        INKML_NAMESPACE.as_bytes(),
        legacy_fragment,
    ) && local.as_ref() == b"trace"
    {
        if let Some(token) = projection.select(SemanticKind::Trace, start) {
            if traces.len() >= MAX_TRACES {
                return Err(limit("InkML traces", MAX_TRACES));
            }
            let mut trace = parse_trace(element, start, end, true, reader)?;
            trace.source = SourceSpan::new(start, end);
            trace.data = SourceSpan::new(end, end);
            reserve_one(traces, "InkML trace records")?;
            traces.push(trace);
            projection.complete(token, SourceSpan::new(start, end))?;
        }
    } else if is_namespace(
        resolved,
        element,
        b"inkml",
        INKML_NAMESPACE.as_bytes(),
        legacy_fragment,
    ) && local.as_ref() == b"brushProperty"
    {
        if let Some(token) = projection.select(SemanticKind::BrushProperty, start) {
            if brush_properties.len() >= MAX_BRUSH_PROPERTIES {
                return Err(limit("InkML brush properties", MAX_BRUSH_PROPERTIES));
            }
            let mut property = parse_brush_property(element, start, end, true, reader)?;
            property.source = SourceSpan::new(start, end);
            reserve_one(brush_properties, "InkML brush-property records")?;
            brush_properties.push(property);
            projection.complete(token, SourceSpan::new(start, end))?;
        }
    }
    Ok(())
}

fn parse_context<R: std::io::BufRead>(
    element: &BytesStart<'_>,
    start: usize,
    reader: &NsReader<R>,
) -> Result<Context> {
    let kind = attr(element, b"type", reader)?
        .ok_or_else(|| value_error(ValueError::MissingAttribute { name: "type" }))?;
    let kind = ContextKind::parse(&kind).map_err(value_error)?;
    let id = attr(element, b"id", reader)?
        .map(|value| Guid::new(value).map_err(value_error))
        .transpose()?;
    if let Some(value) = attr(element, b"customRecognizerId", reader)? {
        let _ = Guid::new(value).map_err(value_error)?;
    }
    for (name, required_type) in [
        (b"rotatedBoundingBox".as_slice(), "rotatedBoundingBox"),
        (b"ascender".as_slice(), "ascender"),
        (b"descender".as_slice(), "descender"),
        (b"baseline".as_slice(), "baseline"),
        (b"midline".as_slice(), "midline"),
        (b"hotPoints".as_slice(), "hotPoints"),
        (b"centroid".as_slice(), "centroid"),
        (b"shapeGeometry".as_slice(), "shapeGeometry"),
    ] {
        if let Some(value) = attr(element, name, reader)? {
            let valid = if name == b"centroid" {
                validate_point(&value)
            } else {
                validate_points(&value)
            };
            valid.map_err(|_| {
                value_error(ValueError::Points {
                    name: required_type,
                })
            })?;
        }
    }
    let alignment_level = parse_int(attr(element, b"alignmentLevel", reader)?, "alignmentLevel")?;
    let content_type = parse_int(attr(element, b"contentType", reader)?, "contentType")?;
    let rotation_angle = parse_int(attr(element, b"rotationAngle", reader)?, "rotationAngle")?;
    let semantic_type = attr(element, b"semanticType", reader)?
        .map(|value| SemanticType::parse(&value).map_err(value_error))
        .transpose()?;
    Ok(Context {
        kind,
        id,
        semantic_type,
        alignment_level,
        content_type,
        rotation_angle,
        source: SourceSpan::new(start, start),
    })
}

fn parse_trace<R: std::io::BufRead>(
    element: &BytesStart<'_>,
    start: usize,
    data_start: usize,
    empty: bool,
    reader: &NsReader<R>,
) -> Result<Trace> {
    let context_ref = attr(element, b"contextRef", reader)?
        .map(validate_reference_text)
        .transpose()?;
    let brush_ref = attr(element, b"brushRef", reader)?
        .map(validate_reference_text)
        .transpose()?;
    let data = if empty {
        SourceSpan::new(data_start, data_start)
    } else {
        SourceSpan::new(data_start, data_start)
    };
    Ok(Trace {
        context_ref,
        brush_ref,
        source: SourceSpan::new(start, start),
        data,
    })
}

fn parse_brush_property<R: std::io::BufRead>(
    element: &BytesStart<'_>,
    start: usize,
    data_start: usize,
    empty: bool,
    reader: &NsReader<R>,
) -> Result<BrushProperty> {
    let name = attr(element, b"name", reader)?
        .ok_or_else(|| value_error(ValueError::MissingAttribute { name: "name" }))?;
    if name.len() > MAX_TOKEN_BYTES {
        return Err(limit("InkML brush property name", MAX_TOKEN_BYTES));
    }
    let value = attr(element, b"value", reader)?.unwrap_or_default();
    let units = attr(element, b"units", reader)?;
    let _ = (data_start, empty);
    Ok(BrushProperty {
        name: BrushPropertyName::parse(&name),
        value: value.into_boxed_str(),
        units: units.map(Into::into),
        source: SourceSpan::new(start, start),
    })
}

fn validate_link<R: std::io::BufRead>(
    element: &BytesStart<'_>,
    reader: &NsReader<R>,
) -> Result<()> {
    if let Some(direction) = attr(element, b"direction", reader)? {
        if !matches!(direction.as_str(), "to" | "from" | "with") {
            return Err(invalid("InkML link direction is invalid"));
        }
    }
    if let Some(reference) = attr(element, b"ref", reader)? {
        if reference.starts_with('{') {
            let _ = Guid::new(reference).map_err(value_error)?;
        } else {
            reference
                .parse::<u32>()
                .map_err(|_| invalid("InkML link ref is invalid"))?;
        }
    }
    Ok(())
}

fn validate_reference_text(value: String) -> Result<Box<str>> {
    if value.is_empty() || value.len() > 4096 || value.bytes().any(|byte| byte == 0) {
        return Err(invalid("InkML reference text is invalid or overlong"));
    }
    Ok(value.into_boxed_str())
}

fn parse_int(value: Option<String>, name: &'static str) -> Result<Option<i32>> {
    value
        .map(|value| {
            value
                .parse::<i32>()
                .map_err(|_| value_error(ValueError::Integer { name }))
        })
        .transpose()
}

fn validate_points(value: &str) -> std::result::Result<(), ()> {
    if value.is_empty() {
        return Err(());
    }
    let mut count = 0usize;
    for point in value.split_whitespace() {
        let mut pieces = point.split(',');
        let x = pieces.next().ok_or(())?;
        let y = pieces.next().ok_or(())?;
        if pieces.next().is_some() || x.parse::<i64>().is_err() || y.parse::<i64>().is_err() {
            return Err(());
        }
        count += 1;
        if count > 4096 {
            return Err(());
        }
    }
    Ok(())
}

fn validate_point(value: &str) -> std::result::Result<(), ()> {
    if value.split_whitespace().count() != 1 {
        return Err(());
    }
    validate_points(value)
}

fn attr<R: std::io::BufRead>(
    element: &BytesStart<'_>,
    name: &[u8],
    reader: &NsReader<R>,
) -> Result<Option<String>> {
    let mut result = None;
    for attribute in element.attributes() {
        let attribute = attribute.map_err(xml_error)?;
        if attribute.key.prefix().is_some() {
            continue;
        }
        if attribute.key.local_name().as_ref() != name {
            continue;
        }
        if result.is_some() {
            return Err(invalid(format!(
                "InkML duplicate attribute '{}'",
                String::from_utf8_lossy(name)
            )));
        }
        if attribute.value.len() > MAX_ATTRIBUTE_VALUE_BYTES {
            return Err(limit(
                "InkML attribute value bytes",
                MAX_ATTRIBUTE_VALUE_BYTES,
            ));
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, reader.decoder())
            .map_err(xml_error)?;
        validate_xml_characters(&value, "InkML attribute")?;
        let mut owned = String::new();
        owned
            .try_reserve_exact(value.len())
            .map_err(|_| invalid("InkML attribute allocation failed"))?;
        owned.push_str(&value);
        result = Some(owned);
    }
    Ok(result)
}

fn require_root(element: &BytesStart<'_>, resolved: &ResolveResult<'_>) -> Result<bool> {
    if element.name().local_name().as_ref() != b"ink" {
        return Err(invalid("InkML root must be inkml:ink"));
    }
    match resolved {
        ResolveResult::Bound(Namespace(value)) if *value == INKML_NAMESPACE.as_bytes() => Ok(false),
        ResolveResult::Unknown(prefix)
            if prefix.as_slice() == b"inkml"
                && !has_namespace_declaration(element, prefix.as_slice()) =>
        {
            Ok(true)
        },
        _ => Err(invalid("InkML root must be inkml:ink")),
    }
}

fn is_namespace(
    resolved: &ResolveResult<'_>,
    element: &BytesStart<'_>,
    prefix: &[u8],
    expected: &[u8],
    legacy_fragment: bool,
) -> bool {
    match resolved {
        ResolveResult::Bound(Namespace(value)) => *value == expected,
        ResolveResult::Unknown(value) => {
            legacy_fragment
                && value.as_slice() == prefix
                && !has_namespace_declaration(element, value.as_slice())
        },
        ResolveResult::Unbound => false,
    }
}

fn namespace_id(resolved: &ResolveResult<'_>) -> Result<NamespaceId> {
    match resolved {
        ResolveResult::Bound(Namespace(value)) => {
            std::str::from_utf8(value).map_err(xml_error)?;
            Ok(match *value {
                value if value == INKML_NAMESPACE.as_bytes() => NamespaceId::Inkml,
                value if value == NAMESPACE.as_bytes() => NamespaceId::Msink,
                _ => NamespaceId::Other,
            })
        },
        ResolveResult::Unknown(prefix) => {
            std::str::from_utf8(prefix).map_err(xml_error)?;
            Ok(NamespaceId::Unknown)
        },
        ResolveResult::Unbound => Ok(NamespaceId::Unbound),
    }
}

fn validate_unknown_prefix(
    resolved: &ResolveResult<'_>,
    element: &BytesStart<'_>,
    legacy_fragment: bool,
) -> Result<()> {
    if let ResolveResult::Unknown(prefix) = resolved {
        std::str::from_utf8(prefix).map_err(xml_error)?;
        if !legacy_fragment
            || has_namespace_declaration(element, prefix.as_slice())
            || !matches!(prefix.as_slice(), b"inkml" | b"msink")
        {
            return Err(invalid("InkML element uses an undeclared namespace prefix"));
        }
    }
    Ok(())
}

fn validate_end_name(
    element: &quick_xml::events::BytesEnd<'_>,
    resolved: &ResolveResult<'_>,
    legacy_fragment: bool,
) -> Result<()> {
    let element_name = element.name();
    let name = std::str::from_utf8(element_name.as_ref()).map_err(xml_error)?;
    if !is_qualified_name(name) {
        return Err(invalid("InkML closing element name is invalid"));
    }
    if let ResolveResult::Unknown(prefix) = resolved {
        std::str::from_utf8(prefix).map_err(xml_error)?;
        if !legacy_fragment || !matches!(prefix.as_slice(), b"inkml" | b"msink") {
            return Err(invalid(
                "InkML closing element uses an undeclared namespace prefix",
            ));
        }
    }
    Ok(())
}

fn validate_element<R: std::io::BufRead>(
    element: &BytesStart<'_>,
    reader: &NsReader<R>,
) -> Result<()> {
    let element_name = element.name();
    let name = std::str::from_utf8(element_name.as_ref()).map_err(xml_error)?;
    if !is_qualified_name(name) {
        return Err(invalid("InkML element name is invalid"));
    }
    let mut attribute_keys = Vec::new();
    for attribute in element.attributes() {
        let attribute = attribute.map_err(xml_error)?;
        if attribute_keys.len() >= MAX_ATTRIBUTES_PER_ELEMENT {
            return Err(limit(
                "InkML attributes per element",
                MAX_ATTRIBUTES_PER_ELEMENT,
            ));
        }
        let name = std::str::from_utf8(attribute.key.as_ref()).map_err(xml_error)?;
        if !is_qualified_name(name) {
            return Err(invalid("InkML attribute name is invalid"));
        }
        if attribute.value.len() > MAX_ATTRIBUTE_VALUE_BYTES {
            return Err(limit(
                "InkML attribute value bytes",
                MAX_ATTRIBUTE_VALUE_BYTES,
            ));
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, reader.decoder())
            .map_err(xml_error)?;
        validate_xml_characters(&value, "InkML attribute")?;
        if let Some(prefix) = attribute.key.prefix()
            && !matches!(prefix.as_ref(), b"xml" | b"xmlns")
            && matches!(
                reader.resolver().resolve_attribute(attribute.key).0,
                ResolveResult::Unknown(_)
            )
        {
            return Err(invalid(
                "InkML attribute uses an undeclared namespace prefix",
            ));
        }
        if attribute_keys
            .iter()
            .copied()
            .any(|key| expanded_attribute_names_equal(key, attribute.key, reader))
        {
            return Err(invalid("InkML element has duplicate expanded attributes"));
        }
        attribute_keys
            .try_reserve(1)
            .map_err(|_| invalid("InkML attribute-key allocation failed"))?;
        attribute_keys.push(attribute.key);
    }
    Ok(())
}

fn validate_declaration(declaration: &BytesDecl<'_>) -> Result<()> {
    let raw = declaration.as_ref();
    std::str::from_utf8(raw).map_err(xml_error)?;
    let mut cursor = 0;
    skip_decl_whitespace(raw, &mut cursor);
    if !consume_decl_token(raw, &mut cursor, b"xml") || !decl_whitespace(raw.get(cursor).copied()) {
        return Err(invalid("InkML XML declaration is malformed"));
    }
    skip_decl_whitespace(raw, &mut cursor);
    let (name, value) = parse_decl_attribute(raw, &mut cursor)?;
    if name != b"version" || value != b"1.0" {
        return Err(invalid("InkML XML declaration must start with version 1.0"));
    }
    let mut previous = b"version".as_slice();
    while {
        skip_decl_whitespace(raw, &mut cursor);
        cursor < raw.len()
    } {
        let (name, value) = parse_decl_attribute(raw, &mut cursor)?;
        let valid_order = match (previous, name) {
            (b"version", b"encoding") => value.eq_ignore_ascii_case(b"utf-8"),
            (b"version" | b"encoding", b"standalone") => {
                matches!(value, b"yes" | b"no")
            },
            _ => false,
        };
        if !valid_order {
            return Err(invalid(
                "InkML XML declaration has an invalid or duplicate attribute",
            ));
        }
        previous = name;
    }
    Ok(())
}

fn decl_whitespace(value: Option<u8>) -> bool {
    matches!(value, Some(b' ' | b'\t' | b'\r' | b'\n'))
}

fn skip_decl_whitespace(raw: &[u8], cursor: &mut usize) {
    while raw
        .get(*cursor)
        .copied()
        .is_some_and(|byte| decl_whitespace(Some(byte)))
    {
        *cursor += 1;
    }
}

fn consume_decl_token(raw: &[u8], cursor: &mut usize, token: &[u8]) -> bool {
    raw.get(*cursor..)
        .is_some_and(|remaining| remaining.starts_with(token))
        .then(|| *cursor += token.len())
        .is_some()
}

fn parse_decl_attribute<'a>(raw: &'a [u8], cursor: &mut usize) -> Result<(&'a [u8], &'a [u8])> {
    let name_start = *cursor;
    while raw
        .get(*cursor)
        .copied()
        .is_some_and(|byte| byte.is_ascii_alphabetic())
    {
        *cursor += 1;
    }
    if *cursor == name_start {
        return Err(invalid("InkML XML declaration attribute name is missing"));
    }
    let name = &raw[name_start..*cursor];
    skip_decl_whitespace(raw, cursor);
    if raw.get(*cursor) != Some(&b'=') {
        return Err(invalid(
            "InkML XML declaration attribute equals sign is missing",
        ));
    }
    *cursor += 1;
    skip_decl_whitespace(raw, cursor);
    let quote = raw
        .get(*cursor)
        .copied()
        .filter(|value| matches!(value, b'\'' | b'"'))
        .ok_or_else(|| invalid("InkML XML declaration attribute quote is missing"))?;
    *cursor += 1;
    let value_start = *cursor;
    while raw
        .get(*cursor)
        .copied()
        .is_some_and(|value| value != quote)
    {
        *cursor += 1;
    }
    if raw.get(*cursor) != Some(&quote) {
        return Err(invalid("InkML XML declaration attribute is unterminated"));
    }
    let value = &raw[value_start..*cursor];
    *cursor += 1;
    Ok((name, value))
}

fn validate_xml_characters(value: &str, what: &str) -> Result<()> {
    if super::xml_characters::valid(value) {
        Ok(())
    } else {
        Err(invalid(format!("{what} contains an invalid XML character")))
    }
}

fn is_xml10_character(value: char) -> bool {
    matches!(value, '\u{9}' | '\u{A}' | '\u{D}' | '\u{20}'..='\u{D7FF}' | '\u{E000}'..='\u{FFFD}' | '\u{10000}'..='\u{10FFFF}')
}

fn validate_text<'a>(value: &'a [u8], what: &str) -> Result<&'a str> {
    let value = std::str::from_utf8(value).map_err(xml_error)?;
    validate_xml_characters(value, what)?;
    Ok(value)
}

fn validate_reference(reference: &BytesRef<'_>) -> Result<()> {
    let value = std::str::from_utf8(reference.as_ref()).map_err(xml_error)?;
    match value {
        "amp" | "lt" | "gt" | "apos" | "quot" => Ok(()),
        value if value.strip_prefix("#x").is_some() => {
            let digits = value.strip_prefix("#x").unwrap_or_default();
            if digits.is_empty() || !digits.bytes().all(|byte| byte.is_ascii_hexdigit()) {
                return Err(invalid("InkML hexadecimal character reference is invalid"));
            }
            let codepoint = u32::from_str_radix(digits, 16)
                .map_err(|_| invalid("InkML hexadecimal character reference is invalid"))?;
            let character = char::from_u32(codepoint)
                .ok_or_else(|| invalid("InkML character reference is invalid"))?;
            if is_xml10_character(character) {
                Ok(())
            } else {
                Err(invalid(
                    "InkML character reference is not an XML 1.0 character",
                ))
            }
        },
        value if value.strip_prefix('#').is_some() => {
            let digits = value.strip_prefix('#').unwrap_or_default();
            if digits.is_empty() || !digits.bytes().all(|byte| byte.is_ascii_digit()) {
                return Err(invalid("InkML decimal character reference is invalid"));
            }
            let codepoint = digits
                .parse::<u32>()
                .map_err(|_| invalid("InkML decimal character reference is invalid"))?;
            let character = char::from_u32(codepoint)
                .ok_or_else(|| invalid("InkML character reference is invalid"))?;
            if is_xml10_character(character) {
                Ok(())
            } else {
                Err(invalid(
                    "InkML character reference is not an XML 1.0 character",
                ))
            }
        },
        _ => Err(invalid("InkML general entity references are not supported")),
    }
}

fn has_namespace_declaration(element: &BytesStart<'_>, prefix: &[u8]) -> bool {
    element.attributes().any(|attribute| {
        let Ok(attribute) = attribute else {
            return false;
        };
        let key = attribute.key.as_ref();
        (prefix.is_empty() && key == b"xmlns") || (key.strip_prefix(b"xmlns:") == Some(prefix))
    })
}

fn expanded_attribute_names_equal<R: std::io::BufRead>(
    left: QName<'_>,
    right: QName<'_>,
    reader: &NsReader<R>,
) -> bool {
    let (left_namespace, left_local) = reader.resolver().resolve_attribute(left);
    let (right_namespace, right_local) = reader.resolver().resolve_attribute(right);
    if left_local != right_local {
        return false;
    }
    match (left_namespace, right_namespace) {
        (ResolveResult::Unbound, ResolveResult::Unbound) => true,
        (ResolveResult::Bound(Namespace(left)), ResolveResult::Bound(Namespace(right))) => {
            left == right
        },
        _ => false,
    }
}

fn reserve_one<T>(values: &mut Vec<T>, resource: &'static str) -> Result<()> {
    values
        .try_reserve(1)
        .map_err(|_| invalid(format!("{resource} allocation failed")))
}

fn enforce_source_limit(source_len: usize) -> Result<()> {
    if source_len > MAX_SOURCE_BYTES {
        Err(limit("InkML source bytes", MAX_SOURCE_BYTES))
    } else {
        Ok(())
    }
}

fn copy_source(xml: &[u8]) -> Result<Arc<Vec<u8>>> {
    let mut source = Vec::new();
    source
        .try_reserve_exact(xml.len())
        .map_err(|_| invalid("InkML source allocation failed"))?;
    source.extend_from_slice(xml);
    Ok(Arc::new(source))
}

fn position<R: std::io::BufRead>(reader: &NsReader<R>, what: &str) -> Result<usize> {
    usize::try_from(reader.buffer_position())
        .map_err(|_| invalid(format!("{what} offset exceeds usize")))
}

fn increment_nodes(nodes: &mut usize) -> Result<()> {
    let next = nodes
        .checked_add(1)
        .ok_or_else(|| limit("InkML XML nodes", MAX_NODES))?;
    enforce_nodes(next)?;
    *nodes = next;
    Ok(())
}

fn enforce_nodes(nodes: usize) -> Result<()> {
    if nodes > MAX_NODES {
        Err(limit("InkML XML nodes", MAX_NODES))
    } else {
        Ok(())
    }
}

fn enforce_depth(depth: usize) -> Result<()> {
    if depth > MAX_DEPTH {
        Err(limit("InkML XML depth", MAX_DEPTH))
    } else {
        Ok(())
    }
}

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}
fn limit(resource: &'static str, limit: usize) -> Error {
    Error::Limit { resource, limit }
}
fn xml_error(error: impl fmt::Display) -> Error {
    Error::Xml(error.to_string())
}
fn value_error(error: impl fmt::Display) -> Error {
    Error::Invalid(error.to_string())
}
