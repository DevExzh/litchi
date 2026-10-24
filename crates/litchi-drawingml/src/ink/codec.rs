//! Bounded namespace-aware InkML XML reader and source-preserving writer.

use std::{fmt, io::Write, sync::Arc};

use litchi_core::xml::ReaderOrigin;
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
use litchi_ooxml_common::xml::attributes::BytesStartExt as _;
use litchi_ooxml_common::xml::attributes::first_wins;

const MAX_NAMESPACE_DECLARATIONS: usize = 256;
const MAX_ATTRIBUTES_PER_ELEMENT: usize = 256;
const MAX_REFERENCE_BYTES: usize = 4096;

/// Caller-selected bounds for raw InkML source and output operations.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
#[must_use]
pub struct Limits {
    /// Maximum retained source bytes accepted by a reader.
    pub max_source_bytes: usize,
    /// Maximum serialized bytes emitted by a writer.
    pub max_output_bytes: usize,
    /// Maximum XML nodes scanned.
    pub max_xml_nodes: usize,
    /// Maximum XML depth scanned.
    pub max_xml_depth: usize,
    /// Maximum contexts retained.
    pub max_contexts: usize,
    /// Maximum traces retained.
    pub max_traces: usize,
    /// Maximum brush properties retained.
    pub max_brush_properties: usize,
    /// Maximum attributes admitted on one element.
    pub max_attributes_per_element: usize,
    /// Maximum decoded bytes in one XML attribute value.
    pub max_attribute_value_bytes: usize,
    /// Maximum bytes in one typed identifier.
    pub max_identifier_bytes: usize,
    /// Maximum bytes in one typed reference value.
    pub max_reference_bytes: usize,
    /// Maximum bytes in one trace lexical value.
    pub max_trace_text_bytes: usize,
}

impl Limits {
    /// Historical bounded InkML policy.
    pub const DEFAULT: Self = Self {
        max_source_bytes: MAX_SOURCE_BYTES,
        max_output_bytes: MAX_SOURCE_BYTES,
        max_xml_nodes: MAX_NODES,
        max_xml_depth: MAX_DEPTH,
        max_contexts: MAX_CONTEXTS,
        max_traces: MAX_TRACES,
        max_brush_properties: MAX_BRUSH_PROPERTIES,
        max_attributes_per_element: MAX_ATTRIBUTES_PER_ELEMENT,
        max_attribute_value_bytes: MAX_ATTRIBUTE_VALUE_BYTES,
        max_identifier_bytes: MAX_TOKEN_BYTES,
        max_reference_bytes: MAX_REFERENCE_BYTES,
        max_trace_text_bytes: MAX_ATTRIBUTE_VALUE_BYTES,
    };

    /// Construct a policy with a caller-selected source/output byte cap.
    #[must_use]
    pub const fn new(max_source_bytes: usize, max_output_bytes: usize) -> Self {
        Self {
            max_source_bytes,
            max_output_bytes,
            ..Self::DEFAULT
        }
    }

    fn validate(self) -> Result<Self> {
        check_limit(
            crate::LimitKind::InkMlSourceBytes,
            self.max_source_bytes,
            MAX_SOURCE_BYTES,
        )?;
        check_limit(
            crate::LimitKind::InkMlOutputBytes,
            self.max_output_bytes,
            MAX_SOURCE_BYTES,
        )?;
        check_limit(
            crate::LimitKind::InkMlXmlNodes,
            self.max_xml_nodes,
            MAX_NODES,
        )?;
        check_limit(
            crate::LimitKind::InkMlXmlDepth,
            self.max_xml_depth,
            MAX_DEPTH,
        )?;
        check_limit(
            crate::LimitKind::InkMlContextNodes,
            self.max_contexts,
            MAX_CONTEXTS,
        )?;
        check_limit(crate::LimitKind::InkMlTraces, self.max_traces, MAX_TRACES)?;
        check_limit(
            crate::LimitKind::InkMlBrushProperties,
            self.max_brush_properties,
            MAX_BRUSH_PROPERTIES,
        )?;
        check_limit(
            crate::LimitKind::InkMlAttributesPerElement,
            self.max_attributes_per_element,
            MAX_ATTRIBUTES_PER_ELEMENT,
        )?;
        check_limit(
            crate::LimitKind::InkMlAttributeValueBytes,
            self.max_attribute_value_bytes,
            MAX_ATTRIBUTE_VALUE_BYTES,
        )?;
        check_limit(
            crate::LimitKind::InkMlIdentifierBytes,
            self.max_identifier_bytes,
            MAX_TOKEN_BYTES,
        )?;
        check_limit(
            crate::LimitKind::InkMlReferenceBytes,
            self.max_reference_bytes,
            MAX_REFERENCE_BYTES,
        )?;
        check_limit(
            crate::LimitKind::InkMlTraceTextBytes,
            self.max_trace_text_bytes,
            MAX_ATTRIBUTE_VALUE_BYTES,
        )?;
        Ok(self)
    }
}

impl Default for Limits {
    fn default() -> Self {
        Self::DEFAULT
    }
}

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
pub fn read(xml: &[u8], limits: Limits) -> Result<Document> {
    let limits = limits.validate()?;
    enforce_source_limit(xml.len(), limits.max_source_bytes)?;
    let metadata = scan(xml, Projection::all(), limits)?.into_metadata();
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
pub fn read_shared(xml: Arc<Vec<u8>>, limits: Limits) -> Result<Document> {
    let limits = limits.validate()?;
    enforce_source_limit(xml.len(), limits.max_source_bytes)?;
    let metadata = scan(xml.as_slice(), Projection::all(), limits)?.into_metadata();
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
pub fn read_metadata(xml: &[u8], limits: Limits) -> Result<Metadata> {
    let limits = limits.validate()?;
    enforce_source_limit(xml.len(), limits.max_source_bytes)?;
    Ok(scan(xml, Projection::all(), limits)?.into_metadata())
}

/// Validate source-backed InkML for a replay policy without rebuilding the
/// retained metadata projection.  The visitor still decodes every recognized
/// typed element and applies all structural, lexical, and resource bounds,
/// but it retains only bounded counters needed for the admission decision.
fn validate_source(xml: &[u8], limits: Limits) -> Result<()> {
    enforce_source_limit(xml.len(), limits.max_source_bytes)?;
    let _ = scan(xml, Projection::validation(), limits)?;
    Ok(())
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
    limits: Limits,
) -> Result<Document> {
    let limits = limits.validate()?;
    enforce_source_limit(xml.len(), limits.max_source_bytes)?;
    let projection = Projection::filtered(xml.len(), contexts, traces, brush_properties, links)?;
    let metadata = scan(xml.as_slice(), projection, limits)?.into_metadata();
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
    limits: Limits,
) -> Result<Metadata> {
    let limits = limits.validate()?;
    enforce_source_limit(xml.len(), limits.max_source_bytes)?;
    let projection = Projection::filtered(xml.len(), contexts, traces, brush_properties, links)?;
    Ok(scan(xml, projection, limits)?.into_metadata())
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
    Validate(SemanticKind),
    Selected(SelectedSpan),
}

impl ProjectionToken {
    fn is_validation(self) -> bool {
        matches!(self, Self::Validate(_))
    }
}

struct Projection<'a> {
    all: bool,
    validate_only: bool,
    contexts: &'a [SourceSpan],
    traces: &'a [SourceSpan],
    brush_properties: &'a [SourceSpan],
    links: &'a [SourceSpan],
    matched_contexts: usize,
    matched_traces: usize,
    matched_brush_properties: usize,
    matched_links: usize,
    validated_contexts: usize,
    validated_traces: usize,
    validated_brush_properties: usize,
}

impl<'a> Projection<'a> {
    fn all() -> Self {
        Self {
            all: true,
            validate_only: false,
            contexts: &[],
            traces: &[],
            brush_properties: &[],
            links: &[],
            matched_contexts: 0,
            matched_traces: 0,
            matched_brush_properties: 0,
            matched_links: 0,
            validated_contexts: 0,
            validated_traces: 0,
            validated_brush_properties: 0,
        }
    }

    fn validation() -> Self {
        Self {
            all: true,
            validate_only: true,
            contexts: &[],
            traces: &[],
            brush_properties: &[],
            links: &[],
            matched_contexts: 0,
            matched_traces: 0,
            matched_brush_properties: 0,
            matched_links: 0,
            validated_contexts: 0,
            validated_traces: 0,
            validated_brush_properties: 0,
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
            validate_only: false,
            contexts,
            traces,
            brush_properties,
            links,
            matched_contexts: 0,
            matched_traces: 0,
            matched_brush_properties: 0,
            matched_links: 0,
            validated_contexts: 0,
            validated_traces: 0,
            validated_brush_properties: 0,
        })
    }

    fn select(&self, kind: SemanticKind, start: usize) -> Option<ProjectionToken> {
        if self.validate_only {
            return Some(ProjectionToken::Validate(kind));
        }
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
        if let ProjectionToken::Validate(kind) = token {
            let matched = match kind {
                SemanticKind::Context => &mut self.validated_contexts,
                SemanticKind::Trace => &mut self.validated_traces,
                SemanticKind::BrushProperty => &mut self.validated_brush_properties,
                SemanticKind::Link => return Ok(()),
            };
            *matched = matched
                .checked_add(1)
                .ok_or_else(|| invalid("InkML validation count overflowed"))?;
            return Ok(());
        }
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

    fn validation_counts(&self) -> (usize, usize, usize) {
        (
            self.validated_contexts,
            self.validated_traces,
            self.validated_brush_properties,
        )
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

fn scan(xml: &[u8], mut projection: Projection<'_>, limits: Limits) -> Result<Parsed> {
    let mut reader = NsReader::from_reader(xml);
    let origin = ReaderOrigin::of(xml);
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
        let start = position(&reader, origin, "InkML")?;
        let event = reader.read_event().map_err(xml_error)?;
        let (resolved, event) = reader.resolver().resolve_event(event);
        let end = position(&reader, origin, "InkML")?;
        let is_declaration = matches!(&event, Event::Decl(_));
        match event {
            Event::Decl(declaration)
                if !root_seen && !declaration_seen && !preamble_content_seen =>
            {
                validate_declaration(&declaration)?;
                declaration_seen = true;
            },
            Event::Start(element) if !root_seen => {
                validate_element(&element, &reader, limits)?;
                legacy_fragment = require_root(&element, &resolved)?;
                root_seen = true;
                increment_nodes(&mut nodes, limits.max_xml_nodes)?;
                enforce_depth(1, limits.max_xml_depth)?;
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
                    limits,
                )?;
                reserve_one(&mut stack, "InkML XML stack")?;
                stack.push(frame);
            },
            Event::Empty(element) if !root_seen => {
                validate_element(&element, &reader, limits)?;
                legacy_fragment = require_root(&element, &resolved)?;
                root_seen = true;
                root_closed = true;
                increment_nodes(&mut nodes, limits.max_xml_nodes)?;
                enforce_depth(1, limits.max_xml_depth)?;
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
                    limits,
                )?;
            },
            Event::Start(element) if root_seen && !root_closed => {
                validate_element(&element, &reader, limits)?;
                validate_unknown_prefix(&resolved, &element, legacy_fragment)?;
                increment_nodes(&mut nodes, limits.max_xml_nodes)?;
                enforce_depth(stack.len().saturating_add(1), limits.max_xml_depth)?;
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
                    limits,
                )?;
                reserve_one(&mut stack, "InkML XML stack")?;
                stack.push(frame);
            },
            Event::Empty(element) if root_seen && !root_closed => {
                validate_element(&element, &reader, limits)?;
                validate_unknown_prefix(&resolved, &element, legacy_fragment)?;
                increment_nodes(&mut nodes, limits.max_xml_nodes)?;
                enforce_depth(stack.len().saturating_add(1), limits.max_xml_depth)?;
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
                    limits,
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
                    ensure_typed_limit(
                        crate::LimitKind::InkMlTraceTextBytes,
                        start.saturating_sub(frame.data_start),
                        limits.max_trace_text_bytes,
                    )?;
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
    let (validated_contexts, validated_traces, validated_brush_properties) =
        projection.validation_counts();
    let context_count = if projection.validate_only {
        validated_contexts
    } else {
        contexts.len()
    };
    let trace_count = if projection.validate_only {
        validated_traces
    } else {
        traces.len()
    };
    let brush_property_count = if projection.validate_only {
        validated_brush_properties
    } else {
        brush_properties.len()
    };
    if context_count > limits.max_contexts {
        return Err(limit(
            crate::LimitKind::InkMlContextNodes,
            context_count,
            limits.max_contexts,
        ));
    }
    if trace_count > limits.max_traces {
        return Err(limit(
            crate::LimitKind::InkMlTraces,
            trace_count,
            limits.max_traces,
        ));
    }
    if brush_property_count > limits.max_brush_properties {
        return Err(limit(
            crate::LimitKind::InkMlBrushProperties,
            brush_property_count,
            limits.max_brush_properties,
        ));
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
pub fn write(document: &Document, limits: Limits) -> Result<Vec<u8>> {
    let limits = limits.validate()?;
    super::validation::validate(document).map_err(value_error)?;
    // A document may have been admitted under a wider policy. Re-run the
    // bounded scanner before reserving the output so every selected source,
    // structural, decoded-value, and typed-payload dimension applies to this
    // writer call rather than only the output-byte cap.
    validate_source(document.source(), limits)?;
    if document.source().len() > limits.max_output_bytes {
        return Err(limit(
            crate::LimitKind::InkMlOutputBytes,
            document.source().len(),
            limits.max_output_bytes,
        ));
    }
    let mut output = Vec::new();
    output
        .try_reserve_exact(document.source().len())
        .map_err(|_| invalid("InkML output allocation failed"))?;
    output.extend_from_slice(document.source());
    Ok(output)
}

/// Write an InkML document to a caller-provided sink.
///
/// # Errors
///
/// Returns an error when source validation fails or the sink write fails.
pub fn write_to<W: Write>(writer: &mut W, document: &Document, limits: Limits) -> Result<()> {
    let limits = limits.validate()?;
    super::validation::validate(document).map_err(value_error)?;
    validate_source(document.source(), limits)?;
    if document.source().len() > limits.max_output_bytes {
        return Err(limit(
            crate::LimitKind::InkMlOutputBytes,
            document.source().len(),
            limits.max_output_bytes,
        ));
    }
    writer.write_all(document.source())?;
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
    limits: Limits,
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
            validate_link(element, reader, limits)?;
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
            if token.is_validation() {
                let _ = parse_context(element, start, reader, limits)?;
                projection_token = Some(token);
                return Ok(Frame {
                    namespace: namespace_id(resolved)?,
                    start,
                    data_start,
                    context: None,
                    trace: None,
                    brush_property: None,
                    projection: projection_token,
                });
            }
            if contexts.len() >= limits.max_contexts {
                return Err(limit(
                    crate::LimitKind::InkMlContextNodes,
                    contexts.len() + 1,
                    limits.max_contexts,
                ));
            }
            let index = contexts.len();
            reserve_one(contexts, "InkML context records")?;
            contexts.push(parse_context(element, start, reader, limits)?);
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
            if token.is_validation() {
                let _ = parse_trace(element, start, data_start, false, reader, limits)?;
                projection_token = Some(token);
                return Ok(Frame {
                    namespace: namespace_id(resolved)?,
                    start,
                    data_start,
                    context: None,
                    trace: None,
                    brush_property: None,
                    projection: projection_token,
                });
            }
            if traces.len() >= limits.max_traces {
                return Err(limit(
                    crate::LimitKind::InkMlTraces,
                    traces.len() + 1,
                    limits.max_traces,
                ));
            }
            let index = traces.len();
            reserve_one(traces, "InkML trace records")?;
            traces.push(parse_trace(
                element, start, data_start, false, reader, limits,
            )?);
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
            if token.is_validation() {
                let _ = parse_brush_property(element, start, data_start, false, reader, limits)?;
                projection_token = Some(token);
                return Ok(Frame {
                    namespace: namespace_id(resolved)?,
                    start,
                    data_start,
                    context: None,
                    trace: None,
                    brush_property: None,
                    projection: projection_token,
                });
            }
            if brush_properties.len() >= limits.max_brush_properties {
                return Err(limit(
                    crate::LimitKind::InkMlBrushProperties,
                    brush_properties.len() + 1,
                    limits.max_brush_properties,
                ));
            }
            let index = brush_properties.len();
            reserve_one(brush_properties, "InkML brush-property records")?;
            brush_properties.push(parse_brush_property(
                element, start, data_start, false, reader, limits,
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
    limits: Limits,
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
            if token.is_validation() {
                let _ = parse_context(element, start, reader, limits)?;
                projection.complete(token, SourceSpan::new(start, end))?;
                return Ok(());
            }
            if contexts.len() >= limits.max_contexts {
                return Err(limit(
                    crate::LimitKind::InkMlContextNodes,
                    contexts.len() + 1,
                    limits.max_contexts,
                ));
            }
            let mut context = parse_context(element, start, reader, limits)?;
            context.source = SourceSpan::new(start, end);
            reserve_one(contexts, "InkML context records")?;
            contexts.push(context);
            projection.complete(token, SourceSpan::new(start, end))?;
        }
    } else if is_link {
        if let Some(token) = projection.select(SemanticKind::Link, start) {
            validate_link(element, reader, limits)?;
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
            if token.is_validation() {
                let _ = parse_trace(element, start, end, true, reader, limits)?;
                projection.complete(token, SourceSpan::new(start, end))?;
                return Ok(());
            }
            if traces.len() >= limits.max_traces {
                return Err(limit(
                    crate::LimitKind::InkMlTraces,
                    traces.len() + 1,
                    limits.max_traces,
                ));
            }
            let mut trace = parse_trace(element, start, end, true, reader, limits)?;
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
            if token.is_validation() {
                let _ = parse_brush_property(element, start, end, true, reader, limits)?;
                projection.complete(token, SourceSpan::new(start, end))?;
                return Ok(());
            }
            if brush_properties.len() >= limits.max_brush_properties {
                return Err(limit(
                    crate::LimitKind::InkMlBrushProperties,
                    brush_properties.len() + 1,
                    limits.max_brush_properties,
                ));
            }
            let mut property = parse_brush_property(element, start, end, true, reader, limits)?;
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
    limits: Limits,
) -> Result<Context> {
    let kind = attr(element, b"type", reader, limits)?
        .ok_or_else(|| value_error(ValueError::MissingAttribute { name: "type" }))?;
    let kind = ContextKind::parse(&kind).map_err(value_error)?;
    let id = attr(element, b"id", reader, limits)?
        .map(|value| {
            ensure_typed_limit(
                crate::LimitKind::InkMlIdentifierBytes,
                value.len(),
                limits.max_identifier_bytes,
            )?;
            Guid::new(value).map_err(value_error)
        })
        .transpose()?;
    if let Some(value) = attr(element, b"customRecognizerId", reader, limits)? {
        ensure_typed_limit(
            crate::LimitKind::InkMlIdentifierBytes,
            value.len(),
            limits.max_identifier_bytes,
        )?;
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
        if let Some(value) = attr(element, name, reader, limits)? {
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
    let alignment_level = parse_int(
        attr(element, b"alignmentLevel", reader, limits)?,
        "alignmentLevel",
    )?;
    let content_type = parse_int(
        attr(element, b"contentType", reader, limits)?,
        "contentType",
    )?;
    let rotation_angle = parse_int(
        attr(element, b"rotationAngle", reader, limits)?,
        "rotationAngle",
    )?;
    let semantic_type = attr(element, b"semanticType", reader, limits)?
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
    limits: Limits,
) -> Result<Trace> {
    let context_ref = attr(element, b"contextRef", reader, limits)?
        .map(|value| validate_reference_text(value, limits.max_reference_bytes))
        .transpose()?;
    let brush_ref = attr(element, b"brushRef", reader, limits)?
        .map(|value| validate_reference_text(value, limits.max_reference_bytes))
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
    limits: Limits,
) -> Result<BrushProperty> {
    let name = attr(element, b"name", reader, limits)?
        .ok_or_else(|| value_error(ValueError::MissingAttribute { name: "name" }))?;
    if name.len() > limits.max_identifier_bytes {
        return Err(limit(
            crate::LimitKind::InkMlIdentifierBytes,
            name.len(),
            limits.max_identifier_bytes,
        ));
    }
    let value = attr(element, b"value", reader, limits)?.unwrap_or_default();
    let units = attr(element, b"units", reader, limits)?;
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
    limits: Limits,
) -> Result<()> {
    if let Some(direction) = attr(element, b"direction", reader, limits)? {
        if !matches!(direction.as_str(), "to" | "from" | "with") {
            return Err(invalid("InkML link direction is invalid"));
        }
    }
    if let Some(reference) = attr(element, b"ref", reader, limits)? {
        ensure_typed_limit(
            crate::LimitKind::InkMlReferenceBytes,
            reference.len(),
            limits.max_reference_bytes,
        )?;
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

fn validate_reference_text(value: String, maximum: usize) -> Result<Box<str>> {
    if value.len() > maximum {
        return Err(limit(
            crate::LimitKind::InkMlReferenceBytes,
            value.len(),
            maximum,
        ));
    }
    if value.is_empty() || value.bytes().any(|byte| byte == 0) {
        return Err(invalid("InkML reference text is invalid"));
    }
    Ok(value.into_boxed_str())
}

fn ensure_typed_limit(kind: crate::LimitKind, observed: usize, maximum: usize) -> Result<()> {
    if observed > maximum {
        return Err(limit(kind, observed, maximum));
    }
    Ok(())
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
    limits: Limits,
) -> Result<Option<String>> {
    let mut result = None;
    for attribute in element.checked_attributes() {
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
                crate::LimitKind::InkMlAttributeValueBytes,
                attribute.value.len(),
                MAX_ATTRIBUTE_VALUE_BYTES,
            ));
        }
        let is_namespace_declaration = attribute.key.as_ref() == b"xmlns"
            || attribute.key.as_ref().strip_prefix(b"xmlns:").is_some();
        crate::preflight_attribute_value_bytes_with_kind(
            attribute.value.as_ref(),
            if is_namespace_declaration {
                MAX_ATTRIBUTE_VALUE_BYTES
            } else {
                limits.max_attribute_value_bytes
            },
            crate::LimitKind::InkMlAttributeValueBytes,
        )?;
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, reader.decoder())
            .map_err(xml_error)?;
        if !is_namespace_declaration && value.len() > limits.max_attribute_value_bytes {
            return Err(limit(
                crate::LimitKind::InkMlAttributeValueBytes,
                value.len(),
                limits.max_attribute_value_bytes,
            ));
        }
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
    limits: Limits,
) -> Result<()> {
    let element_name = element.name();
    let name = std::str::from_utf8(element_name.as_ref()).map_err(xml_error)?;
    if !is_qualified_name(name) {
        return Err(invalid("InkML element name is invalid"));
    }
    let mut attribute_keys = Vec::new();
    for attribute in element.checked_attributes() {
        let attribute = attribute.map_err(xml_error)?;
        if attribute_keys.len() >= limits.max_attributes_per_element {
            return Err(limit(
                crate::LimitKind::InkMlAttributesPerElement,
                attribute_keys.len() + 1,
                limits.max_attributes_per_element,
            ));
        }
        let name = std::str::from_utf8(attribute.key.as_ref()).map_err(xml_error)?;
        if !is_qualified_name(name) {
            return Err(invalid("InkML attribute name is invalid"));
        }
        if attribute.value.len() > MAX_ATTRIBUTE_VALUE_BYTES {
            return Err(limit(
                crate::LimitKind::InkMlAttributeValueBytes,
                attribute.value.len(),
                MAX_ATTRIBUTE_VALUE_BYTES,
            ));
        }
        let is_namespace_declaration = attribute.key.as_ref() == b"xmlns"
            || attribute.key.as_ref().strip_prefix(b"xmlns:").is_some();
        crate::preflight_attribute_value_bytes_with_kind(
            attribute.value.as_ref(),
            if is_namespace_declaration {
                MAX_ATTRIBUTE_VALUE_BYTES
            } else {
                limits.max_attribute_value_bytes
            },
            crate::LimitKind::InkMlAttributeValueBytes,
        )?;
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, reader.decoder())
            .map_err(xml_error)?;
        if !is_namespace_declaration && value.len() > limits.max_attribute_value_bytes {
            return Err(limit(
                crate::LimitKind::InkMlAttributeValueBytes,
                value.len(),
                limits.max_attribute_value_bytes,
            ));
        }
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
    first_wins(element).any(|attribute| {
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

fn enforce_source_limit(source_len: usize, maximum: usize) -> Result<()> {
    if source_len > maximum {
        Err(limit(
            crate::LimitKind::InkMlSourceBytes,
            source_len,
            maximum,
        ))
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

fn position<R: std::io::BufRead>(
    reader: &NsReader<R>,
    origin: ReaderOrigin,
    what: &str,
) -> Result<usize> {
    origin
        .offset(reader.buffer_position())
        .ok_or_else(|| invalid(format!("{what} offset exceeds usize")))
}

fn increment_nodes(nodes: &mut usize, maximum: usize) -> Result<()> {
    let next = nodes
        .checked_add(1)
        .ok_or_else(|| invalid("InkML XML node count overflow"))?;
    enforce_nodes(next, maximum)?;
    *nodes = next;
    Ok(())
}

fn enforce_nodes(nodes: usize, maximum: usize) -> Result<()> {
    if nodes > maximum {
        Err(limit(crate::LimitKind::InkMlXmlNodes, nodes, maximum))
    } else {
        Ok(())
    }
}

fn enforce_depth(depth: usize, maximum: usize) -> Result<()> {
    if depth > maximum {
        Err(limit(crate::LimitKind::InkMlXmlDepth, depth, maximum))
    } else {
        Ok(())
    }
}

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}
fn limit(kind: crate::LimitKind, observed: usize, maximum: usize) -> Error {
    Error::limit(kind, observed, maximum)
}
fn check_limit(kind: crate::LimitKind, value: usize, maximum: usize) -> Result<()> {
    if value == 0 || value > maximum {
        return Err(Error::invalid_limit(kind, value, maximum));
    }
    Ok(())
}
fn xml_error(error: impl fmt::Display) -> Error {
    Error::Xml(error.to_string())
}
fn value_error(error: ValueError) -> Error {
    match error {
        ValueError::Limit {
            kind,
            observed,
            maximum,
        } => Error::limit(kind, observed, maximum),
        error => Error::Invalid(error.to_string()),
    }
}
