use std::collections::HashMap;
use std::sync::Arc;

use litchi_core::Position;
use litchi_core::xml::ReaderOrigin;
use litchi_drawingml::ink as shared;
use litchi_ooxml_common::xml_name::is_ncname;
use litchi_opc::constants::relationship_type as rt;
use litchi_opc::{
    OpcPackage, PackURI, PartData, PartReadSession, PartView, Relationships, SourceBackedPackage,
};
use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, NamespaceResolver, ResolveResult};
use quick_xml::reader::NsReader;

use super::codec::{Form, scan};
use super::model::Payload;
use super::trace::{self, Channel, TraceFormat};
use super::{Annotation, CONTENT_TYPE, Limits, Location, Snapshot};
use crate::package::story::{StoryDialect, StoryKind, capture};
use crate::{Error, Package, Result};
use litchi_ooxml_common::xml::attributes::BytesStartExt as _;

const STRICT_CUSTOM_XML: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships/customXml";
const MAX_PROFILE_ATTRIBUTES: usize = 256;
const EMMA_NAMESPACE: &str = "http://www.w3.org/2003/04/emma";

impl Package {
    /// Inventory active InkML annotations across all reachable Word stories.
    ///
    /// Shared targets are parsed once. Payloads and annotation snapshots retain
    /// immutable source storage after the package is dropped. No links are
    /// fetched, no handwriting is interpreted, and package state is unchanged.
    /// The projection exposes context/brush metadata and trace counts. Generic
    /// non-Ink content is left to its own owner; Word's product-specific
    /// compatibility for those content parts is not validated here.
    ///
    /// # Errors
    ///
    /// Returns an error for dirty facade state, invalid ownership, malformed
    /// Ink anchors/payloads, or exhausted resource budgets.
    pub fn ink(&self) -> Result<Snapshot> {
        self.ink_with_limits(Limits::default())
    }

    /// Inventory InkML using an explicit finite resource policy.
    ///
    /// # Errors
    ///
    /// Returns an error for invalid limits, dirty state, malformed package/XML
    /// content, or exceeded limits. A failed inventory never changes the package.
    pub fn ink_with_limits(&self, limits: Limits) -> Result<Snapshot> {
        let limits = limits.validate()?;
        self.ensure_story_opc_current("ink_with_limits")?;
        load(self.opc_package(), limits)
    }
}

pub(crate) fn load(package: &OpcPackage, limits: Limits) -> Result<Snapshot> {
    let owner = PackageRef::Owned(package);
    check_catalog(owner, limits)?;
    let stories = capture(package, limits.stories)?;
    load_stories(
        owner,
        limits,
        stories.dialect(),
        None,
        stories
            .stories()
            .iter()
            .map(|story| (story.part(), story.kind(), story.source())),
    )
}

pub(crate) fn load_source(package: &SourceBackedPackage, limits: Limits) -> Result<Snapshot> {
    let limits = limits.validate()?;
    let owner = PackageRef::Pinned(package);
    check_catalog(owner, limits)?;
    let mut session = package.read_session();
    let stories = crate::package::story::capture_source(package, limits.stories, &mut session)?;
    load_stories(
        owner,
        limits,
        stories.dialect(),
        Some(&mut session),
        stories
            .stories()
            .iter()
            .map(|story| (story.part(), story.kind(), story.source())),
    )
}

fn check_catalog(package: PackageRef<'_>, limits: Limits) -> Result<()> {
    check(
        "package parts",
        package.part_count(),
        limits.stories.max_package_parts,
    )?;
    let mut relationships = package.rels().len();
    check("relationships", relationships, limits.max_relationships)?;
    for part in package.parts() {
        relationships = relationships
            .checked_add(part.rels().len())
            .ok_or_else(|| exceeded("relationships", usize::MAX, limits.max_relationships))?;
        check("relationships", relationships, limits.max_relationships)?;
    }
    Ok(())
}

/// Select the MS-ODRAWXML Ink content-part profile before generic typed
/// projection. This bounded pass does not replace the shared reader: it
/// supplies the semantic spans that it may decode.
///
/// The shared reader intentionally remains a namespace-aware generic InkML
/// projection.  A Word Ink content part has one additional graph contract:
/// trace references must resolve to local definitions, and Microsoft context
/// records are recognized only in the normative EMMA/annotationXML/traceGroup
/// placement.  Unknown namespaces and extension elements are not interpreted
/// here, so a source-preserving no-op can retain them byte-for-byte. The
/// returned source spans identify the subset that the DOCX owner exposes
/// after the generic DrawingML projection has been parsed. Elements outside
/// those spans remain structurally scanned by the shared reader but are not
/// decoded as typed metadata.
pub(crate) fn validate_content_part(xml: &[u8]) -> Result<ProfileProjection> {
    check(
        "DOCX Ink profile source bytes",
        xml.len(),
        shared::MAX_SOURCE_BYTES,
    )?;
    let mut reader = NsReader::from_reader(xml);
    let origin = ReaderOrigin::of(xml);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = true;
    reader.config_mut().check_comments = true;
    reader.resolver_mut().set_max_declarations_per_element(256);
    let mut stack = Vec::new();
    let mut definitions = Vec::new();
    let mut traces = Vec::new();
    let mut formats = Vec::new();
    let mut channel_properties = Vec::new();
    let mut projection = ProfileProjection::default();
    let mut root_seen = false;
    let mut root_closed = false;
    let mut nodes = 0usize;

    loop {
        let start = position(&reader, origin)?;
        let event = reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        let (namespace, event) = reader.resolver().resolve_event(event);
        let end = position(&reader, origin)?;
        match event {
            Event::Start(element) => {
                if root_closed {
                    return Err(Error::Invalid(
                        "DOCX Ink content part has content after its root".into(),
                    ));
                }
                validate_profile_attributes(&element)?;
                increment_profile_nodes(&mut nodes)?;
                enforce_profile_depth(stack.len().saturating_add(1))?;
                let kind = profile_kind(&element, &namespace);
                let frame = if !root_seen {
                    if kind != ProfileKind::Root {
                        return Err(Error::Invalid(
                            "DOCX Ink content part root is not namespace-bound InkML".into(),
                        ));
                    }
                    root_seen = true;
                    ProfileFrame::recognized(kind, start)
                } else {
                    observe_profile_element(
                        &element,
                        kind,
                        &namespace,
                        reader.resolver(),
                        &mut stack,
                        &mut definitions,
                        &mut traces,
                        &mut formats,
                        &mut channel_properties,
                        start,
                        end,
                    )?
                };
                stack.try_reserve(1).map_err(|source| Error::Allocation {
                    resource: "DOCX Ink profile XML stack",
                    source,
                })?;
                stack.push(frame);
            },
            Event::Empty(element) => {
                if root_closed {
                    return Err(Error::Invalid(
                        "DOCX Ink content part has content after its root".into(),
                    ));
                }
                validate_profile_attributes(&element)?;
                increment_profile_nodes(&mut nodes)?;
                enforce_profile_depth(stack.len().saturating_add(1))?;
                let kind = profile_kind(&element, &namespace);
                let frame = if !root_seen {
                    if kind != ProfileKind::Root {
                        return Err(Error::Invalid(
                            "DOCX Ink content part root is not namespace-bound InkML".into(),
                        ));
                    }
                    root_seen = true;
                    ProfileFrame::recognized(kind, start)
                } else {
                    observe_profile_element(
                        &element,
                        kind,
                        &namespace,
                        reader.resolver(),
                        &mut stack,
                        &mut definitions,
                        &mut traces,
                        &mut formats,
                        &mut channel_properties,
                        start,
                        end,
                    )?
                };
                complete_profile_frame(&frame)?;
                record_projection(
                    frame.projection,
                    shared::SourceSpan::new(start, end),
                    &mut projection,
                )?;
                finish_trace_frame(&frame, end, &mut traces)?;
                if kind == ProfileKind::Root && stack.is_empty() {
                    root_closed = true;
                }
            },
            Event::End(_) => {
                let frame = stack.pop().ok_or_else(|| {
                    Error::Invalid("DOCX Ink content part has an unexpected end".into())
                })?;
                complete_profile_frame(&frame)?;
                record_projection(
                    frame.projection,
                    shared::SourceSpan::new(frame.start, end),
                    &mut projection,
                )?;
                finish_trace_frame(&frame, start, &mut traces)?;
                if frame.kind == ProfileKind::Root && stack.is_empty() {
                    root_closed = true;
                }
            },
            Event::Eof => break,
            Event::Decl(_)
            | Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::DocType(_)
            | Event::PI(_)
            | Event::GeneralRef(_) => {},
        }
    }

    if !root_seen || !root_closed || !stack.is_empty() {
        return Err(Error::Invalid(
            "DOCX Ink content part root is absent or unterminated".into(),
        ));
    }
    for (format_index, format) in formats.iter().enumerate() {
        let Some(context_definition) = format.context_definition else {
            continue;
        };
        let definition = definitions.get_mut(context_definition).ok_or_else(|| {
            Error::Invalid("DOCX Ink traceFormat context owner is out of range".into())
        })?;
        if definition.trace_format.replace(format_index).is_some() {
            return Err(Error::Invalid(
                "DOCX Ink context contains more than one traceFormat".into(),
            ));
        }
    }
    validate_definition_ids(&mut definitions)?;
    for format in &formats {
        if format.regular.is_empty() {
            return Err(Error::Invalid(
                "DOCX Ink traceFormat must declare at least one regular channel".into(),
            ));
        }
    }
    validate_channel_properties(&channel_properties, &formats)?;
    for trace in &traces {
        let context = local_reference(trace.context.as_deref(), "contextRef")?;
        find_definition(&definitions, DefinitionKind::Context, context).ok_or_else(|| {
            Error::Invalid(format!(
                "DOCX Ink trace contextRef has no matching context definition: {context}"
            ))
        })?;
        let brush = local_reference(trace.brush.as_deref(), "brushRef")?;
        if !has_definition(&definitions, DefinitionKind::Brush, brush) {
            return Err(Error::Invalid(format!(
                "DOCX Ink trace brushRef has no matching brush definition: {brush}"
            )));
        }
    }
    validate_trace_streams(xml, &definitions, &formats, &traces)?;
    Ok(projection)
}

fn finish_trace_frame(
    frame: &ProfileFrame,
    data_end: usize,
    traces: &mut [TraceReferences],
) -> Result<()> {
    let Some(index) = frame.trace else {
        return Ok(());
    };
    let trace = traces.get_mut(index).ok_or_else(|| {
        Error::Invalid("DOCX Ink trace frame points outside its trace records".into())
    })?;
    trace.data = shared::SourceSpan::new(frame.data_start, data_end);
    Ok(())
}

fn validate_channel_properties(
    properties: &[ChannelPropertyReference],
    formats: &[TraceFormat],
) -> Result<()> {
    for property in properties {
        let format_index = property.format.ok_or_else(|| {
            Error::Invalid("DOCX Ink channelProperty has no active traceFormat owner".into())
        })?;
        let format = formats.get(format_index).ok_or_else(|| {
            Error::Invalid("DOCX Ink channelProperty traceFormat owner is out of range".into())
        })?;
        if !format
            .regular
            .iter()
            .any(|channel| channel.name == property.channel)
        {
            return Err(Error::Invalid(format!(
                "DOCX Ink channelProperty refers to an undefined channel: {}",
                property.channel
            )));
        }
    }
    Ok(())
}

fn validate_trace_streams(
    xml: &[u8],
    definitions: &[Definition],
    formats: &[TraceFormat],
    traces: &[TraceReferences],
) -> Result<()> {
    let default_format = TraceFormat::default_xy();
    let mut reader = NsReader::from_reader(xml);
    let origin = ReaderOrigin::of(xml);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = true;
    reader.config_mut().check_comments = true;
    reader.resolver_mut().set_max_declarations_per_element(256);
    let mut next_trace = 0usize;
    let mut active: Option<(usize, trace::Validator<'_>)> = None;

    loop {
        let start = position(&reader, origin)?;
        let event = reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        let end = position(&reader, origin)?;
        match event {
            Event::Start(_) => {
                if active.is_some() {
                    return Err(Error::Invalid(
                        "DOCX Ink trace character data cannot contain nested elements".into(),
                    ));
                }
                if let Some(record) = traces.get(next_trace)
                    && record.data.start() == end
                {
                    let format = trace_format_for(record, definitions, formats, &default_format)?;
                    active = Some((next_trace, trace::Validator::new(format)?));
                }
            },
            Event::Empty(_) => {
                if active.is_some() {
                    return Err(Error::Invalid(
                        "DOCX Ink trace character data cannot contain nested elements".into(),
                    ));
                }
                if let Some(record) = traces.get(next_trace)
                    && record.data.start() == end
                {
                    let format = trace_format_for(record, definitions, formats, &default_format)?;
                    let mut validator = trace::Validator::new(format)?;
                    validator.finish()?;
                    next_trace = next_trace
                        .checked_add(1)
                        .ok_or_else(|| Error::Invalid("DOCX Ink trace count overflowed".into()))?;
                }
            },
            Event::End(_) => {
                if let Some((index, mut validator)) = active.take() {
                    let record = traces.get(index).ok_or_else(|| {
                        Error::Invalid("DOCX Ink trace stream record is out of range".into())
                    })?;
                    if record.data.end() != start {
                        return Err(Error::Invalid(
                            "DOCX Ink trace character data has an unexpected closing element"
                                .into(),
                        ));
                    }
                    validator.finish()?;
                    next_trace = next_trace
                        .checked_add(1)
                        .ok_or_else(|| Error::Invalid("DOCX Ink trace count overflowed".into()))?;
                }
            },
            Event::Text(text) => {
                if let Some((_, validator)) = active.as_mut() {
                    let text = text
                        .xml_content(XmlVersion::Explicit1_0)
                        .map_err(|error| Error::Xml(error.to_string()))?;
                    let text = quick_xml::escape::unescape(&text)
                        .map_err(|error| Error::Xml(error.to_string()))?;
                    validator.feed(text.as_bytes())?;
                }
            },
            Event::CData(text) => {
                if let Some((_, validator)) = active.as_mut() {
                    validator.feed(text.as_ref())?;
                }
            },
            Event::GeneralRef(reference) => {
                if let Some((_, validator)) = active.as_mut() {
                    feed_trace_reference(reference.as_ref(), validator)?;
                }
            },
            Event::Eof => break,
            Event::Decl(_) | Event::Comment(_) | Event::DocType(_) | Event::PI(_) => {},
        }
    }
    if active.is_some() || next_trace != traces.len() {
        return Err(Error::Invalid(
            "DOCX Ink trace stream records do not match recognized trace elements".into(),
        ));
    }
    Ok(())
}

fn trace_format_for<'a>(
    trace: &TraceReferences,
    definitions: &[Definition],
    formats: &'a [TraceFormat],
    default: &'a TraceFormat,
) -> Result<&'a TraceFormat> {
    let context = local_reference(trace.context.as_deref(), "contextRef")?;
    let definition =
        find_definition(definitions, DefinitionKind::Context, context).ok_or_else(|| {
            Error::Invalid(format!(
                "DOCX Ink trace contextRef has no matching context definition: {context}"
            ))
        })?;
    Ok(definition
        .trace_format
        .and_then(|index| formats.get(index))
        .unwrap_or(default))
}

fn feed_trace_reference(reference: &[u8], validator: &mut trace::Validator<'_>) -> Result<()> {
    let decoded = match reference {
        b"amp" => u32::from(b'&'),
        b"lt" => u32::from(b'<'),
        b"gt" => u32::from(b'>'),
        b"apos" => u32::from(b'\''),
        b"quot" => u32::from(b'"'),
        value if value.starts_with(b"#x") || value.starts_with(b"#X") => {
            parse_trace_reference(&value[2..], 16)?
        },
        value if value.starts_with(b"#") => parse_trace_reference(&value[1..], 10)?,
        _ => {
            return Err(Error::Invalid(
                "DOCX Ink trace uses an undeclared XML entity reference".into(),
            ));
        },
    };
    let byte = u8::try_from(decoded).map_err(|_| {
        Error::Invalid("DOCX Ink trace entity reference is not an ASCII grammar character".into())
    })?;
    validator.feed(&[byte])
}

fn parse_trace_reference(value: &[u8], radix: u32) -> Result<u32> {
    if value.is_empty() {
        return Err(Error::Invalid(
            "DOCX Ink numeric entity reference has no digits".into(),
        ));
    }
    let mut result = 0u32;
    for &byte in value {
        let digit = match byte {
            b'0'..=b'9' => u32::from(byte - b'0'),
            b'a'..=b'f' if radix == 16 => u32::from(byte - b'a' + 10),
            b'A'..=b'F' if radix == 16 => u32::from(byte - b'A' + 10),
            _ => {
                return Err(Error::Invalid(
                    "DOCX Ink numeric entity reference has an invalid digit".into(),
                ));
            },
        };
        if digit >= radix {
            return Err(Error::Invalid(
                "DOCX Ink numeric entity reference has an invalid digit".into(),
            ));
        }
        result = result
            .checked_mul(radix)
            .and_then(|value| value.checked_add(digit))
            .ok_or_else(|| {
                Error::Invalid("DOCX Ink numeric entity reference is too large".into())
            })?;
    }
    Ok(result)
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum ProfileKind {
    Root,
    Definitions,
    Context,
    Brush,
    InkSource,
    TraceFormat,
    IntermittentChannels,
    Channel,
    ChannelProperties,
    ChannelProperty,
    Trace,
    TraceGroup,
    BrushProperty,
    SourceLink,
    DestinationLink,
    AnnotationXml,
    Emma,
    Interpretation,
    MicrosoftContext,
    Other,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum DefinitionKind {
    Context,
    Brush,
    InkSource,
}

struct Definition {
    kind: DefinitionKind,
    value: String,
    trace_format: Option<usize>,
}

struct TraceReferences {
    context: Option<String>,
    brush: Option<String>,
    data: shared::SourceSpan,
}

struct ChannelPropertyReference {
    format: Option<usize>,
    channel: String,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum ProjectionKind {
    Context,
    Trace,
    BrushProperty,
    Link,
}

#[derive(Default)]
pub(crate) struct ProfileProjection {
    contexts: Vec<shared::SourceSpan>,
    traces: Vec<shared::SourceSpan>,
    brush_properties: Vec<shared::SourceSpan>,
    links: Vec<shared::SourceSpan>,
}

impl ProfileProjection {
    pub(crate) fn contexts(&self) -> &[shared::SourceSpan] {
        &self.contexts
    }

    pub(crate) fn traces(&self) -> &[shared::SourceSpan] {
        &self.traces
    }

    pub(crate) fn brush_properties(&self) -> &[shared::SourceSpan] {
        &self.brush_properties
    }

    pub(crate) fn links(&self) -> &[shared::SourceSpan] {
        &self.links
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
struct ProfileFrame {
    kind: ProfileKind,
    recognized: bool,
    opaque: bool,
    projection: Option<ProjectionKind>,
    start: usize,
    child_elements: usize,
    required_child: bool,
    intermittent_channels: bool,
    definition: Option<usize>,
    context_definition: Option<usize>,
    format: Option<usize>,
    trace: Option<usize>,
    data_start: usize,
}

impl ProfileFrame {
    const fn recognized(kind: ProfileKind, start: usize) -> Self {
        Self {
            kind,
            recognized: true,
            opaque: false,
            projection: None,
            start,
            child_elements: 0,
            required_child: false,
            intermittent_channels: false,
            definition: None,
            context_definition: None,
            format: None,
            trace: None,
            data_start: start,
        }
    }

    const fn recognized_projection(
        kind: ProfileKind,
        start: usize,
        projection: ProjectionKind,
    ) -> Self {
        Self {
            kind,
            recognized: true,
            opaque: false,
            projection: Some(projection),
            start,
            child_elements: 0,
            required_child: false,
            intermittent_channels: false,
            definition: None,
            context_definition: None,
            format: None,
            trace: None,
            data_start: start,
        }
    }

    const fn ignored(kind: ProfileKind, start: usize) -> Self {
        Self {
            kind,
            recognized: false,
            opaque: true,
            projection: None,
            start,
            child_elements: 0,
            required_child: false,
            intermittent_channels: false,
            definition: None,
            context_definition: None,
            format: None,
            trace: None,
            data_start: start,
        }
    }
}

fn profile_kind(element: &BytesStart<'_>, namespace: &ResolveResult<'_>) -> ProfileKind {
    let local = element.local_name();
    if is_bound_namespace(namespace, shared::INKML_NAMESPACE) {
        return match local.as_ref() {
            b"ink" => ProfileKind::Root,
            b"definitions" => ProfileKind::Definitions,
            b"context" => ProfileKind::Context,
            b"brush" => ProfileKind::Brush,
            b"inkSource" => ProfileKind::InkSource,
            b"traceFormat" => ProfileKind::TraceFormat,
            b"intermittentChannels" => ProfileKind::IntermittentChannels,
            b"channel" => ProfileKind::Channel,
            b"channelProperties" => ProfileKind::ChannelProperties,
            b"channelProperty" => ProfileKind::ChannelProperty,
            b"trace" => ProfileKind::Trace,
            b"traceGroup" => ProfileKind::TraceGroup,
            b"brushProperty" => ProfileKind::BrushProperty,
            b"annotationXML" => ProfileKind::AnnotationXml,
            _ => ProfileKind::Other,
        };
    }
    if is_bound_namespace(namespace, "http://www.w3.org/2003/04/emma") {
        return match local.as_ref() {
            b"emma" => ProfileKind::Emma,
            b"interpretation" => ProfileKind::Interpretation,
            _ => ProfileKind::Other,
        };
    }
    if is_bound_namespace(namespace, shared::NAMESPACE) {
        return match local.as_ref() {
            b"context" => ProfileKind::MicrosoftContext,
            b"sourceLink" => ProfileKind::SourceLink,
            b"destinationLink" => ProfileKind::DestinationLink,
            _ => ProfileKind::Other,
        };
    }
    ProfileKind::Other
}

fn is_bound_namespace(namespace: &ResolveResult<'_>, expected: &str) -> bool {
    matches!(
        namespace,
        ResolveResult::Bound(Namespace(value)) if *value == expected.as_bytes()
    )
}

fn observe_profile_element(
    element: &BytesStart<'_>,
    kind: ProfileKind,
    _namespace: &ResolveResult<'_>,
    resolver: &NamespaceResolver,
    stack: &mut [ProfileFrame],
    definitions: &mut Vec<Definition>,
    traces: &mut Vec<TraceReferences>,
    formats: &mut Vec<TraceFormat>,
    channel_properties: &mut Vec<ChannelPropertyReference>,
    start: usize,
    data_start: usize,
) -> Result<ProfileFrame> {
    let Some(parent_index) = stack.len().checked_sub(1) else {
        return Err(Error::Invalid(
            "DOCX Ink content part has an element outside its root".into(),
        ));
    };
    let child_index = stack[parent_index].child_elements;
    stack[parent_index].child_elements = child_index.checked_add(1).ok_or_else(|| {
        exceeded(
            "DOCX Ink profile child elements",
            usize::MAX,
            shared::MAX_NODES,
        )
    })?;
    let parent = stack[parent_index];
    if parent.opaque {
        return Ok(ProfileFrame::ignored(kind, start));
    }
    match kind {
        ProfileKind::Definitions if parent.recognized && parent.kind == ProfileKind::Root => {
            Ok(ProfileFrame::recognized(kind, start))
        },
        ProfileKind::Context | ProfileKind::Brush
            if parent.recognized && parent.kind == ProfileKind::Definitions =>
        {
            if definitions.len() >= shared::MAX_NODES {
                return Err(exceeded(
                    "DOCX Ink profile definition records",
                    definitions.len().saturating_add(1),
                    shared::MAX_NODES,
                ));
            }
            if let Some(value) = xml_id(element)? {
                if !is_ncname(&value) {
                    return Err(Error::Invalid(
                        "DOCX Ink definition xml:id is not an XML NCName".into(),
                    ));
                }
                definitions
                    .try_reserve(1)
                    .map_err(|source| Error::Allocation {
                        resource: "DOCX Ink definition identifiers",
                        source,
                    })?;
                let definition = definitions.len();
                definitions.push(Definition {
                    kind: if kind == ProfileKind::Context {
                        DefinitionKind::Context
                    } else {
                        DefinitionKind::Brush
                    },
                    value,
                    trace_format: None,
                });
                let mut frame = ProfileFrame::recognized(kind, start);
                frame.definition = Some(definition);
                return Ok(frame);
            }
            Ok(ProfileFrame::recognized(kind, start))
        },
        ProfileKind::InkSource if parent.recognized && parent.kind == ProfileKind::Context => {
            let value = xml_id(element)?
                .ok_or_else(|| Error::Invalid("DOCX Ink inkSource xml:id is missing".into()))?;
            if !is_ncname(&value) {
                return Err(Error::Invalid(
                    "DOCX Ink inkSource xml:id is not an XML NCName".into(),
                ));
            }
            let context_definition = parent.definition;
            if definitions.len() >= shared::MAX_NODES {
                return Err(exceeded(
                    "DOCX Ink profile definition records",
                    definitions.len().saturating_add(1),
                    shared::MAX_NODES,
                ));
            }
            definitions
                .try_reserve(1)
                .map_err(|source| Error::Allocation {
                    resource: "DOCX Ink definition identifiers",
                    source,
                })?;
            let definition = definitions.len();
            definitions.push(Definition {
                kind: DefinitionKind::InkSource,
                value,
                trace_format: None,
            });
            let mut frame = ProfileFrame::recognized(kind, start);
            frame.definition = Some(definition);
            frame.context_definition = context_definition;
            Ok(frame)
        },
        ProfileKind::TraceFormat if parent.recognized && parent.kind == ProfileKind::InkSource => {
            if parent.format.is_some() {
                return Err(Error::Invalid(
                    "DOCX Ink inkSource contains more than one traceFormat".into(),
                ));
            }
            let context_definition = parent.context_definition;
            if formats.len() >= shared::MAX_NODES {
                return Err(exceeded(
                    "DOCX Ink traceFormat records",
                    formats.len().saturating_add(1),
                    shared::MAX_NODES,
                ));
            }
            formats.try_reserve(1).map_err(|source| Error::Allocation {
                resource: "DOCX Ink traceFormat records",
                source,
            })?;
            let format = formats.len();
            formats.push(TraceFormat {
                context_definition,
                ..TraceFormat::default()
            });
            stack[parent_index].format = Some(format);
            let mut frame = ProfileFrame::recognized(kind, start);
            frame.format = Some(format);
            frame.context_definition = context_definition;
            Ok(frame)
        },
        ProfileKind::IntermittentChannels
            if parent.recognized && parent.kind == ProfileKind::TraceFormat =>
        {
            if parent.intermittent_channels {
                return Err(Error::Invalid(
                    "DOCX Ink traceFormat contains more than one intermittentChannels section"
                        .into(),
                ));
            }
            let format = parent.format.ok_or_else(|| {
                Error::Invalid("DOCX Ink intermittentChannels owner is out of range".into())
            })?;
            formats
                .get_mut(format)
                .ok_or_else(|| {
                    Error::Invalid("DOCX Ink intermittentChannels owner is out of range".into())
                })?
                .opaque_intermittent = true;
            stack[parent_index].intermittent_channels = true;
            // The Office profile explicitly ignores intermittentChannels and
            // their channel declarations. Keep the subtree opaque while the
            // enclosing traceFormat still tracks its grammar-level position.
            Ok(ProfileFrame::ignored(kind, start))
        },
        ProfileKind::Channel
            if parent.recognized
                && matches!(
                    parent.kind,
                    ProfileKind::TraceFormat | ProfileKind::IntermittentChannels
                ) =>
        {
            if parent.kind == ProfileKind::TraceFormat && parent.intermittent_channels {
                return Err(Error::Invalid(
                    "DOCX Ink traceFormat regular channel appears after intermittentChannels"
                        .into(),
                ));
            }
            if parent.kind == ProfileKind::IntermittentChannels {
                return Ok(ProfileFrame::ignored(kind, start));
            }
            let format = parent.format.ok_or_else(|| {
                Error::Invalid("DOCX Ink channel is missing its traceFormat owner".into())
            })?;
            let name = profile_attr(element, b"name")?.ok_or_else(|| {
                Error::Invalid("DOCX Ink traceFormat channel name is missing".into())
            })?;
            let channel_kind = trace::channel_type(profile_attr(element, b"type")?.as_deref())?;
            let target = formats.get_mut(format).ok_or_else(|| {
                Error::Invalid("DOCX Ink channel traceFormat owner is out of range".into())
            })?;
            if target.total_channels() >= shared::MAX_NODES {
                return Err(exceeded(
                    "DOCX Ink traceFormat channels",
                    target.total_channels().saturating_add(1),
                    shared::MAX_NODES,
                ));
            }
            target
                .regular
                .try_reserve(1)
                .map_err(|source| Error::Allocation {
                    resource: "DOCX Ink regular channels",
                    source,
                })?;
            target.push(Channel {
                name,
                kind: channel_kind,
            });
            Ok(ProfileFrame::recognized(ProfileKind::Channel, start))
        },
        ProfileKind::ChannelProperties
            if parent.recognized && parent.kind == ProfileKind::InkSource =>
        {
            let mut frame = ProfileFrame::recognized(kind, start);
            frame.format = parent.format;
            Ok(frame)
        },
        ProfileKind::ChannelProperty
            if parent.recognized && parent.kind == ProfileKind::ChannelProperties =>
        {
            let channel = profile_attr(element, b"channel")?.ok_or_else(|| {
                Error::Invalid("DOCX Ink channelProperty channel is missing".into())
            })?;
            channel_properties
                .try_reserve(1)
                .map_err(|source| Error::Allocation {
                    resource: "DOCX Ink channel properties",
                    source,
                })?;
            channel_properties.push(ChannelPropertyReference {
                format: parent.format,
                channel,
            });
            Ok(ProfileFrame::recognized(kind, start))
        },
        ProfileKind::Trace
            if parent.recognized
                && matches!(parent.kind, ProfileKind::Root | ProfileKind::TraceGroup) =>
        {
            if traces.len() >= shared::MAX_TRACES {
                return Err(exceeded(
                    "DOCX Ink profile trace references",
                    traces.len().saturating_add(1),
                    shared::MAX_TRACES,
                ));
            }
            traces.try_reserve(1).map_err(|source| Error::Allocation {
                resource: "DOCX Ink trace references",
                source,
            })?;
            traces.push(TraceReferences {
                context: profile_attr(element, b"contextRef")?,
                brush: profile_attr(element, b"brushRef")?,
                data: shared::SourceSpan::new(data_start, data_start),
            });
            let trace = traces.len() - 1;
            let mut frame = ProfileFrame::recognized_projection(kind, start, ProjectionKind::Trace);
            frame.trace = Some(trace);
            frame.data_start = data_start;
            Ok(frame)
        },
        ProfileKind::TraceGroup
            if parent.recognized
                && matches!(parent.kind, ProfileKind::Root | ProfileKind::TraceGroup) =>
        {
            Ok(ProfileFrame::recognized(kind, start))
        },
        ProfileKind::BrushProperty if parent.recognized && parent.kind == ProfileKind::Brush => {
            let name = profile_attr(element, b"name")?.ok_or_else(|| {
                Error::Invalid("DOCX Ink brushProperty name attribute is missing".into())
            })?;
            if !is_profile_brush_property(name.as_str()) {
                return Ok(ProfileFrame::ignored(kind, start));
            }
            Ok(ProfileFrame::recognized_projection(
                kind,
                start,
                ProjectionKind::BrushProperty,
            ))
        },
        ProfileKind::SourceLink | ProfileKind::DestinationLink
            if parent.recognized && parent.kind == ProfileKind::MicrosoftContext =>
        {
            Ok(ProfileFrame::recognized_projection(
                kind,
                start,
                ProjectionKind::Link,
            ))
        },
        ProfileKind::AnnotationXml
            if parent.recognized && parent.kind == ProfileKind::TraceGroup =>
        {
            Ok(ProfileFrame::recognized(kind, start))
        },
        ProfileKind::Emma if parent.recognized && parent.kind == ProfileKind::AnnotationXml => {
            if parent.required_child {
                return Ok(ProfileFrame::ignored(kind, start));
            }
            stack[parent_index].required_child = true;
            Ok(ProfileFrame::recognized(kind, start))
        },
        ProfileKind::Interpretation
            if parent.recognized
                && parent.kind == ProfileKind::Emma
                && child_index == 0
                && emma_mode_is_ink(element, resolver)? =>
        {
            stack[parent_index].required_child = true;
            Ok(ProfileFrame::recognized(kind, start))
        },
        ProfileKind::MicrosoftContext => {
            if stack.len() < 5
                || !stack[stack.len() - 1].recognized
                || stack[stack.len() - 1].kind != ProfileKind::Interpretation
                || !stack[stack.len() - 2].recognized
                || stack[stack.len() - 2].kind != ProfileKind::Emma
                || !stack[stack.len() - 3].recognized
                || stack[stack.len() - 3].kind != ProfileKind::AnnotationXml
                || !stack[stack.len() - 4].recognized
                || stack[stack.len() - 4].kind != ProfileKind::TraceGroup
            {
                return Ok(ProfileFrame::ignored(kind, start));
            }
            stack[parent_index].required_child = true;
            Ok(ProfileFrame::recognized_projection(
                kind,
                start,
                ProjectionKind::Context,
            ))
        },
        // A known InkML construct in a non-recognized parent is ignored by
        // the MS-ODRAWXML subset.  The opaque frame also prevents a nested
        // trace or context record from being mistaken for a recognized one.
        _ => Ok(ProfileFrame::ignored(kind, start)),
    }
}

fn complete_profile_frame(frame: &ProfileFrame) -> Result<()> {
    if !frame.recognized {
        return Ok(());
    }
    let required = match frame.kind {
        ProfileKind::AnnotationXml => "emma:emma",
        ProfileKind::Emma => "emma:interpretation as its first child",
        ProfileKind::Interpretation => "msink:context",
        _ => return Ok(()),
    };
    if frame.required_child {
        Ok(())
    } else {
        Err(Error::Invalid(format!(
            "DOCX Ink {} is missing required {}",
            profile_kind_name(frame.kind),
            required
        )))
    }
}

const fn profile_kind_name(kind: ProfileKind) -> &'static str {
    match kind {
        ProfileKind::AnnotationXml => "annotationXML",
        ProfileKind::Emma => "emma:emma",
        ProfileKind::Interpretation => "emma:interpretation",
        _ => "Ink element",
    }
}

fn is_profile_brush_property(name: &str) -> bool {
    matches!(
        name,
        "width"
            | "height"
            | "color"
            | "transparency"
            | "tip"
            | "rasterOp"
            | "antiAliased"
            | "fitToCurve"
            | "ignorePressure"
            | "inkEffects"
            | "anchorX"
            | "anchorY"
            | "scaleFactor"
    )
}

fn emma_mode_is_ink(element: &BytesStart<'_>, resolver: &NamespaceResolver) -> Result<bool> {
    let mut mode = None;
    for attribute in element.checked_attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        let (namespace, local) = resolver.resolve_attribute(attribute.key);
        if local.as_ref() != b"mode"
            || !matches!(
                namespace,
                ResolveResult::Bound(Namespace(value)) if value == EMMA_NAMESPACE.as_bytes()
            )
        {
            continue;
        }
        if mode.is_some() {
            return Err(Error::Invalid(
                "DOCX Ink emma:interpretation mode attribute is duplicated".into(),
            ));
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, element.decoder())
            .map_err(|error| Error::Xml(error.to_string()))?;
        mode = Some(value.into_owned());
    }
    Ok(mode.as_deref() == Some("ink"))
}

fn validate_profile_attributes(element: &BytesStart<'_>) -> Result<()> {
    let mut count = 0usize;
    for attribute in element.checked_attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        count = count.checked_add(1).ok_or_else(|| {
            exceeded(
                "DOCX Ink profile attributes",
                usize::MAX,
                MAX_PROFILE_ATTRIBUTES,
            )
        })?;
        check("DOCX Ink profile attributes", count, MAX_PROFILE_ATTRIBUTES)?;
        check(
            "DOCX Ink profile attribute bytes",
            attribute.value.len(),
            shared::MAX_ATTRIBUTE_VALUE_BYTES,
        )?;
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, element.decoder())
            .map_err(|error| Error::Xml(error.to_string()))?;
        check(
            "DOCX Ink profile decoded attribute bytes",
            value.len(),
            shared::MAX_ATTRIBUTE_VALUE_BYTES,
        )?;
    }
    Ok(())
}

fn increment_profile_nodes(nodes: &mut usize) -> Result<()> {
    *nodes = nodes
        .checked_add(1)
        .ok_or_else(|| exceeded("DOCX Ink profile XML nodes", usize::MAX, shared::MAX_NODES))?;
    check("DOCX Ink profile XML nodes", *nodes, shared::MAX_NODES)
}

fn enforce_profile_depth(depth: usize) -> Result<()> {
    check("DOCX Ink profile XML depth", depth, shared::MAX_DEPTH)
}

fn record_projection(
    kind: Option<ProjectionKind>,
    span: shared::SourceSpan,
    projection: &mut ProfileProjection,
) -> Result<()> {
    let values = match kind {
        Some(ProjectionKind::Context) => &mut projection.contexts,
        Some(ProjectionKind::Trace) => &mut projection.traces,
        Some(ProjectionKind::BrushProperty) => &mut projection.brush_properties,
        Some(ProjectionKind::Link) => &mut projection.links,
        None => return Ok(()),
    };
    let maximum = match kind {
        Some(ProjectionKind::Context) => shared::MAX_CONTEXTS,
        Some(ProjectionKind::Trace) => shared::MAX_TRACES,
        Some(ProjectionKind::BrushProperty) => shared::MAX_BRUSH_PROPERTIES,
        Some(ProjectionKind::Link) => shared::MAX_NODES,
        None => unreachable!(),
    };
    if values.len() >= maximum {
        return Err(exceeded(
            "DOCX Ink semantic projection",
            values.len().saturating_add(1),
            maximum,
        ));
    }
    values.try_reserve(1).map_err(|source| Error::Allocation {
        resource: "DOCX Ink semantic projection",
        source,
    })?;
    values.push(span);
    Ok(())
}

fn xml_id(element: &BytesStart<'_>) -> Result<Option<String>> {
    let mut result = None;
    for attribute in element.checked_attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        if attribute.key.as_ref() != b"xml:id" {
            continue;
        }
        if result.is_some() {
            return Err(Error::Invalid(
                "DOCX Ink definition xml:id is duplicated".into(),
            ));
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, element.decoder())
            .map_err(|error| Error::Xml(error.to_string()))?;
        result = Some(value.into_owned());
    }
    Ok(result)
}

fn profile_attr(element: &BytesStart<'_>, name: &[u8]) -> Result<Option<String>> {
    let mut result = None;
    for attribute in element.checked_attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        if attribute.key.prefix().is_some() || attribute.key.local_name().as_ref() != name {
            continue;
        }
        if result.is_some() {
            return Err(Error::Invalid(format!(
                "DOCX Ink trace attribute '{}' is duplicated",
                String::from_utf8_lossy(name)
            )));
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, element.decoder())
            .map_err(|error| Error::Xml(error.to_string()))?;
        result = Some(value.into_owned());
    }
    Ok(result)
}

fn validate_definition_ids(definitions: &mut [Definition]) -> Result<()> {
    definitions.sort_unstable_by(|left, right| left.value.cmp(&right.value));
    if definitions
        .windows(2)
        .any(|pair| pair[0].value == pair[1].value)
    {
        return Err(Error::Invalid(
            "DOCX Ink definition identifiers must be unique across contexts and brushes".into(),
        ));
    }
    Ok(())
}

fn find_definition<'a>(
    definitions: &'a [Definition],
    kind: DefinitionKind,
    value: &str,
) -> Option<&'a Definition> {
    definitions
        .binary_search_by(|definition| definition.value.as_str().cmp(value))
        .ok()
        .and_then(|index| {
            let definition = &definitions[index];
            (definition.kind == kind).then_some(definition)
        })
}

fn has_definition(definitions: &[Definition], kind: DefinitionKind, value: &str) -> bool {
    find_definition(definitions, kind, value).is_some()
}

fn local_reference<'a>(value: Option<&'a str>, name: &'static str) -> Result<&'a str> {
    let Some(value) = value else {
        return Err(Error::Invalid(format!("DOCX Ink trace {name} is required")));
    };
    let target = value.strip_prefix('#').ok_or_else(|| {
        Error::Invalid(format!(
            "DOCX Ink trace {name} must target a local # identifier"
        ))
    })?;
    if !is_ncname(target) {
        return Err(Error::Invalid(format!(
            "DOCX Ink trace {name} is not a valid local identifier"
        )));
    }
    Ok(target)
}

fn load_stories<'story>(
    package: PackageRef<'_>,
    limits: Limits,
    dialect: StoryDialect,
    mut session: Option<&mut PartReadSession<'_>>,
    stories: impl Iterator<Item = (&'story PackURI, StoryKind, &'story [u8])>,
) -> Result<Snapshot> {
    let mut payloads: HashMap<PackURI, Option<Arc<Payload>>> = HashMap::new();
    let mut annotations = Vec::new();
    let mut total_payload_bytes = 0usize;
    let mut anchor_count = 0usize;
    let mut roles = [0usize; 7];
    for (story_part, story_kind, story_source) in stories {
        package.check()?;
        let role = role_index(story_kind);
        let location = Location {
            kind: story_kind,
            position: Position::new(roles[role]),
        };
        roles[role] += 1;
        let anchors = scan(
            story_source,
            dialect,
            limits.max_xml_nodes,
            limits.max_xml_depth,
            limits.max_annotations.saturating_sub(anchor_count),
        )?;
        anchor_count = anchor_count
            .checked_add(anchors.len())
            .ok_or_else(|| exceeded("annotations", usize::MAX, limits.max_annotations))?;
        check("annotations", anchor_count, limits.max_annotations)?;
        let owner = package.part(story_part)?;
        for anchor in anchors {
            package.check()?;
            let relationship = owner.rels().get(&anchor.relationship_id).ok_or_else(|| {
                Error::InvalidRelationship("DOCX ink anchor relationship is missing".into())
            })?;
            let expected = if matches!(anchor.form, Form::Base) && dialect == StoryDialect::Strict {
                STRICT_CUSTOM_XML
            } else {
                rt::CUSTOM_XML
            };
            if relationship.is_external() || relationship.reltype() != expected {
                return Err(Error::InvalidRelationship(
                    "DOCX ink content part must use the owning story's internal customXml relationship".into(),
                ));
            }
            check(
                "target reference bytes",
                relationship.target_ref().len(),
                4096,
            )?;
            let target = relationship.target_partname()?;
            let part = package.part(&target)?;
            let is_declared_ink = part.content_type() == CONTENT_TYPE;
            if matches!(anchor.form, Form::GenericDrawing) && !is_declared_ink {
                continue;
            }
            if matches!(anchor.form, Form::Drawing) && !is_declared_ink {
                return Err(Error::ContentType {
                    expected: CONTENT_TYPE.into(),
                    actual: part.content_type().into(),
                });
            }
            if !matches!(anchor.form, Form::Base) && dialect == StoryDialect::Strict {
                return Err(Error::Invalid(
                    "Strict Word drawing Ink requires an explicit extension conformance policy"
                        .into(),
                ));
            }
            // Base contentPart also hosts non-Ink XML, which is preserved but
            // does not become an annotation. Word's text/xml variation applies
            // only to this base form.
            if !is_declared_ink && part.content_type() != "text/xml" {
                continue;
            }
            let document = if let Some(document) = payloads.get(part.partname()) {
                document.clone()
            } else {
                // Central-directory size is a preflight only. PartData checks
                // decoded integrity/length before exposing immutable bytes.
                let length = part.length()?;
                check("payload bytes", length, limits.max_payload_bytes)?;
                total_payload_bytes = total_payload_bytes.checked_add(length).ok_or_else(|| {
                    exceeded(
                        "total payload bytes",
                        usize::MAX,
                        limits.max_total_payload_bytes,
                    )
                })?;
                check(
                    "total payload bytes",
                    total_payload_bytes,
                    limits.max_total_payload_bytes,
                )?;
                let bytes = part.data(session.as_deref_mut())?;
                check(
                    "payload bytes",
                    bytes.as_bytes().len(),
                    limits.max_payload_bytes,
                )?;
                if bytes.as_bytes().len() != length {
                    return Err(Error::Invalid(
                        "DOCX InkML payload length changed during inventory".into(),
                    ));
                }
                let document = if has_ink_root(bytes.as_bytes())? {
                    // A declared Ink Content part cannot carry outgoing
                    // relationships to other parts under ECMA §15.2.4.
                    if !part.rels().is_empty() {
                        return Err(Error::InvalidRelationship(
                            "DOCX InkML payload has unsupported outgoing relationships".into(),
                        ));
                    }
                    Some(Arc::new(match bytes {
                        PayloadBytes::Owned(source) => {
                            let projection = validate_content_part(source.as_slice())?;
                            let document = shared::read_shared_with_source_spans(
                                source,
                                projection.contexts(),
                                projection.traces(),
                                projection.brush_properties(),
                                projection.links(),
                            )?;
                            Payload::Owned(document)
                        },
                        PayloadBytes::Pinned(source) => {
                            let projection = validate_content_part(source.as_bytes())?;
                            let metadata = shared::read_metadata_with_source_spans(
                                source.as_bytes(),
                                projection.contexts(),
                                projection.traces(),
                                projection.brush_properties(),
                                projection.links(),
                            )?;
                            Payload::Pinned {
                                metadata,
                                _source: source,
                            }
                        },
                    }))
                } else if is_declared_ink {
                    return Err(Error::Invalid(
                        "DOCX InkML payload has the wrong root namespace".into(),
                    ));
                } else {
                    None
                };
                payloads
                    .try_reserve(1)
                    .map_err(|source| Error::Allocation {
                        resource: "DOCX ink distinct payloads",
                        source,
                    })?;
                payloads.insert(part.partname().clone(), document.clone());
                document
            };
            if let Some(document) = document {
                annotations
                    .try_reserve(1)
                    .map_err(|source| Error::Allocation {
                        resource: "DOCX ink annotations",
                        source,
                    })?;
                annotations.push(Annotation { location, document });
            }
        }
    }
    Ok(Snapshot {
        annotations: Arc::new(annotations),
        distinct_payloads: payloads
            .values()
            .filter(|document| document.is_some())
            .count(),
    })
}

#[derive(Clone, Copy)]
enum PackageRef<'a> {
    Owned(&'a OpcPackage),
    Pinned(&'a SourceBackedPackage),
}

impl<'a> PackageRef<'a> {
    fn check(self) -> Result<()> {
        if let Self::Pinned(package) = self {
            package.check_execution()?;
            package.source_version()?;
        }
        Ok(())
    }

    fn part_count(self) -> usize {
        match self {
            Self::Owned(package) => package.part_count(),
            Self::Pinned(package) => package.iter_parts().count(),
        }
    }

    fn rels(self) -> &'a Relationships {
        match self {
            Self::Owned(package) => package.rels(),
            Self::Pinned(package) => package.rels(),
        }
    }

    fn parts(self) -> impl Iterator<Item = PartRef<'a>> {
        let (owned, pinned) = match self {
            Self::Owned(package) => (Some(package), None),
            Self::Pinned(package) => (None, Some(package)),
        };
        owned
            .into_iter()
            .flat_map(|package| {
                package.iter_parts().map(move |part| PartRef::Owned {
                    package,
                    partname: part.partname(),
                    content_type: part.content_type(),
                    rels: part.rels(),
                })
            })
            .chain(
                pinned
                    .into_iter()
                    .flat_map(SourceBackedPackage::iter_parts)
                    .map(PartRef::Pinned),
            )
    }

    fn part(self, name: &PackURI) -> Result<PartRef<'a>> {
        match self {
            Self::Owned(package) => {
                let part = package.get_part(name)?;
                Ok(PartRef::Owned {
                    package,
                    partname: part.partname(),
                    content_type: part.content_type(),
                    rels: part.rels(),
                })
            },
            Self::Pinned(package) => Ok(PartRef::Pinned(package.part(name)?)),
        }
    }
}

#[derive(Clone, Copy)]
enum PartRef<'a> {
    /// An owned package part. Its payload is decoded on demand through
    /// `OpcPackage::get_part`, which reports a lazy decode's typed refusal
    /// (ADR 0030); iteration itself never reaches a payload.
    Owned {
        package: &'a OpcPackage,
        partname: &'a PackURI,
        content_type: &'a str,
        rels: &'a Relationships,
    },
    Pinned(PartView<'a>),
}

impl<'a> PartRef<'a> {
    fn partname(self) -> &'a PackURI {
        match self {
            Self::Owned { partname, .. } => partname,
            Self::Pinned(part) => part.partname(),
        }
    }

    fn content_type(self) -> &'a str {
        match self {
            Self::Owned { content_type, .. } => content_type,
            Self::Pinned(part) => part.content_type(),
        }
    }

    fn rels(self) -> &'a Relationships {
        match self {
            Self::Owned { rels, .. } => rels,
            Self::Pinned(part) => part.rels(),
        }
    }

    fn length(self) -> Result<usize> {
        match self {
            Self::Owned {
                package, partname, ..
            } => Ok(package.get_part(partname)?.blob().len()),
            Self::Pinned(part) => {
                usize::try_from(part.declared_uncompressed_size()?).map_err(|_| {
                    Error::Invalid("DOCX InkML declared size exceeds address space".into())
                })
            },
        }
    }

    fn data(self, session: Option<&mut PartReadSession<'_>>) -> Result<PayloadBytes> {
        match self {
            Self::Owned {
                package, partname, ..
            } => {
                let part = package.get_part(partname)?;
                let source = part.blob_arc();
                let visible = part.blob();
                if !std::ptr::eq(source.as_slice(), visible) && source.as_slice() != visible {
                    return Err(Error::Invalid(
                        "DOCX InkML source storage is inconsistent".into(),
                    ));
                }
                Ok(PayloadBytes::Owned(source))
            },
            Self::Pinned(part) => Ok(PayloadBytes::Pinned(match session {
                Some(session) => session.read(part)?,
                None => part.data()?,
            })),
        }
    }
}

enum PayloadBytes {
    Owned(Arc<Vec<u8>>),
    Pinned(PartData),
}

impl PayloadBytes {
    fn as_bytes(&self) -> &[u8] {
        match self {
            Self::Owned(bytes) => bytes.as_slice(),
            Self::Pinned(bytes) => bytes.as_bytes(),
        }
    }
}

pub(crate) fn has_ink_root(xml: &[u8]) -> Result<bool> {
    let mut reader = NsReader::from_reader(xml);
    reader.resolver_mut().set_max_declarations_per_element(256);
    for _ in 0..256 {
        let (namespace, event) = reader.read_resolved_event()?;
        match event {
            Event::Start(element) | Event::Empty(element) => {
                return Ok(element.local_name().as_ref() == b"ink"
                    && matches!(namespace, ResolveResult::Bound(Namespace(uri))
                        if uri == shared::INKML_NAMESPACE.as_bytes()));
            },
            Event::Decl(_) | Event::Comment(_) => {},
            Event::Text(text) if text.as_ref().iter().all(u8::is_ascii_whitespace) => {},
            _ => {
                return Err(Error::Invalid(
                    "DOCX content-part XML has an invalid prolog".into(),
                ));
            },
        }
    }
    Err(exceeded("payload prolog events", 257, 256))
}

fn position<R: std::io::BufRead>(reader: &NsReader<R>, origin: ReaderOrigin) -> Result<usize> {
    origin
        .offset(reader.buffer_position())
        .ok_or_else(|| Error::Invalid("DOCX Ink XML offset exceeds usize".into()))
}

const fn role_index(kind: StoryKind) -> usize {
    match kind {
        StoryKind::Main => 0,
        StoryKind::Header => 1,
        StoryKind::Footer => 2,
        StoryKind::Footnotes => 3,
        StoryKind::Endnotes => 4,
        StoryKind::Comments => 5,
        StoryKind::Glossary => 6,
    }
}

fn exceeded(resource: &'static str, actual: usize, maximum: usize) -> Error {
    Error::InkLimit {
        resource,
        actual,
        maximum,
    }
}

fn check(resource: &'static str, actual: usize, maximum: usize) -> Result<()> {
    if actual > maximum {
        Err(exceeded(resource, actual, maximum))
    } else {
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use litchi_opc::constants::content_type as ct;
    use litchi_opc::{BlobPart, Part};

    #[test]
    fn repeated_anchors_share_projection_and_original_payload_allocation() {
        let mut package = OpcPackage::new();
        let mut main = BlobPart::new(
            PackURI::new("/word/document.xml").unwrap(),
            ct::WML_DOCUMENT_MAIN.into(),
            br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:contentPart r:id="ink"/><w:contentPart r:id="ink"/></w:r></w:p></w:body></w:document>"#.to_vec(),
        );
        Part::rels_mut(&mut main).add_relationship(
            rt::CUSTOM_XML.into(),
            "../ink.xml".into(),
            "ink".into(),
            false,
        );
        package.add_part(Box::new(main));
        package.relate_to("word/document.xml", rt::OFFICE_DOCUMENT);
        let payload = Arc::new(
            br##"<i:ink xmlns:i="http://www.w3.org/2003/InkML"><i:definitions><i:context xml:id="ctx0"/><i:brush xml:id="br0"/></i:definitions><i:trace contextRef="#ctx0" brushRef="#br0">1 2</i:trace></i:ink>"##
                .to_vec(),
        );
        package.add_part(Box::new(BlobPart::new_shared(
            PackURI::new("/ink.xml").unwrap(),
            CONTENT_TYPE.into(),
            Arc::clone(&payload),
        )));
        let snapshot = load(&package, Limits::default()).unwrap();
        assert_eq!(snapshot.annotations().len(), 2);
        let first = &snapshot.annotations()[0].document;
        let second = &snapshot.annotations()[1].document;
        assert!(Arc::ptr_eq(first, second));
        assert_eq!(first.source().as_ptr(), payload.as_ptr());
        assert_eq!(first.source(), payload.as_slice());
        let clone = snapshot.clone();
        assert!(Arc::ptr_eq(&snapshot.annotations, &clone.annotations));
        drop(package);
        drop(payload);
        assert_eq!(clone.annotations()[0].trace_count(), 1);
    }

    #[test]
    fn profile_accepts_forward_references() {
        let source = br##"<i:ink xmlns:i="http://www.w3.org/2003/InkML"><i:trace contextRef="#ctx0" brushRef="#br0">1 2</i:trace><i:definitions><i:context xml:id="ctx0"/><i:brush xml:id="br0"/></i:definitions></i:ink>"##;
        validate_content_part(source).unwrap();
    }

    #[test]
    fn profile_validates_trace_format_channels_and_inkml_lexicals() {
        let valid = br##"<i:ink xmlns:i="http://www.w3.org/2003/InkML"><i:definitions><i:context xml:id="ctx0"><i:inkSource xml:id="source0"><i:traceFormat><i:channel name="X" type="integer"/><i:channel name="Y" type="integer"/><i:channel name="T" type="integer" units="ms"/><i:channel name="F" type="boolean"/><i:intermittentChannels><i:channel name="pressure" type="not-a-type"/></i:intermittentChannels></i:traceFormat><i:channelProperties><i:channelProperty channel="T" name="resolution" value="1"/></i:channelProperties></i:inkSource></i:context><i:brush xml:id="br0"/></i:definitions><i:trace contextRef="#ctx0" brushRef="#br0">1 2 0 T,'3 '4 '5 F,"1 "2 "3 *</i:trace></i:ink>"##;
        validate_content_part(valid).expect("valid imported InkML trace grammar");
        for body in [
            "<![CDATA[1 2 0 T,'3 '4 '5 F,\"1 \"2 \"3 *]]>",
            "1&#x20;2 0 T,3 4 5 F",
            "1 2 0 T<!-- preserved comment -->,3 4 5 F",
        ] {
            let source = format!(
                r##"<i:ink xmlns:i="http://www.w3.org/2003/InkML"><i:definitions><i:context xml:id="ctx0"><i:inkSource xml:id="source0"><i:traceFormat><i:channel name="X" type="integer"/><i:channel name="Y" type="integer"/><i:channel name="T" type="integer"/><i:channel name="F" type="boolean"/><i:intermittentChannels><i:channel name="pressure" type="not-a-type"/></i:intermittentChannels></i:traceFormat></i:inkSource></i:context><i:brush xml:id="br0"/></i:definitions><i:trace contextRef="#ctx0" brushRef="#br0">{body}</i:trace></i:ink>"##
            );
            validate_content_part(source.as_bytes()).expect("split XML character data is valid");
        }

        let valid_text = String::from_utf8(valid.to_vec()).unwrap();
        let opaque_intermittent_data = valid_text
            .replace(
                "1 2 0 T,'3 '4 '5 F,\"1 \"2 \"3 *",
                "1 2 0 T 1e1,'3 '4 '5 F 2e1,\"1 \"2 \"3 * ?",
            )
            .into_bytes();
        validate_content_part(&opaque_intermittent_data)
            .expect("ignored intermittent channels remain generic trace data");
        let replace_data = |data: &str| {
            valid_text
                .replace("1 2 0 T,'3 '4 '5 F,\"1 \"2 \"3 *", data)
                .into_bytes()
        };
        let invalid = [
            replace_data("'1 2 0 T"),
            replace_data("1 2 1.5 T"),
            replace_data("1 2 0 T,\"3 4 5 F"),
            replace_data("1 2 0 T,3 4 5 F #"),
            replace_data(r#"1 2 0 T,'3 '4 '5 F,\"1 \"2 \"3 *,'4 '5 '6 F"#),
        ];
        for source in invalid {
            assert!(
                validate_content_part(&source).is_err(),
                "invalid trace data must be rejected: {}",
                String::from_utf8_lossy(&source)
            );
        }

        let invalid_t = valid_text
            .replace("name=\"T\" type=\"integer\"", "name=\"T\" type=\"decimal\"")
            .replace("1 2 0 T", "1 2 0.5 T")
            .into_bytes();
        assert!(validate_content_part(&invalid_t).is_err());
        let boolean_t = valid_text
            .replace("name=\"T\" type=\"integer\"", "name=\"T\" type=\"boolean\"")
            .replace("1 2 0 T", "1 2 T T")
            .into_bytes();
        assert!(validate_content_part(&boolean_t).is_err());
        let invalid_property = valid_text
            .replace("channel=\"T\"", "channel=\"unknown\"")
            .into_bytes();
        assert!(validate_content_part(&invalid_property).is_err());
    }

    #[test]
    fn profile_rejects_duplicate_or_late_intermittent_channel_sections() {
        let duplicate = br##"<i:ink xmlns:i="http://www.w3.org/2003/InkML"><i:definitions><i:context xml:id="ctx0"><i:inkSource xml:id="source0"><i:traceFormat><i:channel name="X" type="integer"/><i:intermittentChannels><i:channel name="ignored" type="not-a-type"/></i:intermittentChannels><i:intermittentChannels><i:channel name="alsoIgnored" type="not-a-type"/></i:intermittentChannels></i:traceFormat></i:inkSource></i:context><i:brush xml:id="br0"/></i:definitions><i:trace contextRef="#ctx0" brushRef="#br0">1</i:trace></i:ink>"##;
        assert!(validate_content_part(duplicate).is_err());

        let late_regular = br##"<i:ink xmlns:i="http://www.w3.org/2003/InkML"><i:definitions><i:context xml:id="ctx0"><i:inkSource xml:id="source0"><i:traceFormat><i:channel name="X" type="integer"/><i:intermittentChannels><i:channel name="ignored" type="not-a-type"/></i:intermittentChannels><i:channel name="Y" type="integer"/></i:traceFormat></i:inkSource></i:context><i:brush xml:id="br0"/></i:definitions><i:trace contextRef="#ctx0" brushRef="#br0">1 2</i:trace></i:ink>"##;
        assert!(validate_content_part(late_regular).is_err());
    }

    #[test]
    fn profile_accepts_t_default_or_decimal_declarations_with_integer_values() {
        for declaration in [
            r#"<i:channel name="T"/>"#,
            r#"<i:channel name="T" type="decimal"/>"#,
        ] {
            let source = format!(
                r##"<i:ink xmlns:i="http://www.w3.org/2003/InkML"><i:definitions><i:context xml:id="ctx0"><i:inkSource xml:id="source0"><i:traceFormat><i:channel name="X" type="integer"/><i:channel name="Y" type="integer"/>{declaration}</i:traceFormat></i:inkSource></i:context><i:brush xml:id="br0"/></i:definitions><i:trace contextRef="#ctx0" brushRef="#br0">1 2 3</i:trace></i:ink>"##
            );
            validate_content_part(source.as_bytes())
                .expect("integer lexical T values are valid for default/decimal channels");
        }
    }

    #[test]
    fn profile_does_not_reject_boolean_t_declarations_without_t_values() {
        let declaration_only = br##"<i:ink xmlns:i="http://www.w3.org/2003/InkML"><i:definitions><i:context xml:id="ctx0"><i:inkSource xml:id="source0"><i:traceFormat><i:channel name="T" type="boolean"/></i:traceFormat></i:inkSource></i:context></i:definitions></i:ink>"##;
        validate_content_part(declaration_only)
            .expect("a declaration alone does not provide a Boolean T value");

        let ignored_intermittent = br##"<i:ink xmlns:i="http://www.w3.org/2003/InkML"><i:definitions><i:context xml:id="ctx0"><i:inkSource xml:id="source0"><i:traceFormat><i:channel name="X" type="integer"/><i:channel name="Y" type="integer"/><i:intermittentChannels><i:channel name="T" type="boolean"/></i:intermittentChannels></i:traceFormat></i:inkSource></i:context><i:brush xml:id="br0"/></i:definitions><i:trace contextRef="#ctx0" brushRef="#br0">1 2 *</i:trace></i:ink>"##;
        validate_content_part(ignored_intermittent)
            .expect("an ignored Boolean T declaration and first-point wildcard remain generic");
    }

    #[test]
    fn profile_ignores_trace_formats_outside_ink_source_owner() {
        let source = br##"<i:ink xmlns:i="http://www.w3.org/2003/InkML"><i:definitions><i:traceFormat><i:channel type="not-a-type"/></i:traceFormat><i:context xml:id="ctx0"/><i:brush xml:id="br0"/></i:definitions><i:trace contextRef="#ctx0" brushRef="#br0">1 2</i:trace></i:ink>"##;
        validate_content_part(source).expect("misplaced traceFormat is opaque");
    }

    #[test]
    fn profile_rejects_missing_dangling_external_and_duplicate_ids() {
        let missing = br##"<i:ink xmlns:i="http://www.w3.org/2003/InkML"><i:definitions><i:context xml:id="ctx0"/><i:brush xml:id="br0"/></i:definitions><i:trace brushRef="#br0">1 2</i:trace></i:ink>"##;
        assert!(validate_content_part(missing).is_err());

        let dangling = br##"<i:ink xmlns:i="http://www.w3.org/2003/InkML"><i:definitions><i:context xml:id="ctx0"/><i:brush xml:id="br0"/></i:definitions><i:trace contextRef="#other" brushRef="#br0">1 2</i:trace></i:ink>"##;
        assert!(validate_content_part(dangling).is_err());

        let external = br##"<i:ink xmlns:i="http://www.w3.org/2003/InkML"><i:definitions><i:context xml:id="ctx0"/><i:brush xml:id="br0"/></i:definitions><i:trace contextRef="/other.xml#ctx0" brushRef="#br0">1 2</i:trace></i:ink>"##;
        assert!(validate_content_part(external).is_err());

        let duplicate = br##"<i:ink xmlns:i="http://www.w3.org/2003/InkML"><i:definitions><i:context xml:id="same"/><i:brush xml:id="same"/></i:definitions></i:ink>"##;
        assert!(validate_content_part(duplicate).is_err());
    }

    #[test]
    fn profile_ignores_misplaced_context_and_normative_extension_subtrees() {
        let misplaced = br##"<i:ink xmlns:i="http://www.w3.org/2003/InkML" xmlns:m="http://schemas.microsoft.com/ink/2010/main"><i:traceGroup><m:context type="writingRegion"/></i:traceGroup></i:ink>"##;
        validate_content_part(misplaced).unwrap();

        let ignored_trace = br##"<i:ink xmlns:i="http://www.w3.org/2003/InkML"><i:definitions><i:trace>1 2</i:trace></i:definitions></i:ink>"##;
        validate_content_part(ignored_trace).unwrap();

        let ignored_context = br##"<i:ink xmlns:i="http://www.w3.org/2003/InkML" xmlns:m="http://schemas.microsoft.com/ink/2010/main" xmlns:e="http://www.w3.org/2003/04/emma"><i:traceGroup><i:annotationXML><e:emma><e:interpretation e:mode="ink"><e:group><m:context type="writingRegion"/></e:group><m:context type="writingRegion"/></e:interpretation></e:emma></i:annotationXML></i:traceGroup></i:ink>"##;
        validate_content_part(ignored_context).unwrap();
    }

    #[test]
    fn profile_requires_emma_first_interpretation_mode_and_context() {
        let valid = br##"<i:ink xmlns:i="http://www.w3.org/2003/InkML" xmlns:m="http://schemas.microsoft.com/ink/2010/main" xmlns:e="http://www.w3.org/2003/04/emma"><i:traceGroup><i:annotationXML><e:emma><e:interpretation e:mode="ink"><m:context type="writingRegion"/></e:interpretation><e:group><m:context type="line"/></e:group></e:emma></i:annotationXML></i:traceGroup></i:ink>"##;
        validate_content_part(valid).expect("normative EMMA subset");

        let invalid: &[&[u8]] = &[
            br##"<i:ink xmlns:i="http://www.w3.org/2003/InkML" xmlns:m="http://schemas.microsoft.com/ink/2010/main" xmlns:e="http://www.w3.org/2003/04/emma"><i:traceGroup><i:annotationXML><e:emma><e:interpretation><m:context type="writingRegion"/></e:interpretation></e:emma></i:annotationXML></i:traceGroup></i:ink>"##,
            br##"<i:ink xmlns:i="http://www.w3.org/2003/InkML" xmlns:m="http://schemas.microsoft.com/ink/2010/main" xmlns:e="http://www.w3.org/2003/04/emma"><i:traceGroup><i:annotationXML><e:emma><e:interpretation e:mode="speech"><m:context type="writingRegion"/></e:interpretation></e:emma></i:annotationXML></i:traceGroup></i:ink>"##,
            br##"<i:ink xmlns:i="http://www.w3.org/2003/InkML" xmlns:m="http://schemas.microsoft.com/ink/2010/main" xmlns:e="http://www.w3.org/2003/04/emma"><i:traceGroup><i:annotationXML><e:emma><e:group/><e:interpretation e:mode="ink"><m:context type="writingRegion"/></e:interpretation></e:emma></i:annotationXML></i:traceGroup></i:ink>"##,
            br##"<i:ink xmlns:i="http://www.w3.org/2003/InkML" xmlns:m="http://schemas.microsoft.com/ink/2010/main" xmlns:e="http://www.w3.org/2003/04/emma"><i:traceGroup><i:annotationXML><e:emma><e:interpretation e:mode="ink"/></e:emma></i:annotationXML></i:traceGroup></i:ink>"##,
            br##"<i:ink xmlns:i="http://www.w3.org/2003/InkML"><i:traceGroup><i:annotationXML/></i:traceGroup></i:ink>"##,
        ];
        for source in invalid {
            assert!(
                validate_content_part(source).is_err(),
                "profile must reject invalid EMMA boundary: {}",
                String::from_utf8_lossy(source)
            );
        }

        let alternate_prefix = br##"<i:ink xmlns:i="http://www.w3.org/2003/InkML" xmlns:m="http://schemas.microsoft.com/ink/2010/main" xmlns:q="http://www.w3.org/2003/04/emma"><i:traceGroup><i:annotationXML><q:emma><q:interpretation q:mode="ink"><m:context type="writingRegion"/></q:interpretation></q:emma></i:annotationXML></i:traceGroup></i:ink>"##;
        validate_content_part(alternate_prefix)
            .expect("EMMA namespace semantics are prefix independent");
    }

    #[test]
    fn profile_ignores_unknown_brush_properties_but_exposes_effective_defaults() {
        let source = br##"<i:ink xmlns:i="http://www.w3.org/2003/InkML"><i:definitions><i:brush xml:id="br0"><i:brushProperty name="futureProperty" value="opaque"/><i:brushProperty name="width" value="not-a-decimal" units="bogus"/></i:brush></i:definitions></i:ink>"##;
        let projection = validate_content_part(source).expect("profile structure");
        assert_eq!(projection.brush_properties().len(), 1);
        let document = shared::read_shared_with_source_spans(
            Arc::new(source.to_vec()),
            projection.contexts(),
            projection.traces(),
            projection.brush_properties(),
            projection.links(),
        )
        .expect("recognized brush property");
        assert_eq!(document.brush_property_count(), 1);
        let property = &document.brush_properties()[0];
        assert_eq!(property.value(), "not-a-decimal");
        let effective = property.effective().expect("profile property default");
        assert_eq!(effective.value(), ".053");
        assert_eq!(effective.units(), Some("cm"));
        assert!(effective.defaulted());
        assert_eq!(document.source(), source);
    }

    #[test]
    fn profile_applies_normative_units_and_schema_whitespace_to_brush_values() {
        let source = br##"<i:ink xmlns:i="http://www.w3.org/2003/InkML"><i:definitions><i:brush><i:brushProperty name="width" value="&#x20;1.25&#x9;" units="m"/><i:brushProperty name="height" value="1" units="px"/><i:brushProperty name="transparency" value="&#xA;42&#xD;"/><i:brushProperty name="antiAliased" value="&#x20;&#x31;&#x9;"/></i:brush></i:definitions></i:ink>"##;
        let projection = validate_content_part(source).expect("profile structure");
        let document = shared::read_shared_with_source_spans(
            Arc::new(source.to_vec()),
            projection.contexts(),
            projection.traces(),
            projection.brush_properties(),
            projection.links(),
        )
        .expect("recognized brush properties");
        let effective: Vec<_> = document
            .brush_properties()
            .iter()
            .map(|property| property.effective())
            .collect();

        assert_eq!(effective[0].as_ref().unwrap().value(), "1.25");
        assert_eq!(effective[0].as_ref().unwrap().units(), Some("m"));
        assert!(!effective[0].as_ref().unwrap().defaulted());
        assert_eq!(effective[1].as_ref().unwrap().value(), ".001");
        assert_eq!(effective[1].as_ref().unwrap().units(), Some("cm"));
        assert!(effective[1].as_ref().unwrap().defaulted());
        assert_eq!(effective[2].as_ref().unwrap().value(), "42");
        assert!(!effective[2].as_ref().unwrap().defaulted());
        assert_eq!(effective[3].as_ref().unwrap().value(), "true");
        assert!(!effective[3].as_ref().unwrap().defaulted());
        assert_eq!(document.source(), source);
    }

    #[test]
    fn profile_projection_excludes_ignored_typed_elements_and_nested_roots() {
        let source = br##"<i:ink xmlns:i="http://www.w3.org/2003/InkML" xmlns:m="http://schemas.microsoft.com/ink/2010/main" xmlns:e="http://www.w3.org/2003/04/emma"><i:definitions><i:trace>ignored definitions trace</i:trace><i:context xml:id="ctx0"/><i:brush xml:id="br0"><i:brushProperty name="inkEffects" value="pencil"/></i:brush></i:definitions><i:ink/><i:ink><i:trace>ignored nested root trace</i:trace></i:ink><i:traceGroup><i:annotationXML><e:emma><e:interpretation e:mode="ink"><m:context type="writingRegion"/></e:interpretation><e:group><m:context type="line"/></e:group></e:emma></i:annotationXML><i:trace contextRef="#ctx0" brushRef="#br0">1 2</i:trace></i:traceGroup><i:brushProperty name="color" value="#FFFFFF"/></i:ink>"##;
        let projection = validate_content_part(source).unwrap();
        let filtered = shared::read_shared_with_source_spans(
            Arc::new(source.to_vec()),
            projection.contexts(),
            projection.traces(),
            projection.brush_properties(),
            projection.links(),
        )
        .unwrap();
        assert_eq!(filtered.source(), source);
        assert_eq!(filtered.context_count(), 1);
        assert_eq!(filtered.trace_count(), 1);
        assert_eq!(filtered.brush_property_count(), 1);
        assert_eq!(
            filtered.contexts()[0].xml(&filtered),
            br#"<m:context type="writingRegion"/>"#
        );
        assert_eq!(
            filtered.traces()[0].xml(&filtered),
            br##"<i:trace contextRef="#ctx0" brushRef="#br0">1 2</i:trace>"##
        );
        assert_eq!(
            filtered.brush_properties()[0].xml(&filtered),
            br##"<i:brushProperty name="inkEffects" value="pencil"/>"##
        );
    }

    #[test]
    fn ignored_ancestry_skips_invalid_typed_values_but_keeps_bounded_scan() {
        let source = br##"<i:ink xmlns:i="http://www.w3.org/2003/InkML" xmlns:m="http://schemas.microsoft.com/ink/2010/main" xmlns:e="http://www.w3.org/2003/04/emma"><i:definitions><i:opaque><m:sourceLink direction="bad" ref="bad"/><m:context/><m:context type="not-a-guid"/><m:context type="writingRegion" id="not-a-guid"/><m:context type="writingRegion" rotatedBoundingBox="not points"/><i:brushProperty/></i:opaque><i:context xml:id="ctx0"/><i:brush xml:id="br0"><i:brushProperty name="inkEffects" value="pencil"/></i:brush></i:definitions><i:traceGroup><i:annotationXML><e:emma><e:interpretation e:mode="ink"><m:context type="writingRegion"/></e:interpretation></e:emma></i:annotationXML></i:traceGroup></i:ink>"##;
        let projection = validate_content_part(source).unwrap();
        let document = shared::read_shared_with_source_spans(
            Arc::new(source.to_vec()),
            projection.contexts(),
            projection.traces(),
            projection.brush_properties(),
            projection.links(),
        )
        .expect("ignored typed values must remain opaque");
        assert_eq!(document.context_count(), 1);
        assert_eq!(document.trace_count(), 0);
        assert_eq!(document.brush_property_count(), 1);
        assert_eq!(document.source(), source);
    }

    #[test]
    fn recognized_invalid_typed_values_are_still_rejected() {
        let invalid_contexts = [
            r#"<m:context/>"#,
            r#"<m:context type="not-a-guid"/>"#,
            r#"<m:context type="writingRegion" id="not-a-guid"/>"#,
            r#"<m:context type="writingRegion" rotatedBoundingBox="not points"/>"#,
        ];
        for context in invalid_contexts {
            let source = format!(
                r#"<i:ink xmlns:i="http://www.w3.org/2003/InkML" xmlns:m="http://schemas.microsoft.com/ink/2010/main" xmlns:e="http://www.w3.org/2003/04/emma"><i:traceGroup><i:annotationXML><e:emma><e:interpretation e:mode="ink">{context}</e:interpretation></e:emma></i:annotationXML></i:traceGroup></i:ink>"#
            );
            let projection = validate_content_part(source.as_bytes()).unwrap();
            assert!(
                shared::read_metadata_with_source_spans(
                    source.as_bytes(),
                    projection.contexts(),
                    projection.traces(),
                    projection.brush_properties(),
                    projection.links(),
                )
                .is_err(),
                "recognized context must be typed and rejected: {context}"
            );
        }

        let source = br##"<i:ink xmlns:i="http://www.w3.org/2003/InkML" xmlns:m="http://schemas.microsoft.com/ink/2010/main" xmlns:e="http://www.w3.org/2003/04/emma"><i:traceGroup><i:annotationXML><e:emma><e:interpretation e:mode="ink"><m:context type="writingRegion"><m:sourceLink direction="bad" ref="bad"/></m:context></e:interpretation></e:emma></i:annotationXML></i:traceGroup></i:ink>"##;
        let projection = validate_content_part(source).unwrap();
        assert!(
            shared::read_metadata_with_source_spans(
                source,
                projection.contexts(),
                projection.traces(),
                projection.brush_properties(),
                projection.links(),
            )
            .is_err()
        );

        let source = br##"<i:ink xmlns:i="http://www.w3.org/2003/InkML"><i:definitions><i:brush><i:brushProperty/></i:brush></i:definitions></i:ink>"##;
        assert!(validate_content_part(source).is_err());
    }
}
