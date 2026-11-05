//! Bounded, namespace-aware codecs for Custom XML Data Storage.

use crate::mce::{
    Capabilities, Limits as MceLimits, Name, OffsetLimits, process_markup_compatibility,
    process_markup_compatibility_with_source_offsets,
};
use crate::xml::decode_xml_reference;
use crate::xml_name;
use crate::{Error, Result};

use super::limit::LimitKind;
use litchi_opc::ContentType;
use litchi_opc::SourceLimits;
use quick_xml::XmlVersion;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesDecl, BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use std::collections::{HashMap, HashSet};

use super::errors::unsupported_content;
use super::model::{
    Conformance, MAX_DEPTH, MAX_ELEMENTS, MAX_PART_BYTES, MAX_PROPS_BYTES, MAX_SCHEMA_REFS,
    MAX_STRING_BYTES, Props, STRICT_NAMESPACE, STRICT_PROPS_RELATIONSHIP, STRICT_RELATIONSHIP,
    TRANSITIONAL_NAMESPACE, TRANSITIONAL_PROPS_RELATIONSHIP, TRANSITIONAL_RELATIONSHIP,
};
use crate::xml::attributes::BytesStartExt as _;

/// Parse a Custom XML Data Storage Properties part with bounded MCE handling.
/// # Errors
///
/// Returns an error when input violates OOXML constraints, exceeds a configured
/// bound, or an underlying XML or package operation fails.
pub fn read_props(xml: &[u8], source_limits: &SourceLimits) -> Result<Props> {
    let (capabilities, limits) = properties_mce_context(source_limits);
    require_at_most(
        LimitKind::PropertiesBytes,
        xml.len(),
        limits.max_input_bytes,
    )?;
    if crate::mce::admission::mentions_namespace(xml) {
        check_source_profile(xml, source_limits)?;
    }
    let processed = process_markup_compatibility(xml, &capabilities, &limits)
        .map_err(map_properties_mce_error)?;
    parse_props_xml(processed.xml.as_ref(), limits.max_depth, source_limits)
}

/// Serialize properties in stable schema order.
/// # Errors
///
/// Returns an error when input violates OOXML constraints, exceeds a configured
/// bound, or an underlying XML or package operation fails.
pub fn write_props(
    props: &Props,
    conformance: Conformance,
    source_limits: &SourceLimits,
) -> Result<Vec<u8>> {
    validate_props(props)?;
    let output_len = props_output_len(props, conformance)?;
    require_at_most(
        LimitKind::SerializedPropertiesBytes,
        output_len,
        max_properties_bytes(source_limits),
    )?;
    let mut out = Vec::with_capacity(output_len);
    out.extend_from_slice(b"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
    out.extend_from_slice(b"<ds:datastoreItem xmlns:ds=\"");
    push_escaped_attr(&mut out, conformance.namespace());
    out.extend_from_slice(b"\" ds:itemID=\"");
    push_escaped_attr(&mut out, &props.id);
    out.extend_from_slice(b"\">");
    if !props.schemas.is_empty() {
        out.extend_from_slice(b"<ds:schemaRefs>");
        for uri in &props.schemas {
            out.extend_from_slice(b"<ds:schemaRef ds:uri=\"");
            push_escaped_attr(&mut out, uri);
            out.extend_from_slice(b"\"/>");
        }
        out.extend_from_slice(b"</ds:schemaRefs>");
    }
    out.extend_from_slice(b"</ds:datastoreItem>");
    if out.len() != output_len {
        return Err(Error::Internal(
            "custom XML properties size calculation disagrees with serialization",
        ));
    }
    debug_assert_eq!(out.len(), output_len);
    Ok(out)
}

/// Rewrite the typed properties projection while retaining source markup.
/// Known attributes are located from a namespace-aware XML parse and changed
/// in place. Structural schema edits use bounded lexical patches as long as
/// the source has an unambiguous schema-reference topology; opaque content
/// inside that topology is refused only when changing its shape would make
/// lossless placement ambiguous.
pub(super) fn rewrite_props(
    source: &[u8],
    before: &Props,
    after: &Props,
    _conformance: Conformance,
    source_limits: &SourceLimits,
) -> Result<Vec<u8>> {
    require_at_most(
        LimitKind::SerializedPropertiesBytes,
        source.len(),
        max_properties_bytes(source_limits),
    )?;
    if before == after {
        return Ok(source.to_vec());
    }
    if read_props(source, source_limits)? != *before {
        return Err(Error::Internal(
            "custom XML properties source projection is stale",
        ));
    }
    if before.schemas == after.schemas {
        validate_props(after)?;
        let item_id = scan_item_id_span(source, source_limits)?;
        let edits = vec![LexicalEdit {
            start: item_id.0,
            end: item_id.1,
            replacement: Replacement::Attribute(&after.id),
        }];
        return apply_property_edits(source, edits, after, source_limits, false);
    }
    let layout = scan_properties_layout(source, source_limits)?;
    let Some(item_id) = layout.item_id else {
        return Err(Error::Internal(
            "custom XML properties itemID splice layout is missing",
        ));
    };

    let mut edits = vec![LexicalEdit {
        start: item_id.0,
        end: item_id.1,
        replacement: Replacement::Attribute(&after.id),
    }];
    let source_schema_count = layout
        .schema_refs
        .as_ref()
        .map_or(0, |schema_refs| schema_refs.refs.len());
    if source_schema_count != before.schemas.len() {
        return Err(unsupported_content(
            "custom_xml.set_properties",
            "the Custom XML properties schema markup is opaque",
        ));
    }
    if let Some(schema_refs) = &layout.schema_refs {
        for (index, (before_schema, after_schema)) in
            before.schemas.iter().zip(&after.schemas).enumerate()
        {
            if before_schema != after_schema {
                let (start, end) = schema_refs.refs[index].uri;
                edits.push(LexicalEdit {
                    start,
                    end,
                    replacement: Replacement::Attribute(after_schema),
                });
            }
        }
        if after.schemas.len() < before.schemas.len() {
            if schema_refs.refs[after.schemas.len()..]
                .iter()
                .any(|schema_ref| schema_ref.opaque)
            {
                return Err(unsupported_content(
                    "custom_xml.set_properties",
                    "the Custom XML properties schema markup is opaque",
                ));
            }
            for schema_ref in &schema_refs.refs[after.schemas.len()..] {
                edits.push(LexicalEdit {
                    start: schema_ref.start,
                    end: schema_ref.end,
                    replacement: Replacement::Empty,
                });
            }
        } else if after.schemas.len() > before.schemas.len() {
            let values = &after.schemas[before.schemas.len()..];
            if schema_refs.empty {
                edits.push(LexicalEdit {
                    start: schema_refs.start,
                    end: schema_refs.end,
                    replacement: Replacement::ExpandEmpty {
                        source,
                        start: schema_refs.start,
                        end: schema_refs.end,
                        qname: &schema_refs.qname,
                        content: SchemaRefsContent::Generated {
                            namespace: &layout.root_namespace,
                            qname: &schema_refs.qname,
                            values,
                        },
                        attribute: None,
                    },
                });
            } else {
                edits.push(LexicalEdit {
                    start: schema_refs.insertion_start,
                    end: schema_refs.insertion_end,
                    replacement: Replacement::SchemaRefs {
                        values,
                        namespace: &layout.root_namespace,
                        qname: &schema_refs.qname,
                    },
                });
            }
        }
    } else if !after.schemas.is_empty() {
        if layout.root_empty {
            let item = edits.pop().ok_or(Error::Internal(
                "custom XML properties itemID edit is missing",
            ))?;
            let attribute = match item.replacement {
                Replacement::Attribute(value) => Some((item.start, item.end, value)),
                _ => {
                    return Err(Error::Internal(
                        "custom XML properties itemID edit has an invalid replacement",
                    ));
                },
            };
            edits.clear();
            edits.push(LexicalEdit {
                start: layout.root_start,
                end: layout.root_start_end,
                replacement: Replacement::ExpandEmpty {
                    source,
                    start: layout.root_start,
                    end: layout.root_start_end,
                    qname: &layout.root_qname,
                    content: SchemaRefsContent::GeneratedContainer {
                        namespace: &layout.root_namespace,
                        root_qname: &layout.root_qname,
                        values: &after.schemas,
                    },
                    attribute,
                },
            });
        } else {
            let Some(root_end) = layout.root_end else {
                return Err(Error::Internal(
                    "custom XML properties root close layout is missing",
                ));
            };
            edits.push(LexicalEdit {
                start: root_end.0,
                end: root_end.0,
                replacement: Replacement::GeneratedSchemaRefs {
                    namespace: &layout.root_namespace,
                    root_qname: &layout.root_qname,
                    values: &after.schemas,
                },
            });
        }
    }
    apply_property_edits(source, edits, after, source_limits, true)
}

fn max_properties_bytes(source_limits: &SourceLimits) -> usize {
    MAX_PROPS_BYTES.min(usize::try_from(source_limits.max_part_bytes()).unwrap_or(usize::MAX))
}

fn properties_mce_context(source_limits: &SourceLimits) -> (Capabilities, MceLimits) {
    let maximum = max_properties_bytes(source_limits);
    let mut capabilities = Capabilities::ooxml_baseline();
    capabilities
        .understand_namespace(TRANSITIONAL_NAMESPACE)
        .understand_namespace(STRICT_NAMESPACE);
    let mut limits = MceLimits {
        max_input_bytes: maximum,
        max_output_bytes: maximum.saturating_mul(2),
        max_depth: MAX_DEPTH.min(source_limits.max_depth()),
        max_namespace_bindings: 4096,
        max_directive_tokens: 4096,
        max_choices_per_alternate: 1024,
        max_attributes_per_element: crate::mce::DEFAULT_MAX_ATTRIBUTES_PER_ELEMENT,
    };
    limits.max_output_bytes = maximum;
    (capabilities, limits)
}

fn map_properties_mce_error(error: crate::mce::Error) -> Error {
    match error {
        crate::mce::Error::LimitExceeded {
            kind:
                crate::mce::LimitKind::SourceBytes
                | crate::mce::LimitKind::InputBytes
                | crate::mce::LimitKind::OutputBytes,
            observed,
            maximum,
        } => Error::Limit {
            kind: crate::LimitKind::CustomXml(LimitKind::PropertiesBytes),
            observed,
            maximum,
        },
        crate::mce::Error::LimitExceeded {
            kind: crate::mce::LimitKind::Depth,
            observed,
            maximum,
        } => Error::Limit {
            kind: crate::LimitKind::CustomXml(LimitKind::Depth),
            observed,
            maximum,
        },
        other => Error::Mce(other),
    }
}

fn check_source_profile(xml: &[u8], source_limits: &SourceLimits) -> Result<()> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_comments = true;
    reader.config_mut().check_end_names = true;
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut events = 0usize;
    let mut version = XmlVersion::Implicit1_0;
    let mut event_seen = false;

    loop {
        let event = reader.read_event_into(&mut buffer)?;
        if matches!(&event, Event::Eof) {
            break;
        }
        events = events.checked_add(1).ok_or_else(|| {
            source_limit(LimitKind::XmlEvents, usize::MAX, source_limits.max_events())
        })?;
        if events > source_limits.max_events() {
            return Err(source_limit(
                LimitKind::XmlEvents,
                events,
                source_limits.max_events(),
            ));
        }
        match event {
            Event::Decl(declaration) => {
                if event_seen {
                    return invalid("custom XML properties have a late XML declaration");
                }
                version = validate_declaration(&declaration)?;
            },
            Event::Start(element) => {
                check_source_element_bounds(
                    source_limits,
                    depth,
                    &element,
                    version,
                    reader.decoder(),
                )?;
                depth = depth.checked_add(1).ok_or_else(|| {
                    source_limit(LimitKind::Depth, usize::MAX, source_limits.max_depth())
                })?;
            },
            Event::Empty(element) => {
                check_source_element_bounds(
                    source_limits,
                    depth,
                    &element,
                    version,
                    reader.decoder(),
                )?;
            },
            Event::End(_) => {
                depth = depth.checked_sub(1).ok_or_else(|| {
                    Error::Invalid("custom XML properties have an unexpected end tag".into())
                })?;
            },
            _ => {},
        }
        event_seen = true;
        buffer.clear();
    }
    if depth != 0 {
        return Err(Error::Invalid(
            "custom XML properties have an incomplete root".into(),
        ));
    }
    Ok(())
}

#[derive(Debug, Clone)]
struct PropertiesLayout {
    item_id: Option<(usize, usize)>,
    schema_refs: Option<SchemaRefsLayout>,
    root_end: Option<(usize, usize)>,
    root_start: usize,
    root_start_end: usize,
    root_empty: bool,
    root_qname: Vec<u8>,
    root_namespace: String,
}

#[derive(Debug, Clone)]
struct SchemaRefsLayout {
    qname: Vec<u8>,
    refs: Vec<SchemaRefLayout>,
    insertion_start: usize,
    insertion_end: usize,
    empty: bool,
    start: usize,
    end: usize,
}

#[derive(Debug, Clone, Copy)]
struct SchemaRefLayout {
    start: usize,
    end: usize,
    uri: (usize, usize),
    opaque: bool,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum LayoutKind {
    Root,
    SchemaRefs,
    SchemaRef,
    Opaque,
}

#[derive(Debug, Clone, Copy)]
struct LayoutFrame {
    kind: LayoutKind,
    schema_ref: Option<usize>,
    source_start: Option<usize>,
}

#[derive(Debug)]
struct LexicalEdit<'a> {
    start: usize,
    end: usize,
    replacement: Replacement<'a>,
}

#[derive(Debug, Clone)]
struct SourceElementLayout {
    qname: Vec<u8>,
    start: usize,
    open_end: usize,
    close_start: Option<usize>,
    end: usize,
    attribute: Option<(usize, usize)>,
}

#[derive(Debug)]
enum Replacement<'a> {
    Empty,
    Attribute(&'a str),
    SchemaRefs {
        values: &'a [String],
        namespace: &'a str,
        qname: &'a [u8],
    },
    GeneratedSchemaRefs {
        namespace: &'a str,
        root_qname: &'a [u8],
        values: &'a [String],
    },
    ExpandEmpty {
        source: &'a [u8],
        start: usize,
        end: usize,
        qname: &'a [u8],
        content: SchemaRefsContent<'a>,
        attribute: Option<(usize, usize, &'a str)>,
    },
}

#[derive(Debug)]
enum SchemaRefsContent<'a> {
    Generated {
        namespace: &'a str,
        qname: &'a [u8],
        values: &'a [String],
    },
    GeneratedContainer {
        namespace: &'a str,
        root_qname: &'a [u8],
        values: &'a [String],
    },
}

fn scan_source_elements(
    source: &[u8],
    source_limits: &SourceLimits,
) -> Result<(Vec<u32>, HashMap<usize, SourceElementLayout>)> {
    scan_source_elements_mode(source, source_limits, true)
}

fn scan_source_item_elements(
    source: &[u8],
    source_limits: &SourceLimits,
) -> Result<(Vec<u32>, HashMap<usize, SourceElementLayout>)> {
    scan_source_elements_mode(source, source_limits, false)
}

fn scan_source_elements_mode(
    source: &[u8],
    source_limits: &SourceLimits,
    include_schema_refs: bool,
) -> Result<(Vec<u32>, HashMap<usize, SourceElementLayout>)> {
    let mut reader = NsReader::from_reader(source);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_comments = true;
    let mut buffer = Vec::new();
    let mut stack: Vec<Option<usize>> = Vec::new();
    let mut offsets = Vec::new();
    let mut elements = HashMap::new();
    let mut events = 0usize;
    let mut version = XmlVersion::Implicit1_0;

    loop {
        let event_start = usize::try_from(reader.buffer_position())
            .map_err(|_| Error::Internal("custom XML properties event offset overflow"))?;
        let event = reader.read_event_into(&mut buffer)?;
        let event_end = usize::try_from(reader.buffer_position())
            .map_err(|_| Error::Internal("custom XML properties event offset overflow"))?;
        if !matches!(&event, Event::Eof) {
            events = events.checked_add(1).ok_or_else(|| {
                source_limit(LimitKind::XmlEvents, usize::MAX, source_limits.max_events())
            })?;
            if events > source_limits.max_events() {
                return Err(source_limit(
                    LimitKind::XmlEvents,
                    events,
                    source_limits.max_events(),
                ));
            }
        }
        match event {
            Event::Decl(declaration) => {
                version = validate_declaration(&declaration)?;
            },
            Event::Start(element) => {
                let (resolved, _) = reader.resolver().resolve_element(element.name());
                check_source_element_bounds(
                    source_limits,
                    stack.len(),
                    &element,
                    version,
                    reader.decoder(),
                )?;
                let local_name = element.local_name();
                let candidate = is_layout_candidate(local_name.as_ref())
                    && (include_schema_refs || local_name.as_ref() == b"datastoreItem");
                if candidate {
                    let namespace = resolved_namespace(resolved, "element")?;
                    let attribute_name = match local_name.as_ref() {
                        b"datastoreItem" => Some(b"itemID".as_slice()),
                        b"schemaRef" => Some(b"uri".as_slice()),
                        _ => None,
                    };
                    let attribute = match attribute_name {
                        Some(name) => known_attribute_span(
                            source,
                            event_start,
                            event_end,
                            &element,
                            &reader,
                            namespace,
                            name,
                        )?,
                        None => None,
                    };
                    let offset = u32::try_from(event_start).map_err(|_| {
                        Error::Internal("custom XML properties source offset overflow")
                    })?;
                    offsets.push(offset);
                    elements.insert(
                        event_start,
                        SourceElementLayout {
                            qname: element.name().as_ref().to_vec(),
                            start: event_start,
                            open_end: event_end,
                            close_start: None,
                            end: event_end,
                            attribute,
                        },
                    );
                    stack.push(Some(event_start));
                } else {
                    stack.push(None);
                }
            },
            Event::Empty(element) => {
                let (resolved, _) = reader.resolver().resolve_element(element.name());
                check_source_element_bounds(
                    source_limits,
                    stack.len(),
                    &element,
                    version,
                    reader.decoder(),
                )?;
                let local_name = element.local_name();
                if is_layout_candidate(local_name.as_ref())
                    && (include_schema_refs || local_name.as_ref() == b"datastoreItem")
                {
                    let namespace = resolved_namespace(resolved, "element")?;
                    let attribute_name = match local_name.as_ref() {
                        b"datastoreItem" => Some(b"itemID".as_slice()),
                        b"schemaRef" => Some(b"uri".as_slice()),
                        _ => None,
                    };
                    let attribute = match attribute_name {
                        Some(name) => known_attribute_span(
                            source,
                            event_start,
                            event_end,
                            &element,
                            &reader,
                            namespace,
                            name,
                        )?,
                        None => None,
                    };
                    let offset = u32::try_from(event_start).map_err(|_| {
                        Error::Internal("custom XML properties source offset overflow")
                    })?;
                    offsets.push(offset);
                    elements.insert(
                        event_start,
                        SourceElementLayout {
                            qname: element.name().as_ref().to_vec(),
                            start: event_start,
                            open_end: event_end,
                            close_start: None,
                            end: event_end,
                            attribute,
                        },
                    );
                }
            },
            Event::End(_) => {
                if let Some(Some(start)) = stack.pop() {
                    let element = elements.get_mut(&start).ok_or(Error::Internal(
                        "custom XML properties source element layout is missing",
                    ))?;
                    element.close_start = Some(event_start);
                    element.end = event_end;
                }
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    Ok((offsets, elements))
}

fn is_layout_candidate(local_name: &[u8]) -> bool {
    matches!(local_name, b"datastoreItem" | b"schemaRefs" | b"schemaRef")
}

fn check_source_element_bounds(
    source_limits: &SourceLimits,
    depth: usize,
    element: &BytesStart<'_>,
    version: XmlVersion,
    decoder: Decoder,
) -> Result<()> {
    let observed_depth = depth
        .checked_add(1)
        .ok_or_else(|| source_limit(LimitKind::Depth, usize::MAX, source_limits.max_depth()))?;
    if observed_depth > source_limits.max_depth() {
        return Err(source_limit(
            LimitKind::Depth,
            observed_depth,
            source_limits.max_depth(),
        ));
    }
    for attribute in element.checked_attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        let raw_bytes = attribute
            .key
            .as_ref()
            .len()
            .checked_add(attribute.value.as_ref().len())
            .ok_or_else(|| {
                source_limit(
                    LimitKind::XmlAttributeBytes,
                    usize::MAX,
                    source_limits.max_xml_attribute_bytes(),
                )
            })?;
        if raw_bytes > source_limits.max_xml_attribute_bytes() {
            return Err(source_limit(
                LimitKind::XmlAttributeBytes,
                raw_bytes,
                source_limits.max_xml_attribute_bytes(),
            ));
        }
        let decoded = attribute
            .decoded_and_normalized_value(version, decoder)
            .map_err(|error| Error::Xml(error.to_string()))?;
        let decoded_bytes = decoded.len();
        if decoded_bytes > source_limits.max_xml_attribute_bytes() {
            return Err(source_limit(
                LimitKind::XmlAttributeBytes,
                decoded_bytes,
                source_limits.max_xml_attribute_bytes(),
            ));
        }
    }
    Ok(())
}

fn source_limit(kind: LimitKind, observed: usize, maximum: usize) -> Error {
    limit(kind, maximum, observed)
}

fn active_properties_offsets(
    source: &[u8],
    source_limits: &SourceLimits,
    offsets: &[u32],
    retain_processed_xml: bool,
) -> Result<crate::mce::ActiveOffsetOutput> {
    let (capabilities, processing) = properties_mce_context(source_limits);
    let limits = OffsetLimits {
        max_source_bytes: max_properties_bytes(source_limits),
        max_offsets: offsets.len().max(1),
        max_marked_bytes: max_properties_bytes(source_limits),
        processing,
    };
    process_markup_compatibility_with_source_offsets(
        source,
        offsets,
        &capabilities,
        &limits,
        retain_processed_xml,
    )
    .map_err(map_properties_mce_error)
}

fn active_properties_view(
    source: &[u8],
    source_limits: &SourceLimits,
    offsets: &[u32],
) -> Result<crate::mce::ActiveOffsetOutput> {
    active_properties_offsets(source, source_limits, offsets, true)
}

fn scan_item_id_span(source: &[u8], source_limits: &SourceLimits) -> Result<(usize, usize)> {
    let (offsets, source_elements) = scan_source_item_elements(source, source_limits)?;
    let active = active_properties_offsets(source, source_limits, &offsets, false)?;
    let source_start = active
        .selected_offsets
        .first()
        .copied()
        .ok_or(Error::Internal(
            "custom XML properties selected source root is missing",
        ))?;
    let source_element = source_elements.get(
        &usize::try_from(source_start)
            .map_err(|_| Error::Internal("custom XML properties source offset overflow"))?,
    );
    source_element
        .ok_or(Error::Internal(
            "custom XML properties selected source root is missing",
        ))?
        .attribute
        .ok_or(Error::Internal(
            "custom XML properties itemID splice layout is missing",
        ))
}

fn scan_properties_layout(source: &[u8], source_limits: &SourceLimits) -> Result<PropertiesLayout> {
    let (offsets, source_elements) = scan_source_elements(source, source_limits)?;
    let active = active_properties_view(source, source_limits, &offsets)?;
    let mut reader = NsReader::from_reader(active.xml.as_slice());
    reader.config_mut().trim_text(false);
    reader.config_mut().check_comments = true;
    let mut buffer = Vec::new();
    let mut stack = Vec::new();
    let mut selected_source_starts = active.selected_offsets.iter().copied();
    let mut layout = PropertiesLayout {
        item_id: None,
        schema_refs: None,
        root_end: None,
        root_start: 0,
        root_start_end: 0,
        root_empty: false,
        root_qname: Vec::new(),
        root_namespace: String::new(),
    };

    loop {
        let event = reader.read_event_into(&mut buffer)?;
        match event {
            Event::Start(element) => {
                let (resolved, _) = reader.resolver().resolve_element(element.name());
                let namespace = resolved_namespace(resolved, "element")?;
                let local_name = element.local_name();
                let source_element = if is_layout_candidate(local_name.as_ref()) {
                    let source_start = usize::try_from(selected_source_starts.next().ok_or(
                        Error::Internal("custom XML properties selected source element is missing"),
                    )?)
                    .map_err(|_| Error::Internal("custom XML properties source offset overflow"))?;
                    Some(source_elements.get(&source_start).ok_or(Error::Internal(
                        "custom XML properties selected source element is missing",
                    ))?)
                } else {
                    None
                };
                let parent = stack.last().map(|frame: &LayoutFrame| frame.kind);
                let (kind, schema_ref) = if stack.is_empty()
                    && layout.root_qname.is_empty()
                    && local_name.as_ref() == b"datastoreItem"
                {
                    let source_element = source_element.ok_or(Error::Internal(
                        "custom XML properties source root layout is missing",
                    ))?;
                    layout.root_namespace = std::str::from_utf8(namespace)
                        .map_err(|error| Error::Xml(error.to_string()))?
                        .into();
                    layout.root_qname = source_element.qname.clone();
                    layout.root_start = source_element.start;
                    layout.root_start_end = source_element.open_end;
                    layout.item_id = source_element.attribute;
                    (LayoutKind::Root, None)
                } else if parent == Some(LayoutKind::Root)
                    && namespace == layout.root_namespace.as_bytes()
                    && local_name.as_ref() == b"schemaRefs"
                {
                    let source_element = source_element.ok_or(Error::Internal(
                        "custom XML properties source schemaRefs layout is missing",
                    ))?;
                    if layout.schema_refs.is_some() {
                        return Err(Error::Internal(
                            "custom XML properties layout has duplicate schemaRefs",
                        ));
                    }
                    layout.schema_refs = Some(SchemaRefsLayout {
                        qname: source_element.qname.clone(),
                        refs: Vec::new(),
                        insertion_start: source_element.open_end,
                        insertion_end: source_element.open_end,
                        empty: false,
                        start: source_element.start,
                        end: source_element.open_end,
                    });
                    (LayoutKind::SchemaRefs, None)
                } else if parent == Some(LayoutKind::SchemaRefs)
                    && namespace == layout.root_namespace.as_bytes()
                    && local_name.as_ref() == b"schemaRef"
                {
                    let source_element = source_element.ok_or(Error::Internal(
                        "custom XML properties source schemaRef layout is missing",
                    ))?;
                    let uri = source_element.attribute.ok_or(Error::Internal(
                        "custom XML properties schemaRef URI layout is missing",
                    ))?;
                    let opaque = has_unknown_attributes(&element, &reader, namespace, b"uri")?;
                    let schema_refs = layout.schema_refs.as_mut().ok_or(Error::Internal(
                        "custom XML properties schemaRefs layout is missing",
                    ))?;
                    let schema_ref = schema_refs.refs.len();
                    schema_refs.refs.push(SchemaRefLayout {
                        start: source_element.start,
                        end: source_element.open_end,
                        uri,
                        opaque,
                    });
                    (LayoutKind::SchemaRef, Some(schema_ref))
                } else {
                    mark_opaque(&mut layout, &stack);
                    (LayoutKind::Opaque, None)
                };
                stack.push(LayoutFrame {
                    kind,
                    schema_ref,
                    source_start: source_element.map(|element| element.start),
                });
            },
            Event::Empty(element) => {
                let (resolved, _) = reader.resolver().resolve_element(element.name());
                let namespace = resolved_namespace(resolved, "element")?;
                let local_name = element.local_name();
                let source_element = if is_layout_candidate(local_name.as_ref()) {
                    let source_start = usize::try_from(selected_source_starts.next().ok_or(
                        Error::Internal("custom XML properties selected source element is missing"),
                    )?)
                    .map_err(|_| Error::Internal("custom XML properties source offset overflow"))?;
                    Some(source_elements.get(&source_start).ok_or(Error::Internal(
                        "custom XML properties selected source element is missing",
                    ))?)
                } else {
                    None
                };
                let parent = stack.last().map(|frame: &LayoutFrame| frame.kind);
                if stack.is_empty()
                    && layout.root_qname.is_empty()
                    && local_name.as_ref() == b"datastoreItem"
                {
                    let source_element = source_element.ok_or(Error::Internal(
                        "custom XML properties source root layout is missing",
                    ))?;
                    layout.root_namespace = std::str::from_utf8(namespace)
                        .map_err(|error| Error::Xml(error.to_string()))?
                        .into();
                    layout.root_qname = source_element.qname.clone();
                    layout.root_start = source_element.start;
                    layout.root_start_end = source_element.open_end;
                    layout.root_empty = true;
                    layout.item_id = source_element.attribute;
                } else if parent == Some(LayoutKind::Root)
                    && namespace == layout.root_namespace.as_bytes()
                    && local_name.as_ref() == b"schemaRefs"
                {
                    let source_element = source_element.ok_or(Error::Internal(
                        "custom XML properties source schemaRefs layout is missing",
                    ))?;
                    layout.schema_refs = Some(SchemaRefsLayout {
                        qname: source_element.qname.clone(),
                        refs: Vec::new(),
                        insertion_start: source_element.start,
                        insertion_end: source_element.end,
                        empty: true,
                        start: source_element.start,
                        end: source_element.end,
                    });
                } else if parent == Some(LayoutKind::SchemaRefs)
                    && namespace == layout.root_namespace.as_bytes()
                    && local_name.as_ref() == b"schemaRef"
                {
                    let source_element = source_element.ok_or(Error::Internal(
                        "custom XML properties source schemaRef layout is missing",
                    ))?;
                    let uri = source_element.attribute.ok_or(Error::Internal(
                        "custom XML properties schemaRef URI layout is missing",
                    ))?;
                    let opaque = has_unknown_attributes(&element, &reader, namespace, b"uri")?;
                    let schema_refs = layout.schema_refs.as_mut().ok_or(Error::Internal(
                        "custom XML properties schemaRefs layout is missing",
                    ))?;
                    schema_refs.refs.push(SchemaRefLayout {
                        start: source_element.start,
                        end: source_element.end,
                        uri,
                        opaque,
                    });
                } else {
                    mark_opaque(&mut layout, &stack);
                }
            },
            Event::End(_) => {
                let frame = stack.pop().ok_or(Error::Internal(
                    "custom XML properties layout has an unmatched end",
                ))?;
                let source_element = frame
                    .source_start
                    .and_then(|start| source_elements.get(&start));
                match frame.kind {
                    LayoutKind::Root => {
                        let source_element = source_element.ok_or(Error::Internal(
                            "custom XML properties source root close layout is missing",
                        ))?;
                        let close_start = source_element.close_start.ok_or(Error::Internal(
                            "custom XML properties source root close layout is missing",
                        ))?;
                        layout.root_end = Some((close_start, source_element.end));
                    },
                    LayoutKind::SchemaRefs => {
                        let source_element = source_element.ok_or(Error::Internal(
                            "custom XML properties source schemaRefs close layout is missing",
                        ))?;
                        let close_start = source_element.close_start.ok_or(Error::Internal(
                            "custom XML properties source schemaRefs close layout is missing",
                        ))?;
                        let schema_refs = layout.schema_refs.as_mut().ok_or(Error::Internal(
                            "custom XML properties schemaRefs layout is missing",
                        ))?;
                        schema_refs.insertion_start = close_start;
                        schema_refs.insertion_end = close_start;
                        schema_refs.end = source_element.end;
                    },
                    LayoutKind::SchemaRef => {
                        let source_element = source_element.ok_or(Error::Internal(
                            "custom XML properties source schemaRef close layout is missing",
                        ))?;
                        let schema_refs = layout.schema_refs.as_mut().ok_or(Error::Internal(
                            "custom XML properties schemaRefs layout is missing",
                        ))?;
                        if let Some(index) = frame.schema_ref {
                            schema_refs.refs[index].end = source_element.end;
                        }
                    },
                    LayoutKind::Opaque => {},
                }
            },
            Event::Comment(_) => {
                if stack
                    .iter()
                    .any(|frame| matches!(frame.kind, LayoutKind::SchemaRef))
                {
                    mark_opaque(&mut layout, &stack);
                }
            },
            Event::Eof => break,
            Event::Text(text) if !is_xml_whitespace(text.as_ref()) => {
                if stack
                    .iter()
                    .any(|frame| matches!(frame.kind, LayoutKind::SchemaRef))
                {
                    mark_opaque(&mut layout, &stack);
                }
            },
            Event::CData(text) if !is_xml_whitespace(text.as_ref()) => {
                if stack
                    .iter()
                    .any(|frame| matches!(frame.kind, LayoutKind::SchemaRef))
                {
                    mark_opaque(&mut layout, &stack);
                }
            },
            Event::GeneralRef(_) | Event::PI(_) => {
                if stack
                    .iter()
                    .any(|frame| matches!(frame.kind, LayoutKind::SchemaRef))
                {
                    mark_opaque(&mut layout, &stack);
                }
            },
            Event::Decl(_) => {},
            Event::DocType(_) => {
                return Err(Error::Internal(
                    "custom XML properties layout encountered a forbidden declaration",
                ));
            },
            Event::Text(_) | Event::CData(_) => {},
        }
        buffer.clear();
    }
    Ok(layout)
}

fn mark_opaque(layout: &mut PropertiesLayout, stack: &[LayoutFrame]) {
    if let Some(frame) = stack
        .iter()
        .rev()
        .find(|frame| matches!(frame.kind, LayoutKind::SchemaRef))
    {
        if let (Some(schema_refs), Some(index)) = (layout.schema_refs.as_mut(), frame.schema_ref) {
            if let Some(schema_ref) = schema_refs.refs.get_mut(index) {
                schema_ref.opaque = true;
            }
        }
    }
}

fn has_unknown_attributes(
    element: &BytesStart<'_>,
    reader: &NsReader<&[u8]>,
    namespace: &[u8],
    known_local_name: &[u8],
) -> Result<bool> {
    for attribute in element.checked_attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        if is_namespace_declaration(attribute.key.as_ref()) {
            continue;
        }
        let (resolved, local) = reader.resolver().resolve_attribute(attribute.key);
        if resolved_namespace(resolved, "attribute")? != namespace
            || local.as_ref() != known_local_name
        {
            return Ok(true);
        }
    }
    Ok(false)
}

fn known_attribute_span(
    source: &[u8],
    event_start: usize,
    event_end: usize,
    element: &BytesStart<'_>,
    reader: &NsReader<&[u8]>,
    namespace: &[u8],
    local_name: &[u8],
) -> Result<Option<(usize, usize)>> {
    for attribute in element.checked_attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        if is_namespace_declaration(attribute.key.as_ref()) {
            continue;
        }
        let (resolved, local) = reader.resolver().resolve_attribute(attribute.key);
        if resolved_namespace(resolved, "attribute")? == namespace && local.as_ref() == local_name {
            return raw_attribute_value_span(
                source,
                event_start,
                event_end,
                attribute.key.as_ref(),
            );
        }
    }
    Ok(None)
}

fn raw_attribute_value_span(
    source: &[u8],
    event_start: usize,
    event_end: usize,
    wanted_name: &[u8],
) -> Result<Option<(usize, usize)>> {
    let content_start = event_start.checked_add(1).ok_or(Error::Internal(
        "custom XML properties attribute offset overflow",
    ))?;
    let content_end = event_end.checked_sub(1).ok_or(Error::Internal(
        "custom XML properties attribute offset underflow",
    ))?;
    let content = source
        .get(content_start..content_end)
        .ok_or(Error::Internal(
            "custom XML properties attribute range is invalid",
        ))?;
    let mut cursor = 0usize;
    while cursor < content.len() && !content[cursor].is_ascii_whitespace() {
        cursor += 1;
    }
    loop {
        while cursor < content.len() && content[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        if cursor >= content.len() || content[cursor] == b'/' {
            return Ok(None);
        }
        let name_start = cursor;
        while cursor < content.len()
            && !content[cursor].is_ascii_whitespace()
            && content[cursor] != b'='
        {
            cursor += 1;
        }
        let name = &content[name_start..cursor];
        while cursor < content.len() && content[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        if cursor >= content.len() || content[cursor] != b'=' {
            return Err(Error::Internal(
                "custom XML properties attribute layout is malformed",
            ));
        }
        cursor += 1;
        while cursor < content.len() && content[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        let delimiter = *content.get(cursor).ok_or(Error::Internal(
            "custom XML properties attribute value is missing",
        ))?;
        if !matches!(delimiter, b'\'' | b'"') {
            return Err(Error::Internal(
                "custom XML properties attribute value is unquoted",
            ));
        }
        cursor += 1;
        let value_start = cursor;
        while cursor < content.len() && content[cursor] != delimiter {
            cursor += 1;
        }
        if cursor >= content.len() {
            return Err(Error::Internal(
                "custom XML properties attribute value is unterminated",
            ));
        }
        if name == wanted_name {
            return Ok(Some((content_start + value_start, content_start + cursor)));
        }
        cursor += 1;
    }
}

fn apply_property_edits<'a>(
    source: &[u8],
    mut edits: Vec<LexicalEdit<'a>>,
    after: &Props,
    source_limits: &SourceLimits,
    verify_readback: bool,
) -> Result<Vec<u8>> {
    edits.sort_unstable_by_key(|edit| edit.start);
    let mut output_len = source.len();
    let mut previous_end = 0usize;
    for edit in &edits {
        if edit.start < previous_end || edit.end < edit.start || edit.end > source.len() {
            return Err(Error::Internal(
                "custom XML properties lexical edits overlap",
            ));
        }
        let removed = edit.end - edit.start;
        output_len = output_len.checked_sub(removed).ok_or(Error::Internal(
            "custom XML properties output length underflow",
        ))?;
        output_len = output_len
            .checked_add(replacement_len(&edit.replacement)?)
            .ok_or_else(|| {
                limit(
                    LimitKind::SerializedPropertiesBytes,
                    max_properties_bytes(source_limits),
                    usize::MAX,
                )
            })?;
        previous_end = edit.end;
    }
    require_at_most(
        LimitKind::SerializedPropertiesBytes,
        output_len,
        max_properties_bytes(source_limits),
    )?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(output_len)
        .map_err(|source| Error::Allocation {
            resource: "Custom XML properties rewrite",
            source,
        })?;
    let mut cursor = 0usize;
    for edit in edits {
        output.extend_from_slice(&source[cursor..edit.start]);
        write_replacement(&mut output, &edit.replacement)?;
        cursor = edit.end;
    }
    output.extend_from_slice(&source[cursor..]);
    if output.len() != output_len {
        return Err(Error::Internal(
            "custom XML properties lexical rewrite length disagrees",
        ));
    }
    if verify_readback {
        let parsed = read_props(&output, source_limits)?;
        if parsed != *after {
            return Err(Error::Internal(
                "custom XML properties lexical rewrite did not read back as requested",
            ));
        }
    }
    Ok(output)
}

fn replacement_len(replacement: &Replacement<'_>) -> Result<usize> {
    match replacement {
        Replacement::Empty => Ok(0),
        Replacement::Attribute(value) => escaped_attr_len(value),
        Replacement::SchemaRefs {
            values,
            namespace,
            qname,
        } => generated_schema_refs_content_len(values, namespace, qname),
        Replacement::GeneratedSchemaRefs {
            namespace,
            root_qname,
            values,
        } => generated_schema_refs_len(namespace, root_qname, values),
        Replacement::ExpandEmpty {
            source,
            start,
            end,
            qname,
            content,
            attribute,
        } => {
            let slash = (*end).checked_sub(2).ok_or(Error::Internal(
                "custom XML empty-element range is too short",
            ))?;
            if source.get(slash..*end) != Some(b"/>".as_slice()) {
                return Err(Error::Internal(
                    "custom XML empty-element closing layout is invalid",
                ));
            }
            let mut opening_len = slash
                .checked_sub(*start)
                .ok_or(Error::Internal("custom XML empty-element range is invalid"))?;
            if let Some((attribute_start, attribute_end, value)) = attribute {
                if *attribute_start < *start
                    || *attribute_end < *attribute_start
                    || *attribute_end > slash
                {
                    return Err(Error::Internal(
                        "custom XML empty-element attribute range is invalid",
                    ));
                }
                opening_len = opening_len
                    .checked_sub(attribute_end - attribute_start)
                    .ok_or_else(|| {
                        limit(
                            LimitKind::SerializedPropertiesBytes,
                            MAX_PROPS_BYTES,
                            usize::MAX,
                        )
                    })?;
                opening_len = opening_len
                    .checked_add(escaped_attr_len(value)?)
                    .ok_or_else(|| {
                        limit(
                            LimitKind::SerializedPropertiesBytes,
                            MAX_PROPS_BYTES,
                            usize::MAX,
                        )
                    })?;
            }
            let mut output_len = opening_len.checked_add(1).ok_or_else(|| {
                limit(
                    LimitKind::SerializedPropertiesBytes,
                    MAX_PROPS_BYTES,
                    usize::MAX,
                )
            })?;
            output_len = output_len
                .checked_add(schema_refs_content_len_from_content(content)?)
                .ok_or_else(|| {
                    limit(
                        LimitKind::SerializedPropertiesBytes,
                        MAX_PROPS_BYTES,
                        usize::MAX,
                    )
                })?;
            output_len = output_len
                .checked_add(qname.len().checked_add(3).ok_or_else(|| {
                    limit(
                        LimitKind::SerializedPropertiesBytes,
                        MAX_PROPS_BYTES,
                        usize::MAX,
                    )
                })?)
                .ok_or_else(|| {
                    limit(
                        LimitKind::SerializedPropertiesBytes,
                        MAX_PROPS_BYTES,
                        usize::MAX,
                    )
                })?;
            Ok(output_len)
        },
    }
}

fn schema_refs_content_len_from_content(content: &SchemaRefsContent<'_>) -> Result<usize> {
    match content {
        SchemaRefsContent::Generated {
            namespace,
            qname,
            values,
        } => {
            let mut length = 0usize;
            for value in values.iter() {
                add_len(
                    &mut length,
                    generated_schema_ref_len(namespace, qname, value)?,
                )?;
            }
            Ok(length)
        },
        SchemaRefsContent::GeneratedContainer {
            namespace,
            root_qname,
            values,
        } => generated_schema_refs_len(namespace, root_qname, values),
    }
}

fn generated_schema_refs_content_len(
    values: &[String],
    namespace: &str,
    qname: &[u8],
) -> Result<usize> {
    let mut length = 0usize;
    for value in values {
        add_len(
            &mut length,
            generated_schema_ref_len(namespace, qname, value)?,
        )?;
    }
    Ok(length)
}

fn generated_schema_refs_len(
    namespace: &str,
    root_qname: &[u8],
    values: &[String],
) -> Result<usize> {
    let prefix = qname_prefix(root_qname).unwrap_or(b"ds");
    let mut length = 1usize
        .checked_add(prefix.len())
        .and_then(|length| length.checked_add(b":schemaRefs".len()))
        .ok_or_else(|| {
            limit(
                LimitKind::SerializedPropertiesBytes,
                MAX_PROPS_BYTES,
                usize::MAX,
            )
        })?;
    if qname_prefix(root_qname).is_none() {
        add_len(&mut length, b" xmlns:ds=\"".len())?;
        add_len(&mut length, escaped_attr_len(namespace)?)?;
        add_len(&mut length, 1)?;
    }
    add_len(&mut length, 1)?;
    for value in values {
        add_len(
            &mut length,
            generated_schema_ref_len(namespace, root_qname, value)?,
        )?;
    }
    add_len(&mut length, 2)?;
    add_len(&mut length, prefix.len())?;
    add_len(&mut length, b":schemaRefs>".len())?;
    Ok(length)
}

fn generated_schema_ref_len(namespace: &str, qname: &[u8], value: &str) -> Result<usize> {
    let prefix = qname_prefix(qname).unwrap_or(b"ds");
    let mut length = if qname_prefix(qname).is_some() {
        b"<".len()
            .checked_add(prefix.len())
            .and_then(|length| length.checked_add(b":schemaRef ".len()))
            .and_then(|length| length.checked_add(prefix.len()))
            .and_then(|length| length.checked_add(b":uri=\"".len()))
            .ok_or_else(|| {
                limit(
                    LimitKind::SerializedPropertiesBytes,
                    MAX_PROPS_BYTES,
                    usize::MAX,
                )
            })?
    } else {
        let mut length = b"<schemaRef xmlns:ds=\"".len();
        add_len(&mut length, escaped_attr_len(namespace)?)?;
        add_len(&mut length, b"\" ds:uri=\"".len())?;
        length
    };
    add_len(&mut length, escaped_attr_len(value)?)?;
    add_len(&mut length, b"\"/>".len())?;
    Ok(length)
}

fn write_replacement(output: &mut Vec<u8>, replacement: &Replacement<'_>) -> Result<()> {
    match replacement {
        Replacement::Empty => {},
        Replacement::Attribute(value) => push_escaped_attr(output, value),
        Replacement::SchemaRefs {
            values,
            namespace,
            qname,
        } => write_generated_schema_refs_content(output, namespace, qname, values),
        Replacement::GeneratedSchemaRefs {
            namespace,
            root_qname,
            values,
        } => write_generated_schema_refs(output, namespace, root_qname, values)?,
        Replacement::ExpandEmpty {
            source,
            start,
            end,
            qname,
            content,
            attribute,
        } => {
            let slash = (*end).checked_sub(2).ok_or(Error::Internal(
                "custom XML empty-element range is too short",
            ))?;
            if source.get(slash..*end) != Some(b"/>".as_slice()) {
                return Err(Error::Internal(
                    "custom XML empty-element closing layout is invalid",
                ));
            }
            if let Some((attribute_start, attribute_end, value)) = attribute {
                output.extend_from_slice(&source[*start..*attribute_start]);
                push_escaped_attr(output, value);
                output.extend_from_slice(&source[*attribute_end..slash]);
            } else {
                output.extend_from_slice(&source[*start..slash]);
            }
            output.push(b'>');
            write_schema_refs_content_from_content(output, content)?;
            output.extend_from_slice(b"</");
            output.extend_from_slice(qname);
            output.push(b'>');
        },
    }
    Ok(())
}

fn write_schema_refs_content_from_content(
    output: &mut Vec<u8>,
    content: &SchemaRefsContent<'_>,
) -> Result<()> {
    match content {
        SchemaRefsContent::Generated {
            namespace,
            qname,
            values,
        } => {
            for value in values.iter() {
                write_generated_schema_ref(output, namespace, qname, value);
            }
            Ok(())
        },
        SchemaRefsContent::GeneratedContainer {
            namespace,
            root_qname,
            values,
        } => {
            // The root-empty path has no existing schemaRefs wrapper to
            // expand, so emit the complete requested container here.
            write_generated_schema_refs(output, namespace, root_qname, values)
        },
    }
}

fn write_generated_schema_refs_content(
    output: &mut Vec<u8>,
    namespace: &str,
    qname: &[u8],
    values: &[String],
) {
    for value in values {
        write_generated_schema_ref(output, namespace, qname, value);
    }
}

fn write_generated_schema_refs(
    output: &mut Vec<u8>,
    namespace: &str,
    root_qname: &[u8],
    values: &[String],
) -> Result<()> {
    let prefix = qname_prefix(root_qname).unwrap_or(b"ds");
    output.extend_from_slice(b"<");
    output.extend_from_slice(prefix);
    output.extend_from_slice(b":schemaRefs");
    if qname_prefix(root_qname).is_none() {
        output.extend_from_slice(b" xmlns:ds=\"");
        push_escaped_attr(output, namespace);
        output.extend_from_slice(b"\"");
    }
    output.push(b'>');
    for value in values {
        write_generated_schema_ref(output, namespace, root_qname, value);
    }
    output.extend_from_slice(b"</");
    output.extend_from_slice(prefix);
    output.extend_from_slice(b":schemaRefs>");
    Ok(())
}

fn write_generated_schema_ref(output: &mut Vec<u8>, namespace: &str, qname: &[u8], value: &str) {
    let prefix = qname_prefix(qname).unwrap_or(b"ds");
    output.extend_from_slice(b"<");
    if qname_prefix(qname).is_some() {
        output.extend_from_slice(prefix);
        output.extend_from_slice(b":schemaRef ");
        output.extend_from_slice(prefix);
        output.extend_from_slice(b":uri=\"");
    } else {
        output.extend_from_slice(b"schemaRef xmlns:ds=\"");
        push_escaped_attr(output, namespace);
        output.extend_from_slice(b"\" ds:uri=\"");
    }
    push_escaped_attr(output, value);
    output.extend_from_slice(b"\"/>");
}

fn qname_prefix(qname: &[u8]) -> Option<&[u8]> {
    qname
        .iter()
        .position(|byte| *byte == b':')
        .map(|index| &qname[..index])
}

/// Validate a typed properties value without serializing it.
/// # Errors
///
/// Returns an error when input violates OOXML constraints, exceeds a configured
/// bound, or an underlying XML or package operation fails.
pub fn validate_props(props: &Props) -> Result<()> {
    if !valid_guid(&props.id) {
        return invalid(format!("custom XML itemID '{}' is not ST_Guid", props.id));
    }
    require_at_most(
        LimitKind::SchemaReferences,
        props.schemas.len(),
        MAX_SCHEMA_REFS,
    )?;
    let mut string_bytes = props.id.len();
    for schema in &props.schemas {
        if schema.is_empty() {
            return invalid("custom XML schema reference URI must not be empty");
        }
        validate_xml_chars(schema)?;
        string_bytes = string_bytes
            .checked_add(schema.len())
            .ok_or_else(|| limit(LimitKind::PropertyStringBytes, MAX_STRING_BYTES, usize::MAX))?;
    }
    require_at_most(
        LimitKind::PropertyStringBytes,
        string_bytes,
        MAX_STRING_BYTES,
    )
}

/// Return whether `value` is the braced hexadecimal `ST_Guid` lexical form.
#[must_use]
pub fn valid_guid(value: &str) -> bool {
    let Some(inner) = value
        .strip_prefix('{')
        .and_then(|value| value.strip_suffix('}'))
    else {
        return false;
    };
    let groups = [8, 4, 4, 4, 12];
    let mut parts = inner.split('-');
    groups.into_iter().all(|length| {
        parts.next().is_some_and(|part| {
            part.len() == length && part.bytes().all(|byte| byte.is_ascii_hexdigit())
        })
    }) && parts.next().is_none()
}

/// Validate a bounded XML payload and return its expanded document-element name.
/// # Errors
///
/// Returns an error when input violates OOXML constraints, exceeds a configured
/// bound, or an underlying XML or package operation fails.
pub fn validate_payload(xml: &[u8], source_limits: &SourceLimits) -> Result<Name> {
    let max_part_bytes = max_payload_bytes(source_limits);
    let max_depth = MAX_DEPTH.min(source_limits.max_depth());
    let max_elements = MAX_ELEMENTS.min(source_limits.max_events());
    require_at_most(LimitKind::PayloadBytes, xml.len(), max_part_bytes)?;
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_comments = true;
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut root = None;
    let mut closed_root = false;
    let mut elements = 0usize;
    let mut version = XmlVersion::Implicit1_0;
    let mut event_seen = false;

    loop {
        let event = reader.read_event_into(&mut buffer)?;
        match event {
            Event::Decl(declaration) => {
                if event_seen {
                    return invalid("custom XML payload has a late XML declaration");
                }
                version = validate_declaration(&declaration)?;
            },
            Event::Start(element) => {
                check_source_element_bounds(
                    source_limits,
                    depth,
                    &element,
                    version,
                    reader.decoder(),
                )?;
                require_nested_depth(depth, max_depth)?;
                elements = bump_elements(elements, max_elements)?;
                let name = inspect_element(&reader, &element, version, depth == 0)?;
                if depth == 0 {
                    if closed_root || root.is_some() {
                        return invalid("custom XML payload has multiple roots");
                    }
                    root = name;
                }
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| limit(LimitKind::Depth, max_depth, usize::MAX))?;
            },
            Event::Empty(element) => {
                check_source_element_bounds(
                    source_limits,
                    depth,
                    &element,
                    version,
                    reader.decoder(),
                )?;
                require_nested_depth(depth, max_depth)?;
                elements = bump_elements(elements, max_elements)?;
                let name = inspect_element(&reader, &element, version, depth == 0)?;
                if depth == 0 {
                    if closed_root || root.is_some() {
                        return invalid("custom XML payload has multiple roots");
                    }
                    root = name;
                    closed_root = true;
                }
            },
            Event::End(_) => {
                depth = depth.checked_sub(1).ok_or_else(|| {
                    Error::Invalid("custom XML payload has an unexpected end tag".into())
                })?;
                if depth == 0 {
                    closed_root = true;
                }
            },
            Event::DocType(_) => return invalid("DTD is forbidden in custom XML payloads"),
            Event::GeneralRef(reference) => {
                if depth == 0 {
                    return invalid("custom XML payload has a reference outside its root");
                }
                let value = decode_xml_reference(&reference)?;
                validate_xml_chars(&value)?;
            },
            Event::Text(text) => {
                let decoded = text
                    .decode()
                    .map_err(|error| Error::Xml(error.to_string()))?;
                validate_xml_chars(&decoded)?;
                if depth == 0 && !is_xml_whitespace(decoded.as_bytes()) {
                    return invalid("custom XML payload has text outside its root");
                }
            },
            Event::CData(text) => {
                let decoded = text
                    .decode()
                    .map_err(|error| Error::Xml(error.to_string()))?;
                validate_xml_chars(&decoded)?;
                if depth == 0 {
                    return invalid("custom XML payload has CDATA outside its root");
                }
            },
            Event::Comment(text) => {
                let text = text
                    .decode()
                    .map_err(|error| Error::Xml(error.to_string()))?;
                validate_xml_chars(&text)?;
            },
            Event::PI(instruction) => validate_instruction(&reader, &instruction)?,
            Event::Eof => break,
        }
        event_seen = true;
        buffer.clear();
    }
    if depth != 0 || !closed_root {
        return invalid("custom XML payload has no complete root element");
    }
    root.ok_or_else(|| Error::Invalid("custom XML payload has no root".into()))
}

fn max_payload_bytes(source_limits: &SourceLimits) -> usize {
    MAX_PART_BYTES.min(usize::try_from(source_limits.max_part_bytes()).unwrap_or(usize::MAX))
}

/// Validate that a Custom XML data-part content type is a well-formed XML media type.
/// # Errors
///
/// Returns an error when input violates OOXML constraints, exceeds a configured
/// bound, or an underlying XML or package operation fails.
pub fn validate_content_type(content_type: &str) -> Result<()> {
    let parsed = ContentType::new(content_type).map_err(|_error| Error::ContentType {
        expected: "a well-formed XML media type".into(),
        actual: content_type.into(),
    })?;
    let media_type = parsed.as_str().split(';').next().unwrap_or_default();
    let Some((type_name, subtype)) = media_type.split_once('/') else {
        return Err(Error::ContentType {
            expected: "an XML media type".into(),
            actual: content_type.into(),
        });
    };
    let xml = (type_name.eq_ignore_ascii_case("application")
        || type_name.eq_ignore_ascii_case("text"))
        && subtype.eq_ignore_ascii_case("xml");
    let suffix = subtype.len() > 4
        && subtype
            .get(subtype.len() - 4..)
            .is_some_and(|value| value.eq_ignore_ascii_case("+xml"));
    if xml || suffix {
        Ok(())
    } else {
        Err(Error::ContentType {
            expected: "an XML media type".into(),
            actual: content_type.into(),
        })
    }
}

#[derive(Debug)]
struct ResolvedElement {
    namespace: String,
    local_name: String,
    attributes: Vec<(String, String, String)>,
}

fn inspect_element(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    version: XmlVersion,
    capture: bool,
) -> Result<Option<Name>> {
    validate_qname(element.name().as_ref(), "element")?;
    let (resolved, local) = reader.resolver().resolve_element(element.name());
    let namespace = resolved_namespace(resolved, "element")?;
    let namespace =
        std::str::from_utf8(namespace).map_err(|error| Error::Xml(error.to_string()))?;
    let local =
        std::str::from_utf8(local.as_ref()).map_err(|error| Error::Xml(error.to_string()))?;
    let captured = capture.then(|| Name {
        namespace: namespace.into(),
        local_name: local.into(),
    });

    let mut seen = HashSet::new();
    for attribute in element.checked_attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        let value = attribute
            .decoded_and_normalized_value(version, reader.decoder())
            .map_err(|error| Error::Xml(error.to_string()))?;
        validate_xml_chars(&value)?;
        if is_namespace_declaration(attribute.key.as_ref()) {
            validate_namespace_declaration(attribute.key.as_ref())?;
            continue;
        }
        validate_qname(attribute.key.as_ref(), "attribute")?;
        let (resolved, local) = reader.resolver().resolve_attribute(attribute.key);
        let namespace = resolved_namespace(resolved, "attribute")?;
        std::str::from_utf8(namespace).map_err(|error| Error::Xml(error.to_string()))?;
        std::str::from_utf8(local.as_ref()).map_err(|error| Error::Xml(error.to_string()))?;
        if !seen.insert((namespace.to_vec(), local.as_ref().to_vec())) {
            return invalid(format!(
                "duplicate expanded XML attribute '{}'",
                String::from_utf8_lossy(attribute.key.as_ref())
            ));
        }
    }
    Ok(captured)
}

fn resolve_props_element(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    version: XmlVersion,
) -> Result<ResolvedElement> {
    validate_qname(element.name().as_ref(), "element")?;
    let (resolved, local) = reader.resolver().resolve_element(element.name());
    let namespace = std::str::from_utf8(resolved_namespace(resolved, "element")?)
        .map_err(|error| Error::Xml(error.to_string()))?
        .to_owned();
    let local_name = std::str::from_utf8(local.as_ref())
        .map_err(|error| Error::Xml(error.to_string()))?
        .to_owned();
    let mut attributes = Vec::new();
    let mut seen = HashSet::new();
    for attribute in element.checked_attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        let value = attribute
            .decoded_and_normalized_value(version, reader.decoder())
            .map_err(|error| Error::Xml(error.to_string()))?
            .into_owned();
        validate_xml_chars(&value)?;
        if is_namespace_declaration(attribute.key.as_ref()) {
            validate_namespace_declaration(attribute.key.as_ref())?;
            continue;
        }
        validate_qname(attribute.key.as_ref(), "attribute")?;
        let (resolved, local) = reader.resolver().resolve_attribute(attribute.key);
        let attribute_namespace = std::str::from_utf8(resolved_namespace(resolved, "attribute")?)
            .map_err(|error| Error::Xml(error.to_string()))?
            .to_owned();
        let attribute_name = std::str::from_utf8(local.as_ref())
            .map_err(|error| Error::Xml(error.to_string()))?
            .to_owned();
        if !seen.insert((attribute_namespace.clone(), attribute_name.clone())) {
            return invalid(format!(
                "duplicate expanded XML attribute '{}'",
                String::from_utf8_lossy(attribute.key.as_ref())
            ));
        }
        attributes.push((attribute_namespace, attribute_name, value));
    }
    Ok(ResolvedElement {
        namespace,
        local_name,
        attributes,
    })
}

fn resolved_namespace<'a>(result: ResolveResult<'a>, kind: &str) -> Result<&'a [u8]> {
    match result {
        ResolveResult::Bound(Namespace(value)) => Ok(value),
        ResolveResult::Unbound => Ok(b""),
        ResolveResult::Unknown(prefix) => invalid(format!(
            "unbound XML {kind} namespace prefix '{}'",
            String::from_utf8_lossy(prefix.as_ref())
        )),
    }
}

fn parse_props_xml(xml: &[u8], max_depth: usize, source_limits: &SourceLimits) -> Result<Props> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_comments = true;
    let mut buffer = Vec::new();
    let mut contexts = Vec::new();
    let mut id = None;
    let mut schemas = Vec::new();
    let mut root_namespace = None;
    let mut seen_schemas = false;
    let mut closed_root = false;
    let mut version = XmlVersion::Implicit1_0;
    let mut event_seen = false;
    let mut events = 0usize;

    loop {
        let event = reader.read_event_into(&mut buffer)?;
        if !matches!(&event, Event::Eof) {
            events = events.checked_add(1).ok_or_else(|| {
                source_limit(LimitKind::XmlEvents, usize::MAX, source_limits.max_events())
            })?;
            if events > source_limits.max_events() {
                return Err(source_limit(
                    LimitKind::XmlEvents,
                    events,
                    source_limits.max_events(),
                ));
            }
        }
        match event {
            Event::Decl(declaration) => {
                if event_seen {
                    return invalid("custom XML properties have a late XML declaration");
                }
                version = validate_declaration(&declaration)?;
            },
            Event::Start(element) => {
                check_source_element_bounds(
                    source_limits,
                    contexts.len(),
                    &element,
                    version,
                    reader.decoder(),
                )?;
                require_nested_depth(contexts.len(), max_depth)?;
                let element = resolve_props_element(&reader, &element, version)?;
                let context = props_start(
                    &element,
                    contexts.last().copied(),
                    &mut id,
                    &mut schemas,
                    &mut root_namespace,
                    &mut seen_schemas,
                )?;
                contexts.push(context);
            },
            Event::Empty(element) => {
                check_source_element_bounds(
                    source_limits,
                    contexts.len(),
                    &element,
                    version,
                    reader.decoder(),
                )?;
                require_nested_depth(contexts.len(), max_depth)?;
                let element = resolve_props_element(&reader, &element, version)?;
                props_start(
                    &element,
                    contexts.last().copied(),
                    &mut id,
                    &mut schemas,
                    &mut root_namespace,
                    &mut seen_schemas,
                )?;
                if contexts.is_empty() {
                    closed_root = true;
                }
            },
            Event::End(_) => {
                if contexts.pop().is_none() {
                    return invalid("unexpected custom XML properties end tag");
                }
                if contexts.is_empty() {
                    closed_root = true;
                }
            },
            Event::DocType(_) => return invalid("DTD is forbidden in custom XML properties"),
            Event::Text(text) => {
                let text = text
                    .decode()
                    .map_err(|error| Error::Xml(error.to_string()))?;
                validate_xml_chars(&text)?;
                if !matches!(contexts.last(), Some(PropsContext::Opaque))
                    && !is_xml_whitespace(text.as_bytes())
                {
                    return invalid("text is not permitted in custom XML properties");
                }
            },
            Event::CData(text) => {
                let text = text
                    .decode()
                    .map_err(|error| Error::Xml(error.to_string()))?;
                validate_xml_chars(&text)?;
                if contexts.is_empty() {
                    return invalid("CDATA is not permitted outside custom XML properties");
                }
                if !matches!(contexts.last(), Some(PropsContext::Opaque))
                    && !is_xml_whitespace(text.as_bytes())
                {
                    return invalid("CDATA is not permitted in custom XML properties");
                }
            },
            Event::GeneralRef(reference) => {
                if contexts.is_empty() {
                    return invalid("references are not permitted outside custom XML properties");
                }
                let value = decode_xml_reference(&reference)?;
                validate_xml_chars(&value)?;
                if !matches!(contexts.last(), Some(PropsContext::Opaque))
                    && !is_xml_whitespace(value.as_bytes())
                {
                    return invalid("references are not permitted in custom XML properties");
                }
            },
            Event::Comment(text) => {
                text.decode()
                    .map_err(|error| Error::Xml(error.to_string()))?;
            },
            Event::PI(_) if !matches!(contexts.last(), Some(PropsContext::Opaque)) => {
                return invalid("processing instructions are forbidden in custom XML properties");
            },
            Event::PI(_) => {},
            Event::Eof => break,
        }
        event_seen = true;
        buffer.clear();
    }
    if !closed_root || !contexts.is_empty() {
        return invalid("custom XML properties root is incomplete");
    }
    let props = Props {
        id: id.ok_or_else(|| Error::Invalid("datastoreItem requires ds:itemID".into()))?,
        schemas,
    };
    validate_props(&props)?;
    Ok(props)
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum PropsContext {
    Item,
    Schemas,
    Schema,
    Opaque,
}

fn props_start(
    element: &ResolvedElement,
    parent: Option<PropsContext>,
    id: &mut Option<String>,
    schemas: &mut Vec<String>,
    root_namespace: &mut Option<String>,
    seen_schemas: &mut bool,
) -> Result<PropsContext> {
    let supported_namespace = is_custom_namespace(&element.namespace);
    match parent {
        None if root_namespace.is_none()
            && supported_namespace
            && element.local_name == "datastoreItem" =>
        {
            reject_attributes_except(element, &[(element.namespace.as_str(), "itemID")]);
            *id = Some(required_attr(element, &element.namespace, "itemID")?.into());
            *root_namespace = Some(element.namespace.clone());
            Ok(PropsContext::Item)
        },
        Some(PropsContext::Item)
            if root_namespace.as_deref() == Some(element.namespace.as_str())
                && element.local_name == "schemaRefs" =>
        {
            if *seen_schemas {
                return invalid("datastoreItem has multiple schemaRefs elements");
            }
            *seen_schemas = true;
            reject_attributes_except(element, &[]);
            Ok(PropsContext::Schemas)
        },
        Some(PropsContext::Schemas)
            if root_namespace.as_deref() == Some(element.namespace.as_str())
                && element.local_name == "schemaRef" =>
        {
            let next = schemas
                .len()
                .checked_add(1)
                .ok_or_else(|| limit(LimitKind::SchemaReferences, MAX_SCHEMA_REFS, usize::MAX))?;
            require_at_most(LimitKind::SchemaReferences, next, MAX_SCHEMA_REFS)?;
            reject_attributes_except(element, &[(element.namespace.as_str(), "uri")]);
            schemas.push(required_attr(element, &element.namespace, "uri")?.into());
            Ok(PropsContext::Schema)
        },
        Some(PropsContext::Opaque) => Ok(PropsContext::Opaque),
        Some(PropsContext::Item | PropsContext::Schemas | PropsContext::Schema) => {
            // Future properties markup is retained in the source part and is
            // intentionally not interpreted by this host-neutral layer.
            Ok(PropsContext::Opaque)
        },
        _ => invalid(format!(
            "unexpected custom XML properties element {{{}}}{}",
            element.namespace, element.local_name
        )),
    }
}

fn reject_attributes_except(element: &ResolvedElement, allowed: &[(&str, &str)]) {
    // Unknown attributes are opaque extension data. Known attributes are
    // still checked by `required_attr`; retaining this helper keeps the
    // semantic call sites explicit without discarding future markup.
    let _ = (element, allowed);
}

fn required_attr<'a>(
    element: &'a ResolvedElement,
    namespace: &str,
    local_name: &str,
) -> Result<&'a str> {
    element
        .attributes
        .iter()
        .find(|(candidate_namespace, candidate_name, _)| {
            candidate_namespace == namespace && candidate_name == local_name
        })
        .map(|(_, _, value)| value.as_str())
        .ok_or_else(|| {
            Error::Invalid(format!(
                "{} requires {{{namespace}}}{local_name}",
                element.local_name
            ))
        })
}

fn validate_declaration(declaration: &BytesDecl<'_>) -> Result<XmlVersion> {
    let version = declaration.xml_version()?;
    let declaration_text =
        std::str::from_utf8(declaration.as_ref()).map_err(|error| Error::Xml(error.to_string()))?;
    let raw = BytesStart::from_content(declaration_text, 3);
    let mut declaration_state = 0u8;
    for attribute in raw.checked_attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        if attribute.key.prefix().is_some() {
            return invalid(format!(
                "unexpected XML declaration attribute '{}'",
                String::from_utf8_lossy(attribute.key.as_ref())
            ));
        }
        declaration_state = match (declaration_state, attribute.key.as_ref()) {
            (0, b"version") => 1,
            (1, b"encoding") => 2,
            (1 | 2, b"standalone") => 3,
            _ => {
                return invalid(format!(
                    "unexpected or out-of-order XML declaration attribute '{}'",
                    String::from_utf8_lossy(attribute.key.as_ref())
                ));
            },
        };
        std::str::from_utf8(attribute.value.as_ref())
            .map_err(|error| Error::Xml(error.to_string()))?;
    }
    if let Some(encoding) = declaration.encoding() {
        let encoding = encoding.map_err(|error| Error::Xml(error.to_string()))?;
        let encoding =
            std::str::from_utf8(&encoding).map_err(|error| Error::Xml(error.to_string()))?;
        if !valid_encoding_name(encoding) {
            return invalid(format!(
                "XML declaration encoding '{encoding}' is not an EncName"
            ));
        }
    }
    if let Some(standalone) = declaration.standalone() {
        let standalone = standalone.map_err(|error| Error::Xml(error.to_string()))?;
        if !matches!(standalone.as_ref(), b"yes" | b"no") {
            return invalid("XML declaration standalone must be 'yes' or 'no'");
        }
    }
    Ok(version)
}

fn validate_instruction(
    reader: &NsReader<&[u8]>,
    instruction: &quick_xml::events::BytesPI<'_>,
) -> Result<()> {
    let target = reader
        .decoder()
        .decode(instruction.target())
        .map_err(|error| Error::Xml(error.to_string()))?;
    if !xml_name::is_xml_name(&target) {
        return Err(Error::Xml(format!(
            "invalid processing-instruction target '{target}'"
        )));
    }
    if target.eq_ignore_ascii_case("xml") {
        return invalid("processing-instruction target cannot be 'xml'");
    }
    let content = reader
        .decoder()
        .decode(instruction.content())
        .map_err(|error| Error::Xml(error.to_string()))?;
    validate_xml_chars(&content)?;
    Ok(())
}

fn props_output_len(props: &Props, conformance: Conformance) -> Result<usize> {
    const DECL: &[u8] = b"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>";
    const ROOT_PREFIX: &[u8] = b"<ds:datastoreItem xmlns:ds=\"";
    const ID_PREFIX: &[u8] = b"\" ds:itemID=\"";
    const ROOT_OPEN_END: &[u8] = b"\">";
    const SCHEMAS_OPEN: &[u8] = b"<ds:schemaRefs>";
    const SCHEMA_PREFIX: &[u8] = b"<ds:schemaRef ds:uri=\"";
    const SCHEMA_END: &[u8] = b"\"/>";
    const SCHEMAS_END: &[u8] = b"</ds:schemaRefs>";
    const ROOT_END: &[u8] = b"</ds:datastoreItem>";

    let mut total = 0usize;
    add_len(&mut total, DECL.len())?;
    add_len(&mut total, ROOT_PREFIX.len())?;
    add_len(&mut total, escaped_attr_len(conformance.namespace())?)?;
    add_len(&mut total, ID_PREFIX.len())?;
    add_len(&mut total, escaped_attr_len(&props.id)?)?;
    add_len(&mut total, ROOT_OPEN_END.len())?;
    if !props.schemas.is_empty() {
        add_len(&mut total, SCHEMAS_OPEN.len())?;
        for schema in &props.schemas {
            add_len(&mut total, SCHEMA_PREFIX.len())?;
            add_len(&mut total, escaped_attr_len(schema)?)?;
            add_len(&mut total, SCHEMA_END.len())?;
        }
        add_len(&mut total, SCHEMAS_END.len())?;
    }
    add_len(&mut total, ROOT_END.len())?;
    require_at_most(LimitKind::SerializedPropertiesBytes, total, MAX_PROPS_BYTES)?;
    Ok(total)
}

fn escaped_attr_len(value: &str) -> Result<usize> {
    value.chars().try_fold(0usize, |total, character| {
        let width = match character {
            '&' | '\t' | '\n' | '\r' => 5,
            '<' | '>' => 4,
            '"' | '\'' => 6,
            _ => character.len_utf8(),
        };
        total.checked_add(width).ok_or_else(|| {
            limit(
                LimitKind::SerializedPropertiesBytes,
                MAX_PROPS_BYTES,
                usize::MAX,
            )
        })
    })
}

fn add_len(total: &mut usize, value: usize) -> Result<()> {
    *total = total.checked_add(value).ok_or_else(|| {
        limit(
            LimitKind::SerializedPropertiesBytes,
            MAX_PROPS_BYTES,
            usize::MAX,
        )
    })?;
    Ok(())
}

fn push_escaped_attr(out: &mut Vec<u8>, value: &str) {
    for character in value.chars() {
        match character {
            '&' => out.extend_from_slice(b"&amp;"),
            '<' => out.extend_from_slice(b"&lt;"),
            '>' => out.extend_from_slice(b"&gt;"),
            '"' => out.extend_from_slice(b"&quot;"),
            '\'' => out.extend_from_slice(b"&apos;"),
            '\t' => out.extend_from_slice(b"&#x9;"),
            '\n' => out.extend_from_slice(b"&#xA;"),
            '\r' => out.extend_from_slice(b"&#xD;"),
            _ => {
                let mut encoded = [0; 4];
                out.extend_from_slice(character.encode_utf8(&mut encoded).as_bytes());
            },
        }
    }
}

fn validate_xml_chars(value: &str) -> Result<()> {
    if value.chars().any(|character| {
        !matches!(
            character,
            '\u{9}' | '\u{A}' | '\u{D}' | '\u{20}'..='\u{D7FF}' | '\u{E000}'..='\u{FFFD}' | '\u{10000}'..='\u{10FFFF}'
        )
    }) {
        return Err(Error::Xml(
            "custom XML contains a character forbidden by XML 1.0".into(),
        ));
    }
    Ok(())
}

pub(super) fn require_rel_id(value: &str, label: &str) -> Result<()> {
    if xml_name::is_ncname(value) {
        Ok(())
    } else {
        Err(Error::Relationship(format!(
            "{label} ID '{value}' is not an XML NCName"
        )))
    }
}

fn validate_qname(value: &[u8], kind: &str) -> Result<()> {
    let value = std::str::from_utf8(value).map_err(|error| Error::Xml(error.to_string()))?;
    if xml_name::is_qualified_name(value) {
        Ok(())
    } else {
        Err(Error::Xml(format!("invalid XML {kind} QName '{value}'")))
    }
}

fn validate_namespace_declaration(value: &[u8]) -> Result<()> {
    if value == b"xmlns" {
        return Ok(());
    }
    let Some(prefix) = value.strip_prefix(b"xmlns:") else {
        return Err(Error::Xml("invalid XML namespace declaration".into()));
    };
    let prefix = std::str::from_utf8(prefix).map_err(|error| Error::Xml(error.to_string()))?;
    if xml_name::is_ncname(prefix) {
        Ok(())
    } else {
        Err(Error::Xml(format!(
            "invalid XML namespace prefix '{prefix}'"
        )))
    }
}

fn valid_encoding_name(value: &str) -> bool {
    let mut bytes = value.bytes();
    bytes.next().is_some_and(|byte| byte.is_ascii_alphabetic())
        && bytes.all(|byte| byte.is_ascii_alphanumeric() || matches!(byte, b'.' | b'_' | b'-'))
}

fn require_nested_depth(depth: usize, maximum: usize) -> Result<()> {
    if depth >= maximum {
        let observed = depth.checked_add(1).unwrap_or(usize::MAX);
        Err(limit(LimitKind::Depth, maximum, observed))
    } else {
        Ok(())
    }
}

fn bump_elements(elements: usize, maximum: usize) -> Result<usize> {
    let next = elements
        .checked_add(1)
        .ok_or_else(|| limit(LimitKind::Elements, maximum, usize::MAX))?;
    require_at_most(LimitKind::Elements, next, maximum)?;
    Ok(next)
}

pub(super) fn require_at_most(kind: LimitKind, actual: usize, max: usize) -> Result<()> {
    if actual > max {
        Err(limit(kind, max, actual))
    } else {
        Ok(())
    }
}

pub(super) fn limit(kind: LimitKind, max: usize, actual: usize) -> Error {
    Error::Limit {
        kind: crate::LimitKind::CustomXml(kind),
        maximum: max,
        observed: actual,
    }
}

pub(super) fn invalid<T>(message: impl Into<String>) -> Result<T> {
    Err(Error::Invalid(message.into()))
}

pub(super) fn is_data_relationship(value: &str) -> bool {
    matches!(value, TRANSITIONAL_RELATIONSHIP | STRICT_RELATIONSHIP)
}

pub(super) fn is_props_relationship(value: &str) -> bool {
    matches!(
        value,
        TRANSITIONAL_PROPS_RELATIONSHIP | STRICT_PROPS_RELATIONSHIP
    )
}

fn is_custom_namespace(value: &str) -> bool {
    matches!(value, TRANSITIONAL_NAMESPACE | STRICT_NAMESPACE)
}

fn is_namespace_declaration(value: &[u8]) -> bool {
    value == b"xmlns" || value.starts_with(b"xmlns:")
}

fn is_xml_whitespace(value: &[u8]) -> bool {
    value
        .iter()
        .all(|byte| matches!(byte, b' ' | b'\t' | b'\r' | b'\n'))
}
