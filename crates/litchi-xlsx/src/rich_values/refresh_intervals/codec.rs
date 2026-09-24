//! Bounded XML inspection and source-local rewriting for rich-value refresh metadata.

use std::collections::TryReserveError;
use std::ops::Range;

use litchi_core::xml::ReaderOrigin;
use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::ResolveResult;
use quick_xml::reader::NsReader;

use crate::error::{Error, Result, invalid};
use crate::rich_values::codec::xml::{
    Node, no_attributes, optional, parse_document, require, required, whitespace,
};

use super::super::MAX_DEPTH;
use super::model::{
    REFRESH_INTERVALS_NAMESPACE, RefreshInterval, RefreshIntervals, TypeRefreshIntervals,
    validate_intervals,
};
use super::{MAX_INTERVALS, MAX_TYPES, valid_xml10};
use crate::rich_values::{
    MAX_OUTPUT_BYTES, MAX_STRING_BYTES, MAX_XML_BYTES, RICH_DATA_2, SPREADSHEETML,
};

const TYPE_NAME_MAX_CHARS: usize = 255;

#[derive(Clone, Debug)]
pub(crate) struct Inspection {
    pub(crate) types: Vec<TypeInspection>,
}

#[derive(Clone, Debug)]
pub(crate) struct TypeInspection {
    pub(crate) model: TypeRefreshIntervals,
    pub(crate) refresh: Option<Range<usize>>,
    pub(crate) extensions: Vec<ExtensionLocation>,
    pub(crate) extension_list_seen: bool,
}

#[derive(Clone, Debug)]
pub(crate) struct ExtensionLocation {
    pub(crate) start_tag: Range<usize>,
    pub(crate) close_tag: Option<Range<usize>>,
    pub(crate) name: Vec<u8>,
}

/// Parse one standalone `refreshIntervals` element.
pub fn parse_refresh_intervals(xml: &[u8]) -> Result<RefreshIntervals> {
    let root = parse_document(xml)?;
    parse_refresh_node(&root)
}

/// Serialize one bounded `refreshIntervals` element.
pub fn write_refresh_intervals(value: &RefreshIntervals) -> Result<Vec<u8>> {
    validate_intervals(value.intervals())?;
    let size = refresh_xml_len(value)?;
    if size > MAX_OUTPUT_BYTES {
        return Err(super::super::limit("refresh interval XML output bytes"));
    }
    let mut output = Vec::new();
    output
        .try_reserve_exact(size)
        .map_err(|source| allocation("refresh interval XML output", source))?;
    output.extend_from_slice(b"<rr:refreshIntervals xmlns:rr=\"");
    output.extend_from_slice(REFRESH_INTERVALS_NAMESPACE.as_bytes());
    output.extend_from_slice(b"\">");
    for interval in value.intervals() {
        output.extend_from_slice(b"<rr:refreshInterval");
        if let Some(resource_id_int) = interval.resource_id_int() {
            append_attribute(&mut output, "resourceIdInt", &resource_id_int.to_string());
        }
        if let Some(resource_id_str) = interval.resource_id_str() {
            append_attribute(&mut output, "resourceIdStr", resource_id_str);
        }
        append_attribute(&mut output, "interval", &interval.interval().to_string());
        output.extend_from_slice(b"/>");
    }
    output.extend_from_slice(b"</rr:refreshIntervals>");
    debug_assert_eq!(output.len(), size);
    Ok(output)
}

/// Inspect `rvTypesInfo` while retaining source spans for a later local edit.
pub(crate) fn inspect(xml: &[u8]) -> Result<Inspection> {
    // The common DOM pass enforces the package XML byte, depth, node, string,
    // entity, and document-shape limits before the span scanner retains any
    // source-bound metadata.
    let _ = parse_document(xml)?;
    if xml.len() > MAX_XML_BYTES {
        return Err(super::super::limit("rich-value types XML bytes"));
    }

    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = true;
    let origin = ReaderOrigin::of(xml);
    let mut buffer = Vec::new();
    let mut stack = Vec::<Frame>::new();
    let mut types = Vec::<TypeInspection>::new();
    let mut root_seen = false;
    let mut root_closed = false;
    let mut types_seen = false;

    loop {
        let start = checked_position(origin, reader.buffer_position(), xml.len())?;
        let event = reader
            .read_event_into(&mut buffer)
            .map_err(super::super::xml_error)?;
        let end = checked_position(origin, reader.buffer_position(), xml.len())?;
        match event {
            Event::Start(element) => {
                let info = ElementInfo::new(&reader, &element)?;
                if root_closed && stack.is_empty() {
                    return Err(invalid("rich-value types document has multiple roots"));
                }
                let parent = stack.last();
                if parent
                    .is_some_and(|parent| matches!(parent.kind, FrameKind::RefreshInterval(_, _)))
                {
                    return Err(invalid("refreshInterval must not contain child elements"));
                }
                let kind = if !root_seen {
                    if info.namespace != RICH_DATA_2 || info.name != "rvTypesInfo" {
                        return Err(invalid("rich-value types root must be rvTypesInfo"));
                    }
                    no_non_namespace_attributes(&reader, &element, "rvTypesInfo")?;
                    root_seen = true;
                    FrameKind::Root
                } else if parent.is_some_and(|parent| {
                    matches!(parent.kind, FrameKind::Root)
                        && info.namespace == RICH_DATA_2
                        && info.name == "types"
                }) {
                    if types_seen {
                        return Err(invalid("rich-value types has duplicate types containers"));
                    }
                    types_seen = true;
                    FrameKind::Types
                } else if parent.is_some_and(|parent| {
                    matches!(parent.kind, FrameKind::Types)
                        && info.namespace == RICH_DATA_2
                        && info.name == "type"
                }) {
                    let name = type_name(&reader, &element)?;
                    push_bounded(
                        &mut types,
                        MAX_TYPES,
                        TypeInspection {
                            model: TypeRefreshIntervals {
                                name,
                                intervals: None,
                            },
                            refresh: None,
                            extensions: Vec::new(),
                            extension_list_seen: false,
                        },
                        "rich-value types",
                    )?;
                    FrameKind::Type(types.len() - 1)
                } else if parent.is_some_and(|parent| {
                    matches!(parent.kind, FrameKind::Type(_))
                        && parent.namespace == RICH_DATA_2
                        && info.namespace == SPREADSHEETML
                        && info.name == "extLst"
                }) {
                    let type_index = nearest_type(&stack)
                        .ok_or_else(|| invalid("rich-value extension lost its type owner"))?;
                    if types[type_index].extension_list_seen {
                        return Err(invalid(format!(
                            "rich-value type '{}' has duplicate extLst containers",
                            types[type_index].model.name
                        )));
                    }
                    types[type_index].extension_list_seen = true;
                    FrameKind::ExtensionList(type_index)
                } else if parent.is_some_and(|parent| {
                    matches!(parent.kind, FrameKind::ExtensionList(_))
                        && info.namespace == SPREADSHEETML
                        && info.name == "ext"
                }) {
                    let type_index = nearest_type(&stack)
                        .ok_or_else(|| invalid("rich-value extension lost its type owner"))?;
                    let extension_index = push_extension(
                        &mut types[type_index].extensions,
                        ExtensionLocation {
                            start_tag: start..end,
                            close_tag: None,
                            name: bounded_qname(&element)?,
                        },
                    )?;
                    FrameKind::Extension(type_index, extension_index)
                } else if let Some((type_index, extension_index)) = direct_refresh_parent(&stack)
                    .filter(|_| {
                        info.namespace == REFRESH_INTERVALS_NAMESPACE
                            && info.name == "refreshIntervals"
                    })
                {
                    no_non_namespace_attributes(&reader, &element, "refreshIntervals")?;
                    if types[type_index].refresh.is_some() {
                        return Err(invalid(format!(
                            "rich-value type '{}' has duplicate refreshIntervals payloads",
                            types[type_index].model.name
                        )));
                    }
                    let _ = extension_index;
                    FrameKind::Refresh(type_index)
                } else if let Some(type_index) = direct_refresh_owner(&stack) {
                    if info.namespace != REFRESH_INTERVALS_NAMESPACE
                        || info.name != "refreshInterval"
                    {
                        return Err(invalid(
                            "refreshIntervals contains an unexpected child element",
                        ));
                    }
                    let interval = parse_interval_attributes(&reader, &element)?;
                    FrameKind::RefreshInterval(type_index, interval)
                } else if info.namespace == REFRESH_INTERVALS_NAMESPACE {
                    return Err(invalid(
                        "rich-value refresh payload is outside its type ext owner",
                    ));
                } else {
                    FrameKind::Other
                };
                push_bounded(
                    &mut stack,
                    MAX_DEPTH,
                    Frame {
                        kind,
                        namespace: info.namespace,
                        start,
                        non_whitespace_text: false,
                        intervals: Vec::new(),
                    },
                    "rich-value XML depth",
                )?;
            },
            Event::Empty(element) => {
                let info = ElementInfo::new(&reader, &element)?;
                if root_closed && stack.is_empty() {
                    return Err(invalid("rich-value types document has multiple roots"));
                }
                if !root_seen {
                    if info.namespace != RICH_DATA_2 || info.name != "rvTypesInfo" {
                        return Err(invalid("rich-value types root must be rvTypesInfo"));
                    }
                    no_non_namespace_attributes(&reader, &element, "rvTypesInfo")?;
                    root_seen = true;
                    root_closed = true;
                } else if let Some(parent) = stack.last() {
                    if matches!(parent.kind, FrameKind::RefreshInterval(_, _)) {
                        return Err(invalid("refreshInterval must not contain child elements"));
                    }
                    if matches!(parent.kind, FrameKind::Root)
                        && info.namespace == RICH_DATA_2
                        && info.name == "types"
                    {
                        if types_seen {
                            return Err(invalid("rich-value types has duplicate types containers"));
                        }
                        types_seen = true;
                    } else if matches!(parent.kind, FrameKind::Types)
                        && info.namespace == RICH_DATA_2
                        && info.name == "type"
                    {
                        let name = type_name(&reader, &element)?;
                        push_bounded(
                            &mut types,
                            MAX_TYPES,
                            TypeInspection {
                                model: TypeRefreshIntervals {
                                    name,
                                    intervals: None,
                                },
                                refresh: None,
                                extensions: Vec::new(),
                                extension_list_seen: false,
                            },
                            "rich-value types",
                        )?;
                    } else if matches!(parent.kind, FrameKind::Type(_))
                        && info.namespace == SPREADSHEETML
                        && info.name == "extLst"
                    {
                        let type_index = nearest_type(&stack)
                            .ok_or_else(|| invalid("rich-value extension lost its type owner"))?;
                        if types[type_index].extension_list_seen {
                            return Err(invalid(format!(
                                "rich-value type '{}' has duplicate extLst containers",
                                types[type_index].model.name
                            )));
                        }
                        types[type_index].extension_list_seen = true;
                    } else if let Some(type_index) = direct_refresh_owner(&stack) {
                        if info.namespace != REFRESH_INTERVALS_NAMESPACE
                            || info.name != "refreshInterval"
                        {
                            return Err(invalid(
                                "refreshIntervals contains an unexpected child element",
                            ));
                        }
                        let interval = parse_interval_attributes(&reader, &element)?;
                        let refresh = stack
                            .iter_mut()
                            .rev()
                            .find(|frame| matches!(frame.kind, FrameKind::Refresh(_)))
                            .ok_or_else(|| invalid("refreshInterval lost its parent"))?;
                        push_bounded(
                            &mut refresh.intervals,
                            MAX_INTERVALS,
                            interval,
                            "rich-value refresh intervals",
                        )?;
                        let _ = type_index;
                    } else if direct_refresh_parent(&stack).is_some_and(|(_, _)| {
                        info.namespace == REFRESH_INTERVALS_NAMESPACE
                            && info.name == "refreshIntervals"
                    }) {
                        no_non_namespace_attributes(&reader, &element, "refreshIntervals")?;
                        return Err(invalid(
                            "refreshIntervals requires at least one refreshInterval",
                        ));
                    }
                }
                // Empty extension elements are still eligible insertion owners.
                if let Some(parent) = stack.last()
                    && matches!(parent.kind, FrameKind::ExtensionList(_))
                    && info.namespace == SPREADSHEETML
                    && info.name == "ext"
                {
                    let type_index = nearest_type(&stack)
                        .ok_or_else(|| invalid("rich-value extension lost its type owner"))?;
                    push_extension(
                        &mut types[type_index].extensions,
                        ExtensionLocation {
                            start_tag: start..end,
                            close_tag: None,
                            name: bounded_qname(&element)?,
                        },
                    )?;
                }
                if info.namespace == REFRESH_INTERVALS_NAMESPACE
                    && !empty_target_is_in_context(&stack, &info)
                {
                    return Err(invalid(
                        "rich-value refresh payload is outside its type ext owner",
                    ));
                }
            },
            Event::Text(text) => {
                if !text
                    .decode()
                    .map_err(super::super::xml_error)?
                    .trim()
                    .is_empty()
                    && let Some(frame) = stack.last_mut()
                {
                    frame.non_whitespace_text = true;
                }
            },
            Event::CData(text) => {
                if !std::str::from_utf8(text.as_ref())
                    .map_err(super::super::xml_error)?
                    .trim()
                    .is_empty()
                    && let Some(frame) = stack.last_mut()
                {
                    frame.non_whitespace_text = true;
                }
            },
            Event::GeneralRef(reference) => {
                if !reference
                    .decode()
                    .map_err(super::super::xml_error)?
                    .is_empty()
                    && let Some(frame) = stack.last_mut()
                {
                    frame.non_whitespace_text = true;
                }
            },
            Event::End(_) => {
                let frame = stack
                    .pop()
                    .ok_or_else(|| invalid("unexpected rich-value types closing element"))?;
                let range = frame.start..end;
                match frame.kind {
                    FrameKind::RefreshInterval(type_index, interval) => {
                        if frame.non_whitespace_text {
                            return Err(invalid("refreshInterval must not contain character data"));
                        }
                        let refresh = stack
                            .iter_mut()
                            .rev()
                            .find(|parent| matches!(parent.kind, FrameKind::Refresh(_)))
                            .ok_or_else(|| invalid("refreshInterval lost its parent"))?;
                        push_bounded(
                            &mut refresh.intervals,
                            MAX_INTERVALS,
                            interval,
                            "rich-value refresh intervals",
                        )?;
                        let _ = type_index;
                    },
                    FrameKind::Refresh(type_index) => {
                        if frame.non_whitespace_text {
                            return Err(invalid(
                                "refreshIntervals must contain only refreshInterval children",
                            ));
                        }
                        let intervals = RefreshIntervals::new(frame.intervals)?;
                        types[type_index].model.intervals = Some(intervals);
                        types[type_index].refresh = Some(range);
                    },
                    FrameKind::Extension(type_index, extension_index) => {
                        types[type_index].extensions[extension_index].close_tag = Some(start..end);
                    },
                    FrameKind::Root => root_closed = true,
                    FrameKind::Types
                    | FrameKind::Type(_)
                    | FrameKind::ExtensionList(_)
                    | FrameKind::Other => {},
                }
            },
            Event::Decl(_) | Event::Comment(_) => {},
            Event::PI(_) | Event::DocType(_) => {
                return Err(invalid(
                    "rich-value types DTDs and processing instructions are rejected",
                ));
            },
            Event::Eof => break,
        }
        buffer.clear();
    }
    if !root_seen || !root_closed || !stack.is_empty() {
        return Err(invalid("unterminated rich-value types XML"));
    }
    Ok(Inspection { types })
}

/// Rewrite only the selected `refreshIntervals` payload spans.
pub(crate) fn rewrite(
    source: &[u8],
    before: &Inspection,
    after: &[TypeRefreshIntervals],
) -> Result<Vec<u8>> {
    if before.types.len() != after.len() {
        return Err(invalid("rich-value type count changed during refresh edit"));
    }
    // First calculate every replacement span and byte length without building
    // any replacement buffer.  A package may contain many types; checking the
    // aggregate cap only after materializing their payloads would allow a
    // failed edit to allocate far beyond the bounded output budget.
    let mut planned = Vec::<PlannedEdit>::new();
    planned
        .try_reserve(before.types.len().min(MAX_TYPES))
        .map_err(|source| allocation("rich-value refresh edit plan", source))?;
    let mut output_len = source.len();
    for (index, (old, new)) in before.types.iter().zip(after).enumerate() {
        if old.model != *new {
            let (range, replacement_len) = preflight_type_edit(source, index, old, new)?;
            if range.end > source.len() {
                return Err(invalid("rich-value refresh edit span exceeds source"));
            }
            output_len = output_len
                .checked_sub(range.len())
                .and_then(|size| size.checked_add(replacement_len))
                .ok_or_else(|| invalid("rich-value refresh output size overflows"))?;
            planned.push(PlannedEdit {
                range,
                replacement_len,
            });
        }
    }
    if output_len > MAX_OUTPUT_BYTES {
        return Err(super::super::limit("rich-value refresh XML output bytes"));
    }

    let mut edits = Vec::<Edit>::new();
    edits
        .try_reserve(planned.len())
        .map_err(|source| allocation("rich-value refresh edits", source))?;
    let mut planned = planned.into_iter();
    for (index, (old, new)) in before.types.iter().zip(after).enumerate() {
        if old.model != *new {
            let planned = planned
                .next()
                .ok_or_else(|| invalid("rich-value refresh edit plan is incomplete"))?;
            let range = planned.range;
            plan_type_edit(source, index, old, new, &mut edits)?;
            let edit = edits
                .last()
                .ok_or_else(|| invalid("rich-value refresh edit plan was not materialized"))?;
            if edit.range != range || edit.replacement.len() != planned.replacement_len {
                return Err(invalid(
                    "rich-value refresh edit plan changed during materialization",
                ));
            }
        }
    }
    if planned.next().is_some() {
        return Err(invalid("rich-value refresh edit plan has extra entries"));
    }
    if edits.is_empty() {
        return Ok(source.to_vec());
    }
    edits.sort_by_key(|edit| edit.range.start);
    let mut previous_end = 0usize;
    for edit in &edits {
        if edit.range.start < previous_end || edit.range.end > source.len() {
            return Err(invalid("overlapping rich-value refresh edit spans"));
        }
        previous_end = edit.range.end;
    }
    let mut output = Vec::new();
    output
        .try_reserve_exact(output_len)
        .map_err(|source| allocation("rich-value refresh XML output", source))?;
    let mut cursor = 0usize;
    for edit in edits {
        output.extend_from_slice(&source[cursor..edit.range.start]);
        output.extend_from_slice(&edit.replacement);
        cursor = edit.range.end;
    }
    output.extend_from_slice(&source[cursor..]);
    debug_assert_eq!(output.len(), output_len);
    Ok(output)
}

#[derive(Clone, Debug)]
struct Edit {
    range: Range<usize>,
    replacement: Vec<u8>,
}

#[derive(Clone, Debug)]
struct PlannedEdit {
    range: Range<usize>,
    replacement_len: usize,
}

fn preflight_type_edit(
    source: &[u8],
    type_index: usize,
    old: &TypeInspection,
    new: &TypeRefreshIntervals,
) -> Result<(Range<usize>, usize)> {
    let _ = type_index;
    match (&old.refresh, &new.intervals) {
        (Some(range), Some(intervals)) => Ok((range.clone(), refresh_xml_len(intervals)?)),
        (Some(range), None) => Ok((range.clone(), 0)),
        (None, Some(intervals)) => {
            let extension = old.extensions.first().ok_or_else(|| {
                invalid(
                    "cannot author refreshIntervals without an existing rich-value type ext owner",
                )
            })?;
            let payload_len = refresh_xml_len(intervals)?;
            if let Some(close) = &extension.close_tag {
                Ok((close.start..close.start, payload_len))
            } else {
                let start = extension.start_tag.clone();
                let tag = source.get(start.clone()).ok_or_else(|| {
                    invalid("rich-value extension self-closing span exceeds source")
                })?;
                let close = tag
                    .iter()
                    .rposition(|byte| *byte == b'>')
                    .ok_or_else(|| invalid("rich-value extension self-closing tag is malformed"))?;
                let before_close = tag[..close]
                    .iter()
                    .rposition(|byte| !byte.is_ascii_whitespace());
                let slash = before_close
                    .filter(|position| tag[*position] == b'/')
                    .ok_or_else(|| invalid("rich-value extension self-closing tag is malformed"))?;
                let replacement_len = slash
                    .checked_add(1)
                    .and_then(|size| size.checked_add(payload_len))
                    .and_then(|size| size.checked_add(extension.name.len()))
                    .and_then(|size| size.checked_add(3))
                    .ok_or_else(|| invalid("rich-value extension output size overflows"))?;
                Ok((start, replacement_len))
            }
        },
        (None, None) => Ok((0..0, 0)),
    }
}

fn plan_type_edit(
    source: &[u8],
    type_index: usize,
    old: &TypeInspection,
    new: &TypeRefreshIntervals,
    edits: &mut Vec<Edit>,
) -> Result<()> {
    let _ = type_index;
    match (&old.refresh, &new.intervals) {
        (Some(range), Some(intervals)) => edits.push(Edit {
            range: range.clone(),
            replacement: write_refresh_intervals(intervals)?,
        }),
        (Some(range), None) => edits.push(Edit {
            range: range.clone(),
            replacement: Vec::new(),
        }),
        (None, Some(intervals)) => {
            let extension = old.extensions.first().ok_or_else(|| {
                invalid(
                    "cannot author refreshIntervals without an existing rich-value type ext owner",
                )
            })?;
            let payload = write_refresh_intervals(intervals)?;
            if let Some(close) = &extension.close_tag {
                edits.push(Edit {
                    range: close.start..close.start,
                    replacement: payload,
                });
            } else {
                let start = extension.start_tag.clone();
                let tag = &source[start.clone()];
                let close = tag
                    .iter()
                    .rposition(|byte| *byte == b'>')
                    .ok_or_else(|| invalid("rich-value extension self-closing tag is malformed"))?;
                let before_close = tag[..close]
                    .iter()
                    .rposition(|byte| !byte.is_ascii_whitespace());
                let slash = before_close
                    .filter(|position| tag[*position] == b'/')
                    .ok_or_else(|| invalid("rich-value extension self-closing tag is malformed"))?;
                let mut replacement = Vec::new();
                let qname = &extension.name;
                let size = slash
                    .checked_add(1)
                    .and_then(|size| size.checked_add(payload.len()))
                    .and_then(|size| size.checked_add(qname.len()))
                    .and_then(|size| size.checked_add(3))
                    .ok_or_else(|| invalid("rich-value extension output size overflows"))?;
                replacement
                    .try_reserve_exact(size)
                    .map_err(|source| allocation("rich-value extension output", source))?;
                replacement.extend_from_slice(&tag[..slash]);
                replacement.push(b'>');
                replacement.extend_from_slice(&payload);
                replacement.extend_from_slice(b"</");
                replacement.extend_from_slice(qname);
                replacement.push(b'>');
                edits.push(Edit {
                    range: start,
                    replacement,
                });
            }
        },
        (None, None) => {},
    }
    Ok(())
}

fn parse_refresh_node(node: &Node) -> Result<RefreshIntervals> {
    require(node, REFRESH_INTERVALS_NAMESPACE, "refreshIntervals")?;
    no_attributes(node, &[])?;
    whitespace_around_children(node)?;
    let mut intervals = Vec::new();
    for child in &node.children {
        if child.namespace != REFRESH_INTERVALS_NAMESPACE || child.name != "refreshInterval" {
            return Err(invalid(
                "refreshIntervals contains an unexpected child element",
            ));
        }
        push_bounded(
            &mut intervals,
            MAX_INTERVALS,
            parse_interval_node(child)?,
            "refresh intervals",
        )?;
    }
    RefreshIntervals::new(intervals)
}

fn parse_interval_node(node: &Node) -> Result<RefreshInterval> {
    no_attributes(
        node,
        &[
            ("", "resourceIdInt"),
            ("", "resourceIdStr"),
            ("", "interval"),
        ],
    )?;
    whitespace(node)?;
    let resource_id_int = optional(node, "", "resourceIdInt")
        .map(|value| parse_i32(value, "resourceIdInt"))
        .transpose()?;
    let resource_id_str = optional(node, "", "resourceIdStr").map(str::to_owned);
    let interval = parse_i32(required(node, "", "interval")?, "interval")?;
    RefreshInterval::new(resource_id_int, resource_id_str, interval)
}

fn whitespace_around_children(node: &Node) -> Result<()> {
    if node.text.trim().is_empty() {
        Ok(())
    } else {
        Err(invalid(
            "refreshIntervals contains unexpected character data",
        ))
    }
}

fn parse_i32(value: &str, field: &str) -> Result<i32> {
    value
        .trim_matches(|character| matches!(character, ' ' | '\t' | '\n' | '\r'))
        .parse::<i32>()
        .map_err(|_| invalid(format!("refresh interval '{field}' is not an xsd:int")))
}

fn refresh_xml_len(value: &RefreshIntervals) -> Result<usize> {
    let mut size = b"<rr:refreshIntervals xmlns:rr=\"\">"
        .len()
        .checked_add(REFRESH_INTERVALS_NAMESPACE.len())
        .and_then(|size| size.checked_add(b"</rr:refreshIntervals>".len()))
        .ok_or_else(|| invalid("refresh interval XML size overflows"))?;
    for interval in value.intervals() {
        size = size
            .checked_add(b"<rr:refreshInterval".len())
            .and_then(|size| {
                interval.resource_id_int().map_or(Some(size), |value| {
                    checked_attribute_len(size, "resourceIdInt", &value.to_string())
                })
            })
            .and_then(|size| {
                interval.resource_id_str().map_or(Some(size), |value| {
                    checked_attribute_len(size, "resourceIdStr", value)
                })
            })
            .and_then(|size| {
                checked_attribute_len(size, "interval", &interval.interval().to_string())
            })
            .and_then(|size| size.checked_add(2))
            .ok_or_else(|| invalid("refresh interval XML size overflows"))?;
    }
    Ok(size)
}

fn checked_attribute_len(current: usize, name: &str, value: &str) -> Option<usize> {
    let encoded = escaped_attr_len(value).ok()?;
    current
        .checked_add(1 + name.len() + 3)
        .and_then(|size| size.checked_add(encoded))
}

fn escaped_attr_len(value: &str) -> Result<usize> {
    value.chars().try_fold(0usize, |length, character| {
        if !valid_xml10(character) {
            return Err(invalid(
                "refresh interval attribute contains an XML 1.0 invalid character",
            ));
        }
        let encoded = match character {
            '&' => 5,
            '<' => 4,
            '"' => 6,
            '\t' | '\n' | '\r' => 5,
            _ => character.len_utf8(),
        };
        length
            .checked_add(encoded)
            .ok_or_else(|| invalid("refresh interval XML size overflows"))
    })
}

fn append_attribute(output: &mut Vec<u8>, name: &str, value: &str) {
    output.push(b' ');
    output.extend_from_slice(name.as_bytes());
    output.extend_from_slice(b"=\"");
    append_escaped_attribute(output, value);
    output.push(b'"');
}

fn append_escaped_attribute(output: &mut Vec<u8>, value: &str) {
    for character in value.chars() {
        match character {
            '&' => output.extend_from_slice(b"&amp;"),
            '<' => output.extend_from_slice(b"&lt;"),
            '"' => output.extend_from_slice(b"&quot;"),
            '\t' => output.extend_from_slice(b"&#x9;"),
            '\n' => output.extend_from_slice(b"&#xA;"),
            '\r' => output.extend_from_slice(b"&#xD;"),
            _ => {
                let mut bytes = [0; 4];
                output.extend_from_slice(character.encode_utf8(&mut bytes).as_bytes());
            },
        }
    }
}

#[derive(Clone, Debug)]
struct ElementInfo {
    namespace: String,
    name: String,
}

impl ElementInfo {
    fn new(reader: &NsReader<&[u8]>, element: &BytesStart<'_>) -> Result<Self> {
        let namespace = resolve_namespace(reader.resolver().resolve_element(element.name()).0)?;
        let name = std::str::from_utf8(element.local_name().as_ref())
            .map_err(super::super::xml_error)?
            .to_owned();
        Ok(Self { namespace, name })
    }
}

#[derive(Clone, Debug)]
struct Frame {
    kind: FrameKind,
    namespace: String,
    start: usize,
    non_whitespace_text: bool,
    intervals: Vec<RefreshInterval>,
}

#[derive(Clone, Debug)]
enum FrameKind {
    Root,
    Types,
    Type(usize),
    ExtensionList(usize),
    Extension(usize, usize),
    Refresh(usize),
    RefreshInterval(usize, RefreshInterval),
    Other,
}

fn nearest_type(stack: &[Frame]) -> Option<usize> {
    stack.iter().rev().find_map(|frame| match frame.kind {
        FrameKind::Type(index)
        | FrameKind::ExtensionList(index)
        | FrameKind::Extension(index, _)
        | FrameKind::Refresh(index)
        | FrameKind::RefreshInterval(index, _) => Some(index),
        FrameKind::Root | FrameKind::Types | FrameKind::Other => None,
    })
}

fn direct_refresh_parent(stack: &[Frame]) -> Option<(usize, usize)> {
    let parent = stack.last()?;
    if let FrameKind::Extension(type_index, extension_index) = parent.kind {
        Some((type_index, extension_index))
    } else {
        None
    }
}

fn direct_refresh_owner(stack: &[Frame]) -> Option<usize> {
    stack.last().and_then(|frame| match frame.kind {
        FrameKind::Refresh(type_index) => Some(type_index),
        _ => None,
    })
}

fn empty_target_is_in_context(stack: &[Frame], info: &ElementInfo) -> bool {
    match stack.last().map(|frame| &frame.kind) {
        Some(FrameKind::Refresh(_)) => info.name == "refreshInterval",
        Some(FrameKind::Extension(_, _)) => info.name == "refreshIntervals",
        _ => false,
    }
}

fn type_name(reader: &NsReader<&[u8]>, element: &BytesStart<'_>) -> Result<String> {
    let mut result = None;
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(super::super::xml_error)?;
        if attribute.key.as_ref() == b"xmlns" || attribute.key.as_ref().starts_with(b"xmlns:") {
            continue;
        }
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        let namespace = resolve_namespace(namespace)?;
        let name = std::str::from_utf8(local.as_ref()).map_err(super::super::xml_error)?;
        if !namespace.is_empty() || name != "name" {
            return Err(invalid("rich-value type has an unexpected attribute"));
        }
        if result.is_some() {
            return Err(invalid("rich-value type has duplicate name attributes"));
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
            .map_err(super::super::xml_error)?
            .into_owned();
        if value.chars().count() > TYPE_NAME_MAX_CHARS {
            return Err(invalid("rich-value type name exceeds 255 characters"));
        }
        if !value.chars().all(valid_xml10) {
            return Err(invalid(
                "rich-value type name contains an XML 1.0 invalid character",
            ));
        }
        result = Some(value);
    }
    result.ok_or_else(|| invalid("rich-value type is missing its name attribute"))
}

fn parse_interval_attributes(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
) -> Result<RefreshInterval> {
    let mut resource_id_int = None;
    let mut resource_id_str = None;
    let mut interval = None;
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(super::super::xml_error)?;
        if attribute.key.as_ref() == b"xmlns" || attribute.key.as_ref().starts_with(b"xmlns:") {
            continue;
        }
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        let namespace = resolve_namespace(namespace)?;
        if !namespace.is_empty() {
            return Err(invalid("refreshInterval has a namespaced attribute"));
        }
        let name = std::str::from_utf8(local.as_ref()).map_err(super::super::xml_error)?;
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
            .map_err(super::super::xml_error)?
            .into_owned();
        match name {
            "resourceIdInt" if resource_id_int.is_none() => {
                resource_id_int = Some(parse_i32(&value, "resourceIdInt")?);
            },
            "resourceIdStr" if resource_id_str.is_none() => resource_id_str = Some(value),
            "interval" if interval.is_none() => interval = Some(parse_i32(&value, "interval")?),
            _ => {
                return Err(invalid(format!(
                    "refreshInterval has unexpected attribute '{name}'"
                )));
            },
        }
    }
    RefreshInterval::new(
        resource_id_int,
        resource_id_str,
        interval.ok_or_else(|| invalid("refreshInterval is missing its interval attribute"))?,
    )
}

fn no_non_namespace_attributes(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    owner: &str,
) -> Result<()> {
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(super::super::xml_error)?;
        if attribute.key.as_ref() == b"xmlns" || attribute.key.as_ref().starts_with(b"xmlns:") {
            continue;
        }
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        let namespace = resolve_namespace(namespace)?;
        return Err(invalid(format!(
            "{owner} has unexpected {}attribute '{}'",
            if namespace.is_empty() {
                ""
            } else {
                "namespaced "
            },
            String::from_utf8_lossy(local.as_ref())
        )));
    }
    Ok(())
}

fn resolve_namespace(value: ResolveResult<'_>) -> Result<String> {
    match value {
        ResolveResult::Bound(namespace) => Ok(std::str::from_utf8(namespace.0)
            .map_err(super::super::xml_error)?
            .to_owned()),
        ResolveResult::Unbound => Ok(String::new()),
        ResolveResult::Unknown(prefix) => Err(invalid(format!(
            "unbound XML namespace prefix '{}'",
            String::from_utf8_lossy(prefix.as_ref())
        ))),
    }
}

fn bounded_qname(element: &BytesStart<'_>) -> Result<Vec<u8>> {
    let name = element.name();
    if name.as_ref().len() > MAX_STRING_BYTES {
        return Err(super::super::limit("rich-value extension name"));
    }
    Ok(name.as_ref().to_vec())
}

/// The byte offset in the source of reader position `position`; `origin` is
/// the source's [`ReaderOrigin`].
fn checked_position(origin: ReaderOrigin, position: u64, source_len: usize) -> Result<usize> {
    let position = origin
        .offset(position)
        .ok_or_else(|| invalid("rich-value XML source position overflows usize"))?;
    if position > source_len {
        return Err(invalid(
            "rich-value XML source position exceeds source bytes",
        ));
    }
    Ok(position)
}

fn push_extension(values: &mut Vec<ExtensionLocation>, value: ExtensionLocation) -> Result<usize> {
    if values.len() >= MAX_INTERVALS {
        return Err(super::super::limit("rich-value extension locations"));
    }
    values
        .try_reserve(1)
        .map_err(|source| allocation("rich-value extension locations", source))?;
    values.push(value);
    Ok(values.len() - 1)
}

fn push_bounded<T>(
    values: &mut Vec<T>,
    maximum: usize,
    value: T,
    resource: &'static str,
) -> Result<()> {
    if values.len() >= maximum {
        return Err(super::super::limit(resource));
    }
    values
        .try_reserve(1)
        .map_err(|source| allocation(resource, source))?;
    values.push(value);
    Ok(())
}

fn allocation(resource: &'static str, source: TryReserveError) -> Error {
    Error::Allocation { resource, source }
}
