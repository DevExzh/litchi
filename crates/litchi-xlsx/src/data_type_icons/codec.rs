//! Bounded extension inspection and source-local replacement.

use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;

use super::model::ShowDataTypeIcons;
use super::{
    MAX_DEPTH, MAX_EXTENSION_BYTES, MAX_NODES, MAX_PART_BYTES, SHOW_DATA_TYPE_ICONS_NAMESPACE,
    Target, invalid,
};

const CORE: &[u8] = b"http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const STRICT: &[u8] = b"http://purl.oclc.org/ooxml/spreadsheetml/main";
const NAMED: &[u8] = b"http://schemas.microsoft.com/office/spreadsheetml/2019/namedsheetviews";

/// Parse an already captured SpreadsheetML `ext` element.
pub(crate) fn inspect_extension(
    markup: &[u8],
    target: Target,
) -> crate::Result<Option<ShowDataTypeIcons>> {
    if markup.is_empty() || markup.len() > MAX_EXTENSION_BYTES {
        return Err(invalid("data-type-icon extension exceeds its byte limit"));
    }
    let mut reader = NsReader::from_reader(markup);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = true;
    let mut buffer = Vec::new();
    let mut stack = Vec::<Frame>::new();
    let mut root_seen = false;
    let mut root_closed = false;
    let mut nodes = 0usize;
    let mut found = None;

    loop {
        let event = reader
            .read_event_into(&mut buffer)
            .map_err(|error| invalid(error.to_string()))?;
        if !matches!(event, Event::Eof) {
            nodes = nodes
                .checked_add(1)
                .ok_or_else(|| invalid("data-type-icon extension node count overflow"))?;
            if nodes > MAX_NODES {
                return Err(invalid("data-type-icon extension node count exceeds limit"));
            }
        }
        let resolver = reader.resolver().clone();
        match event {
            Event::Start(element) => {
                let (namespace, local) = resolve(&resolver, element.name());
                if !root_seen {
                    if local != b"ext" {
                        return Err(invalid(
                            "data-type-icon markup must have a SpreadsheetML ext root",
                        ));
                    }
                    root_seen = true;
                    stack.push(Frame::Ext);
                    continue;
                }
                if root_closed && stack.is_empty() {
                    return Err(invalid("data-type-icon extension has multiple roots"));
                }
                if stack.len() >= MAX_DEPTH {
                    return Err(invalid("data-type-icon extension depth exceeds limit"));
                }
                let parent = stack.last().copied();
                if matches!(parent, Some(Frame::Ext)) && namespace.is_target() {
                    if local != target.local_name() {
                        return Err(invalid(format!(
                            "unexpected data-type-icon namespace element '{}', expected {}",
                            String::from_utf8_lossy(local),
                            target.display_name()
                        )));
                    }
                    if found.is_some() {
                        return Err(invalid(format!(
                            "duplicate {} payload",
                            target.display_name()
                        )));
                    }
                    found = Some(parse_target_attributes(&element, reader.decoder(), target)?);
                    stack.push(Frame::Target);
                } else if matches!(parent, Some(Frame::Target)) {
                    return Err(invalid(format!(
                        "{} must not contain child elements",
                        target.display_name()
                    )));
                } else {
                    stack.push(Frame::Other);
                }
            },
            Event::Empty(element) => {
                let (namespace, local) = resolve(&resolver, element.name());
                if !root_seen {
                    if local != b"ext" {
                        return Err(invalid(
                            "data-type-icon markup must have a SpreadsheetML ext root",
                        ));
                    }
                    root_seen = true;
                    root_closed = true;
                    continue;
                }
                let parent = stack.last().copied();
                if matches!(parent, Some(Frame::Ext)) && namespace.is_target() {
                    if local != target.local_name() {
                        return Err(invalid(format!(
                            "unexpected data-type-icon namespace element '{}', expected {}",
                            String::from_utf8_lossy(local),
                            target.display_name()
                        )));
                    }
                    if found.is_some() {
                        return Err(invalid(format!(
                            "duplicate {} payload",
                            target.display_name()
                        )));
                    }
                    found = Some(parse_target_attributes(&element, reader.decoder(), target)?);
                } else if matches!(parent, Some(Frame::Target)) {
                    return Err(invalid(format!(
                        "{} must not contain child elements",
                        target.display_name()
                    )));
                }
            },
            Event::Text(_) => {
                if matches!(stack.last(), Some(Frame::Target)) {
                    return Err(invalid(format!(
                        "{} must not contain character content",
                        target.display_name()
                    )));
                }
            },
            Event::CData(_) => {
                if matches!(stack.last(), Some(Frame::Target)) {
                    return Err(invalid(format!(
                        "{} must not contain character content",
                        target.display_name()
                    )));
                }
            },
            Event::End(_) => {
                let frame = stack
                    .pop()
                    .ok_or_else(|| invalid("data-type-icon extension has an unmatched end"))?;
                if matches!(frame, Frame::Ext) {
                    root_closed = true;
                }
            },
            Event::Eof => break,
            Event::Decl(_) | Event::Comment(_) => {},
            Event::PI(_) | Event::DocType(_) | Event::GeneralRef(_) => {
                return Err(invalid(
                    "DTD, processing instructions, and entities are rejected",
                ));
            },
        }
    }
    if !root_seen || !root_closed || !stack.is_empty() {
        return Err(invalid("unterminated data-type-icon extension"));
    }
    Ok(found)
}

/// Inspect an event while its parent `ext` is still in the reader's namespace
/// scope.  Captured extension fragments intentionally retain their original
/// lexical bytes and therefore may not carry inherited `xmlns` declarations;
/// this hook keeps namespace validation correct without rebuilding the DOM.
pub(crate) fn observe_event(
    event: &Event<'_>,
    resolver: &quick_xml::name::NamespaceResolver,
    decoder: quick_xml::encoding::Decoder,
    target: Target,
    depth: usize,
    seen: &mut bool,
    open_depth: &mut Option<usize>,
) -> crate::Result<Option<ShowDataTypeIcons>> {
    if let Some(open) = *open_depth {
        match event {
            Event::Start(_) | Event::Empty(_) => {
                return Err(invalid(format!(
                    "{} must not contain child elements",
                    target.display_name()
                )));
            },
            Event::Text(_) => {
                if *open_depth == Some(open) {
                    return Err(invalid(format!(
                        "{} must not contain character content",
                        target.display_name()
                    )));
                }
            },
            Event::CData(_) => {
                if *open_depth == Some(open) {
                    return Err(invalid(format!(
                        "{} must not contain character content",
                        target.display_name()
                    )));
                }
            },
            Event::GeneralRef(_) => {
                return Err(invalid(format!(
                    "{} must not contain entity references or text",
                    target.display_name()
                )));
            },
            Event::End(_) if depth == open => *open_depth = None,
            _ => {},
        }
    }
    let direct = depth == 1;
    let element = match event {
        Event::Start(element) | Event::Empty(element) => Some(element),
        _ => None,
    };
    let Some(element) = element else {
        return Ok(None);
    };
    let (namespace, local) = resolve(resolver, element.name());
    if !namespace.is_target() {
        return Ok(None);
    }
    if !direct {
        return Err(invalid(format!(
            "{} is outside its direct ext owner",
            target.display_name()
        )));
    }
    if local != target.local_name() {
        return Err(invalid(format!(
            "unexpected data-type-icon namespace element '{}', expected {}",
            String::from_utf8_lossy(local),
            target.display_name()
        )));
    }
    if *seen {
        return Err(invalid(format!(
            "duplicate {} payload",
            target.display_name()
        )));
    }
    *seen = true;
    let value = parse_target_attributes(element, decoder, target)?;
    if matches!(event, Event::Start(_)) {
        *open_depth = Some(depth + 1);
    }
    Ok(Some(value))
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum Frame {
    Ext,
    Other,
    Target,
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum NamespaceKind {
    Core,
    Strict,
    Target,
    Named,
    Other,
}

impl NamespaceKind {
    const fn is_core(self) -> bool {
        matches!(self, Self::Core | Self::Strict)
    }

    const fn is_target(self) -> bool {
        matches!(self, Self::Target)
    }

    const fn is_named(self) -> bool {
        matches!(self, Self::Named)
    }
}

fn resolve<'a>(
    resolver: &quick_xml::name::NamespaceResolver,
    name: quick_xml::name::QName<'a>,
) -> (NamespaceKind, &'a [u8]) {
    let (resolved, local) = resolver.resolve_element(name);
    let namespace = match resolved {
        ResolveResult::Bound(Namespace(value)) if value == CORE => NamespaceKind::Core,
        ResolveResult::Bound(Namespace(value)) if value == STRICT => NamespaceKind::Strict,
        ResolveResult::Bound(Namespace(value))
            if value == SHOW_DATA_TYPE_ICONS_NAMESPACE.as_bytes() =>
        {
            NamespaceKind::Target
        },
        ResolveResult::Bound(Namespace(value)) if value == NAMED => NamespaceKind::Named,
        ResolveResult::Bound(_) | ResolveResult::Unbound | ResolveResult::Unknown(_) => {
            NamespaceKind::Other
        },
    };
    (namespace, local.into_inner())
}

fn parse_target_attributes(
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
    target: Target,
) -> crate::Result<ShowDataTypeIcons> {
    let mut visible = true;
    let mut seen = false;
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(|error| invalid(error.to_string()))?;
        let key = attribute.key.as_ref();
        if key == b"visible" {
            if seen {
                return Err(invalid(format!(
                    "duplicate {} visible attribute",
                    target.display_name()
                )));
            }
            seen = true;
            let value = attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
                .map_err(|error| invalid(error.to_string()))?;
            visible = ShowDataTypeIcons::parse_boolean(&value)?;
        } else if key == b"xmlns" || key.starts_with(b"xmlns:") {
            // Namespace declarations are source syntax, not CT attributes.
        } else {
            return Err(invalid(format!(
                "unexpected {} attribute '{}'",
                target.display_name(),
                String::from_utf8_lossy(key)
            )));
        }
    }
    Ok(ShowDataTypeIcons::new(visible))
}

pub(crate) fn write_target(value: ShowDataTypeIcons, target: Target) -> Vec<u8> {
    let local = std::str::from_utf8(target.local_name()).unwrap_or("showDataTypeIcons");
    let visible = if value.visible() { "1" } else { "0" };
    format!(r#"<sdt:{local} xmlns:sdt="{SHOW_DATA_TYPE_ICONS_NAMESPACE}" visible="{visible}"/>"#)
        .into_bytes()
}

/// Rewrite one already captured `x:ext` fragment.  This is used by authored
/// model values where the fragment is kept in a named-sheet-view extension.
pub(crate) fn rewrite_extension(
    markup: &[u8],
    target: Target,
    value: Option<ShowDataTypeIcons>,
) -> crate::Result<Vec<u8>> {
    if markup.is_empty() || markup.len() > MAX_EXTENSION_BYTES {
        return Err(invalid("data-type-icon extension exceeds its byte limit"));
    }
    let owner = Span {
        start: 0,
        end: markup.len(),
    };
    let found = find_target(markup, owner, target)?;
    match (found, value) {
        (Some(span), Some(value)) => replace_span(markup, span, &write_target(value, target)),
        (Some(span), None) => replace_span(markup, span, &[]),
        (None, None) => Ok(markup.to_vec()),
        (None, Some(value)) => {
            let replacement = write_target(value, target);
            let mut output = Vec::new();
            if markup.ends_with(b"/>\n") || markup.ends_with(b"/>") {
                let slash = markup
                    .iter()
                    .rposition(|byte| *byte == b'/')
                    .ok_or_else(|| invalid("data-type-icon extension opening tag is malformed"))?;
                let name = opening_name(markup)?;
                let size = markup
                    .len()
                    .checked_add(replacement.len())
                    .and_then(|size| size.checked_add(name.len() + 3))
                    .ok_or_else(|| invalid("data-type-icon output size overflow"))?;
                reserve_output(&mut output, size)?;
                output.extend_from_slice(&markup[..slash]);
                output.push(b'>');
                output.extend_from_slice(&replacement);
                output.extend_from_slice(b"</");
                output.extend_from_slice(name);
                output.push(b'>');
                output.extend_from_slice(&markup[slash + 2..]);
                Ok(output)
            } else {
                let close_start = markup
                    .iter()
                    .rposition(|byte| *byte == b'<')
                    .ok_or_else(|| invalid("data-type-icon extension closing tag is malformed"))?;
                let size = markup
                    .len()
                    .checked_add(replacement.len())
                    .ok_or_else(|| invalid("data-type-icon output size overflow"))?;
                reserve_output(&mut output, size)?;
                output.extend_from_slice(&markup[..close_start]);
                output.extend_from_slice(&replacement);
                output.extend_from_slice(&markup[close_start..]);
                Ok(output)
            }
        },
    }
}

#[derive(Clone, Copy, Debug)]
pub(crate) enum OwnerKind {
    WorksheetView,
    CustomSheetView,
}

#[derive(Clone, Copy, Debug)]
struct Span {
    start: usize,
    end: usize,
}

#[derive(Clone, Copy, Debug)]
struct OwnerSpan {
    span: Span,
    tag_end: usize,
    close_start: Option<usize>,
    owner_index: usize,
    target_span: Option<Span>,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum NodeKind {
    Worksheet,
    SheetViews,
    SheetView,
    NamedViews,
    NamedView,
    ExtList,
    Ext,
    Target,
    Other,
}

#[derive(Clone, Copy, Debug)]
struct Node {
    kind: NodeKind,
    owner_index: usize,
    start: usize,
    tag_end: usize,
    target_span: Option<Span>,
}

/// Replace one direct extension payload while retaining every other source
/// byte.  New payloads require an existing, unambiguous `ext` owner because
/// [MS-XLSX] does not assign a canonical URI to this feature.
pub(crate) fn rewrite(
    xml: &[u8],
    owner_kind: OwnerKind,
    owner_index: usize,
    value: Option<ShowDataTypeIcons>,
) -> crate::Result<Vec<u8>> {
    if xml.len() > MAX_PART_BYTES {
        return Err(invalid("data-type-icon source part exceeds its byte limit"));
    }
    let owners = scan_owners(xml, owner_kind)?;
    let matching = owners
        .iter()
        .filter(|owner| owner.owner_index == owner_index)
        .copied()
        .collect::<Vec<_>>();
    if matching.is_empty() {
        return Err(invalid(
            "data-type-icon edit requires one direct source owner (an existing ext owner); MCE-selected or unsupported owner markup is not span-stable",
        ));
    }
    let target = match owner_kind {
        OwnerKind::WorksheetView => Target::Worksheet,
        OwnerKind::CustomSheetView => Target::CustomSheetView,
    };
    let mut found = None;
    let mut selected_owner = None;
    for owner in matching.iter().copied() {
        if let Some(span) = owner.target_span {
            if found.is_some() {
                return Err(invalid(format!(
                    "duplicate {} payload",
                    target.display_name()
                )));
            }
            found = Some(span);
            selected_owner = Some(owner);
        }
    }
    if let (Some(span), Some(value)) = (found, value) {
        return replace_span(xml, span, &write_target(value, target));
    }
    if let (Some(span), None) = (found, value) {
        return replace_span(xml, span, &[]);
    }
    if value.is_none() {
        return Ok(xml.to_vec());
    }
    if matching.len() != 1 {
        return Err(invalid(
            "cannot add data-type-icon payload without one existing ext owner",
        ));
    }
    let owner = selected_owner.unwrap_or(matching[0]);
    let Some(value) = value else {
        return Ok(xml.to_vec());
    };
    let replacement = write_target(value, target);
    if let Some(close_start) = owner.close_start {
        let mut output = Vec::new();
        let size = xml
            .len()
            .checked_add(replacement.len())
            .ok_or_else(|| invalid("data-type-icon output size overflow"))?;
        reserve_output(&mut output, size)?;
        output.extend_from_slice(&xml[..close_start]);
        output.extend_from_slice(&replacement);
        output.extend_from_slice(&xml[close_start..]);
        Ok(output)
    } else {
        let opening = &xml[owner.span.start..owner.tag_end];
        let name = opening_name(opening)?;
        let mut output = Vec::new();
        let size = xml
            .len()
            .checked_add(replacement.len())
            .and_then(|size| size.checked_add(name.len() + 3))
            .ok_or_else(|| invalid("data-type-icon output size overflow"))?;
        reserve_output(&mut output, size)?;
        output.extend_from_slice(&xml[..owner.span.start]);
        output.extend_from_slice(&opening[..opening.len() - 2]);
        output.push(b'>');
        output.extend_from_slice(&replacement);
        output.extend_from_slice(b"</");
        output.extend_from_slice(name);
        output.push(b'>');
        output.extend_from_slice(&xml[owner.tag_end..]);
        Ok(output)
    }
}

fn reserve_output(output: &mut Vec<u8>, size: usize) -> crate::Result<()> {
    if size > MAX_PART_BYTES {
        return Err(invalid("data-type-icon output exceeds its byte limit"));
    }
    output
        .try_reserve_exact(size)
        .map_err(|_| invalid("data-type-icon output allocation failed"))
}

fn replace_span(xml: &[u8], span: Span, replacement: &[u8]) -> crate::Result<Vec<u8>> {
    let size = xml
        .len()
        .checked_sub(span.end.saturating_sub(span.start))
        .and_then(|size| size.checked_add(replacement.len()))
        .ok_or_else(|| invalid("data-type-icon output size overflow"))?;
    let mut output = Vec::new();
    reserve_output(&mut output, size)?;
    output.extend_from_slice(&xml[..span.start]);
    output.extend_from_slice(replacement);
    output.extend_from_slice(&xml[span.end..]);
    Ok(output)
}

fn opening_name(opening: &[u8]) -> crate::Result<&[u8]> {
    let start = opening
        .iter()
        .position(|byte| *byte == b'<')
        .ok_or_else(|| invalid("data-type-icon owner opening tag is malformed"))?
        + 1;
    let end = opening[start..]
        .iter()
        .position(|byte| matches!(byte, b' ' | b'>' | b'/'))
        .map(|offset| start + offset)
        .unwrap_or(opening.len());
    Ok(&opening[start..end])
}

fn find_target(xml: &[u8], owner: Span, target: Target) -> crate::Result<Option<Span>> {
    let mut reader = NsReader::from_reader(&xml[owner.start..owner.end]);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = true;
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut nodes = 0usize;
    let mut candidate_start = None;
    let mut found = None;
    loop {
        let start = reader.buffer_position() as usize + owner.start;
        let event = reader
            .read_event_into(&mut buffer)
            .map_err(|error| invalid(error.to_string()))?;
        let end = reader.buffer_position() as usize + owner.start;
        if !matches!(event, Event::Eof) {
            nodes = nodes
                .checked_add(1)
                .ok_or_else(|| invalid("data-type-icon target node count overflow"))?;
            if nodes > MAX_NODES {
                return Err(invalid("data-type-icon target node count exceeds limit"));
            }
        }
        let resolver = reader.resolver().clone();
        match event {
            Event::Start(element) => {
                let (namespace, local) = resolve(&resolver, element.name());
                if depth == 1 && namespace.is_target() {
                    if local != target.local_name() {
                        return Err(invalid("unexpected data-type-icon payload"));
                    }
                    if found.is_some() {
                        return Err(invalid(format!(
                            "duplicate {} payload",
                            target.display_name()
                        )));
                    }
                    parse_target_attributes(&element, reader.decoder(), target)?;
                    candidate_start = Some(start);
                } else if candidate_start.is_some() && depth >= 2 {
                    return Err(invalid(format!(
                        "{} must not contain child elements",
                        target.display_name()
                    )));
                }
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| invalid("data-type-icon target depth overflow"))?;
            },
            Event::Empty(element) => {
                let (namespace, local) = resolve(&resolver, element.name());
                if depth == 1 && namespace.is_target() {
                    if local != target.local_name() {
                        return Err(invalid("unexpected data-type-icon payload"));
                    }
                    if found.is_some() {
                        return Err(invalid(format!(
                            "duplicate {} payload",
                            target.display_name()
                        )));
                    }
                    parse_target_attributes(&element, reader.decoder(), target)?;
                    found = Some(Span { start, end });
                }
            },
            Event::End(_) => {
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("data-type-icon target depth underflow"))?;
                if let Some(candidate) = candidate_start
                    && depth == 1
                {
                    found = Some(Span {
                        start: candidate,
                        end,
                    });
                    candidate_start = None;
                }
            },
            Event::Eof => break,
            Event::Text(_) | Event::CData(_) if candidate_start.is_some() => {
                return Err(invalid(format!(
                    "{} must not contain character content",
                    target.display_name()
                )));
            },
            Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::DocType(_) => {},
            Event::GeneralRef(_) if candidate_start.is_some() => {
                return Err(invalid(format!(
                    "{} must not contain entity references or text",
                    target.display_name()
                )));
            },
            Event::GeneralRef(_) => {},
        }
    }
    Ok(found)
}

fn scan_owners(xml: &[u8], owner_kind: OwnerKind) -> crate::Result<Vec<OwnerSpan>> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = true;
    let mut buffer = Vec::new();
    let mut stack = Vec::<Node>::new();
    let mut owners = Vec::new();
    let mut next_owner_index = 0usize;
    let mut nodes = 0usize;

    loop {
        let start = reader.buffer_position() as usize;
        let event = reader
            .read_event_into(&mut buffer)
            .map_err(|error| invalid(error.to_string()))?;
        let end = reader.buffer_position() as usize;
        if !matches!(event, Event::Eof) {
            nodes = nodes
                .checked_add(1)
                .ok_or_else(|| invalid("data-type-icon source node count overflow"))?;
            if nodes > MAX_NODES {
                return Err(invalid("data-type-icon source node count exceeds limit"));
            }
        }
        let resolver = reader.resolver().clone();
        match event {
            Event::Start(element) => {
                let (namespace, local) = resolve(&resolver, element.name());
                let parent = stack.last().copied();
                if parent.is_some_and(|parent| parent.kind == NodeKind::Target) {
                    return Err(invalid(format!(
                        "{} must not contain child elements",
                        target_for(owner_kind).display_name()
                    )));
                }
                let node_kind = classify(owner_kind, parent, namespace, local);
                if node_kind == NodeKind::Target && local != target_for(owner_kind).local_name() {
                    return Err(invalid(format!(
                        "unexpected data-type-icon namespace element '{}', expected {}",
                        String::from_utf8_lossy(local),
                        target_for(owner_kind).display_name()
                    )));
                }
                if node_kind == NodeKind::Target {
                    let owner = parent
                        .ok_or_else(|| invalid("data-type-icon target lost its ext owner"))?;
                    if owner.target_span.is_some() {
                        return Err(invalid(format!(
                            "duplicate {} payload",
                            target_for(owner_kind).display_name()
                        )));
                    }
                    parse_target_attributes(&element, reader.decoder(), target_for(owner_kind))?;
                }
                let owner_index = if matches!(node_kind, NodeKind::SheetView | NodeKind::NamedView)
                {
                    let index = next_owner_index;
                    next_owner_index = next_owner_index
                        .checked_add(1)
                        .ok_or_else(|| invalid("data-type-icon owner count overflow"))?;
                    index
                } else if matches!(node_kind, NodeKind::ExtList | NodeKind::Ext) {
                    parent.map_or(0, |parent| parent.owner_index)
                } else {
                    parent.map_or(0, |parent| parent.owner_index)
                };
                stack.push(Node {
                    kind: node_kind,
                    owner_index,
                    start,
                    tag_end: end,
                    target_span: None,
                });
            },
            Event::Empty(element) => {
                let (namespace, local) = resolve(&resolver, element.name());
                let parent = stack.last().copied();
                if parent.is_some_and(|parent| parent.kind == NodeKind::Target) {
                    return Err(invalid(format!(
                        "{} must not contain child elements",
                        target_for(owner_kind).display_name()
                    )));
                }
                let node_kind = classify(owner_kind, parent, namespace, local);
                if node_kind == NodeKind::Ext {
                    owners.push(OwnerSpan {
                        span: Span { start, end },
                        tag_end: end,
                        close_start: None,
                        owner_index: parent.map_or(0, |parent| parent.owner_index),
                        target_span: None,
                    });
                } else if node_kind == NodeKind::Target {
                    if local != target_for(owner_kind).local_name() {
                        return Err(invalid(format!(
                            "unexpected data-type-icon namespace element '{}', expected {}",
                            String::from_utf8_lossy(local),
                            target_for(owner_kind).display_name()
                        )));
                    }
                    parse_target_attributes(&element, reader.decoder(), target_for(owner_kind))?;
                    let owner = stack
                        .last_mut()
                        .ok_or_else(|| invalid("data-type-icon target lost its ext owner"))?;
                    if owner.kind != NodeKind::Ext || owner.target_span.is_some() {
                        return Err(invalid(format!(
                            "duplicate {} payload",
                            target_for(owner_kind).display_name()
                        )));
                    }
                    owner.target_span = Some(Span { start, end });
                }
            },
            Event::End(_) => {
                let node = stack
                    .pop()
                    .ok_or_else(|| invalid("data-type-icon source has unmatched end"))?;
                if node.kind == NodeKind::Target {
                    let parent = stack
                        .last_mut()
                        .ok_or_else(|| invalid("data-type-icon target lost its ext owner"))?;
                    if parent.kind != NodeKind::Ext {
                        return Err(invalid("data-type-icon target has an invalid owner"));
                    }
                    parent.target_span = Some(Span {
                        start: node.start,
                        end,
                    });
                } else if node.kind == NodeKind::Ext {
                    owners.push(OwnerSpan {
                        span: Span {
                            start: node.start,
                            end,
                        },
                        tag_end: node.tag_end,
                        close_start: Some(start),
                        owner_index: node.owner_index,
                        target_span: node.target_span,
                    });
                }
            },
            Event::Eof => break,
            Event::Text(_) => {
                if stack
                    .last()
                    .is_some_and(|node| node.kind == NodeKind::Target)
                {
                    return Err(invalid(format!(
                        "{} must not contain character content",
                        target_for(owner_kind).display_name()
                    )));
                }
            },
            Event::CData(_) => {
                if stack
                    .last()
                    .is_some_and(|node| node.kind == NodeKind::Target)
                {
                    return Err(invalid(format!(
                        "{} must not contain character content",
                        target_for(owner_kind).display_name()
                    )));
                }
            },
            Event::Comment(_) | Event::Decl(_) | Event::PI(_) | Event::DocType(_) => {},
            Event::GeneralRef(_)
                if stack
                    .last()
                    .is_some_and(|node| node.kind == NodeKind::Target) =>
            {
                return Err(invalid(format!(
                    "{} must not contain entity references or text",
                    target_for(owner_kind).display_name()
                )));
            },
            Event::GeneralRef(_) => {},
        }
    }
    Ok(owners)
}

fn classify(
    owner_kind: OwnerKind,
    parent: Option<Node>,
    namespace: NamespaceKind,
    local: &[u8],
) -> NodeKind {
    let core = namespace.is_core();
    let named = namespace.is_named();
    if namespace.is_target() && parent.is_some_and(|node| node.kind == NodeKind::Ext) {
        return NodeKind::Target;
    }
    match owner_kind {
        OwnerKind::WorksheetView => match (parent.map(|node| node.kind), local) {
            (None, b"worksheet") if core => NodeKind::Worksheet,
            (Some(NodeKind::Worksheet), b"sheetViews") if core => NodeKind::SheetViews,
            (Some(NodeKind::SheetViews), b"sheetView") if core => NodeKind::SheetView,
            (Some(NodeKind::SheetView), b"extLst") if core => NodeKind::ExtList,
            (Some(NodeKind::ExtList), b"ext") if core => NodeKind::Ext,
            _ => NodeKind::Other,
        },
        OwnerKind::CustomSheetView => match (parent.map(|node| node.kind), local) {
            (None, b"namedSheetViews") if named => NodeKind::NamedViews,
            (Some(NodeKind::NamedViews), b"namedSheetView") if named => NodeKind::NamedView,
            (Some(NodeKind::NamedView), b"extLst") if named => NodeKind::ExtList,
            (Some(NodeKind::ExtList), b"ext") if core => NodeKind::Ext,
            _ => NodeKind::Other,
        },
    }
}

const fn target_for(owner_kind: OwnerKind) -> Target {
    match owner_kind {
        OwnerKind::WorksheetView => Target::Worksheet,
        OwnerKind::CustomSheetView => Target::CustomSheetView,
    }
}
