//! Lossless direct-slide transition splicing for deferred presentations.

use std::ops::Range;

use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, QName, ResolveResult};
use quick_xml::reader::NsReader;

use crate::transition::{Kind, Transition};
use crate::{Error, Result};

const PML: &[u8] = b"http://schemas.openxmlformats.org/presentationml/2006/main";
const STRICT_PML: &[u8] = b"http://purl.oclc.org/ooxml/presentationml/main";
const MAX_DEPTH: usize = 128;
const MAX_NODES: usize = 1_000_000;

#[derive(Debug)]
struct Layout {
    transition: Option<Range<usize>>,
    insertion: usize,
    prefix: Box<str>,
    namespace: Box<str>,
}

pub(super) fn read_direct(xml: &[u8]) -> Result<Option<Transition>> {
    let layout = locate(xml, "read_transition")?;
    let Some(range) = layout.transition else {
        return Ok(None);
    };
    validate_direct_subtree(&xml[range], &layout.prefix, &layout.namespace)?;
    crate::transition::read(xml)
}

pub(super) fn stage(
    xml: &[u8],
    presentation_xml: &[u8],
    target: Option<&Transition>,
    max_output_bytes: usize,
    operation: &'static str,
) -> Result<(Option<Vec<u8>>, bool)> {
    let layout = locate(xml, operation)?;
    let current = if let Some(range) = layout.transition.as_ref() {
        let bytes = xml
            .get(range.clone())
            .ok_or_else(|| invalid("direct transition range is outside its slide"))?;
        validate_direct_subtree(bytes, &layout.prefix, &layout.namespace)?;
        crate::transition::read(xml)?
            .ok_or_else(|| invalid("direct transition disappeared during semantic readback"))?
            .into()
    } else {
        None
    };

    if current.as_ref() == target {
        return Ok((None, false));
    }
    if let Some(target) = target {
        validate_target(target, operation)?;
    }
    if same_semantics(current.as_ref(), target) {
        return Ok((None, false));
    }
    if presentation_is_protected(presentation_xml)? {
        return Err(Error::UnsafeEdit {
            operation,
            reason: "source-backed transition edits refuse modification-protected presentations",
        });
    }

    let replacement = target
        .map(|value| direct_fragment(value, &layout.prefix, operation))
        .transpose()?;
    let replaced_len = layout.transition.as_ref().map_or(0, Range::len);
    let replacement_len = replacement.as_ref().map_or(0, Vec::len);
    let output_len = xml
        .len()
        .checked_sub(replaced_len)
        .and_then(|len| len.checked_add(replacement_len))
        .ok_or_else(|| invalid("direct transition output size overflow"))?;
    if output_len > max_output_bytes {
        return Err(Error::Limit {
            resource: "source-backed slide transition output bytes",
            limit: max_output_bytes,
        });
    }

    let at = layout
        .transition
        .as_ref()
        .map_or(layout.insertion, |range| range.start);
    let end = layout.transition.as_ref().map_or(at, |range| range.end);
    let mut output = Vec::new();
    output
        .try_reserve_exact(output_len)
        .map_err(|source| Error::Allocation {
            resource: "source-backed slide transition XML",
            source,
        })?;
    output.extend_from_slice(&xml[..at]);
    if let Some(replacement) = replacement {
        output.extend_from_slice(&replacement);
    }
    output.extend_from_slice(&xml[end..]);
    if output.len() != output_len {
        return Err(invalid(
            "direct transition output length changed during emission",
        ));
    }

    let published_layout = locate(&output, operation)?;
    let published = if let Some(range) = published_layout.transition {
        validate_direct_subtree(
            &output[range],
            &published_layout.prefix,
            &published_layout.namespace,
        )?;
        crate::transition::read(&output)?
            .ok_or_else(|| invalid("staged direct transition disappeared during readback"))?
            .into()
    } else {
        None
    };
    if !same_semantics(published.as_ref(), target) {
        return Err(invalid(
            "staged direct transition did not round-trip semantically",
        ));
    }
    let scene = crate::shape::Scene::read(&output)?;
    if scene.is_rewritten() {
        return Err(Error::UnsafeEdit {
            operation,
            reason: "source-backed transition edits do not support markup-compatibility branch selection",
        });
    }
    Ok((Some(output), true))
}

fn same_semantics(current: Option<&Transition>, target: Option<&Transition>) -> bool {
    match (current, target) {
        (None, None) => true,
        (Some(current), Some(target)) => current.same_semantics(target),
        _ => false,
    }
}

fn direct_fragment(value: &Transition, prefix: &str, operation: &'static str) -> Result<Vec<u8>> {
    validate_target(value, operation)?;
    let canonical = crate::transition::write(&crate::transition::semantic_clone(value))?;
    if canonical.contains("mc:AlternateContent") || canonical.contains("p14:") {
        return Err(Error::UnsafeEdit {
            operation,
            reason: "source-backed transition edits do not publish markup-compatibility or extension transitions",
        });
    }
    let lexical = if prefix == "p" {
        canonical
    } else if prefix.is_empty() {
        canonical.replace("p:", "")
    } else {
        canonical.replace("p:", &format!("{prefix}:"))
    };
    Ok(lexical.into_bytes())
}

fn validate_target(value: &Transition, operation: &'static str) -> Result<()> {
    if value.duration().is_some()
        || value.preserved_len() != 0
        || matches!(value.kind(), Kind::Ripple(_) | Kind::Raw(_))
    {
        return Err(Error::UnsafeEdit {
            operation,
            reason: "source-backed transition edits support only standard direct transition values without retained extensions",
        });
    }
    if let Some(effect) = crate::transition::preserved_effect_xml(value) {
        validate_effect_fragment(effect.as_bytes())?;
    }
    Ok(())
}

fn locate(xml: &[u8], operation: &'static str) -> Result<Layout> {
    let mut reader = NsReader::from_reader(xml);
    let mut depth = 0usize;
    let mut nodes = 0usize;
    let mut root_seen = false;
    let mut root_prefix = None;
    let mut root_namespace = None;
    let mut root_close = None;
    let mut transition = None;
    let mut open_transition = None;
    let mut previous_rank = 0u8;
    let mut insertion = None;
    let mut child_counts = [0u8; 5];

    loop {
        let start = position(&reader)?;
        let event = reader.read_event()?.into_owned();
        let end = position(&reader)?;
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        match event {
            Event::Start(element) => {
                nodes = add_node(nodes)?;
                depth = enter_depth(depth)?;
                if depth == 1 {
                    if root_seen || !is_pml(&namespace, element.name(), b"sld") {
                        return Err(invalid("slide XML must contain exactly one p:sld root"));
                    }
                    root_seen = true;
                    root_prefix = Some(qname_prefix(element.name())?);
                    root_namespace = Some(pml_namespace(&namespace)?);
                } else if depth == 2 {
                    let rank = child_rank(&namespace, element.name()).ok_or(Error::UnsafeEdit {
                        operation,
                        reason: "source-backed transition edits refuse unknown direct slide children",
                    })?;
                    record_child(
                        rank,
                        start,
                        &mut previous_rank,
                        &mut insertion,
                        &mut child_counts,
                    )?;
                    if rank == 3 {
                        if transition.is_some() || open_transition.is_some() {
                            return Err(invalid("slide contains duplicate direct transitions"));
                        }
                        open_transition = Some(start);
                    }
                } else if is_pml(&namespace, element.name(), b"transition") {
                    return Err(Error::UnsafeEdit {
                        operation,
                        reason: "source-backed transition edits refuse nested or markup-compatibility transitions",
                    });
                }
            },
            Event::Empty(element) => {
                nodes = add_node(nodes)?;
                let event_depth = enter_depth(depth)?;
                if event_depth == 1 {
                    return Err(invalid("slide root cannot be empty"));
                }
                if event_depth == 2 {
                    let rank = child_rank(&namespace, element.name()).ok_or(Error::UnsafeEdit {
                        operation,
                        reason: "source-backed transition edits refuse unknown direct slide children",
                    })?;
                    record_child(
                        rank,
                        start,
                        &mut previous_rank,
                        &mut insertion,
                        &mut child_counts,
                    )?;
                    if rank == 3 {
                        if transition.replace(start..end).is_some() || open_transition.is_some() {
                            return Err(invalid("slide contains duplicate direct transitions"));
                        }
                    }
                } else if is_pml(&namespace, element.name(), b"transition") {
                    return Err(Error::UnsafeEdit {
                        operation,
                        reason: "source-backed transition edits refuse nested or markup-compatibility transitions",
                    });
                }
            },
            Event::End(element) => {
                if depth == 0 {
                    return Err(invalid("slide XML contains an unmatched end element"));
                }
                if depth == 2 && is_pml(&namespace, element.name(), b"transition") {
                    let open = open_transition.take().ok_or_else(|| {
                        invalid("direct transition close has no matching open element")
                    })?;
                    transition = Some(open..end);
                }
                if depth == 1 {
                    if !is_pml(&namespace, element.name(), b"sld") {
                        return Err(invalid("slide root close does not match p:sld"));
                    }
                    root_close = Some(start);
                }
                depth -= 1;
            },
            Event::DocType(_) => {
                return Err(invalid("DOCTYPE is forbidden in source-backed slide XML"));
            },
            Event::Text(text) => {
                if depth == 0 && !text.as_ref().iter().all(u8::is_ascii_whitespace) {
                    return Err(invalid("slide XML has text outside its root"));
                }
            },
            Event::CData(_) | Event::GeneralRef(_) if depth == 0 => {
                return Err(invalid("slide XML has data outside its root"));
            },
            Event::Eof => break,
            _ => {},
        }
    }
    if depth != 0 || !root_seen || open_transition.is_some() || child_counts[0] != 1 {
        return Err(invalid(
            "slide XML has an incomplete direct-child structure",
        ));
    }
    let root_close = root_close.ok_or_else(|| invalid("slide root is not closed"))?;
    Ok(Layout {
        transition,
        insertion: insertion.unwrap_or(root_close),
        prefix: root_prefix
            .ok_or_else(|| invalid("slide root prefix is unavailable"))?
            .into_boxed_str(),
        namespace: root_namespace
            .ok_or_else(|| invalid("slide root namespace is unavailable"))?
            .into_boxed_str(),
    })
}

fn child_rank(namespace: &ResolveResult<'_>, name: QName<'_>) -> Option<u8> {
    if !is_pml_namespace(namespace) {
        return None;
    }
    match name.local_name().as_ref() {
        b"cSld" => Some(1),
        b"clrMapOvr" => Some(2),
        b"transition" => Some(3),
        b"timing" => Some(4),
        b"extLst" => Some(5),
        _ => None,
    }
}

fn record_child(
    rank: u8,
    start: usize,
    previous_rank: &mut u8,
    insertion: &mut Option<usize>,
    counts: &mut [u8; 5],
) -> Result<()> {
    if rank < *previous_rank {
        return Err(invalid("slide direct children are outside schema order"));
    }
    let slot = usize::from(rank - 1);
    counts[slot] = counts[slot]
        .checked_add(1)
        .ok_or_else(|| invalid("slide direct-child count overflow"))?;
    if counts[slot] > 1 {
        return Err(invalid("slide contains a duplicate direct child"));
    }
    if rank > 3 && insertion.is_none() {
        *insertion = Some(start);
    }
    *previous_rank = rank;
    Ok(())
}

fn validate_direct_subtree(xml: &[u8], prefix: &str, expected_namespace: &str) -> Result<()> {
    let mut reader = NsReader::from_reader(xml);
    let mut depth = 0usize;
    let mut nodes = 0usize;
    let mut effect_seen = false;
    loop {
        let event = reader.read_event()?.into_owned();
        let resolver = reader.resolver().clone();
        let (resolved, event) = resolver.resolve_event(event);
        match event {
            Event::Start(element) => {
                nodes = add_node(nodes)?;
                depth = enter_depth(depth)?;
                validate_transition_element(
                    &resolved,
                    &element,
                    depth,
                    &mut effect_seen,
                    prefix,
                    expected_namespace,
                )?;
            },
            Event::Empty(element) => {
                nodes = add_node(nodes)?;
                let event_depth = enter_depth(depth)?;
                validate_transition_element(
                    &resolved,
                    &element,
                    event_depth,
                    &mut effect_seen,
                    prefix,
                    expected_namespace,
                )?;
            },
            Event::End(_) => {
                if depth == 0 {
                    return Err(invalid(
                        "transition subtree contains an unmatched end element",
                    ));
                }
                depth -= 1;
            },
            Event::Text(text) if !text.as_ref().iter().all(u8::is_ascii_whitespace) => {
                return Err(unsupported_transition());
            },
            Event::Comment(_) | Event::CData(_) | Event::PI(_) | Event::GeneralRef(_) => {
                return Err(unsupported_transition());
            },
            Event::DocType(_) => return Err(invalid("DOCTYPE is forbidden in transition XML")),
            Event::Eof => break,
            _ => {},
        }
    }
    if depth != 0 {
        return Err(invalid("transition subtree is not closed"));
    }
    Ok(())
}

fn validate_effect_fragment(xml: &[u8]) -> Result<()> {
    let mut reader = NsReader::from_reader(xml);
    let mut depth = 0usize;
    let mut nodes = 0usize;
    let mut root_seen = false;
    let mut effect_seen = false;
    loop {
        let event = reader.read_event()?.into_owned();
        let resolver = reader.resolver().clone();
        let (resolved, event) = resolver.resolve_event(event);
        match event {
            Event::Start(element) => {
                nodes = add_node(nodes)?;
                depth = enter_depth(depth)?;
                if depth != 1 || root_seen {
                    return Err(unsupported_transition());
                }
                root_seen = true;
                validate_detached_effect(&resolved, &element, &mut effect_seen)?;
            },
            Event::Empty(element) => {
                nodes = add_node(nodes)?;
                if enter_depth(depth)? != 1 || root_seen {
                    return Err(unsupported_transition());
                }
                root_seen = true;
                validate_detached_effect(&resolved, &element, &mut effect_seen)?;
            },
            Event::End(_) => {
                if depth == 0 {
                    return Err(invalid("detached transition effect has an unmatched close"));
                }
                depth -= 1;
            },
            Event::Text(text) if !text.as_ref().iter().all(u8::is_ascii_whitespace) => {
                return Err(unsupported_transition());
            },
            Event::Comment(_)
            | Event::CData(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::GeneralRef(_) => return Err(unsupported_transition()),
            Event::DocType(_) => return Err(invalid("DOCTYPE is forbidden in transition XML")),
            Event::Eof => break,
            _ => {},
        }
    }
    if depth != 0 || !root_seen || !effect_seen {
        return Err(invalid("detached transition effect is incomplete"));
    }
    Ok(())
}

fn validate_detached_effect(
    resolved: &ResolveResult<'_>,
    element: &BytesStart<'_>,
    effect_seen: &mut bool,
) -> Result<()> {
    let prefix = qname_prefix(element.name())?;
    let namespace = match resolved {
        ResolveResult::Bound(Namespace(value)) if *value == PML || *value == STRICT_PML => {
            std::str::from_utf8(value).map_err(|error| {
                Error::Xml(format!("transition effect namespace is not UTF-8: {error}"))
            })?
        },
        ResolveResult::Unknown(value)
            if !prefix.is_empty() && value.as_slice() == prefix.as_bytes() =>
        {
            std::str::from_utf8(PML).map_err(|error| {
                Error::Xml(format!("PresentationML namespace is not UTF-8: {error}"))
            })?
        },
        ResolveResult::Unbound if prefix.is_empty() && element.name().prefix().is_none() => {
            std::str::from_utf8(PML).map_err(|error| {
                Error::Xml(format!("PresentationML namespace is not UTF-8: {error}"))
            })?
        },
        _ => return Err(unsupported_transition()),
    };
    validate_transition_element(resolved, element, 2, effect_seen, &prefix, namespace)
}

fn validate_transition_element(
    resolved: &ResolveResult<'_>,
    element: &BytesStart<'_>,
    depth: usize,
    effect_seen: &mut bool,
    prefix: &str,
    namespace: &str,
) -> Result<()> {
    if !is_expected_subtree_namespace(resolved, element.name(), prefix, namespace) {
        return Err(unsupported_transition());
    }
    let local = element.local_name();
    let allowed = if depth == 1 && local.as_ref() == b"transition" {
        &[
            b"spd".as_slice(),
            b"advClick".as_slice(),
            b"advTm".as_slice(),
        ][..]
    } else if depth == 2 && is_standard_effect(local.as_ref()) {
        if *effect_seen {
            return Err(invalid("transition contains more than one visual effect"));
        }
        *effect_seen = true;
        effect_attributes(local.as_ref())
    } else {
        return Err(unsupported_transition());
    };
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        if attribute.key.prefix().is_some()
            || !allowed
                .iter()
                .any(|name| *name == attribute.key.local_name().as_ref())
        {
            return Err(unsupported_transition());
        }
    }
    Ok(())
}

fn is_expected_subtree_namespace(
    resolved: &ResolveResult<'_>,
    name: QName<'_>,
    prefix: &str,
    namespace: &str,
) -> bool {
    match resolved {
        ResolveResult::Bound(Namespace(value)) => *value == namespace.as_bytes(),
        ResolveResult::Unknown(value) => {
            !prefix.is_empty() && value.as_slice() == prefix.as_bytes()
        },
        ResolveResult::Unbound => prefix.is_empty() && name.prefix().is_none(),
    }
}

fn is_standard_effect(local: &[u8]) -> bool {
    matches!(
        local,
        b"cut"
            | b"fade"
            | b"push"
            | b"wipe"
            | b"split"
            | b"pull"
            | b"cover"
            | b"dissolve"
            | b"blinds"
            | b"checker"
            | b"randomBar"
            | b"circle"
            | b"diamond"
            | b"plus"
            | b"wedge"
            | b"zoom"
            | b"random"
            | b"wheel"
            | b"newsflash"
            | b"strips"
            | b"comb"
    )
}

fn effect_attributes(local: &[u8]) -> &'static [&'static [u8]] {
    match local {
        b"cut" | b"fade" => &[b"thruBlk"],
        b"push" | b"wipe" | b"pull" | b"cover" | b"blinds" | b"checker" | b"randomBar"
        | b"zoom" | b"strips" | b"comb" => &[b"dir"],
        b"split" => &[b"orient", b"dir"],
        b"wheel" => &[b"spokes"],
        _ => &[],
    }
}

fn presentation_is_protected(xml: &[u8]) -> Result<bool> {
    let text = std::str::from_utf8(xml)
        .map_err(|error| Error::Xml(format!("presentation XML is not UTF-8: {error}")))?;
    Ok(
        crate::presentation_properties::metadata::protection::Settings::parse_xml(text)?
            .is_protected(),
    )
}

fn unsupported_transition() -> Error {
    Error::UnsafeEdit {
        operation: "edit_transition",
        reason: "source-backed transition edits refuse extension, sound-action, or unknown transition markup",
    }
}

fn is_pml(namespace: &ResolveResult<'_>, name: QName<'_>, local: &[u8]) -> bool {
    name.local_name().as_ref() == local && is_pml_namespace(namespace)
}

fn is_pml_namespace(namespace: &ResolveResult<'_>) -> bool {
    matches!(namespace, ResolveResult::Bound(Namespace(value)) if *value == PML || *value == STRICT_PML)
}

fn pml_namespace(namespace: &ResolveResult<'_>) -> Result<String> {
    match namespace {
        ResolveResult::Bound(Namespace(value)) if *value == PML || *value == STRICT_PML => {
            std::str::from_utf8(value)
                .map(str::to_owned)
                .map_err(|error| Error::Xml(format!("slide namespace is not UTF-8: {error}")))
        },
        _ => Err(invalid("slide root has no PresentationML namespace")),
    }
}

fn qname_prefix(name: QName<'_>) -> Result<String> {
    let raw = name.as_ref();
    let prefix = raw
        .iter()
        .position(|byte| *byte == b':')
        .map_or(&[][..], |colon| &raw[..colon]);
    std::str::from_utf8(prefix)
        .map(str::to_owned)
        .map_err(|error| Error::Xml(format!("slide root prefix is not UTF-8: {error}")))
}

fn add_node(nodes: usize) -> Result<usize> {
    let nodes = nodes.checked_add(1).ok_or(Error::Limit {
        resource: "source-backed transition XML nodes",
        limit: MAX_NODES,
    })?;
    if nodes > MAX_NODES {
        Err(Error::Limit {
            resource: "source-backed transition XML nodes",
            limit: MAX_NODES,
        })
    } else {
        Ok(nodes)
    }
}

fn enter_depth(depth: usize) -> Result<usize> {
    let depth = depth.checked_add(1).ok_or(Error::Limit {
        resource: "source-backed transition XML depth",
        limit: MAX_DEPTH,
    })?;
    if depth > MAX_DEPTH {
        Err(Error::Limit {
            resource: "source-backed transition XML depth",
            limit: MAX_DEPTH,
        })
    } else {
        Ok(depth)
    }
}

fn position(reader: &NsReader<&[u8]>) -> Result<usize> {
    usize::try_from(reader.buffer_position())
        .map_err(|_error| invalid("source-backed transition XML position exceeds usize"))
}

fn invalid(message: &str) -> Error {
    Error::Invalid(format!("source-backed slide transition: {message}"))
}
