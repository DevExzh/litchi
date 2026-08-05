//! Bounded namespace-aware XML codec for auto-mark-file references.

use super::model::AlphabeticalIndexAutoMarkFile;
use super::{
    MAX_AGGREGATE_BYTES, MAX_DEPTH, MAX_OCCURRENCES, MAX_VALUE_BYTES, OFFICE, TEXT, XLINK, invalid,
    make_error,
};
use crate::variable_declaration::{Body, Part, Scope};
use litchi_core::Result;
use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use std::collections::{HashMap, HashSet};

#[derive(Clone)]
struct Frame {
    namespace: Option<String>,
    local: String,
}

struct TextScope {
    depth: usize,
    /// Highest preface rank seen so far; document content starts at
    /// `RANK_OTHER_CONTENT`.
    last_rank: u8,
    seen_auto_mark_file: bool,
}

struct PendingElement {
    depth: usize,
}

/// Preface rank of an `office:text` child in normative ODF 1.3 order.
const RANK_OTHER_CONTENT: u8 = 4;

fn child_rank(namespace: Option<&str>, local: &str) -> u8 {
    if namespace != Some(TEXT) {
        return RANK_OTHER_CONTENT;
    }
    match local {
        "forms" => 0,
        "tracked-changes" | "change-tracking-config" => 1,
        "dde-connection-decls" => 2,
        "alphabetical-index-auto-mark-file" => 3,
        _ => RANK_OTHER_CONTENT,
    }
}

type Attributes = HashMap<(String, String), String>;

pub(super) fn parse_part(
    xml: &str,
    part: Part,
    references: &mut Vec<AlphabeticalIndexAutoMarkFile>,
    scopes: &mut HashSet<(Part, Scope)>,
    aggregate: &mut usize,
) -> Result<()> {
    let mut reader = NsReader::from_str(xml);
    reader.config_mut().check_end_names = true;
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut stack = Vec::<Frame>::new();
    let mut scope: Option<TextScope> = None;
    let mut pending: Option<PendingElement> = None;

    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| make_error(format!("invalid auto-mark-file XML: {error}")))?;
        match event {
            Event::Start(ref element) => {
                if pending.is_some() {
                    return invalid(
                        "text:alphabetical-index-auto-mark-file cannot contain elements",
                    );
                }
                let namespace = namespace_uri(&namespace)?;
                let local = decode(element.local_name().as_ref(), "element name")?;
                reject_spoofed_name(namespace.as_deref(), &local)?;
                if is_auto_mark_file(namespace.as_deref(), &local) {
                    let scope = scope.as_mut().ok_or_else(|| {
                        make_error(
                            "text:alphabetical-index-auto-mark-file occurs outside office:text",
                        )
                    })?;
                    if scope.depth != depth {
                        return invalid(
                            "text:alphabetical-index-auto-mark-file must be a direct office:text child",
                        );
                    }
                    register_reference(
                        &reader, element, part, scope, references, scopes, aggregate,
                    )?;
                    pending = Some(PendingElement { depth: depth + 1 });
                }
                if namespace.as_deref() == Some(OFFICE)
                    && local == "text"
                    && stack.last().is_some_and(|frame| {
                        frame.namespace.as_deref() == Some(OFFICE) && frame.local == "body"
                    })
                {
                    if scope.is_some() {
                        return invalid("nested office:text element");
                    }
                    scope = Some(TextScope {
                        depth: depth + 1,
                        last_rank: 0,
                        seen_auto_mark_file: false,
                    });
                } else if let Some(active) = scope.as_mut()
                    && active.depth == depth
                {
                    let rank = child_rank(namespace.as_deref(), &local);
                    if local != "alphabetical-index-auto-mark-file" {
                        active.last_rank = active.last_rank.max(rank);
                    }
                }
                stack.push(Frame { namespace, local });
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| make_error("auto-mark-file XML depth overflow"))?;
                if depth > MAX_DEPTH {
                    return invalid(format!(
                        "auto-mark-file XML nesting exceeds {MAX_DEPTH} levels"
                    ));
                }
            },
            Event::Empty(ref element) => {
                if pending.is_some() {
                    return invalid(
                        "text:alphabetical-index-auto-mark-file cannot contain elements",
                    );
                }
                let namespace = namespace_uri(&namespace)?;
                let local = decode(element.local_name().as_ref(), "element name")?;
                reject_spoofed_name(namespace.as_deref(), &local)?;
                if is_auto_mark_file(namespace.as_deref(), &local) {
                    let scope = scope.as_mut().ok_or_else(|| {
                        make_error(
                            "text:alphabetical-index-auto-mark-file occurs outside office:text",
                        )
                    })?;
                    if scope.depth != depth {
                        return invalid(
                            "text:alphabetical-index-auto-mark-file must be a direct office:text child",
                        );
                    }
                    register_reference(
                        &reader, element, part, scope, references, scopes, aggregate,
                    )?;
                }
                if let Some(active) = scope.as_mut()
                    && active.depth == depth
                {
                    let rank = child_rank(namespace.as_deref(), &local);
                    if !is_auto_mark_file(namespace.as_deref(), &local) {
                        active.last_rank = active.last_rank.max(rank);
                    }
                }
            },
            Event::End(_) => {
                if pending
                    .as_ref()
                    .is_some_and(|pending| pending.depth == depth)
                {
                    pending = None;
                }
                if scope.as_ref().is_some_and(|active| active.depth == depth) {
                    scope = None;
                }
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| make_error("auto-mark-file XML stack underflow"))?;
                stack
                    .pop()
                    .ok_or_else(|| make_error("auto-mark-file XML frame stack underflow"))?;
            },
            Event::Text(ref text) => {
                let value = text
                    .decode()
                    .map_err(|error| make_error(format!("invalid auto-mark-file text: {error}")))?;
                if pending.is_some() && !value.is_empty() {
                    return invalid("text:alphabetical-index-auto-mark-file must be empty");
                }
            },
            Event::CData(_) | Event::GeneralRef(_) if pending.is_some() => {
                return invalid("text:alphabetical-index-auto-mark-file must be empty");
            },
            Event::DocType(_) | Event::PI(_) => {
                return invalid("DTDs and processing instructions are not allowed");
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    if depth != 0 || !stack.is_empty() || scope.is_some() || pending.is_some() {
        return invalid("incomplete auto-mark-file XML structure");
    }
    Ok(())
}

fn is_auto_mark_file(namespace: Option<&str>, local: &str) -> bool {
    namespace == Some(TEXT) && local == "alphabetical-index-auto-mark-file"
}

fn register_reference(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    part: Part,
    scope: &mut TextScope,
    references: &mut Vec<AlphabeticalIndexAutoMarkFile>,
    scopes: &mut HashSet<(Part, Scope)>,
    aggregate: &mut usize,
) -> Result<()> {
    if scope.seen_auto_mark_file {
        return invalid("duplicate text:alphabetical-index-auto-mark-file in one office:text");
    }
    if scope.last_rank >= RANK_OTHER_CONTENT {
        return invalid("text:alphabetical-index-auto-mark-file must precede document content");
    }
    scope.seen_auto_mark_file = true;

    let attributes = collect_attributes(reader, element, aggregate)?;
    reject_unexpected(&attributes)?;
    let href = required_nonempty(&attributes, XLINK, "href")?;
    match get(&attributes, XLINK, "type") {
        Some("simple") => {},
        _ => {
            return invalid(
                "text:alphabetical-index-auto-mark-file requires xlink:type=\"simple\"",
            );
        },
    }

    let scope_value = Scope::Body(Body::Text);
    if !scopes.insert((part, scope_value.clone())) {
        return invalid("duplicate text:alphabetical-index-auto-mark-file in one document part");
    }
    if references.len() >= MAX_OCCURRENCES {
        return invalid(format!(
            "document exceeds {MAX_OCCURRENCES} auto-mark-file references"
        ));
    }
    references.push(AlphabeticalIndexAutoMarkFile {
        part,
        scope: scope_value,
        href,
    });
    Ok(())
}

fn collect_attributes(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    aggregate: &mut usize,
) -> Result<Attributes> {
    let mut attributes = HashMap::new();
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute
            .map_err(|error| make_error(format!("invalid auto-mark-file attribute: {error}")))?;
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        let namespace = namespace_uri(&namespace)?.unwrap_or_default();
        let local = decode(local.as_ref(), "attribute name")?;
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, reader.decoder())
            .map_err(|error| make_error(format!("invalid auto-mark-file attribute: {error}")))?
            .into_owned();
        if value.len() > MAX_VALUE_BYTES {
            return invalid("auto-mark-file attribute exceeds 64 KiB");
        }
        *aggregate = aggregate
            .checked_add(value.len())
            .ok_or_else(|| make_error("auto-mark-file attribute size overflow"))?;
        if *aggregate > MAX_AGGREGATE_BYTES {
            return invalid("auto-mark-file metadata exceeds 16 MiB");
        }
        if attributes.insert((namespace, local), value).is_some() {
            return invalid("duplicate expanded auto-mark-file attribute");
        }
    }
    Ok(attributes)
}

fn reject_unexpected(attributes: &Attributes) -> Result<()> {
    let allowed = [(XLINK, "href"), (XLINK, "type")];
    for (namespace, local) in attributes.keys() {
        if !allowed.iter().any(|(allowed_namespace, allowed_local)| {
            namespace == allowed_namespace && local == allowed_local
        }) && matches!(namespace.as_str(), OFFICE | TEXT | XLINK)
        {
            return invalid(format!(
                "unexpected auto-mark-file attribute {namespace}:{local}"
            ));
        }
    }
    Ok(())
}

fn reject_spoofed_name(namespace: Option<&str>, local: &str) -> Result<()> {
    if local == "alphabetical-index-auto-mark-file" && namespace != Some(TEXT) {
        return invalid("alphabetical-index-auto-mark-file uses the wrong namespace");
    }
    Ok(())
}

fn get<'a>(attributes: &'a Attributes, namespace: &str, local: &str) -> Option<&'a str> {
    attributes
        .get(&(namespace.to_string(), local.to_string()))
        .map(String::as_str)
}

fn required_nonempty(attributes: &Attributes, namespace: &str, local: &str) -> Result<String> {
    match get(attributes, namespace, local) {
        Some(value) if !value.is_empty() => Ok(value.to_string()),
        _ => invalid(format!("auto-mark-file requires non-empty xlink:{local}")),
    }
}

fn namespace_uri(result: &ResolveResult<'_>) -> Result<Option<String>> {
    match result {
        ResolveResult::Bound(Namespace(value)) => Ok(Some(decode(value, "namespace URI")?)),
        ResolveResult::Unbound => Ok(None),
        ResolveResult::Unknown(prefix) => Err(make_error(format!(
            "unbound namespace prefix '{}'",
            String::from_utf8_lossy(prefix)
        ))),
    }
}

fn decode(value: &[u8], description: &str) -> Result<String> {
    std::str::from_utf8(value)
        .map(str::to_string)
        .map_err(|_| make_error(format!("invalid UTF-8 {description}")))
}
