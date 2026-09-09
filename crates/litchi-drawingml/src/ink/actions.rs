//! Source-backed PowerPoint 2014 ink-action metadata.

use std::{fmt, io::Write, sync::Arc};

use litchi_ooxml_common::xml_name::is_qualified_name;
use quick_xml::{
    XmlVersion,
    events::{BytesDecl, BytesRef, BytesStart, Event},
    name::{Namespace, QName, ResolveResult},
    reader::NsReader,
};

use crate::{Error, Result};

use super::{
    ACTION_NAMESPACE, MAX_ATTRIBUTE_VALUE_BYTES, MAX_DEPTH, MAX_NODES, MAX_SOURCE_BYTES,
    MAX_TOKEN_BYTES,
};

/// Maximum action records retained by one action part.
pub const MAX_ACTIONS: usize = 65_536;
/// Maximum action-group records retained by one action part.
pub const MAX_ACTION_GROUPS: usize = 16_384;

const MAX_NAMESPACE_DECLARATIONS: usize = 256;
const MAX_ATTRIBUTES_PER_ELEMENT: usize = 256;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum NamespaceId {
    Action,
    Other,
    Unbound,
    Unknown,
}

/// Reserved or future PowerPoint ink action type.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum ActionType {
    /// Add ink data.
    Add,
    /// Remove ink data.
    Remove,
    /// Transform ink data.
    Transform,
    /// A future or user-defined action type.
    Custom(Box<str>),
}

impl ActionType {
    fn parse(value: &str) -> Result<Self> {
        if value.is_empty() || value.len() > 256 || value.bytes().any(|byte| byte == 0) {
            return Err(invalid("ink action type is empty or overlong"));
        }
        Ok(match value {
            "add" => Self::Add,
            "remove" => Self::Remove,
            "transform" => Self::Transform,
            _ => Self::Custom(value.into()),
        })
    }
    /// Return the source lexical action type.
    #[must_use]
    pub fn as_str(&self) -> &str {
        match self {
            Self::Add => "add",
            Self::Remove => "remove",
            Self::Transform => "transform",
            Self::Custom(value) => value,
        }
    }
}

/// Typed metadata for one CT_Action element.
#[derive(Debug, Clone, PartialEq, Eq)]
#[must_use]
pub struct Action {
    action_type: ActionType,
    start_time: Box<str>,
    source_start: usize,
    source_end: usize,
}

impl Action {
    /// The action type.
    #[must_use]
    pub const fn action_type(&self) -> &ActionType {
        &self.action_type
    }
    /// The exact decimal start-time lexical value.
    #[must_use]
    pub fn start_time(&self) -> &str {
        &self.start_time
    }
    /// Source span of the complete action element.
    #[must_use]
    pub const fn source_span(&self) -> (usize, usize) {
        (self.source_start, self.source_end)
    }
    /// Borrow the exact action element XML.
    #[must_use]
    pub fn xml<'a>(&self, document: &'a Actions) -> &'a [u8] {
        document
            .source
            .get(self.source_start..self.source_end)
            .unwrap_or_default()
    }
}

/// Immutable, source-backed PowerPoint ink-actions part.
#[derive(Debug, Clone)]
#[must_use]
pub struct Actions {
    source: Arc<[u8]>,
    length_unit: Box<str>,
    time_unit: Box<str>,
    action_groups: usize,
    actions: Vec<Action>,
}

impl PartialEq for Actions {
    fn eq(&self, other: &Self) -> bool {
        self.source.as_ref() == other.source.as_ref()
            && self.length_unit == other.length_unit
            && self.time_unit == other.time_unit
            && self.action_groups == other.action_groups
            && self.actions == other.actions
    }
}
impl Eq for Actions {}

impl Actions {
    /// Borrow exact source bytes.
    #[must_use]
    pub fn source(&self) -> &[u8] {
        &self.source
    }
    /// Required InkML length unit.
    #[must_use]
    pub fn length_unit(&self) -> &str {
        &self.length_unit
    }
    /// Required InkML time unit.
    #[must_use]
    pub fn time_unit(&self) -> &str {
        &self.time_unit
    }
    /// Number of action groups.
    #[must_use]
    pub const fn action_group_count(&self) -> usize {
        self.action_groups
    }
    /// Borrow actions in source order.
    #[must_use]
    pub fn actions(&self) -> &[Action] {
        &self.actions
    }
}

/// Read one complete iact:actions element.
///
/// # Errors
///
/// Returns an error for malformed XML, invalid required attributes, or
/// exhausted resource limits.
pub fn read(xml: &[u8]) -> Result<Actions> {
    if xml.len() > MAX_SOURCE_BYTES {
        return Err(limit("ink actions source bytes", MAX_SOURCE_BYTES));
    }
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = true;
    reader.config_mut().check_comments = true;
    reader
        .resolver_mut()
        .set_max_declarations_per_element(MAX_NAMESPACE_DECLARATIONS);
    let mut stack = Vec::new();
    let mut root_seen = false;
    let mut root_closed = false;
    let mut nodes = 0usize;
    let mut groups = 0usize;
    let mut actions = Vec::new();
    let mut length_unit = None;
    let mut time_unit = None;
    let mut declaration_seen = false;
    let mut preamble_content_seen = false;
    let mut legacy_fragment = false;

    loop {
        let start = pos(&reader)?;
        let event = reader.read_event().map_err(xml_error)?;
        let (resolved, event) = reader.resolver().resolve_event(event);
        let end = pos(&reader)?;
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
                length_unit = Some(required_attr(&element, b"lengthUnit", &reader)?);
                time_unit = Some(required_attr(&element, b"timeUnit", &reader)?);
                validate_unit(length_unit.as_deref().unwrap_or_default(), "lengthUnit")?;
                validate_unit(time_unit.as_deref().unwrap_or_default(), "timeUnit")?;
                reserve_one(&mut stack, "ink actions XML stack")?;
                stack.push(Frame {
                    namespace: namespace_id(&resolved)?,
                    action: None,
                    action_group: false,
                    direct_actions: 0,
                });
            },
            Event::Empty(element) if !root_seen => {
                validate_element(&element, &reader)?;
                legacy_fragment = require_root(&element, &resolved)?;
                root_seen = true;
                root_closed = true;
                increment_nodes(&mut nodes)?;
                enforce_depth(1)?;
                length_unit = Some(required_attr(&element, b"lengthUnit", &reader)?);
                time_unit = Some(required_attr(&element, b"timeUnit", &reader)?);
                validate_unit(length_unit.as_deref().unwrap_or_default(), "lengthUnit")?;
                validate_unit(time_unit.as_deref().unwrap_or_default(), "timeUnit")?;
            },
            Event::Start(element) if root_seen && !root_closed => {
                validate_element(&element, &reader)?;
                validate_unknown_prefix(&resolved, &element, legacy_fragment)?;
                increment_nodes(&mut nodes)?;
                enforce_depth(stack.len().saturating_add(1))?;
                let local = element.name().local_name();
                let action_group = is_action_namespace(&resolved, &element, legacy_fragment)
                    && local.as_ref() == b"actionGroup";
                let action_element = is_action_namespace(&resolved, &element, legacy_fragment)
                    && local.as_ref() == b"action";
                let action = if action_group {
                    increment_groups(&mut groups)?;
                    require_action_attrs(&element, &reader)?;
                    None
                } else if action_element {
                    ensure_action_capacity(&actions)?;
                    record_group_action(&mut stack);
                    let index = actions.len();
                    reserve_one(&mut actions, "ink action records")?;
                    actions.push(parse_action(&element, start, end, &reader)?);
                    Some(index)
                } else {
                    None
                };
                reserve_one(&mut stack, "ink actions XML stack")?;
                stack.push(Frame {
                    namespace: namespace_id(&resolved)?,
                    action,
                    action_group,
                    direct_actions: 0,
                });
            },
            Event::Empty(element) if root_seen && !root_closed => {
                validate_element(&element, &reader)?;
                validate_unknown_prefix(&resolved, &element, legacy_fragment)?;
                increment_nodes(&mut nodes)?;
                enforce_depth(stack.len().saturating_add(1))?;
                let local = element.name().local_name();
                if is_action_namespace(&resolved, &element, legacy_fragment)
                    && local.as_ref() == b"actionGroup"
                {
                    increment_groups(&mut groups)?;
                    require_action_attrs(&element, &reader)?;
                    return Err(invalid("ink actionGroup requires an action child"));
                } else if is_action_namespace(&resolved, &element, legacy_fragment)
                    && local.as_ref() == b"action"
                {
                    ensure_action_capacity(&actions)?;
                    record_group_action(&mut stack);
                    reserve_one(&mut actions, "ink action records")?;
                    actions.push(parse_action(&element, start, end, &reader)?);
                }
            },
            Event::End(element) if root_seen && !root_closed => {
                let frame = stack
                    .pop()
                    .ok_or_else(|| invalid("ink actions has an unexpected end"))?;
                if frame.action_group && frame.direct_actions == 0 {
                    return Err(invalid("ink actionGroup requires an action child"));
                }
                validate_end_name(&element, &resolved, legacy_fragment)?;
                if frame.namespace != namespace_id(&resolved)? {
                    return Err(invalid("ink actions has mismatched closing elements"));
                }
                if let Some(index) = frame.action {
                    actions[index].source_end = end;
                }
                if stack.is_empty() {
                    root_closed = true;
                }
            },
            Event::Start(_) | Event::Empty(_) | Event::End(_) if root_closed => {
                return Err(invalid("ink actions has content after its root"));
            },
            Event::Start(_) | Event::Empty(_) | Event::End(_) => {
                return Err(invalid("ink actions has an invalid root transition"));
            },
            Event::Text(text)
                if (!root_seen || root_closed)
                    && !validate_text(&text, "ink actions text")?
                        .bytes()
                        .all(|byte| byte.is_ascii_whitespace()) =>
            {
                return Err(invalid("ink actions has text outside its root"));
            },
            Event::Text(text) => {
                validate_text(&text, "ink actions text")?;
            },
            Event::CData(data) => {
                validate_text(&data, "ink actions CDATA")?;
                if !root_seen || root_closed {
                    return Err(invalid("ink actions CDATA is not allowed outside its root"));
                }
            },
            Event::GeneralRef(reference) => {
                validate_reference(&reference)?;
                if !root_seen || root_closed {
                    return Err(invalid("ink actions has a reference outside its root"));
                }
            },
            Event::Comment(comment) => {
                validate_text(&comment, "ink actions comment")?;
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid(
                    "ink actions rejects DTDs and processing instructions",
                ));
            },
            Event::Decl(_) => {
                return Err(invalid(
                    "ink actions has a duplicate or late XML declaration",
                ));
            },
            Event::Eof => break,
        }
        if !is_declaration {
            preamble_content_seen = true;
        }
    }

    if !root_seen || !root_closed || !stack.is_empty() {
        return Err(invalid("ink actions root is absent or unterminated"));
    }
    let length_unit = length_unit.ok_or_else(|| invalid("ink actions lengthUnit is missing"))?;
    let time_unit = time_unit.ok_or_else(|| invalid("ink actions timeUnit is missing"))?;
    let source = copy_source(xml)?;
    Ok(Actions {
        source,
        length_unit: length_unit.into(),
        time_unit: time_unit.into(),
        action_groups: groups,
        actions,
    })
}

/// Write one action part using exact retained source bytes.
///
/// # Errors
///
/// Returns an error only when the caller-provided sink fails.
pub fn write_to<W: Write>(writer: &mut W, actions: &Actions) -> Result<()> {
    writer.write_all(actions.source())?;
    Ok(())
}

/// Write one action part to a byte vector.
///
/// # Errors
///
/// Returns an error when source validation fails.
pub fn write(actions: &Actions) -> Result<Vec<u8>> {
    if actions.source().len() > MAX_SOURCE_BYTES {
        return Err(limit("ink actions source bytes", MAX_SOURCE_BYTES));
    }
    Ok(actions.source().to_vec())
}

#[derive(Debug)]
struct Frame {
    namespace: NamespaceId,
    action: Option<usize>,
    action_group: bool,
    direct_actions: usize,
}

fn record_group_action(stack: &mut [Frame]) {
    if let Some(frame) = stack.last_mut()
        && frame.action_group
    {
        frame.direct_actions = frame.direct_actions.saturating_add(1);
    }
}

fn parse_action<R: std::io::BufRead>(
    element: &BytesStart<'_>,
    start: usize,
    end: usize,
    reader: &NsReader<R>,
) -> Result<Action> {
    let action_type = required_attr(element, b"type", reader)?;
    let start_time = required_attr(element, b"startTime", reader)?;
    validate_decimal(&start_time)?;
    Ok(Action {
        action_type: ActionType::parse(&action_type)?,
        start_time: start_time.into_boxed_str(),
        source_start: start,
        source_end: end,
    })
}

fn require_action_attrs<R: std::io::BufRead>(
    element: &BytesStart<'_>,
    reader: &NsReader<R>,
) -> Result<()> {
    let kind = required_attr(element, b"type", reader)?;
    let time = required_attr(element, b"startTime", reader)?;
    ActionType::parse(&kind)?;
    validate_decimal(&time)
}

fn required_attr<R: std::io::BufRead>(
    element: &BytesStart<'_>,
    name: &[u8],
    reader: &NsReader<R>,
) -> Result<String> {
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
            return Err(invalid("ink actions has a duplicate typed attribute"));
        }
        if attribute.value.len() > MAX_ATTRIBUTE_VALUE_BYTES {
            return Err(limit(
                "ink actions attribute value bytes",
                MAX_ATTRIBUTE_VALUE_BYTES,
            ));
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, reader.decoder())
            .map_err(xml_error)?;
        validate_xml_characters(&value, "ink actions attribute")?;
        let mut owned = String::new();
        owned
            .try_reserve_exact(value.len())
            .map_err(|_| invalid("ink actions attribute allocation failed"))?;
        owned.push_str(&value);
        result = Some(owned);
    }
    result.ok_or_else(|| {
        invalid(format!(
            "ink actions attribute '{}' is missing",
            String::from_utf8_lossy(name)
        ))
    })
}

fn validate_decimal(value: &str) -> Result<()> {
    if value.is_empty() || value.len() > 256 {
        return Err(invalid("ink action decimal is empty or overlong"));
    }
    let mut digits = 0usize;
    let mut dot = false;
    for (index, byte) in value.bytes().enumerate() {
        if (byte == b'+' || byte == b'-') && index == 0 {
            continue;
        }
        if byte == b'.' && !dot {
            dot = true;
            continue;
        }
        if !byte.is_ascii_digit() {
            return Err(invalid("ink action startTime is not xsd:decimal"));
        }
        digits += 1;
    }
    if digits == 0 {
        Err(invalid("ink action startTime has no digits"))
    } else {
        Ok(())
    }
}

fn validate_unit(value: &str, name: &'static str) -> Result<()> {
    if value.is_empty()
        || value.len() > MAX_TOKEN_BYTES
        || value
            .bytes()
            .any(|byte| byte == 0 || byte.is_ascii_whitespace())
    {
        return Err(invalid(format!(
            "ink actions {name} is empty, overlong, or contains whitespace"
        )));
    }
    Ok(())
}

fn require_root(element: &BytesStart<'_>, resolved: &ResolveResult<'_>) -> Result<bool> {
    if element.name().local_name().as_ref() != b"actions" {
        return Err(invalid("ink actions root must be iact:actions"));
    }
    match resolved {
        ResolveResult::Bound(Namespace(value)) if *value == ACTION_NAMESPACE.as_bytes() => {
            Ok(false)
        },
        ResolveResult::Unknown(prefix)
            if prefix.as_slice() == b"iact"
                && !has_namespace_declaration(element, prefix.as_slice()) =>
        {
            Ok(true)
        },
        _ => Err(invalid("ink actions root must be iact:actions")),
    }
}

fn is_action_namespace(
    resolved: &ResolveResult<'_>,
    element: &BytesStart<'_>,
    legacy_fragment: bool,
) -> bool {
    match resolved {
        ResolveResult::Bound(Namespace(value)) => *value == ACTION_NAMESPACE.as_bytes(),
        ResolveResult::Unknown(prefix) => {
            legacy_fragment
                && prefix.as_slice() == b"iact"
                && !has_namespace_declaration(element, prefix.as_slice())
        },
        ResolveResult::Unbound => false,
    }
}

fn namespace_id(resolved: &ResolveResult<'_>) -> Result<NamespaceId> {
    match resolved {
        ResolveResult::Bound(Namespace(value)) => {
            std::str::from_utf8(value).map_err(xml_error)?;
            Ok(if *value == ACTION_NAMESPACE.as_bytes() {
                NamespaceId::Action
            } else {
                NamespaceId::Other
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
            || prefix.as_slice() != b"iact"
        {
            return Err(invalid(
                "ink actions element uses an undeclared namespace prefix",
            ));
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
        return Err(invalid("ink actions closing element name is invalid"));
    }
    if let ResolveResult::Unknown(prefix) = resolved {
        std::str::from_utf8(prefix).map_err(xml_error)?;
        if !legacy_fragment || prefix.as_slice() != b"iact" {
            return Err(invalid(
                "ink actions closing element uses an undeclared namespace prefix",
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
        return Err(invalid("ink actions element name is invalid"));
    }
    let mut attribute_keys = Vec::new();
    for attribute in element.attributes() {
        let attribute = attribute.map_err(xml_error)?;
        if attribute_keys.len() >= MAX_ATTRIBUTES_PER_ELEMENT {
            return Err(limit(
                "ink actions attributes per element",
                MAX_ATTRIBUTES_PER_ELEMENT,
            ));
        }
        let name = std::str::from_utf8(attribute.key.as_ref()).map_err(xml_error)?;
        if !is_qualified_name(name) {
            return Err(invalid("ink actions attribute name is invalid"));
        }
        if attribute.value.len() > MAX_ATTRIBUTE_VALUE_BYTES {
            return Err(limit(
                "ink actions attribute value bytes",
                MAX_ATTRIBUTE_VALUE_BYTES,
            ));
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, reader.decoder())
            .map_err(xml_error)?;
        validate_xml_characters(&value, "ink actions attribute")?;
        if let Some(prefix) = attribute.key.prefix()
            && !matches!(prefix.as_ref(), b"xml" | b"xmlns")
            && matches!(
                reader.resolver().resolve_attribute(attribute.key).0,
                ResolveResult::Unknown(_)
            )
        {
            return Err(invalid(
                "ink actions attribute uses an undeclared namespace prefix",
            ));
        }
        if attribute_keys
            .iter()
            .copied()
            .any(|key| expanded_attribute_names_equal(key, attribute.key, reader))
        {
            return Err(invalid(
                "ink actions element has duplicate expanded attributes",
            ));
        }
        attribute_keys
            .try_reserve(1)
            .map_err(|_| invalid("ink actions attribute-key allocation failed"))?;
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
        return Err(invalid("ink actions XML declaration is malformed"));
    }
    skip_decl_whitespace(raw, &mut cursor);
    let (name, value) = parse_decl_attribute(raw, &mut cursor)?;
    if name != b"version" || value != b"1.0" {
        return Err(invalid(
            "ink actions XML declaration must start with version 1.0",
        ));
    }
    let mut previous = b"version".as_slice();
    while {
        skip_decl_whitespace(raw, &mut cursor);
        cursor < raw.len()
    } {
        let (name, value) = parse_decl_attribute(raw, &mut cursor)?;
        let valid_order = match (previous, name) {
            (b"version", b"encoding") => value.eq_ignore_ascii_case(b"utf-8"),
            (b"version" | b"encoding", b"standalone") => matches!(value, b"yes" | b"no"),
            _ => false,
        };
        if !valid_order {
            return Err(invalid(
                "ink actions XML declaration has an invalid or duplicate attribute",
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
        return Err(invalid(
            "ink actions XML declaration attribute name is missing",
        ));
    }
    let name = &raw[name_start..*cursor];
    skip_decl_whitespace(raw, cursor);
    if raw.get(*cursor) != Some(&b'=') {
        return Err(invalid(
            "ink actions XML declaration attribute equals sign is missing",
        ));
    }
    *cursor += 1;
    skip_decl_whitespace(raw, cursor);
    let quote = raw
        .get(*cursor)
        .copied()
        .filter(|value| matches!(value, b'\'' | b'"'))
        .ok_or_else(|| invalid("ink actions XML declaration attribute quote is missing"))?;
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
        return Err(invalid(
            "ink actions XML declaration attribute is unterminated",
        ));
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
                return Err(invalid(
                    "ink actions hexadecimal character reference is invalid",
                ));
            }
            let codepoint = u32::from_str_radix(digits, 16)
                .map_err(|_| invalid("ink actions hexadecimal character reference is invalid"))?;
            let character = char::from_u32(codepoint)
                .ok_or_else(|| invalid("ink actions character reference is invalid"))?;
            if is_xml10_character(character) {
                Ok(())
            } else {
                Err(invalid(
                    "ink actions character reference is not an XML 1.0 character",
                ))
            }
        },
        value if value.strip_prefix('#').is_some() => {
            let digits = value.strip_prefix('#').unwrap_or_default();
            if digits.is_empty() || !digits.bytes().all(|byte| byte.is_ascii_digit()) {
                return Err(invalid(
                    "ink actions decimal character reference is invalid",
                ));
            }
            let codepoint = digits
                .parse::<u32>()
                .map_err(|_| invalid("ink actions decimal character reference is invalid"))?;
            let character = char::from_u32(codepoint)
                .ok_or_else(|| invalid("ink actions character reference is invalid"))?;
            if is_xml10_character(character) {
                Ok(())
            } else {
                Err(invalid(
                    "ink actions character reference is not an XML 1.0 character",
                ))
            }
        },
        _ => Err(invalid(
            "ink actions general entity references are not supported",
        )),
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

fn copy_source(xml: &[u8]) -> Result<Arc<[u8]>> {
    let mut source = Vec::new();
    source
        .try_reserve_exact(xml.len())
        .map_err(|_| invalid("ink actions source allocation failed"))?;
    source.extend_from_slice(xml);
    Ok(Arc::from(source.into_boxed_slice()))
}
fn pos<R: std::io::BufRead>(reader: &NsReader<R>) -> Result<usize> {
    usize::try_from(reader.buffer_position())
        .map_err(|_| invalid("ink actions offset exceeds usize"))
}
fn increment_nodes(nodes: &mut usize) -> Result<()> {
    let next = nodes
        .checked_add(1)
        .ok_or_else(|| limit("ink actions XML nodes", MAX_NODES))?;
    enforce_nodes(next)?;
    *nodes = next;
    Ok(())
}
fn enforce_nodes(nodes: usize) -> Result<()> {
    if nodes > MAX_NODES {
        Err(limit("ink actions XML nodes", MAX_NODES))
    } else {
        Ok(())
    }
}
fn enforce_depth(depth: usize) -> Result<()> {
    if depth > MAX_DEPTH {
        Err(limit("ink actions XML depth", MAX_DEPTH))
    } else {
        Ok(())
    }
}
fn increment_groups(groups: &mut usize) -> Result<()> {
    let next = groups
        .checked_add(1)
        .ok_or_else(|| limit("ink action groups", MAX_ACTION_GROUPS))?;
    if next > MAX_ACTION_GROUPS {
        return Err(limit("ink action groups", MAX_ACTION_GROUPS));
    }
    *groups = next;
    Ok(())
}
fn ensure_action_capacity(actions: &[Action]) -> Result<()> {
    if actions.len() >= MAX_ACTIONS {
        Err(limit("ink actions", MAX_ACTIONS))
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
