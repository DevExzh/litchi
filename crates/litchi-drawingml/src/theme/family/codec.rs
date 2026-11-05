//! Bounded, namespace-aware XML codec for `thm15:themeFamily`.

use std::{fmt, io::Write, sync::Arc};

use litchi_core::xml::ReaderOrigin;
use litchi_ooxml_common::xml_name::is_qualified_name;
use quick_xml::{
    XmlVersion,
    escape::resolve_xml_entity,
    events::{BytesDecl, BytesRef, BytesStart, Event},
    name::{Namespace, QName, ResolveResult},
    reader::NsReader,
};

use crate::{Error, LimitKind, Result};

use super::model::{Family, Guid, Source, checked_name, validate_xml_text};
use super::{
    DRAWINGML_NAMESPACE, DRAWINGML_NAMESPACE_STRICT, MAX_ATTRIBUTE_VALUE_BYTES, MAX_ATTRIBUTES,
    MAX_DEPTH, MAX_NAMESPACE_BYTES, MAX_NAMESPACE_DECLARATIONS, MAX_NODES, MAX_XML_BYTES,
    NAMESPACE, XML_NAMESPACE, XMLNS_NAMESPACE,
};
use litchi_ooxml_common::xml::attributes::BytesStartExt as _;

/// Caller-selected bounds for a standalone `themeFamily` fragment.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
#[must_use]
pub struct Limits {
    /// Maximum complete fragment bytes accepted or emitted.
    pub max_xml_bytes: usize,
    /// Maximum element nodes scanned.
    pub max_nodes: usize,
    /// Maximum element nesting depth.
    pub max_depth: usize,
    /// Maximum attributes on one element.
    pub max_attributes: usize,
    /// Maximum namespace declarations on one element.
    pub max_namespace_declarations: usize,
    /// Maximum decoded attribute value bytes.
    pub max_attribute_value_bytes: usize,
    /// Maximum namespace prefix or URI bytes.
    pub max_namespace_bytes: usize,
}

impl Limits {
    /// The historical bounded fragment policy.
    pub const DEFAULT: Self = Self {
        max_xml_bytes: MAX_XML_BYTES,
        max_nodes: MAX_NODES,
        max_depth: MAX_DEPTH,
        max_attributes: MAX_ATTRIBUTES,
        max_namespace_declarations: MAX_NAMESPACE_DECLARATIONS,
        max_attribute_value_bytes: MAX_ATTRIBUTE_VALUE_BYTES,
        max_namespace_bytes: MAX_NAMESPACE_BYTES,
    };

    /// Construct a policy with a caller-selected complete-fragment byte cap.
    #[must_use]
    pub const fn new(max_xml_bytes: usize) -> Self {
        Self {
            max_xml_bytes,
            ..Self::DEFAULT
        }
    }

    fn validate(self) -> Result<Self> {
        check_limit(LimitKind::XmlBytes, self.max_xml_bytes, MAX_XML_BYTES)?;
        check_limit(LimitKind::XmlNodes, self.max_nodes, MAX_NODES)?;
        check_limit(LimitKind::XmlDepth, self.max_depth, MAX_DEPTH)?;
        check_limit(
            LimitKind::XmlAttributes,
            self.max_attributes,
            MAX_ATTRIBUTES,
        )?;
        check_limit(
            LimitKind::NamespaceDeclarations,
            self.max_namespace_declarations,
            MAX_NAMESPACE_DECLARATIONS,
        )?;
        check_limit(
            LimitKind::AttributeValueBytes,
            self.max_attribute_value_bytes,
            MAX_ATTRIBUTE_VALUE_BYTES,
        )?;
        check_limit(
            LimitKind::NamespaceBytes,
            self.max_namespace_bytes,
            MAX_NAMESPACE_BYTES,
        )?;
        Ok(self)
    }
}

impl Default for Limits {
    fn default() -> Self {
        Self::DEFAULT
    }
}

/// Read one complete `themeFamily` fragment while retaining its source bytes.
///
/// The input may contain an XML declaration or comments around the root.  The
/// returned source includes those bytes, so a no-op write is byte-for-byte
/// identical to the input.  Every element is still scanned for namespace,
/// nesting, node, attribute, and value bounds, including opaque extension
/// children.
pub fn read(xml: &[u8], limits: Limits) -> Result<Family> {
    let limits = limits.validate()?;
    if xml.len() > limits.max_xml_bytes {
        return Err(limit(LimitKind::XmlBytes, xml.len(), limits.max_xml_bytes));
    }
    let parsed = scan(xml, limits)?;
    let source = Arc::<[u8]>::from(xml);
    Ok(Family::from_source(Source {
        xml: source,
        root_start: parsed.root_start,
        root_end: parsed.root_end,
        name: parsed.name,
        id: parsed.id,
        variant_id: parsed.variant_id,
    }))
}

/// Read a complete fragment from an already shared immutable source allocation.
///
/// The caller's `Arc<[u8]>` is retained after validation, avoiding a second
/// source copy for package owners that already use shared part storage.
pub fn read_shared(xml: Arc<[u8]>, limits: Limits) -> Result<Family> {
    let limits = limits.validate()?;
    if xml.len() > limits.max_xml_bytes {
        return Err(limit(LimitKind::XmlBytes, xml.len(), limits.max_xml_bytes));
    }
    let parsed = scan(xml.as_ref(), limits)?;
    Ok(Family::from_source(Source {
        xml,
        root_start: parsed.root_start,
        root_end: parsed.root_end,
        name: parsed.name,
        id: parsed.id,
        variant_id: parsed.variant_id,
    }))
}

/// Serialize one `themeFamily` fragment.
///
/// A parsed value whose typed fields are unchanged returns its exact source.
/// Known-field edits patch only `name`, `id`, and `vid`; unknown attributes,
/// namespace choices, comments, and the optional family-namespace `extLst`
/// subtree remain byte-for-byte intact.
pub fn write(value: &Family, limits: Limits) -> Result<Vec<u8>> {
    let limits = limits.validate()?;
    validate(value, limits)?;
    if source_is_exact(value) {
        if let Some(source) = value.source_state() {
            if source.xml.len() > limits.max_xml_bytes {
                return Err(limit(
                    LimitKind::XmlBytes,
                    source.xml.len(),
                    limits.max_xml_bytes,
                ));
            }
            let mut output = Vec::new();
            output
                .try_reserve_exact(source.xml.len())
                .map_err(|_| invalid("theme family output allocation failed"))?;
            output.extend_from_slice(source.xml.as_ref());
            return Ok(output);
        }
    }
    if let Some(source) = value.source_state() {
        return rewrite_source(source, value, limits);
    }
    write_detached(value, limits)
}

/// Serialize one `themeFamily` fragment to a caller-provided sink.
pub fn write_to<W: Write>(writer: &mut W, value: &Family, limits: Limits) -> Result<()> {
    let limits = limits.validate()?;
    validate(value, limits)?;
    if source_is_exact(value) {
        if let Some(source) = value.source_state() {
            if source.xml.len() > limits.max_xml_bytes {
                return Err(limit(
                    LimitKind::XmlBytes,
                    source.xml.len(),
                    limits.max_xml_bytes,
                ));
            }
            writer.write_all(source.xml.as_ref())?;
            return Ok(());
        }
    } else if let Some(source) = value.source_state() {
        let rewrite = source_rewrite(source, value, limits)?;
        write_source_rewrite_to(writer, source, &rewrite)?;
    } else {
        write_detached_to(writer, value, limits)?;
    }
    Ok(())
}

#[derive(Debug)]
struct Parsed {
    root_start: usize,
    root_end: usize,
    name: Arc<str>,
    id: Guid,
    variant_id: Guid,
}

#[derive(Debug)]
struct Frame {
    local: Vec<u8>,
    namespace: Vec<u8>,
    kind: FrameKind,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum FrameKind {
    Root,
    FamilyExtensionList,
    DrawingExtension,
    Opaque,
}

fn classify_child(local: &[u8], namespace: &[u8], stack: &[Frame]) -> Result<FrameKind> {
    let parent = stack
        .last()
        .ok_or_else(|| invalid("theme family child has no parent"))?;
    match parent.kind {
        FrameKind::Root => {
            if local == b"extLst" && namespace == NAMESPACE.as_bytes() {
                return Ok(FrameKind::FamilyExtensionList);
            }
            if namespace == NAMESPACE.as_bytes() {
                return Err(invalid(
                    "theme family root has an unexpected family-namespace child",
                ));
            }
            Ok(FrameKind::Opaque)
        },
        FrameKind::FamilyExtensionList => {
            if local == b"ext" && is_drawingml_namespace(namespace) {
                Ok(FrameKind::DrawingExtension)
            } else {
                Err(invalid(
                    "theme family extLst contains a non-DrawingML ext child",
                ))
            }
        },
        FrameKind::DrawingExtension | FrameKind::Opaque => Ok(FrameKind::Opaque),
    }
}

fn is_drawingml_namespace(namespace: &[u8]) -> bool {
    namespace == DRAWINGML_NAMESPACE.as_bytes()
        || namespace == DRAWINGML_NAMESPACE_STRICT.as_bytes()
}

fn validate_extension_attributes(
    element: &BytesStart<'_>,
    reader: &NsReader<&[u8]>,
    limits: Limits,
) -> Result<()> {
    let mut uri_seen = false;
    for attribute in element.checked_attributes() {
        let attribute = attribute.map_err(xml_error)?;
        if is_namespace_attribute(attribute.key) {
            continue;
        }
        if attribute.key.prefix().is_none() && attribute.key.local_name().as_ref() == b"uri" {
            if uri_seen {
                return Err(invalid("theme family a:ext has duplicate uri attributes"));
            }
            // The generic attribute pass has already bounded and decoded this
            // value. Decode again only for the required-field presence check;
            // no normalized URI is exposed by the focused family model.
            decode_attribute(&attribute, reader.decoder(), limits)?;
            uri_seen = true;
        }
    }
    if !uri_seen {
        return Err(invalid("theme family a:ext lacks required uri"));
    }
    Ok(())
}

fn validate_element_name(element: &BytesStart<'_>, limits: Limits) -> Result<()> {
    let qualified_name = element.name();
    validate_qname_prefix(qualified_name, limits)?;
    let name = std::str::from_utf8(qualified_name.as_ref()).map_err(xml_error)?;
    if !is_qualified_name(name) {
        return Err(invalid("theme family element name is invalid"));
    }
    Ok(())
}

fn validate_end_name(element: &quick_xml::events::BytesEnd<'_>) -> Result<()> {
    let qualified_name = element.name();
    let name = std::str::from_utf8(qualified_name.as_ref()).map_err(xml_error)?;
    if !is_qualified_name(name) {
        return Err(invalid("theme family closing element name is invalid"));
    }
    Ok(())
}

fn in_element_only_context(stack: &[Frame]) -> bool {
    stack.last().is_some_and(|frame| {
        matches!(
            frame.kind,
            FrameKind::Root | FrameKind::FamilyExtensionList | FrameKind::DrawingExtension
        )
    })
}

fn is_xml_whitespace(value: &[u8]) -> bool {
    value
        .iter()
        .all(|byte| matches!(byte, b' ' | b'\t' | b'\r' | b'\n'))
}

fn is_xml_whitespace_character(character: char) -> bool {
    matches!(character, ' ' | '\t' | '\r' | '\n')
}

pub(crate) fn validate_general_ref(
    reference: &BytesRef<'_>,
    limits: Limits,
) -> Result<Option<char>> {
    if reference.len() > limits.max_attribute_value_bytes {
        return Err(limit(
            LimitKind::AttributeValueBytes,
            reference.len(),
            limits.max_attribute_value_bytes,
        ));
    }
    if reference.is_char_ref() {
        let character = reference
            .resolve_char_ref()
            .map_err(xml_error)?
            .ok_or_else(|| invalid("theme family character reference is invalid"))?;
        let mut encoded = [0u8; 4];
        validate_xml_text(
            character.encode_utf8(&mut encoded),
            "theme family character reference",
        )?;
        return Ok(Some(character));
    }
    let entity = reference.decode().map_err(xml_error)?;
    let replacement = resolve_xml_entity(entity.as_ref())
        .ok_or_else(|| invalid("theme family contains an unknown general entity"))?;
    let mut characters = replacement.chars();
    let character = characters.next();
    if characters.next().is_some() {
        return Err(invalid("theme family predefined entity is not scalar text"));
    }
    Ok(character)
}

fn scan(xml: &[u8], limits: Limits) -> Result<Parsed> {
    let mut reader = NsReader::from_reader(xml);
    let origin = ReaderOrigin::of(xml);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = true;
    reader.config_mut().check_comments = true;
    reader
        .resolver_mut()
        .set_max_declarations_per_element(limits.max_namespace_declarations);

    let mut buffer = Vec::new();
    let mut stack = Vec::<Frame>::new();
    let mut root_seen = false;
    let mut root_closed = false;
    let mut root_start = None;
    let mut root_end = None;
    let mut parsed = None;
    let mut ext_list_seen = false;
    let mut nodes = 0usize;
    let mut declaration_seen = false;
    let mut pre_root_event_seen = false;

    loop {
        let event_start = position(&reader, origin, "theme family")?;
        let event = reader
            .read_event_into(&mut buffer)
            .map_err(xml_error)?
            .into_owned();
        let resolved = match &event {
            Event::Start(element) | Event::Empty(element) => {
                validate_qname_prefix(element.name(), limits)?;
                reader.resolver().resolve_element(element.name()).0
            },
            Event::End(element) => {
                validate_qname_prefix(element.name(), limits)?;
                reader.resolver().resolve_element(element.name()).0
            },
            _ => ResolveResult::Unbound,
        };
        let event_namespace = resolved_namespace(&resolved, limits)?;
        let event_end = position(&reader, origin, "theme family")?;
        if !root_seen && !matches!(&event, Event::Decl(_) | Event::Eof) {
            pre_root_event_seen = true;
        }
        match event {
            Event::Decl(declaration) => {
                if root_seen || declaration_seen || pre_root_event_seen {
                    return Err(invalid("theme family XML declaration is misplaced"));
                }
                validate_declaration(&declaration)?;
                declaration_seen = true;
            },
            Event::Start(element) => {
                nodes = nodes
                    .checked_add(1)
                    .ok_or_else(|| invalid("theme family node count overflow"))?;
                if nodes > limits.max_nodes {
                    return Err(limit(LimitKind::XmlNodes, nodes, limits.max_nodes));
                }
                let namespace = event_namespace.clone();
                validate_element_name(&element, limits)?;
                validate_element_attributes(&element, &reader, limits)?;
                let local = element.local_name().as_ref().to_vec();
                let kind = if !root_seen {
                    require_root(&local, &namespace)?;
                    let root = parse_root_attributes(&element, &reader, limits)?;
                    root_seen = true;
                    root_start = Some(event_start);
                    parsed = Some(root);
                    FrameKind::Root
                } else {
                    if root_closed || stack.is_empty() {
                        return Err(invalid("theme family fragment has more than one root"));
                    }
                    let kind = classify_child(&local, &namespace, &stack)?;
                    if kind == FrameKind::FamilyExtensionList {
                        if ext_list_seen {
                            return Err(invalid("theme family contains duplicate extLst elements"));
                        }
                        ext_list_seen = true;
                    }
                    if kind == FrameKind::DrawingExtension {
                        validate_extension_attributes(&element, &reader, limits)?;
                    }
                    kind
                };
                if stack.len() >= limits.max_depth {
                    return Err(limit(
                        LimitKind::XmlDepth,
                        stack.len() + 1,
                        limits.max_depth,
                    ));
                }
                stack.push(Frame {
                    local,
                    namespace,
                    kind,
                });
            },
            Event::Empty(element) => {
                nodes = nodes
                    .checked_add(1)
                    .ok_or_else(|| invalid("theme family node count overflow"))?;
                if nodes > limits.max_nodes {
                    return Err(limit(LimitKind::XmlNodes, nodes, limits.max_nodes));
                }
                let namespace = event_namespace.clone();
                validate_element_name(&element, limits)?;
                validate_element_attributes(&element, &reader, limits)?;
                let local = element.local_name().as_ref().to_vec();
                let _kind = if !root_seen {
                    require_root(&local, &namespace)?;
                    let root = parse_root_attributes(&element, &reader, limits)?;
                    root_seen = true;
                    root_closed = true;
                    root_start = Some(event_start);
                    root_end = Some(event_end);
                    parsed = Some(root);
                    FrameKind::Root
                } else {
                    if root_closed || stack.is_empty() {
                        return Err(invalid("theme family fragment has more than one root"));
                    }
                    let kind = classify_child(&local, &namespace, &stack)?;
                    if kind == FrameKind::FamilyExtensionList {
                        if ext_list_seen {
                            return Err(invalid("theme family contains duplicate extLst elements"));
                        }
                        ext_list_seen = true;
                    }
                    if kind == FrameKind::DrawingExtension {
                        validate_extension_attributes(&element, &reader, limits)?;
                    }
                    kind
                };
                if !root_closed && stack.len() >= limits.max_depth {
                    return Err(limit(
                        LimitKind::XmlDepth,
                        stack.len() + 1,
                        limits.max_depth,
                    ));
                }
            },
            Event::End(element) => {
                if !root_seen || root_closed {
                    return Err(invalid("theme family has markup outside its root"));
                }
                validate_end_name(&element)?;
                let namespace = event_namespace;
                let Some(frame) = stack.pop() else {
                    return Err(invalid("theme family has an unexpected closing element"));
                };
                if frame.local.as_slice() != element.local_name().as_ref()
                    || frame.namespace != namespace
                {
                    return Err(invalid("theme family closing element does not match"));
                }
                if stack.is_empty() {
                    root_closed = true;
                    root_end = Some(event_end);
                }
            },
            Event::Text(text) => {
                validate_text_node(text.as_ref(), "theme family text")?;
                if (!root_seen || root_closed) && !is_xml_whitespace(text.as_ref()) {
                    return Err(invalid("theme family has text outside its root"));
                }
                if in_element_only_context(&stack) && !is_xml_whitespace(text.as_ref()) {
                    return Err(invalid("theme family element-only container contains text"));
                }
            },
            Event::CData(text) => {
                validate_event_text(text.as_ref(), "theme family CDATA")?;
                if !root_seen || root_closed {
                    return Err(invalid("theme family has CDATA outside its root"));
                }
                if in_element_only_context(&stack) && !is_xml_whitespace(text.as_ref()) {
                    return Err(invalid(
                        "theme family element-only container contains CDATA",
                    ));
                }
            },
            Event::GeneralRef(reference) => {
                if !root_seen || root_closed {
                    return Err(invalid("theme family has a reference outside its root"));
                }
                let character = validate_general_ref(&reference, limits)?;
                if in_element_only_context(&stack)
                    && !character.is_some_and(is_xml_whitespace_character)
                {
                    return Err(invalid(
                        "theme family element-only container contains a reference",
                    ));
                }
            },
            Event::Comment(comment) => {
                validate_event_text(comment.as_ref(), "theme family comment")?;
            },
            Event::PI(_) | Event::DocType(_) => {
                return Err(invalid("theme family contains forbidden document markup"));
            },
            Event::Eof => break,
        }
        buffer.clear();
    }

    if !root_seen || !root_closed || !stack.is_empty() {
        return Err(invalid("theme family fragment is unterminated"));
    }
    let parsed = parsed.ok_or_else(|| invalid("theme family has no root"))?;
    Ok(Parsed {
        root_start: root_start.ok_or_else(|| invalid("theme family root start is missing"))?,
        root_end: root_end.ok_or_else(|| invalid("theme family root end is missing"))?,
        name: parsed.name,
        id: parsed.id,
        variant_id: parsed.variant_id,
    })
}

struct RootAttributes {
    name: Arc<str>,
    id: Guid,
    variant_id: Guid,
}

fn parse_root_attributes(
    element: &BytesStart<'_>,
    reader: &NsReader<&[u8]>,
    limits: Limits,
) -> Result<RootAttributes> {
    let mut name = None;
    let mut id = None;
    let mut variant_id = None;
    for attribute in element.checked_attributes() {
        let attribute = attribute.map_err(xml_error)?;
        if is_namespace_attribute(attribute.key) {
            continue;
        }
        let raw_name = attribute.key.as_ref();
        if attribute.key.prefix().is_none() {
            let value = decode_attribute(&attribute, reader.decoder(), limits)?;
            match raw_name {
                b"name" => {
                    if name.replace(checked_name(&value)?).is_some() {
                        return Err(invalid("theme family has duplicate name attributes"));
                    }
                },
                b"id" => {
                    if id.replace(Guid::new(value).map_err(Error::from)?).is_some() {
                        return Err(invalid("theme family has duplicate id attributes"));
                    }
                },
                b"vid" => {
                    if variant_id
                        .replace(Guid::new(value).map_err(Error::from)?)
                        .is_some()
                    {
                        return Err(invalid("theme family has duplicate vid attributes"));
                    }
                },
                _ => {},
            }
        }
    }
    Ok(RootAttributes {
        name: name.ok_or_else(|| invalid("theme family lacks required name"))?,
        id: id.ok_or_else(|| invalid("theme family lacks required id"))?,
        variant_id: variant_id.ok_or_else(|| invalid("theme family lacks required vid"))?,
    })
}

fn validate_element_attributes(
    element: &BytesStart<'_>,
    reader: &NsReader<&[u8]>,
    limits: Limits,
) -> Result<()> {
    let mut count = 0usize;
    let mut namespaces = 0usize;
    let mut seen = Vec::<(Vec<u8>, Vec<u8>)>::new();
    for attribute in element.checked_attributes() {
        let attribute = attribute.map_err(xml_error)?;
        count = count
            .checked_add(1)
            .ok_or_else(|| invalid("theme family attribute count overflow"))?;
        if count > limits.max_attributes {
            return Err(limit(
                LimitKind::XmlAttributes,
                count,
                limits.max_attributes,
            ));
        }
        let raw_name = attribute.key.as_ref();
        let name = std::str::from_utf8(raw_name).map_err(xml_error)?;
        if !is_qualified_name(name) {
            return Err(invalid("theme family attribute name is invalid"));
        }
        validate_qname_prefix(attribute.key, limits)?;
        if is_namespace_attribute(attribute.key) {
            namespaces = namespaces
                .checked_add(1)
                .ok_or_else(|| invalid("theme family namespace count overflow"))?;
            if namespaces > limits.max_namespace_declarations {
                return Err(limit(
                    LimitKind::NamespaceDeclarations,
                    namespaces,
                    limits.max_namespace_declarations,
                ));
            }
            let value = decode_attribute(&attribute, reader.decoder(), limits)?;
            let prefix = attribute
                .key
                .as_ref()
                .strip_prefix(b"xmlns:")
                .unwrap_or(&[]);
            if prefix.len() > limits.max_namespace_bytes {
                return Err(limit(
                    LimitKind::NamespacePrefixBytes,
                    prefix.len(),
                    limits.max_namespace_bytes,
                ));
            }
            validate_namespace_binding(prefix, &value)?;
            if value.len() > limits.max_namespace_bytes {
                return Err(limit(
                    LimitKind::NamespaceBytes,
                    value.len(),
                    limits.max_namespace_bytes,
                ));
            }
            continue;
        }
        let value = decode_attribute(&attribute, reader.decoder(), limits)?;
        let (namespace, _) = reader.resolver().resolve_attribute(attribute.key);
        let namespace = match namespace {
            ResolveResult::Bound(Namespace(value)) => value.to_vec(),
            ResolveResult::Unbound => Vec::new(),
            ResolveResult::Unknown(prefix) => {
                if prefix.len() > limits.max_namespace_bytes {
                    return Err(limit(
                        LimitKind::NamespacePrefixBytes,
                        prefix.len(),
                        limits.max_namespace_bytes,
                    ));
                }
                return Err(invalid(format!(
                    "theme family attribute uses undeclared namespace prefix '{}'",
                    String::from_utf8_lossy(prefix.as_ref())
                )));
            },
        };
        if attribute
            .key
            .prefix()
            .is_some_and(|prefix| prefix.as_ref() != b"xml")
            && namespace.is_empty()
        {
            return Err(invalid("theme family attribute prefix is not bound"));
        }
        let local = attribute.key.local_name().as_ref().to_vec();
        if seen.iter().any(|(known_namespace, known_local)| {
            known_namespace == &namespace && known_local == &local
        }) {
            return Err(invalid("theme family has duplicate expanded attributes"));
        }
        seen.try_reserve(1)
            .map_err(|_| invalid("theme family attribute-key allocation failed"))?;
        seen.push((namespace, local));
        if value.len() > limits.max_attribute_value_bytes {
            return Err(limit(
                LimitKind::AttributeValueBytes,
                value.len(),
                limits.max_attribute_value_bytes,
            ));
        }
    }
    Ok(())
}

fn decode_attribute(
    attribute: &quick_xml::events::attributes::Attribute<'_>,
    decoder: quick_xml::encoding::Decoder,
    limits: Limits,
) -> Result<String> {
    // XML 1.0 attribute-value normalization applies before the schema type is
    // interpreted. Literal tabs, CR, and LF therefore become spaces even for
    // xsd:string; character references retain their represented characters.
    validate_raw_attribute_value(attribute.value.as_ref(), "theme family attribute")?;
    crate::preflight_attribute_value_bytes(
        attribute.value.as_ref(),
        limits.max_attribute_value_bytes,
    )?;
    let value = attribute
        .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
        .map_err(xml_error)?
        .into_owned();
    if value.len() > limits.max_attribute_value_bytes {
        return Err(limit(
            LimitKind::AttributeValueBytes,
            value.len(),
            limits.max_attribute_value_bytes,
        ));
    }
    validate_xml_text(&value, "theme family attribute")?;
    Ok(value)
}

fn require_root(local: &[u8], namespace: &[u8]) -> Result<()> {
    if local != b"themeFamily" || namespace != NAMESPACE.as_bytes() {
        return Err(invalid(
            "theme family root must be themeFamily in the 2012 theme namespace",
        ));
    }
    Ok(())
}

pub(crate) fn validate_declaration(declaration: &BytesDecl<'_>) -> Result<()> {
    let source = declaration.as_ref();
    if !source.starts_with(b"xml") {
        return Err(invalid("theme family XML declaration name is invalid"));
    }
    let mut cursor = 3usize;
    if source.get(cursor).is_none_or(|byte| !is_whitespace(*byte)) {
        return Err(invalid("theme family XML declaration lacks whitespace"));
    }

    let mut version_seen = false;
    let mut encoding_seen = false;
    let mut standalone_seen = false;
    while cursor < source.len() {
        while cursor < source.len() && is_whitespace(source[cursor]) {
            cursor += 1;
        }
        if cursor == source.len() {
            break;
        }

        let key_start = cursor;
        while cursor < source.len() && !is_whitespace(source[cursor]) && source[cursor] != b'=' {
            cursor += 1;
        }
        if key_start == cursor {
            return Err(invalid("theme family XML declaration attribute is missing"));
        }
        let key = &source[key_start..cursor];
        while cursor < source.len() && is_whitespace(source[cursor]) {
            cursor += 1;
        }
        if source.get(cursor) != Some(&b'=') {
            return Err(invalid(
                "theme family XML declaration attribute lacks equals",
            ));
        }
        cursor += 1;
        while cursor < source.len() && is_whitespace(source[cursor]) {
            cursor += 1;
        }
        let quote = match source.get(cursor).copied() {
            Some(quote @ (b'\'' | b'"')) => quote,
            Some(_) => {
                return Err(invalid(
                    "theme family XML declaration attribute is not quoted",
                ));
            },
            None => {
                return Err(invalid(
                    "theme family XML declaration attribute value is missing",
                ));
            },
        };
        cursor += 1;
        let value_start = cursor;
        while cursor < source.len() && source[cursor] != quote {
            cursor += 1;
        }
        if cursor == source.len() {
            return Err(invalid(
                "theme family XML declaration attribute value is unterminated",
            ));
        }
        let value = &source[value_start..cursor];
        cursor += 1;
        if cursor < source.len() && !is_whitespace(source[cursor]) {
            return Err(invalid(
                "theme family XML declaration attributes are not separated",
            ));
        }

        match key {
            b"version" if !version_seen && !encoding_seen && !standalone_seen => {
                if value != b"1.0" {
                    return Err(invalid(
                        "theme family XML declaration must use XML version 1.0",
                    ));
                }
                version_seen = true;
            },
            b"encoding" if version_seen && !encoding_seen && !standalone_seen => {
                if !value.eq_ignore_ascii_case(b"UTF-8") {
                    return Err(invalid(
                        "theme family XML declaration must use UTF-8 encoding",
                    ));
                }
                encoding_seen = true;
            },
            b"standalone" if version_seen && !standalone_seen => {
                if value != b"yes" && value != b"no" {
                    return Err(invalid(
                        "theme family XML declaration standalone value is invalid",
                    ));
                }
                standalone_seen = true;
            },
            b"version" | b"encoding" | b"standalone" => {
                return Err(invalid(
                    "theme family XML declaration has a duplicate or misplaced attribute",
                ));
            },
            _ => {
                return Err(invalid(
                    "theme family XML declaration has an unsupported attribute",
                ));
            },
        }
    }
    if !version_seen {
        return Err(invalid("theme family XML declaration lacks version"));
    }
    Ok(())
}

fn resolved_namespace(resolved: &ResolveResult<'_>, limits: Limits) -> Result<Vec<u8>> {
    match resolved {
        ResolveResult::Bound(Namespace(value)) => {
            if value.len() > limits.max_namespace_bytes {
                return Err(limit(
                    LimitKind::NamespaceBytes,
                    value.len(),
                    limits.max_namespace_bytes,
                ));
            }
            std::str::from_utf8(value).map_err(xml_error)?;
            Ok(value.to_vec())
        },
        ResolveResult::Unbound => Ok(Vec::new()),
        ResolveResult::Unknown(prefix) => {
            if prefix.len() > limits.max_namespace_bytes {
                return Err(limit(
                    LimitKind::NamespacePrefixBytes,
                    prefix.len(),
                    limits.max_namespace_bytes,
                ));
            }
            Err(invalid(format!(
                "theme family element uses undeclared namespace prefix '{}'",
                String::from_utf8_lossy(prefix.as_ref())
            )))
        },
    }
}

fn is_namespace_attribute(key: QName<'_>) -> bool {
    key.as_ref() == b"xmlns" || key.as_ref().starts_with(b"xmlns:")
}

fn validate_qname_prefix(name: QName<'_>, limits: Limits) -> Result<()> {
    if name
        .prefix()
        .is_some_and(|prefix| prefix.as_ref().len() > limits.max_namespace_bytes)
    {
        return Err(limit(
            LimitKind::NamespacePrefixBytes,
            name.prefix().map_or(0, |prefix| prefix.as_ref().len()),
            limits.max_namespace_bytes,
        ));
    }
    Ok(())
}

pub(crate) fn validate_event_text(value: &[u8], field: &str) -> Result<()> {
    let value = std::str::from_utf8(value).map_err(xml_error)?;
    validate_xml_text(value, field)
}

fn validate_text_node(value: &[u8], field: &str) -> Result<()> {
    validate_event_text(value, field)?;
    if value.windows(3).any(|window| window == b"]]>") {
        return Err(invalid(format!(
            "{field} contains the forbidden raw ']]>' delimiter"
        )));
    }
    Ok(())
}

pub(crate) fn validate_raw_attribute_value(value: &[u8], field: &str) -> Result<()> {
    if value.contains(&b'<') {
        return Err(invalid(format!("{field} contains a raw '<' delimiter")));
    }
    Ok(())
}

pub(crate) fn validate_namespace_binding(prefix: &[u8], value: &str) -> Result<()> {
    if prefix == b"xmlns" {
        return Err(invalid("theme family rebinds the reserved xmlns prefix"));
    }
    if value == XMLNS_NAMESPACE {
        return Err(invalid("theme family binds the reserved XMLNS namespace"));
    }
    if value == XML_NAMESPACE && prefix != b"xml" {
        return Err(invalid(
            "theme family binds the reserved XML namespace to another prefix",
        ));
    }
    if !prefix.is_empty() && value.is_empty() {
        return Err(invalid(
            "theme family uses an empty prefixed namespace binding",
        ));
    }
    if prefix == b"xml" && value != XML_NAMESPACE {
        return Err(invalid("theme family rebinds the xml namespace"));
    }
    Ok(())
}

fn detached_len(
    name_len: usize,
    id_len: usize,
    variant_id_len: usize,
    limits: Limits,
) -> Result<usize> {
    let output_len = b"<thm15:themeFamily xmlns:thm15=\""
        .len()
        .checked_add(NAMESPACE.len())
        .and_then(|length| length.checked_add(b"\" name=\"".len()))
        .and_then(|length| length.checked_add(name_len))
        .and_then(|length| length.checked_add(b"\" id=\"".len()))
        .and_then(|length| length.checked_add(id_len))
        .and_then(|length| length.checked_add(b"\" vid=\"".len()))
        .and_then(|length| length.checked_add(variant_id_len))
        .and_then(|length| length.checked_add(b"\"/>".len()))
        .ok_or_else(|| invalid("theme family output length overflows"))?;
    if output_len > limits.max_xml_bytes {
        return Err(limit(LimitKind::XmlBytes, output_len, limits.max_xml_bytes));
    }
    Ok(output_len)
}

/// Size the serialized fragment without allocating its complete output buffer.
pub(crate) fn serialized_len(value: &Family, limits: Limits) -> Result<usize> {
    let limits = limits.validate()?;
    validate(value, limits)?;
    if let Some(source) = value.source_state() {
        if source_is_exact(value) {
            if source.xml.len() > limits.max_xml_bytes {
                return Err(limit(
                    LimitKind::XmlBytes,
                    source.xml.len(),
                    limits.max_xml_bytes,
                ));
            }
            return Ok(source.xml.len());
        }
        return Ok(source_rewrite(source, value, limits)?.output_len);
    }
    detached_len(
        escaped_attribute_len(value.name())?,
        escaped_xml_len(value.id().as_str())?,
        escaped_xml_len(value.variant_id().as_str())?,
        limits,
    )
}

fn write_detached(value: &Family, limits: Limits) -> Result<Vec<u8>> {
    // Compute the exact escaped length before constructing any proportional
    // serialized strings. This makes a tight output cap reject before the
    // attribute values themselves are allocated.
    let output_len = detached_len(
        escaped_attribute_len(value.name())?,
        escaped_xml_len(value.id().as_str())?,
        escaped_xml_len(value.variant_id().as_str())?,
        limits,
    )?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(output_len)
        .map_err(|_| invalid("theme family output allocation failed"))?;
    write_detached_to(&mut output, value, limits)?;
    bounded_output(output, limits)
}

fn write_detached_to<W: Write>(writer: &mut W, value: &Family, limits: Limits) -> Result<()> {
    let _output_len = detached_len(
        escaped_attribute_len(value.name())?,
        escaped_xml_len(value.id().as_str())?,
        escaped_xml_len(value.variant_id().as_str())?,
        limits,
    )?;
    writer.write_all(b"<thm15:themeFamily xmlns:thm15=\"")?;
    writer.write_all(NAMESPACE.as_bytes())?;
    writer.write_all(b"\" name=\"")?;
    write_escaped_attribute(writer, value.name())?;
    writer.write_all(b"\" id=\"")?;
    write_escaped_xml(writer, value.id().as_str())?;
    writer.write_all(b"\" vid=\"")?;
    write_escaped_xml(writer, value.variant_id().as_str())?;
    writer.write_all(b"\"/>")?;
    Ok(())
}

fn escaped_attribute_len(value: &str) -> Result<usize> {
    escaped_len(value, true)
}

fn escaped_xml_len(value: &str) -> Result<usize> {
    escaped_len(value, false)
}

fn escaped_len(value: &str, attribute: bool) -> Result<usize> {
    value.chars().try_fold(0usize, |length, character| {
        let added = match character {
            '&' => 5,
            '<' | '>' => 4,
            '"' | '\'' => 6,
            '\t' if attribute => 4,
            '\n' | '\r' if attribute => 5,
            _ => character.len_utf8(),
        };
        length
            .checked_add(added)
            .ok_or_else(|| invalid("theme family escaped value length overflows"))
    })
}

struct SourceRewrite<'a> {
    output_len: usize,
    replacements: Vec<SourceReplacement<'a>>,
}

struct SourceReplacement<'a> {
    range: std::ops::Range<usize>,
    kind: u8,
    value: &'a str,
}

fn source_rewrite<'a>(
    source: &Source,
    value: &'a Family,
    limits: Limits,
) -> Result<SourceRewrite<'a>> {
    let root = source
        .xml
        .get(source.root_start..source.root_end)
        .ok_or_else(|| invalid("theme family source root range is invalid"))?;
    let tag_end = open_tag_end(root)?;
    let attributes = root_attributes(root, tag_end)?;
    let mut ranges = Vec::with_capacity(3);
    if source.name.as_ref() != value.name() {
        let range = attributes
            .name
            .clone()
            .ok_or_else(|| invalid("theme family source name attribute is missing"))?;
        ranges.push((range, escaped_attribute_len(value.name())?, 0_u8));
    }
    if source.id != *value.id() {
        let range = attributes
            .id
            .clone()
            .ok_or_else(|| invalid("theme family source id attribute is missing"))?;
        ranges.push((range, escaped_xml_len(value.id().as_str())?, 1_u8));
    }
    if source.variant_id != *value.variant_id() {
        let range = attributes
            .variant_id
            .clone()
            .ok_or_else(|| invalid("theme family source vid attribute is missing"))?;
        ranges.push((range, escaped_xml_len(value.variant_id().as_str())?, 2_u8));
    }
    let mut output_len = source.xml.len();
    for (range, replacement_len, _) in &ranges {
        output_len = output_len
            .checked_sub(
                range
                    .end
                    .checked_sub(range.start)
                    .ok_or_else(|| invalid("theme family replacement range underflow"))?,
            )
            .and_then(|length| length.checked_add(*replacement_len))
            .ok_or_else(|| invalid("theme family patched XML length overflows"))?;
    }
    if output_len > limits.max_xml_bytes {
        return Err(limit(LimitKind::XmlBytes, output_len, limits.max_xml_bytes));
    }
    let mut absolute = ranges
        .into_iter()
        .map(|(range, _, kind)| {
            let replacement = match kind {
                0 => value.name(),
                1 => value.id().as_str(),
                _ => value.variant_id().as_str(),
            };
            SourceReplacement {
                range: (source.root_start + range.start)..(source.root_start + range.end),
                kind,
                value: replacement,
            }
        })
        .collect::<Vec<_>>();
    absolute.sort_by_key(|replacement| replacement.range.start);
    Ok(SourceRewrite {
        output_len,
        replacements: absolute,
    })
}

fn rewrite_source(source: &Source, value: &Family, limits: Limits) -> Result<Vec<u8>> {
    let SourceRewrite {
        output_len,
        replacements: absolute,
    } = source_rewrite(source, value, limits)?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(output_len)
        .map_err(|_| invalid("theme family output allocation failed"))?;
    let rewrite = SourceRewrite {
        output_len,
        replacements: absolute,
    };
    write_source_rewrite_to(&mut output, source, &rewrite)?;
    bounded_output(output, limits)
}

fn write_source_rewrite_to<W: Write>(
    writer: &mut W,
    source: &Source,
    rewrite: &SourceRewrite<'_>,
) -> Result<()> {
    let mut cursor = 0usize;
    for replacement in &rewrite.replacements {
        writer.write_all(
            source
                .xml
                .get(cursor..replacement.range.start)
                .ok_or_else(|| invalid("theme family replacement range is invalid"))?,
        )?;
        match replacement.kind {
            0 => write_escaped_attribute(writer, replacement.value)?,
            1 | 2 => write_escaped_xml(writer, replacement.value)?,
            _ => return Err(invalid("theme family replacement kind is invalid")),
        }
        cursor = replacement.range.end;
    }
    writer.write_all(
        source
            .xml
            .get(cursor..)
            .ok_or_else(|| invalid("theme family source suffix range is invalid"))?,
    )?;
    Ok(())
}

struct RootAttributeRanges {
    name: Option<std::ops::Range<usize>>,
    id: Option<std::ops::Range<usize>>,
    variant_id: Option<std::ops::Range<usize>>,
}

fn root_attributes(root: &[u8], tag_end: usize) -> Result<RootAttributeRanges> {
    let mut cursor = 1usize;
    while cursor < tag_end && !is_name_start(root[cursor]) {
        cursor += 1;
    }
    while cursor < tag_end && !is_whitespace(root[cursor]) && root[cursor] != b'/' {
        cursor += 1;
    }
    let mut ranges = RootAttributeRanges {
        name: None,
        id: None,
        variant_id: None,
    };
    while cursor < tag_end {
        while cursor < tag_end && (is_whitespace(root[cursor]) || root[cursor] == b'/') {
            cursor += 1;
        }
        if cursor >= tag_end || root[cursor] == b'>' {
            break;
        }
        let key_start = cursor;
        while cursor < tag_end
            && !is_whitespace(root[cursor])
            && !matches!(root[cursor], b'=' | b'/' | b'>')
        {
            cursor += 1;
        }
        let key = root
            .get(key_start..cursor)
            .ok_or_else(|| invalid("theme family source attribute name range is invalid"))?;
        while cursor < tag_end && is_whitespace(root[cursor]) {
            cursor += 1;
        }
        if root.get(cursor) != Some(&b'=') {
            return Err(invalid("theme family source attribute has no value"));
        }
        cursor += 1;
        while cursor < tag_end && is_whitespace(root[cursor]) {
            cursor += 1;
        }
        let quote = *root
            .get(cursor)
            .ok_or_else(|| invalid("theme family source attribute value is missing"))?;
        if quote != b'\'' && quote != b'"' {
            return Err(invalid("theme family source attribute value is not quoted"));
        }
        cursor += 1;
        let value_start = cursor;
        while cursor < tag_end && root[cursor] != quote {
            cursor += 1;
        }
        let value_end = cursor;
        if cursor >= tag_end {
            return Err(invalid(
                "theme family source attribute value is unterminated",
            ));
        }
        if key == b"name" {
            if ranges.name.replace(value_start..value_end).is_some() {
                return Err(invalid("theme family source has duplicate name attributes"));
            }
        } else if key == b"id" {
            if ranges.id.replace(value_start..value_end).is_some() {
                return Err(invalid("theme family source has duplicate id attributes"));
            }
        } else if key == b"vid" && ranges.variant_id.replace(value_start..value_end).is_some() {
            return Err(invalid("theme family source has duplicate vid attributes"));
        }
        cursor += 1;
    }
    Ok(ranges)
}

fn open_tag_end(root: &[u8]) -> Result<usize> {
    let mut quote = None;
    for (index, byte) in root.iter().copied().enumerate() {
        match quote {
            Some(current) if byte == current => quote = None,
            Some(_) => {},
            None if byte == b'\'' || byte == b'"' => quote = Some(byte),
            None if byte == b'>' => return Ok(index + 1),
            None => {},
        }
    }
    Err(invalid(
        "theme family source root start tag is unterminated",
    ))
}

fn is_name_start(byte: u8) -> bool {
    byte.is_ascii_alphabetic() || byte == b'_' || byte == b':'
}

fn is_whitespace(byte: u8) -> bool {
    matches!(byte, b' ' | b'\t' | b'\r' | b'\n')
}

fn validate(value: &Family, limits: Limits) -> Result<()> {
    checked_name(value.name())?;
    let _ = Guid::new(value.id().as_str()).map_err(Error::from)?;
    let _ = Guid::new(value.variant_id().as_str()).map_err(Error::from)?;
    validate_attribute_budget(value.name(), limits.max_attribute_value_bytes)?;
    validate_attribute_budget(value.id().as_str(), limits.max_attribute_value_bytes)?;
    validate_attribute_budget(
        value.variant_id().as_str(),
        limits.max_attribute_value_bytes,
    )?;
    if let Some(source) = value.source_state() {
        if source.xml.len() > limits.max_xml_bytes {
            return Err(limit(
                LimitKind::XmlBytes,
                source.xml.len(),
                limits.max_xml_bytes,
            ));
        }
        if source.root_start >= source.root_end || source.root_end > source.xml.len() {
            return Err(invalid("theme family source root range is invalid"));
        }
        // A source parsed under a wider policy may later be written through a
        // narrower caller policy. Re-scan before any sink or output
        // allocation so structural, namespace, and decoded-value dimensions
        // are enforced for the actual writer policy as well.
        scan(source.xml.as_ref(), limits)?;
    } else {
        validate_detached_structure(limits)?;
    }
    Ok(())
}

fn validate_attribute_budget(value: &str, maximum: usize) -> Result<()> {
    if value.len() > maximum {
        return Err(limit(LimitKind::AttributeValueBytes, value.len(), maximum));
    }
    Ok(())
}

fn validate_detached_structure(limits: Limits) -> Result<()> {
    const ROOT_ATTRIBUTES: usize = 4;
    const ROOT_NODES: usize = 1;
    const ROOT_DEPTH: usize = 1;
    const ROOT_NAMESPACE_DECLARATIONS: usize = 1;
    const ROOT_PREFIX_BYTES: usize = b"thm15".len();
    if ROOT_NODES > limits.max_nodes {
        return Err(limit(LimitKind::XmlNodes, ROOT_NODES, limits.max_nodes));
    }
    if ROOT_DEPTH > limits.max_depth {
        return Err(limit(LimitKind::XmlDepth, ROOT_DEPTH, limits.max_depth));
    }
    if ROOT_ATTRIBUTES > limits.max_attributes {
        return Err(limit(
            LimitKind::XmlAttributes,
            ROOT_ATTRIBUTES,
            limits.max_attributes,
        ));
    }
    if ROOT_NAMESPACE_DECLARATIONS > limits.max_namespace_declarations {
        return Err(limit(
            LimitKind::NamespaceDeclarations,
            ROOT_NAMESPACE_DECLARATIONS,
            limits.max_namespace_declarations,
        ));
    }
    if ROOT_PREFIX_BYTES > limits.max_namespace_bytes {
        return Err(limit(
            LimitKind::NamespacePrefixBytes,
            ROOT_PREFIX_BYTES,
            limits.max_namespace_bytes,
        ));
    }
    if NAMESPACE.len() > limits.max_namespace_bytes {
        return Err(limit(
            LimitKind::NamespaceBytes,
            NAMESPACE.len(),
            limits.max_namespace_bytes,
        ));
    }
    Ok(())
}

fn source_is_exact(value: &Family) -> bool {
    value.source_state().is_some_and(|source| {
        source.name.as_ref() == value.name()
            && source.id == *value.id()
            && source.variant_id == *value.variant_id()
    })
}

fn bounded_output(output: Vec<u8>, limits: Limits) -> Result<Vec<u8>> {
    if output.len() > limits.max_xml_bytes {
        return Err(limit(
            LimitKind::XmlBytes,
            output.len(),
            limits.max_xml_bytes,
        ));
    }
    Ok(output)
}

fn write_escaped_attribute<W: Write>(writer: &mut W, value: &str) -> Result<()> {
    write_escaped_with(writer, value, |character| match character {
        '\t' => Some("&#9;"),
        '\n' => Some("&#10;"),
        '\r' => Some("&#13;"),
        '&' => Some("&amp;"),
        '<' => Some("&lt;"),
        '>' => Some("&gt;"),
        '"' => Some("&quot;"),
        '\'' => Some("&apos;"),
        _ => None,
    })
}

fn write_escaped_xml<W: Write>(writer: &mut W, value: &str) -> Result<()> {
    write_escaped_with(writer, value, |character| match character {
        '&' => Some("&amp;"),
        '<' => Some("&lt;"),
        '>' => Some("&gt;"),
        '"' => Some("&quot;"),
        '\'' => Some("&apos;"),
        _ => None,
    })
}

fn write_escaped_with<W, F>(writer: &mut W, value: &str, replacement: F) -> Result<()>
where
    W: Write,
    F: Fn(char) -> Option<&'static str>,
{
    let mut cursor = 0usize;
    for (index, character) in value.char_indices() {
        let Some(replacement) = replacement(character) else {
            continue;
        };
        writer.write_all(&value.as_bytes()[cursor..index])?;
        writer.write_all(replacement.as_bytes())?;
        cursor = index + character.len_utf8();
    }
    writer.write_all(&value.as_bytes()[cursor..])?;
    Ok(())
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

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}

fn limit(kind: LimitKind, observed: usize, maximum: usize) -> Error {
    Error::limit(kind, observed, maximum)
}

fn check_limit(kind: LimitKind, value: usize, maximum: usize) -> Result<()> {
    if value == 0 || value > maximum {
        return Err(Error::invalid_limit(kind, value, maximum));
    }
    Ok(())
}

fn xml_error(error: impl fmt::Display) -> Error {
    Error::Xml(error.to_string())
}
