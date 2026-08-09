//! Bounded namespace-aware XML tree and lossless extension-fragment codec.

use super::super::model::Limits;
use super::super::{
    DRAWINGML_NAMESPACE, STRICT_DRAWINGML_NAMESPACE, TASK_PANES_NAMESPACE, WEB_EXTENSION_NAMESPACE,
};
use super::super::{Event, Reader, XmlVersion};
use super::semantic::{enforce_count_with, escape_attr, invalid, limit, parse_bool};
use crate::mce::process_markup_compatibility;
use crate::{Error, Result};
use std::collections::{HashMap, HashSet};
use std::sync::Arc;

pub(in crate::web) fn effective_namespaces(
    scope: &NamespaceScope,
) -> Result<Vec<(&String, &String)>> {
    let mut namespaces = Vec::new();
    namespaces
        .try_reserve(scope.binding_count)
        .map_err(|_| Error::Limit {
            resource: "retained web extension namespace entries",
            max: scope.binding_count,
            actual: scope.binding_count,
        })?;
    let mut seen = HashSet::new();
    seen.try_reserve(scope.binding_count)
        .map_err(|_| Error::Limit {
            resource: "retained web extension namespace entries",
            max: scope.binding_count,
            actual: scope.binding_count,
        })?;
    let mut current = Some(scope);
    while let Some(value) = current {
        for (prefix, namespace) in &value.local {
            if seen.insert(prefix.as_str()) {
                namespaces.push((prefix, namespace));
            }
        }
        current = value.parent.as_deref();
    }
    Ok(namespaces)
}

pub(in crate::web) fn retained_namespace_bytes(
    namespaces: &[(&String, &String)],
    declared_prefixes: &HashSet<String>,
) -> Result<usize> {
    let mut total = 0usize;
    for (prefix, namespace) in namespaces {
        if prefix.as_str() == "xml" || declared_prefixes.contains(prefix.as_str()) {
            continue;
        }
        let head = if prefix.is_empty() {
            " xmlns=\"".len()
        } else {
            " xmlns:"
                .len()
                .checked_add(prefix.len())
                .and_then(|value| value.checked_add("=\"".len()))
                .ok_or(Error::Limit {
                    resource: "retained web extension namespace bytes",
                    max: usize::MAX,
                    actual: usize::MAX,
                })?
        };
        let value = escaped_attr_bytes(namespace)?;
        total = total
            .checked_add(head)
            .and_then(|total| total.checked_add(value))
            .and_then(|total| total.checked_add(1))
            .ok_or(Error::Limit {
                resource: "retained web extension namespace bytes",
                max: usize::MAX,
                actual: usize::MAX,
            })?;
    }
    Ok(total)
}

pub(in crate::web) fn escaped_attr_bytes(value: &str) -> Result<usize> {
    value.chars().try_fold(0usize, |total, character| {
        let bytes = match character {
            '&' => "&amp;".len(),
            '<' => "&lt;".len(),
            '>' => "&gt;".len(),
            '"' => "&quot;".len(),
            '\'' => "&apos;".len(),
            _ => character.len_utf8(),
        };
        total.checked_add(bytes).ok_or(Error::Limit {
            resource: "retained web extension namespace bytes",
            max: usize::MAX,
            actual: usize::MAX,
        })
    })
}

pub(in crate::web) fn canonical_node_xml(node: &Node) -> String {
    pub(in crate::web) fn write_node(out: &mut String, node: &Node) {
        out.push('<');
        out.push_str(&node.local_name);
        out.push_str(" xmlns=\"");
        escape_attr(out, &node.namespace);
        out.push('"');
        for (index, attribute) in node.attributes.iter().enumerate() {
            if attribute.namespace.is_empty() {
                out.push(' ');
                out.push_str(&attribute.local_name);
            } else if attribute.namespace == "http://www.w3.org/XML/1998/namespace" {
                out.push_str(" xml:");
                out.push_str(&attribute.local_name);
            } else {
                out.push_str(" xmlns:n");
                out.push_str(&index.to_string());
                out.push_str("=\"");
                escape_attr(out, &attribute.namespace);
                out.push_str("\" n");
                out.push_str(&index.to_string());
                out.push(':');
                out.push_str(&attribute.local_name);
            }
            out.push_str("=\"");
            escape_attr(out, &attribute.value);
            out.push('"');
        }
        if node.children.is_empty() {
            out.push_str("/>");
            return;
        }
        out.push('>');
        for child in &node.children {
            write_node(out, child);
        }
        out.push_str("</");
        out.push_str(&node.local_name);
        out.push('>');
    }

    let mut out = String::new();
    write_node(&mut out, node);
    out
}

#[derive(Debug)]
pub(in crate::web) struct Attribute {
    pub(in crate::web) namespace: String,
    pub(in crate::web) local_name: String,
    pub(in crate::web) value: String,
}

#[derive(Debug)]
pub(in crate::web) struct Node {
    pub(in crate::web) namespace: String,
    pub(in crate::web) local_name: String,
    pub(in crate::web) attributes: Vec<Attribute>,
    pub(in crate::web) children: Vec<Node>,
    pub(in crate::web) raw_fragment: Option<RawFragment>,
}

#[derive(Debug)]
pub(in crate::web) struct RawFragment {
    pub(in crate::web) start: usize,
    pub(in crate::web) start_tag_end: usize,
    pub(in crate::web) end: usize,
    pub(in crate::web) namespaces: Arc<NamespaceScope>,
    pub(in crate::web) declared_prefixes: HashSet<String>,
}

#[derive(Debug)]
pub(in crate::web) struct NamespaceScope {
    pub(in crate::web) parent: Option<Arc<NamespaceScope>>,
    pub(in crate::web) local: HashMap<String, String>,
    pub(in crate::web) binding_count: usize,
}

impl NamespaceScope {
    pub(in crate::web) fn xml() -> Arc<Self> {
        Arc::new(Self {
            parent: None,
            local: HashMap::from([("xml".into(), "http://www.w3.org/XML/1998/namespace".into())]),
            binding_count: 1,
        })
    }

    pub(in crate::web) fn get(&self, prefix: &str) -> Option<&str> {
        self.local
            .get(prefix)
            .map(String::as_str)
            .or_else(|| self.parent.as_deref().and_then(|parent| parent.get(prefix)))
    }
}

#[derive(Debug)]
pub(in crate::web) struct NodeFrame {
    pub(in crate::web) node: Node,
    pub(in crate::web) namespaces: Arc<NamespaceScope>,
    pub(in crate::web) extension_depth: Option<usize>,
    pub(in crate::web) direct_extension_count: usize,
}

#[derive(Debug, Default)]
pub(in crate::web) struct XmlBuildState {
    pub(in crate::web) root: Option<Node>,
    pub(in crate::web) stack: Vec<NodeFrame>,
    pub(in crate::web) string_bytes: usize,
    pub(in crate::web) nodes: usize,
}

#[derive(Debug)]
pub(in crate::web) struct XmlDocument {
    pub(in crate::web) root: Option<Node>,
    pub(in crate::web) xml: Vec<u8>,
    pub(in crate::web) string_bytes: usize,
}

impl XmlDocument {
    pub(in crate::web) fn root(&self) -> Result<&Node> {
        self.root
            .as_ref()
            .ok_or_else(|| Error::Invalid("missing XML root".into()))
    }

    pub(in crate::web) fn self_contained_fragment(&self, node: &Node) -> Result<String> {
        let fragment = node
            .raw_fragment
            .as_ref()
            .ok_or_else(|| Error::Invalid("XML node has no retained fragment bounds".into()))?;
        if fragment.start > fragment.start_tag_end
            || fragment.start_tag_end > fragment.end
            || fragment.end > self.xml.len()
        {
            return invalid("invalid retained XML fragment bounds".into());
        }
        let raw = &self.xml[fragment.start..fragment.end];
        let start_tag_end = fragment.start_tag_end - fragment.start;
        if start_tag_end == 0 || raw.get(start_tag_end - 1) != Some(&b'>') {
            return invalid("retained XML fragment has an invalid start tag".into());
        }
        let mut insert_at = start_tag_end - 1;
        let mut cursor = insert_at;
        while cursor > 0 && raw[cursor - 1].is_ascii_whitespace() {
            cursor -= 1;
        }
        if cursor > 0 && raw[cursor - 1] == b'/' {
            insert_at = cursor - 1;
        }

        let raw = std::str::from_utf8(raw)
            .map_err(|error| Error::Xml(format!("non-UTF-8 extension fragment: {error}")))?;
        let mut namespaces = effective_namespaces(&fragment.namespaces)?;
        namespaces.sort_unstable_by(|left, right| left.0.cmp(right.0));
        let extra = retained_namespace_bytes(&namespaces, &fragment.declared_prefixes)?;
        let capacity = raw.len().checked_add(extra).ok_or(Error::Limit {
            resource: "retained web extension fragment bytes",
            max: usize::MAX,
            actual: usize::MAX,
        })?;
        let mut out = String::new();
        out.try_reserve(capacity).map_err(|_| Error::Limit {
            resource: "retained web extension fragment bytes",
            max: capacity,
            actual: capacity,
        })?;
        out.push_str(&raw[..insert_at]);
        for (prefix, namespace) in namespaces {
            if prefix == "xml" || fragment.declared_prefixes.contains(prefix) {
                continue;
            }
            if prefix.is_empty() {
                out.push_str(" xmlns=\"");
            } else {
                out.push_str(" xmlns:");
                out.push_str(prefix);
                out.push_str("=\"");
            }
            escape_attr(&mut out, namespace);
            out.push('"');
        }
        out.push_str(&raw[insert_at..]);
        Ok(out)
    }
}

pub(in crate::web) fn parse_mce_xml(
    xml: &[u8],
    namespaces: &[&str],
    limits: &Limits,
) -> Result<XmlDocument> {
    if xml.len() > limits.xml_bytes {
        return limit("web extension XML bytes", limits.xml_bytes, xml.len());
    }
    let mut capabilities = crate::mce::Capabilities::ooxml_baseline();
    for namespace in namespaces {
        capabilities.understand_namespace(*namespace);
    }
    let mce_limits = crate::mce::Limits {
        max_input_bytes: limits.xml_bytes,
        max_output_bytes: limits.xml_bytes,
        max_depth: limits.depth,
        max_namespace_bindings: 4096,
        max_directive_tokens: 4096,
        max_choices_per_alternate: 1024,
    };
    let processed = process_markup_compatibility(xml, &capabilities, &mce_limits)?;
    parse_xml_owned(processed.xml.into_owned(), limits)
}

pub(in crate::web) fn parse_xml(xml: &[u8]) -> Result<XmlDocument> {
    let limits = Limits::standard();
    if xml.len() > limits.xml_bytes {
        return limit("web extension XML bytes", limits.xml_bytes, xml.len());
    }
    parse_xml_owned(xml.to_vec(), &limits)
}

pub(in crate::web) fn parse_xml_owned(xml: Vec<u8>, limits: &Limits) -> Result<XmlDocument> {
    if xml.len() > limits.xml_bytes {
        return limit("web extension XML bytes", limits.xml_bytes, xml.len());
    }
    let mut reader = Reader::from_reader(xml.as_slice());
    reader.config_mut().trim_text(false);
    reader.config_mut().check_comments = true;
    let mut buffer = Vec::new();
    let mut state = XmlBuildState::default();
    let mut xml_version = XmlVersion::Implicit1_0;
    let mut declaration_seen = false;
    let mut content_seen = false;
    loop {
        let event_start = reader.buffer_position() as usize;
        let event = reader.read_event_into(&mut buffer)?;
        let event_end = reader.buffer_position() as usize;
        let declaration_or_eof = matches!(&event, Event::Decl(_) | Event::Eof);
        match event {
            Event::Decl(declaration) => {
                if declaration_seen || content_seen {
                    return invalid("XML declaration must appear once at the beginning".into());
                }
                declaration_seen = true;
                xml_version = declaration.xml_version()?;
                if xml_version == XmlVersion::Explicit1_1 {
                    return invalid("XML 1.1 is not supported for web extension parts".into());
                }
            },
            Event::Start(element) => push_element(
                &reader,
                &element,
                &mut state,
                xml_version,
                ElementEvent {
                    empty: false,
                    start: event_start,
                    end: event_end,
                },
                limits,
            )?,
            Event::Empty(element) => push_element(
                &reader,
                &element,
                &mut state,
                xml_version,
                ElementEvent {
                    empty: true,
                    start: event_start,
                    end: event_end,
                },
                limits,
            )?,
            Event::Eof => break,
            Event::DocType(_) => return invalid("DTD is forbidden in web extension XML".into()),
            Event::Text(text)
                if !extension_text_is_allowed(&state.stack)
                    && !text.as_ref().iter().all(u8::is_ascii_whitespace) =>
            {
                return invalid("text is not permitted in web extension structures".into());
            },
            Event::CData(text)
                if !extension_text_is_allowed(&state.stack)
                    && !text.as_ref().iter().all(u8::is_ascii_whitespace) =>
            {
                return invalid("CDATA is not permitted in web extension structures".into());
            },
            Event::GeneralRef(_) => {
                return invalid(
                    "general entity references are forbidden in web extension XML".into(),
                );
            },
            Event::End(_) if state.stack.is_empty() => {
                return invalid("unexpected XML end tag".into());
            },
            Event::End(_) => {
                let mut frame = state
                    .stack
                    .pop()
                    .ok_or_else(|| Error::Invalid("unexpected XML end tag".into()))?;
                if let Some(fragment) = frame.node.raw_fragment.as_mut() {
                    fragment.end = event_end;
                }
                attach_node(&mut state.root, &mut state.stack, frame.node)?;
            },
            _ => {},
        }
        if !declaration_or_eof {
            content_seen = true;
        }
        buffer.clear();
    }
    if !state.stack.is_empty() {
        return invalid("unclosed XML element".into());
    }
    if state.string_bytes > limits.string_bytes {
        return limit(
            "web extension decoded string bytes",
            limits.string_bytes,
            state.string_bytes,
        );
    }
    drop(reader);
    Ok(XmlDocument {
        root: state.root,
        xml,
        string_bytes: state.string_bytes,
    })
}

#[derive(Debug, Clone, Copy)]
pub(in crate::web) struct ElementEvent {
    pub(in crate::web) empty: bool,
    pub(in crate::web) start: usize,
    pub(in crate::web) end: usize,
}

pub(in crate::web) fn push_element(
    reader: &Reader<&[u8]>,
    element: &quick_xml::events::BytesStart<'_>,
    state: &mut XmlBuildState,
    xml_version: XmlVersion,
    event: ElementEvent,
    limits: &Limits,
) -> Result<()> {
    if state.stack.len() >= limits.depth {
        return limit(
            "web extension XML depth",
            limits.depth,
            state.stack.len().saturating_add(1),
        );
    }
    state.nodes = state
        .nodes
        .checked_add(1)
        .ok_or_else(|| Error::Invalid("web extension node count overflow".into()))?;
    if state.nodes > limits.nodes {
        return limit("web extension XML nodes", limits.nodes, state.nodes);
    }
    let parent_namespaces = state
        .stack
        .last()
        .map_or_else(NamespaceScope::xml, |frame| Arc::clone(&frame.namespaces));
    let mut local_namespaces = HashMap::new();
    let mut raw_attributes = Vec::new();
    let mut declared_prefixes = HashSet::new();
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        let name = std::str::from_utf8(attribute.key.as_ref())
            .map_err(|error| Error::Xml(error.to_string()))?
            .to_owned();
        let value = attribute
            .decoded_and_normalized_value(xml_version, reader.decoder())
            .map_err(|error| Error::Xml(error.to_string()))?
            .into_owned();
        state.string_bytes = state
            .string_bytes
            .checked_add(name.len().saturating_add(value.len()))
            .ok_or(Error::Limit {
                resource: "web extension decoded string bytes",
                max: limits.string_bytes,
                actual: usize::MAX,
            })?;
        if state.string_bytes > limits.string_bytes {
            return limit(
                "web extension decoded string bytes",
                limits.string_bytes,
                state.string_bytes,
            );
        }
        if name == "xmlns" {
            if !declared_prefixes.insert(String::new()) {
                return invalid("duplicate default namespace declaration".into());
            }
            local_namespaces.insert(String::new(), value);
        } else if let Some(prefix) = name.strip_prefix("xmlns:") {
            if prefix == "xmlns"
                || (prefix == "xml" && value != "http://www.w3.org/XML/1998/namespace")
                || value.is_empty()
            {
                return invalid(format!(
                    "invalid namespace declaration for prefix '{prefix}'"
                ));
            }
            if !declared_prefixes.insert(prefix.to_owned()) {
                return invalid(format!(
                    "duplicate namespace declaration for prefix '{prefix}'"
                ));
            }
            local_namespaces.insert(prefix.to_owned(), value);
        } else {
            raw_attributes.push((name, value));
        }
    }
    let new_bindings = local_namespaces
        .keys()
        .filter(|prefix| parent_namespaces.get(prefix).is_none())
        .count();
    let binding_count = parent_namespaces
        .binding_count
        .checked_add(new_bindings)
        .ok_or(Error::Limit {
            resource: "web extension XML namespace bindings",
            max: 4096,
            actual: usize::MAX,
        })?;
    if binding_count > 4096 {
        return invalid("web extension XML namespace bindings exceed 4096".into());
    }
    let namespaces = if local_namespaces.is_empty() {
        parent_namespaces
    } else {
        Arc::new(NamespaceScope {
            parent: Some(parent_namespaces),
            local: local_namespaces,
            binding_count,
        })
    };
    let element_name = element.name();
    let raw_name = std::str::from_utf8(element_name.as_ref())
        .map_err(|error| Error::Xml(error.to_string()))?;
    let (prefix, local_name) = split_qname(raw_name);
    let namespace = if prefix.is_empty() {
        namespaces.get(prefix).unwrap_or_default().to_owned()
    } else {
        namespaces
            .get(prefix)
            .map(str::to_owned)
            .ok_or_else(|| Error::Invalid(format!("unbound XML namespace prefix '{prefix}'")))?
    };
    state.string_bytes = state
        .string_bytes
        .checked_add(namespace.len().saturating_add(local_name.len()))
        .ok_or(Error::Limit {
            resource: "web extension decoded string bytes",
            max: limits.string_bytes,
            actual: usize::MAX,
        })?;
    if state.string_bytes > limits.string_bytes {
        return limit(
            "web extension decoded string bytes",
            limits.string_bytes,
            state.string_bytes,
        );
    }
    let mut attributes = Vec::with_capacity(raw_attributes.len());
    let mut seen = HashSet::new();
    for (raw_name, value) in raw_attributes {
        let (prefix, local_name) = split_qname(&raw_name);
        let namespace = if prefix.is_empty() {
            String::new()
        } else {
            namespaces
                .get(prefix)
                .map(str::to_owned)
                .ok_or_else(|| Error::Invalid(format!("unbound attribute prefix '{prefix}'")))?
        };
        if !seen.insert((namespace.clone(), local_name.to_owned())) {
            return invalid(format!("duplicate attribute {{{namespace}}}{local_name}"));
        }
        attributes.push(Attribute {
            namespace,
            local_name: local_name.to_owned(),
            value,
        });
    }
    let capture_fragment = should_capture_extension_list(
        state.stack.last().map(|frame| &frame.node),
        &namespace,
        local_name,
    );
    let raw_fragment = if capture_fragment {
        let inherited = effective_namespaces(&namespaces)?;
        let retained_bytes = declared_prefixes.iter().try_fold(
            retained_namespace_bytes(&inherited, &declared_prefixes)?,
            |total, prefix| {
                total.checked_add(prefix.len()).ok_or(Error::Limit {
                    resource: "web extension decoded string bytes",
                    max: limits.string_bytes,
                    actual: usize::MAX,
                })
            },
        )?;
        state.string_bytes =
            state
                .string_bytes
                .checked_add(retained_bytes)
                .ok_or(Error::Limit {
                    resource: "web extension decoded string bytes",
                    max: limits.string_bytes,
                    actual: usize::MAX,
                })?;
        if state.string_bytes > limits.string_bytes {
            return limit(
                "web extension decoded string bytes",
                limits.string_bytes,
                state.string_bytes,
            );
        }
        Some(RawFragment {
            start: event.start,
            start_tag_end: event.end,
            end: if event.empty { event.end } else { 0 },
            namespaces: Arc::clone(&namespaces),
            declared_prefixes,
        })
    } else {
        None
    };
    let node = Node {
        namespace,
        local_name: local_name.to_owned(),
        attributes,
        children: Vec::new(),
        raw_fragment,
    };
    if state
        .stack
        .last()
        .is_some_and(|frame| frame.extension_depth == Some(0))
    {
        let parent = state
            .stack
            .last_mut()
            .ok_or_else(|| Error::Invalid("extension-list child has no parent element".into()))?;
        let expected_namespace = if parent.node.namespace == STRICT_DRAWINGML_NAMESPACE {
            STRICT_DRAWINGML_NAMESPACE
        } else {
            DRAWINGML_NAMESPACE
        };
        require_name(&node, expected_namespace, "ext")?;
        reject_unknown_attributes(&node, &[("", "uri")])?;
        required_attr(&node, "", "uri")?;
        parent.direct_extension_count = parent
            .direct_extension_count
            .checked_add(1)
            .ok_or_else(|| Error::Invalid("extLst count overflow".into()))?;
        enforce_count_with("OfficeArt extension", parent.direct_extension_count, limits)?;
    }
    let extension_depth = if capture_fragment {
        Some(0)
    } else {
        state
            .stack
            .last()
            .and_then(|frame| frame.extension_depth)
            .map(|depth| depth + 1)
    };
    if event.empty {
        attach_node(&mut state.root, &mut state.stack, node)?;
    } else {
        state.stack.push(NodeFrame {
            node,
            namespaces,
            extension_depth,
            direct_extension_count: 0,
        });
    }
    Ok(())
}

pub(in crate::web) fn attach_node(
    root: &mut Option<Node>,
    stack: &mut [NodeFrame],
    node: Node,
) -> Result<()> {
    if let Some(parent) = stack.last_mut() {
        if parent.extension_depth.is_none() {
            parent.node.children.push(node);
        }
    } else if root.replace(node).is_some() {
        return invalid("multiple XML root elements".into());
    }
    Ok(())
}

pub(in crate::web) fn should_capture_extension_list(
    parent: Option<&Node>,
    namespace: &str,
    local_name: &str,
) -> bool {
    if local_name != "extLst" {
        return false;
    }
    let allowed_namespace = matches!(
        namespace,
        WEB_EXTENSION_NAMESPACE
            | TASK_PANES_NAMESPACE
            | DRAWINGML_NAMESPACE
            | STRICT_DRAWINGML_NAMESPACE
    );
    if !allowed_namespace {
        return false;
    }
    let Some(parent) = parent else {
        return true;
    };
    matches!(
        (
            parent.namespace.as_str(),
            parent.local_name.as_str(),
            namespace
        ),
        (
            WEB_EXTENSION_NAMESPACE,
            "webextension" | "reference" | "binding",
            WEB_EXTENSION_NAMESPACE
        ) | (
            WEB_EXTENSION_NAMESPACE,
            "snapshot",
            DRAWINGML_NAMESPACE | STRICT_DRAWINGML_NAMESPACE
        ) | (TASK_PANES_NAMESPACE, "taskpane", TASK_PANES_NAMESPACE)
    )
}

pub(in crate::web) fn extension_text_is_allowed(stack: &[NodeFrame]) -> bool {
    stack
        .last()
        .and_then(|frame| frame.extension_depth)
        .is_some_and(|depth| depth >= 2)
}

pub(in crate::web) fn split_qname(name: &str) -> (&str, &str) {
    name.split_once(':').unwrap_or(("", name))
}

pub(in crate::web) fn element_children(node: &Node) -> Vec<&Node> {
    node.children.iter().collect()
}

pub(in crate::web) fn require_name(node: &Node, namespace: &str, local_name: &str) -> Result<()> {
    if node.namespace == namespace && node.local_name == local_name {
        Ok(())
    } else {
        invalid(format!(
            "expected {{{namespace}}}{local_name}, got {{{}}}{}",
            node.namespace, node.local_name
        ))
    }
}

pub(in crate::web) fn attr<'a>(
    node: &'a Node,
    namespace: &str,
    local_name: &str,
) -> Option<&'a str> {
    node.attributes
        .iter()
        .find(|attribute| attribute.namespace == namespace && attribute.local_name == local_name)
        .map(|attribute| attribute.value.as_str())
}

pub(in crate::web) fn required_attr<'a>(
    node: &'a Node,
    namespace: &str,
    local_name: &str,
) -> Result<&'a str> {
    attr(node, namespace, local_name).ok_or_else(|| {
        Error::Invalid(format!(
            "{} requires attribute {{{namespace}}}{local_name}",
            node.local_name
        ))
    })
}

pub(in crate::web) fn is_drawingml_namespace(namespace: &str) -> bool {
    matches!(namespace, DRAWINGML_NAMESPACE | STRICT_DRAWINGML_NAMESPACE)
}

pub(in crate::web) fn optional_bool_attr(
    node: &Node,
    namespace: &str,
    local_name: &str,
) -> Result<Option<bool>> {
    attr(node, namespace, local_name)
        .map(parse_bool)
        .transpose()
}

pub(in crate::web) fn reject_unknown_attributes(
    node: &Node,
    allowed: &[(&str, &str)],
) -> Result<()> {
    for attribute in &node.attributes {
        if !allowed.iter().any(|(namespace, local_name)| {
            attribute.namespace == *namespace && attribute.local_name == *local_name
        }) {
            return invalid(format!(
                "unexpected attribute {{{}}}{} on {}",
                attribute.namespace, attribute.local_name, node.local_name
            ));
        }
    }
    Ok(())
}

pub(in crate::web) fn is_next(
    children: &[&Node],
    position: usize,
    namespace: &str,
    local_name: &str,
) -> bool {
    children
        .get(position)
        .is_some_and(|child| child.namespace == namespace && child.local_name == local_name)
}

pub(in crate::web) fn next_required<'a>(
    children: &[&'a Node],
    position: &mut usize,
    namespace: &str,
    local_name: &str,
) -> Result<&'a Node> {
    if !is_next(children, *position, namespace, local_name) {
        return invalid(format!("missing or misplaced {local_name}"));
    }
    let node = children[*position];
    *position += 1;
    Ok(node)
}

pub(in crate::web) fn ensure_consumed(
    children: &[&Node],
    position: usize,
    parent: &str,
) -> Result<()> {
    if position == children.len() {
        Ok(())
    } else {
        invalid(format!(
            "unexpected child {} in {parent}",
            children[position].local_name
        ))
    }
}
