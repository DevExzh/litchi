//! Format-agnostic SmartArt/diagram model and `DrawingML` part generators.
//!
//! `SmartArt` graphics are represented as diagrams in OOXML. This module holds
//! the semantic model (`SmartArt`, `DiagramNode`, `DiagramType`), the builder,
//! and the generators for the five diagram parts (data, layout, quick style,
//! colors, pre-rendered drawing). The model is shared by the format modules
//! (`docx`, `pptx`); format-specific anchoring lives in those modules.

use crate::diagram::Limits;
use crate::{Error, LimitKind, Result};
use litchi_core::xml::escape_xml;
use litchi_ooxml_common::xml::attributes::BytesStartExt as _;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use std::collections::HashMap;

const DGM_NAMESPACE_BYTES: &[u8] = b"http://schemas.openxmlformats.org/drawingml/2006/diagram";
const DGM_NAMESPACE_STRICT_BYTES: &[u8] = b"http://purl.oclc.org/ooxml/drawingml/diagram";
const DML_NAMESPACE_BYTES: &[u8] = b"http://schemas.openxmlformats.org/drawingml/2006/main";
const DML_NAMESPACE_STRICT_BYTES: &[u8] = b"http://purl.oclc.org/ooxml/drawingml/main";

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum SmartArtNamespace {
    Diagram,
    Drawing,
    Other,
}

fn namespace_kind(resolved: &ResolveResult<'_>) -> SmartArtNamespace {
    if is_dgm_namespace(resolved) {
        SmartArtNamespace::Diagram
    } else if is_dml_namespace(resolved) {
        SmartArtNamespace::Drawing
    } else {
        SmartArtNamespace::Other
    }
}

/// `SmartArt` diagram type.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DiagramType {
    /// List diagram
    List,
    /// Process diagram
    Process,
    /// Cycle diagram
    Cycle,
    /// Hierarchy diagram
    Hierarchy,
    /// Relationship diagram
    Relationship,
    /// Matrix diagram
    Matrix,
    /// Pyramid diagram
    Pyramid,
    /// Picture diagram
    Picture,
    /// Unknown diagram type
    Unknown,
}

impl DiagramType {
    /// Parse diagram type from layout type URI.
    #[must_use]
    pub fn from_layout_uri(uri: &str) -> Self {
        let uri_lower = uri.to_lowercase();
        if uri_lower.contains("list") {
            DiagramType::List
        } else if uri_lower.contains("process") {
            DiagramType::Process
        } else if uri_lower.contains("cycle") {
            DiagramType::Cycle
        } else if uri_lower.contains("hierarchy") || uri_lower.contains("orgchart") {
            DiagramType::Hierarchy
        } else if uri_lower.contains("relationship") || uri_lower.contains("venn") {
            DiagramType::Relationship
        } else if uri_lower.contains("matrix") {
            DiagramType::Matrix
        } else if uri_lower.contains("pyramid") {
            DiagramType::Pyramid
        } else if uri_lower.contains("picture") {
            DiagramType::Picture
        } else {
            DiagramType::Unknown
        }
    }
}

/// A `SmartArt` diagram node/item.
#[derive(Debug, Clone)]
pub struct DiagramNode {
    /// Node text content
    pub text: String,
    /// Child nodes
    pub children: Vec<DiagramNode>,
    /// Node depth level (0 = root)
    pub depth: u32,
}

impl DiagramNode {
    /// Create a new diagram node.
    pub fn new(text: impl Into<String>) -> Self {
        Self {
            text: text.into(),
            children: Vec::new(),
            depth: 0,
        }
    }

    /// Add a child node.
    pub fn add_child(&mut self, child: DiagramNode) {
        self.children.push(child);
    }

    /// Get all text from this node and its children.
    #[must_use]
    pub fn all_text(&self) -> String {
        let mut result = self.text.clone();
        for child in &self.children {
            if !result.is_empty() && !child.text.is_empty() {
                result.push('\n');
            }
            result.push_str(&child.all_text());
        }
        result
    }
}

/// `SmartArt` diagram information.
#[derive(Debug, Clone)]
pub struct SmartArt {
    /// Diagram type
    pub diagram_type: DiagramType,
    /// Root nodes of the diagram
    pub nodes: Vec<DiagramNode>,
    /// Layout name/description if available
    pub layout_name: Option<String>,
    /// Unique ID of the diagram
    pub id: Option<String>,
}

impl SmartArt {
    /// Create a new `SmartArt` diagram.
    #[must_use]
    pub fn new(diagram_type: DiagramType) -> Self {
        Self {
            diagram_type,
            nodes: Vec::new(),
            layout_name: None,
            id: None,
        }
    }

    /// Add a root node.
    pub fn add_node(&mut self, node: DiagramNode) {
        self.nodes.push(node);
    }

    /// Get all text content from the diagram.
    #[must_use]
    pub fn text(&self) -> String {
        self.nodes
            .iter()
            .map(DiagramNode::all_text)
            .collect::<Vec<_>>()
            .join("\n")
    }

    /// Get the number of root nodes.
    #[must_use]
    pub fn node_count(&self) -> usize {
        self.nodes.len()
    }

    /// Parse `SmartArt` data XML (dgm:data).
    /// # Errors
    ///
    /// Returns an error when input violates DrawingML constraints, exceeds a configured
    /// bound, or an underlying XML, MCE, I/O, or formatting operation fails.
    pub fn parse_data_xml(xml: &str, limits: &Limits) -> Result<Vec<DiagramNode>> {
        limits.validate()?;
        let max_input = limits.input_bytes_value().min(16 * 1024 * 1024);
        if xml.len() > max_input {
            return Err(Error::limit(
                LimitKind::DiagramDataModelBytes,
                xml.len(),
                max_input,
            ));
        }
        let xml = litchi_ooxml_common::mce::process_str(xml)?;
        if xml.len() > max_input {
            return Err(Error::limit(
                LimitKind::DiagramDataModelBytes,
                xml.len(),
                max_input,
            ));
        }
        let mut reader = NsReader::from_reader(xml.as_ref().as_bytes());
        reader.config_mut().trim_text(false);

        let mut buffer = Vec::new();
        let mut points = Vec::new();
        let mut connections = Vec::new();
        let mut current_point = None;
        let mut current_text = String::new();
        let mut in_text_body = false;
        let mut in_text_run = false;
        let mut text_bytes = 0_usize;
        let mut depth = 0_usize;
        let mut xml_nodes = 0_usize;
        let mut stack: Vec<(SmartArtNamespace, Vec<u8>)> = Vec::new();
        let mut root_seen = false;
        let mut root_closed = false;
        let max_depth = limits.xml_depth_value().min(128);
        let max_xml_nodes = limits.xml_nodes_value().min(200_000);
        let max_points = limits.points_value().min(100_000);
        let max_connections = limits.connections_value().min(100_000);
        let max_nodes = limits.nodes_value().min(200_000);
        let max_node_depth = limits.node_depth_value().min(128);
        let max_text = limits.data_text_bytes_value().min(1024 * 1024);
        let mut point_list_seen = false;
        let mut connection_list_seen = false;

        loop {
            buffer.clear();
            let (resolved, event) = reader
                .read_resolved_event_into(&mut buffer)
                .map_err(|error| Error::Xml(error.to_string()))?;
            match event {
                Event::Start(ref element) => {
                    xml_nodes = xml_nodes.checked_add(1).ok_or_else(|| {
                        Error::Invalid("SmartArt XML node count overflows".to_owned())
                    })?;
                    if xml_nodes > max_xml_nodes {
                        return Err(Error::limit(LimitKind::XmlNodes, xml_nodes, max_xml_nodes));
                    }
                    if root_closed {
                        return Err(Error::Invalid(
                            "trailing markup follows the SmartArt data model".to_owned(),
                        ));
                    }
                    let event_depth = depth
                        .checked_add(1)
                        .ok_or_else(|| Error::Invalid("SmartArt XML depth overflows".to_owned()))?;
                    depth = event_depth;
                    if depth > max_depth {
                        return Err(Error::limit(LimitKind::XmlDepth, depth, max_depth));
                    }
                    let namespace = namespace_kind(&resolved);
                    let dgm = namespace == SmartArtNamespace::Diagram;
                    let dml = namespace == SmartArtNamespace::Drawing;
                    let local = element.local_name().as_ref().to_vec();
                    if depth == 1 {
                        if root_seen {
                            return Err(Error::Invalid(
                                "the SmartArt data model has multiple roots".to_owned(),
                            ));
                        }
                        if !dgm || local != b"dataModel" {
                            return Err(Error::Invalid(
                                "the SmartArt data model root is invalid".to_owned(),
                            ));
                        }
                        root_seen = true;
                    } else if dgm && local == b"dataModel" {
                        return Err(Error::Invalid(
                            "the SmartArt data model has a nested root".to_owned(),
                        ));
                    }
                    let parent = stack.last().map(|(kind, name)| (*kind, name.as_slice()));
                    if dgm && local == b"ptLst" {
                        if parent != Some((SmartArtNamespace::Diagram, b"dataModel".as_slice()))
                            || point_list_seen
                        {
                            return Err(Error::Invalid(
                                "the SmartArt point list is misplaced or duplicated".to_owned(),
                            ));
                        }
                        point_list_seen = true;
                    } else if dgm && local == b"cxnLst" {
                        if parent != Some((SmartArtNamespace::Diagram, b"dataModel".as_slice()))
                            || !point_list_seen
                            || connection_list_seen
                        {
                            return Err(Error::Invalid(
                                "the SmartArt connection list is misplaced or duplicated"
                                    .to_owned(),
                            ));
                        }
                        connection_list_seen = true;
                    } else if dgm && local == b"pt" {
                        if parent != Some((SmartArtNamespace::Diagram, b"ptLst".as_slice())) {
                            return Err(Error::Invalid(
                                "the SmartArt point is outside its point list".to_owned(),
                            ));
                        }
                        if current_point.is_some() {
                            return Err(Error::Invalid(
                                "nested SmartArt points are invalid".to_owned(),
                            ));
                        }
                        if points.len() >= max_points {
                            return Err(Error::limit(
                                LimitKind::DiagramPoints,
                                points.len() + 1,
                                max_points,
                            ));
                        }
                        let id = attribute_value(element, b"modelId")?
                            .unwrap_or_else(|| format!("__point_{}", points.len()));
                        let kind =
                            attribute_value(element, b"type")?.unwrap_or_else(|| "node".to_owned());
                        points.push(RawPoint {
                            id,
                            kind,
                            text: None,
                        });
                        current_point = Some(points.len() - 1);
                    } else if dgm && local == b"cxn" {
                        if parent != Some((SmartArtNamespace::Diagram, b"cxnLst".as_slice())) {
                            return Err(Error::Invalid(
                                "the SmartArt connection is outside its connection list".to_owned(),
                            ));
                        }
                        if connections.len() >= max_connections {
                            return Err(Error::limit(
                                LimitKind::DiagramConnections,
                                connections.len() + 1,
                                max_connections,
                            ));
                        }
                        let source_order = attribute_value(element, b"srcOrd")?
                            .map(|value| {
                                value.parse::<usize>().map_err(|_| {
                                    Error::Invalid(
                                        "the SmartArt connection source order is invalid"
                                            .to_owned(),
                                    )
                                })
                            })
                            .transpose()?
                            .unwrap_or(connections.len());
                        connections.push(RawConnection {
                            kind: attribute_value(element, b"type")?
                                .unwrap_or_else(|| "parOf".to_owned()),
                            source: attribute_value(element, b"srcId")?.unwrap_or_default(),
                            destination: attribute_value(element, b"destId")?.unwrap_or_default(),
                            source_order,
                        });
                    } else if dgm
                        && local == b"t"
                        && parent == Some((SmartArtNamespace::Diagram, b"pt".as_slice()))
                        && current_point.is_some()
                    {
                        in_text_body = true;
                        in_text_run = false;
                        current_text.clear();
                        text_bytes = 0;
                    } else if dml && local == b"t" && in_text_body {
                        in_text_run = true;
                    }
                    stack.push((namespace, local));
                },
                Event::Empty(ref element) => {
                    xml_nodes = xml_nodes.checked_add(1).ok_or_else(|| {
                        Error::Invalid("SmartArt XML node count overflows".to_owned())
                    })?;
                    if xml_nodes > max_xml_nodes {
                        return Err(Error::limit(LimitKind::XmlNodes, xml_nodes, max_xml_nodes));
                    }
                    if root_closed {
                        return Err(Error::Invalid(
                            "trailing markup follows the SmartArt data model".to_owned(),
                        ));
                    }
                    let child_depth = depth
                        .checked_add(1)
                        .ok_or_else(|| Error::Invalid("SmartArt XML depth overflows".to_owned()))?;
                    if child_depth > max_depth {
                        return Err(Error::limit(LimitKind::XmlDepth, child_depth, max_depth));
                    }
                    let namespace = namespace_kind(&resolved);
                    let dgm = namespace == SmartArtNamespace::Diagram;
                    let local = element.local_name().as_ref().to_vec();
                    if depth == 0 {
                        if root_seen {
                            return Err(Error::Invalid(
                                "the SmartArt data model has multiple roots".to_owned(),
                            ));
                        }
                        if !dgm || local != b"dataModel" {
                            return Err(Error::Invalid(
                                "the SmartArt data model root is invalid".to_owned(),
                            ));
                        }
                        return Err(Error::Invalid(
                            "the SmartArt data model root is empty".to_owned(),
                        ));
                    }
                    let parent = stack.last().map(|(kind, name)| (*kind, name.as_slice()));
                    if dgm && local == b"dataModel" {
                        return Err(Error::Invalid(
                            "the SmartArt data model has a nested root".to_owned(),
                        ));
                    } else if dgm && local == b"ptLst" {
                        if parent != Some((SmartArtNamespace::Diagram, b"dataModel".as_slice()))
                            || point_list_seen
                        {
                            return Err(Error::Invalid(
                                "the SmartArt point list is misplaced or duplicated".to_owned(),
                            ));
                        }
                        point_list_seen = true;
                    } else if dgm && local == b"cxnLst" {
                        if parent != Some((SmartArtNamespace::Diagram, b"dataModel".as_slice()))
                            || !point_list_seen
                            || connection_list_seen
                        {
                            return Err(Error::Invalid(
                                "the SmartArt connection list is misplaced or duplicated"
                                    .to_owned(),
                            ));
                        }
                        connection_list_seen = true;
                    } else if dgm && local == b"pt" {
                        if parent != Some((SmartArtNamespace::Diagram, b"ptLst".as_slice())) {
                            return Err(Error::Invalid(
                                "the SmartArt point is outside its point list".to_owned(),
                            ));
                        }
                        if points.len() >= max_points {
                            return Err(Error::limit(
                                LimitKind::DiagramPoints,
                                points.len() + 1,
                                max_points,
                            ));
                        }
                        let id = attribute_value(element, b"modelId")?
                            .unwrap_or_else(|| format!("__point_{}", points.len()));
                        let kind =
                            attribute_value(element, b"type")?.unwrap_or_else(|| "node".to_owned());
                        points.push(RawPoint {
                            id,
                            kind,
                            text: None,
                        });
                    } else if dgm && local == b"cxn" {
                        if parent != Some((SmartArtNamespace::Diagram, b"cxnLst".as_slice())) {
                            return Err(Error::Invalid(
                                "the SmartArt connection is outside its connection list".to_owned(),
                            ));
                        }
                        if connections.len() >= max_connections {
                            return Err(Error::limit(
                                LimitKind::DiagramConnections,
                                connections.len() + 1,
                                max_connections,
                            ));
                        }
                        let source_order = attribute_value(element, b"srcOrd")?
                            .map(|value| {
                                value.parse::<usize>().map_err(|_| {
                                    Error::Invalid(
                                        "the SmartArt connection source order is invalid"
                                            .to_owned(),
                                    )
                                })
                            })
                            .transpose()?
                            .unwrap_or(connections.len());
                        connections.push(RawConnection {
                            kind: attribute_value(element, b"type")?
                                .unwrap_or_else(|| "parOf".to_owned()),
                            source: attribute_value(element, b"srcId")?.unwrap_or_default(),
                            destination: attribute_value(element, b"destId")?.unwrap_or_default(),
                            source_order,
                        });
                    }
                },
                Event::End(ref element) => {
                    if depth == 0 || stack.is_empty() {
                        return Err(Error::Invalid(
                            "unexpected SmartArt data-model closing element".to_owned(),
                        ));
                    }
                    let namespace = namespace_kind(&resolved);
                    let dgm = namespace == SmartArtNamespace::Diagram;
                    let local = element.local_name();
                    let Some((expected_namespace, expected_local)) = stack.last() else {
                        return Err(Error::Invalid(
                            "unexpected SmartArt data-model closing element".to_owned(),
                        ));
                    };
                    if *expected_namespace != namespace
                        || expected_local.as_slice() != local.as_ref()
                    {
                        return Err(Error::Invalid(
                            "SmartArt data-model element names are unbalanced".to_owned(),
                        ));
                    }
                    if depth == 1 && (!dgm || local.as_ref() != b"dataModel") {
                        return Err(Error::Invalid(
                            "the SmartArt data model root is not closed".to_owned(),
                        ));
                    }
                    if dgm && local.as_ref() == b"pt" {
                        current_point = None;
                    } else if dgm && local.as_ref() == b"t" && in_text_body {
                        if let Some(index) = current_point {
                            points[index].text = Some(current_text.trim().to_owned());
                        }
                        in_text_body = false;
                        in_text_run = false;
                    } else if namespace == SmartArtNamespace::Drawing
                        && local.as_ref() == b"t"
                        && in_text_run
                    {
                        in_text_run = false;
                    }
                    stack.pop();
                    depth -= 1;
                    if depth == 0 {
                        root_closed = true;
                    }
                },
                Event::Text(text) => {
                    if in_text_body {
                        let decoded = text
                            .decode()
                            .map_err(|error| Error::Xml(error.to_string()))?;
                        let decoded = quick_xml::escape::unescape(&decoded)
                            .map_err(|error| Error::Xml(error.to_string()))?;
                        append_text(
                            &mut current_text,
                            &mut text_bytes,
                            decoded.as_ref(),
                            max_text,
                            LimitKind::DiagramDataModelTextBytes,
                        )?;
                    } else if text.as_ref().iter().any(|byte| !byte.is_ascii_whitespace()) {
                        return Err(Error::Invalid(
                            "non-whitespace text is outside the SmartArt data model".to_owned(),
                        ));
                    }
                },
                Event::CData(data) if in_text_body => {
                    let text = String::from_utf8_lossy(data.as_ref());
                    append_text(
                        &mut current_text,
                        &mut text_bytes,
                        &text,
                        max_text,
                        LimitKind::DiagramDataModelTextBytes,
                    )?;
                },
                Event::CData(_) => {
                    return Err(Error::Invalid("CDATA is outside SmartArt text".to_owned()));
                },
                Event::GeneralRef(reference) if in_text_body => {
                    let name = reference
                        .decode()
                        .map_err(|error| Error::Xml(error.to_string()))?;
                    let entity = format!("&{name};");
                    let text = quick_xml::escape::unescape(&entity)
                        .map_err(|error| Error::Xml(error.to_string()))?;
                    append_text(
                        &mut current_text,
                        &mut text_bytes,
                        &text,
                        max_text,
                        LimitKind::DiagramDataModelTextBytes,
                    )?;
                },
                Event::GeneralRef(_) => {
                    return Err(Error::Invalid(
                        "entity reference is outside the SmartArt data model".to_owned(),
                    ));
                },
                Event::Eof => break,
                Event::DocType(_) => return Err(Error::Invalid("DTDs are rejected".to_owned())),
                Event::Comment(_) | Event::Decl(_) | Event::PI(_) => {},
            }
        }

        if !root_seen
            || !root_closed
            || !point_list_seen
            || depth != 0
            || !stack.is_empty()
            || in_text_body
        {
            return Err(Error::Invalid("unterminated SmartArt data XML".to_owned()));
        }
        let mut point_ids = HashMap::with_capacity(points.len());
        for point in &points {
            if point_ids.insert(point.id.clone(), ()).is_some() {
                return Err(Error::Invalid(
                    "duplicate SmartArt point identifier".to_owned(),
                ));
            }
        }
        let document_id = points
            .iter()
            .find(|point| point.kind == "doc")
            .map(|point| point.id.clone());
        let mut content = Vec::new();
        let mut index_by_id = HashMap::new();
        for point in points.into_iter().filter(|point| {
            !matches!(
                point.kind.as_str(),
                "doc" | "pres" | "parTrans" | "sibTrans"
            )
        }) {
            if content.len() >= max_nodes {
                return Err(Error::limit(
                    LimitKind::DiagramNodes,
                    content.len() + 1,
                    max_nodes,
                ));
            }
            if index_by_id
                .insert(point.id.clone(), content.len())
                .is_some()
            {
                return Err(Error::Invalid(
                    "duplicate SmartArt point identifier".to_owned(),
                ));
            }
            content.push(point);
        }
        let mut parents = vec![None; content.len()];
        let mut parent_seen = vec![false; content.len()];
        let mut parent_orders = vec![None; content.len()];
        for connection in connections
            .into_iter()
            .filter(|connection| connection.kind == "parOf")
        {
            let Some(&child) = index_by_id.get(&connection.destination) else {
                return Err(Error::Invalid(
                    "a SmartArt parent connection has a dangling destination".to_owned(),
                ));
            };
            let parent = if document_id.as_deref() == Some(connection.source.as_str()) {
                None
            } else {
                Some(*index_by_id.get(&connection.source).ok_or_else(|| {
                    Error::Invalid("a SmartArt parent connection has a dangling source".to_owned())
                })?)
            };
            if parent_seen[child] {
                return Err(Error::Invalid(
                    "a SmartArt point has multiple parents".to_owned(),
                ));
            }
            parent_seen[child] = true;
            parents[child] = parent;
            parent_orders[child] = Some(connection.source_order);
        }
        let mut children = vec![Vec::new(); content.len()];
        let mut roots = Vec::new();
        for (index, parent) in parents.iter().copied().enumerate() {
            if let Some(parent) = parent {
                // The source order on a parOf connection is the semantic
                // sibling order.  Preserve the XML order only as a stable
                // tie-breaker for malformed-but-readable inputs that repeat
                // an order value.
                children[parent].push((parent_orders[index].unwrap_or(usize::MAX), index));
            } else {
                roots.push(index);
            }
        }
        roots.sort_unstable_by_key(|index| (parent_orders[*index].unwrap_or(usize::MAX), *index));
        for siblings in &mut children {
            siblings.sort_unstable_by_key(|(order, index)| (*order, *index));
        }
        let mut visiting = vec![false; content.len()];
        let mut visited = vec![false; content.len()];
        fn build(
            index: usize,
            depth: usize,
            content: &[RawPoint],
            children: &[Vec<(usize, usize)>],
            visiting: &mut [bool],
            visited: &mut [bool],
            max_depth: usize,
        ) -> Result<DiagramNode> {
            if depth > max_depth {
                return Err(Error::limit(LimitKind::DiagramNodeDepth, depth, max_depth));
            }
            if visiting[index] {
                return Err(Error::Invalid(
                    "the SmartArt point graph contains a cycle".to_owned(),
                ));
            }
            visiting[index] = true;
            let mut node = DiagramNode::new(content[index].text.as_deref().unwrap_or_default());
            node.depth = u32::try_from(depth)
                .map_err(|_| Error::Invalid("SmartArt node depth overflows".to_owned()))?;
            for &(_, child) in &children[index] {
                let child_depth = depth
                    .checked_add(1)
                    .ok_or_else(|| Error::Invalid("SmartArt node depth overflows".to_owned()))?;
                node.children.push(build(
                    child,
                    child_depth,
                    content,
                    children,
                    visiting,
                    visited,
                    max_depth,
                )?);
            }
            visiting[index] = false;
            visited[index] = true;
            Ok(node)
        }
        let mut nodes = Vec::with_capacity(roots.len());
        for index in roots {
            nodes.push(build(
                index,
                0,
                &content,
                &children,
                &mut visiting,
                &mut visited,
                max_node_depth,
            )?);
        }
        if visited.iter().any(|visited| !visited) {
            return Err(Error::Invalid(
                "the SmartArt point graph contains an unreachable cycle".to_owned(),
            ));
        }
        Ok(nodes)
    }
}

#[derive(Debug)]
struct RawPoint {
    id: String,
    kind: String,
    text: Option<String>,
}

#[derive(Debug)]
struct RawConnection {
    kind: String,
    source: String,
    destination: String,
    source_order: usize,
}

fn is_dgm_namespace(resolved: &ResolveResult<'_>) -> bool {
    matches!(
        resolved,
        ResolveResult::Bound(Namespace(namespace))
            if *namespace == DGM_NAMESPACE_BYTES || *namespace == DGM_NAMESPACE_STRICT_BYTES
    )
}

fn is_dml_namespace(resolved: &ResolveResult<'_>) -> bool {
    matches!(
        resolved,
        ResolveResult::Bound(Namespace(namespace))
            if *namespace == DML_NAMESPACE_BYTES || *namespace == DML_NAMESPACE_STRICT_BYTES
    )
}

fn attribute_value(element: &BytesStart<'_>, name: &[u8]) -> Result<Option<String>> {
    for attribute in element.checked_attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        if attribute.key.local_name().as_ref() == name {
            let value = std::str::from_utf8(attribute.value.as_ref())
                .map_err(|error| Error::Xml(error.to_string()))?;
            return Ok(Some(
                quick_xml::escape::unescape(value)
                    .map_err(|error| Error::Xml(error.to_string()))?
                    .into_owned(),
            ));
        }
    }
    Ok(None)
}

fn append_text(
    output: &mut String,
    bytes: &mut usize,
    text: &str,
    maximum: usize,
    kind: LimitKind,
) -> Result<()> {
    let observed = bytes
        .checked_add(text.len())
        .ok_or_else(|| Error::Invalid("SmartArt text length overflows".to_owned()))?;
    if observed > maximum {
        return Err(Error::limit(kind, observed, maximum));
    }
    *bytes = observed;
    output.push_str(text);
    Ok(())
}

/// Builder for creating `SmartArt` diagrams.
pub struct SmartArtBuilder {
    diagram_type: DiagramType,
    nodes: Vec<DiagramNode>,
    layout_name: Option<String>,
}

impl SmartArtBuilder {
    /// Create a new `SmartArt` builder with the specified diagram type.
    #[must_use]
    pub fn new(diagram_type: DiagramType) -> Self {
        Self {
            diagram_type,
            nodes: Vec::new(),
            layout_name: None,
        }
    }

    /// Set the layout name.
    pub fn layout_name(mut self, name: impl Into<String>) -> Self {
        self.layout_name = Some(name.into());
        self
    }

    /// Add a text item to the diagram.
    pub fn add_item(mut self, text: impl Into<String>) -> Self {
        self.nodes.push(DiagramNode::new(text));
        self
    }

    /// Add multiple text items.
    pub fn add_items<I, S>(mut self, items: I) -> Self
    where
        I: IntoIterator<Item = S>,
        S: Into<String>,
    {
        for item in items {
            self.nodes.push(DiagramNode::new(item));
        }
        self
    }

    /// Build the `SmartArt` diagram.
    #[must_use]
    pub fn build(self) -> SmartArt {
        SmartArt {
            diagram_type: self.diagram_type,
            nodes: self.nodes,
            layout_name: self.layout_name,
            id: None,
        }
    }
}

/// Generate `SmartArt` data XML under the caller's bounded policy.
pub fn generate_smartart_data_xml(smartart: &SmartArt, limits: &Limits) -> Result<String> {
    limits.validate()?;
    let (node_count, text_bytes) = validate_smartart(smartart, limits)?;
    let estimate = estimate_smartart_output(node_count, text_bytes, 2048, 2048)?;
    ensure_output(estimate, limits, LimitKind::DiagramDataModelBytes)?;
    let xml = generate_smartart_data_xml_unchecked(smartart, node_count, estimate);
    ensure_output(xml.len(), limits, LimitKind::DiagramDataModelBytes)?;
    Ok(xml)
}

fn generate_smartart_data_xml_unchecked(
    smartart: &SmartArt,
    node_count: usize,
    capacity: usize,
) -> String {
    let mut xml = String::with_capacity(capacity);

    xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
    xml.push_str(
        r#"<dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram">"#,
    );
    xml.push_str("<dgm:ptLst>");

    // Root point: declare layout, quick style, and color identifiers via prSet.
    // These identifiers follow the OOXML SmartArt semantics used by Office and
    // reference implementations. We choose layout based on DiagramType, while
    // reusing the same quick style and color scheme that we generate in
    // quickStyle/colors parts.
    //
    // loTypeId values are aligned with the built-in layout IDs used by
    // PowerPoint (e.g. process1, cycle2, orgChart1, venn1, matrix3, pyramid1)
    // so that each diagram type can use its distinct default layout. For these
    // built-in layouts, PowerPoint typically leaves loCatId empty and derives
    // category from the layout definition's catLst instead, so we follow that
    // pattern and use an empty string for loCatId.
    let (lo_type_id, lo_cat_id) = match smartart.diagram_type {
        DiagramType::Process => (
            "urn:microsoft.com/office/officeart/2005/8/layout/process1",
            "",
        ),
        DiagramType::Cycle => (
            "urn:microsoft.com/office/officeart/2005/8/layout/cycle2",
            "",
        ),
        DiagramType::Hierarchy => (
            "urn:microsoft.com/office/officeart/2005/8/layout/orgChart1",
            "",
        ),
        DiagramType::Relationship => ("urn:microsoft.com/office/officeart/2005/8/layout/venn1", ""),
        DiagramType::Matrix => (
            "urn:microsoft.com/office/officeart/2005/8/layout/matrix3",
            "",
        ),
        DiagramType::Pyramid => (
            "urn:microsoft.com/office/officeart/2005/8/layout/pyramid1",
            "",
        ),
        DiagramType::List | DiagramType::Picture | DiagramType::Unknown => (
            "urn:microsoft.com/office/officeart/2005/8/layout/default",
            "",
        ),
    };

    // Quick style and colors: use the same built-in identifiers as in our
    // quickStyle/colors parts (simple1, accent1_1).
    let qs_type_id = "urn:microsoft.com/office/officeart/2005/8/quickstyle/simple1";
    let qs_cat_id = "simple";
    let cs_type_id = "urn:microsoft.com/office/officeart/2005/8/colors/accent1_1";
    let cs_cat_id = "accent1";

    xml.push_str("<dgm:pt modelId=\"0\" type=\"doc\">");
    xml.push_str(&format!(
        "<dgm:prSet loTypeId=\"{lo_type_id}\" loCatId=\"{lo_cat_id}\" qsTypeId=\"{qs_type_id}\" qsCatId=\"{qs_cat_id}\" csTypeId=\"{cs_type_id}\" csCatId=\"{cs_cat_id}\"/>"
    ));
    xml.push_str(
        r#"<dgm:spPr/><dgm:t><a:bodyPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/><a:lstStyle xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/><a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:endParaRPr/></a:p></dgm:t></dgm:pt>"#,
    );

    // Diagram types for which we emit a full presentation model (pres points,
    // transitions, and a SmartArt-style layout). Picture/Unknown remain
    // minimal for now.
    let has_presentation_model = matches!(
        smartart.diagram_type,
        DiagramType::List
            | DiagramType::Process
            | DiagramType::Cycle
            | DiagramType::Hierarchy
            | DiagramType::Relationship
            | DiagramType::Matrix
            | DiagramType::Pyramid
    );

    // Flatten diagram nodes (including hierarchy children) into a linear
    // sequence with parent references so that we can build the SmartArt data
    // graph.
    let mut flat_nodes: Vec<(i32, i32, &DiagramNode)> = Vec::with_capacity(node_count);
    let mut next_model_id: i32 = 1;

    fn flatten_nodes<'a>(
        out: &mut Vec<(i32, i32, &'a DiagramNode)>,
        next_id: &mut i32,
        parent_id: i32,
        node: &'a DiagramNode,
    ) {
        let id = *next_id;
        *next_id += 1;
        out.push((id, parent_id, node));
        for child in &node.children {
            flatten_nodes(out, next_id, id, child);
        }
    }

    for root in &smartart.nodes {
        flatten_nodes(&mut flat_nodes, &mut next_model_id, 0, root);
    }

    let node_count = flat_nodes.len();

    // Add content nodes (type defaults to "node").
    for (model_id, _, node) in &flat_nodes {
        xml.push_str(&format!(
            r#"<dgm:pt modelId="{}"><dgm:prSet/><dgm:spPr/><dgm:t><a:bodyPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/><a:lstStyle xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/><a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:r><a:rPr lang="en-US"/><a:t>{}</a:t></a:r></a:p></dgm:t></dgm:pt>"#,
            model_id,
            escape_xml(&node.text)
        ));
    }

    // For supported diagrams, add minimal presentation points (type="pres")
    // that map logical nodes to presentation nodes, plus a diagram-level
    // presentation point. This models the basic pres/presParOf graph used by
    // SmartArt layouts without copying concrete data from sample files.
    if has_presentation_model && node_count > 0 {
        let base_pres_id = 1000_i32;
        let diagram_pres_id = base_pres_id;

        // Diagram-level presentation point associated with the document root.
        xml.push_str(&format!(
            r#"<dgm:pt modelId="{diagram_pres_id}" type="pres"><dgm:prSet presAssocID="0" presName="diagram" presStyleCnt="0"><dgm:presLayoutVars><dgm:dir/><dgm:resizeHandles val="exact"/></dgm:presLayoutVars></dgm:prSet><dgm:spPr/></dgm:pt>"#,
        ));

        // One presentation point per logical node, including descendants,
        // each associated with that node and using a simple bullet-enabled
        // layout. Keeping the complete flattened graph avoids dangling
        // `presOf` connections for nested content points.
        for (node_id, _, _) in &flat_nodes {
            let pres_id = base_pres_id.saturating_add(*node_id);
            xml.push_str(&format!(
                r#"<dgm:pt modelId="{pres_id}" type="pres"><dgm:prSet presAssocID="{node_id}" presName="node" presStyleLbl="node0" presStyleIdx="0" presStyleCnt="1"><dgm:presLayoutVars><dgm:bulletEnabled val="1"/></dgm:presLayoutVars></dgm:prSet><dgm:spPr/></dgm:pt>"#,
            ));
        }

        // For each structural connection (doc -> node), define transition
        // points of type parTrans and sibTrans. These are referenced from the
        // CT_Cxn entries via parTransId/sibTransId and are used by layouts
        // when iterating over siblings (axis="followSib").
        let base_partrans_id = 2000_i32;
        let base_sibtrans_id = 2100_i32;
        for (idx, _) in flat_nodes.iter().enumerate() {
            let cxn_model_id = 100_i32.saturating_add(index_i32(idx));
            let par_pt_id = base_partrans_id.saturating_add(index_i32(idx));
            let sib_pt_id = base_sibtrans_id.saturating_add(index_i32(idx));

            xml.push_str(&format!(
                r#"<dgm:pt modelId="{par_pt_id}" type="parTrans" cxnId="{cxn_model_id}"><dgm:prSet/><dgm:spPr/></dgm:pt>"#,
            ));
            xml.push_str(&format!(
                r#"<dgm:pt modelId="{sib_pt_id}" type="sibTrans" cxnId="{cxn_model_id}"><dgm:prSet/><dgm:spPr/></dgm:pt>"#,
            ));
        }
    }

    xml.push_str("</dgm:ptLst>");

    // Connection list: structural parent-of connections from the document root
    // to each node, plus (for supported diagrams) basic presentation mappings.
    xml.push_str("<dgm:cxnLst>");

    // parOf connections: doc (0) -> each node, or parent -> child for
    // hierarchical diagrams. We use the flattened node list and connect each
    // node to its recorded parent (0 = document root).
    for (idx, (node_id, parent_id, _)) in flat_nodes.iter().enumerate() {
        let cxn_model_id = 100_i32.saturating_add(index_i32(idx));

        // Compute srcOrd as the index of this child among siblings with the
        // same parent.
        let src_ord = flat_nodes[..idx]
            .iter()
            .filter(|(_, p, _)| p == parent_id)
            .count();

        // For diagrams with a full presentation model, attach
        // parTransId/sibTransId pointing to the corresponding transition
        // points. For minimal diagram types, we omit these attributes.
        if has_presentation_model && node_count > 0 {
            let base_partrans_id = 2000_i32;
            let base_sibtrans_id = 2100_i32;
            let par_pt_id = base_partrans_id.saturating_add(index_i32(idx));
            let sib_pt_id = base_sibtrans_id.saturating_add(index_i32(idx));

            xml.push_str(&format!(
                r#"<dgm:cxn modelId="{cxn_model_id}" srcId="{parent_id}" destId="{node_id}" srcOrd="{src_ord}" destOrd="0" parTransId="{par_pt_id}" sibTransId="{sib_pt_id}"/>"#,
            ));
        } else {
            xml.push_str(&format!(
                r#"<dgm:cxn modelId="{cxn_model_id}" srcId="{parent_id}" destId="{node_id}" srcOrd="{src_ord}" destOrd="0"/>"#,
            ));
        }
    }

    if has_presentation_model && node_count > 0 {
        let base_pres_id = 1000_i32;
        let diagram_pres_id = base_pres_id;
        let pres_layout_id = lo_type_id;

        // presOf connection: document root -> diagram presentation node. This
        // ties the top-level layout node to the diagram-level presentation
        // point.
        xml.push_str(&format!(
            r#"<dgm:cxn modelId="{}" type="presOf" srcId="0" destId="{}" srcOrd="0" destOrd="0" presId="{}"/>"#,
            250,
            diagram_pres_id,
            pres_layout_id,
        ));

        // presOf connections: each logical node -> its presentation node.
        for (idx, (node_id, _, _)) in flat_nodes.iter().enumerate() {
            let pres_id = base_pres_id.saturating_add(*node_id);
            let cxn_id = 300_i32.saturating_add(index_i32(idx));
            xml.push_str(&format!(
                r#"<dgm:cxn modelId="{cxn_id}" type="presOf" srcId="{node_id}" destId="{pres_id}" srcOrd="0" destOrd="0" presId="{pres_layout_id}"/>"#,
            ));
        }

        // presParOf connections preserve the same parent/child relationship
        // as the content graph, with top-level nodes attached to the diagram
        // presentation point.
        for (idx, (node_id, parent_id, _)) in flat_nodes.iter().enumerate() {
            let pres_id = base_pres_id.saturating_add(*node_id);
            let source = if *parent_id == 0 {
                diagram_pres_id
            } else {
                base_pres_id.saturating_add(*parent_id)
            };
            let src_ord = flat_nodes[..idx]
                .iter()
                .filter(|(_, parent, _)| parent == parent_id)
                .count();
            let cxn_id = 500_i32.saturating_add(index_i32(idx));
            xml.push_str(&format!(
                r#"<dgm:cxn modelId="{cxn_id}" type="presParOf" srcId="{source}" destId="{pres_id}" srcOrd="{src_ord}" destOrd="0" presId="{pres_layout_id}"/>"#,
            ));
        }
    }

    xml.push_str("</dgm:cxnLst>");

    xml.push_str("<dgm:bg/><dgm:whole/>");
    xml.push_str("</dgm:dataModel>");

    xml
}

/// Generate `SmartArt` layout XML under the caller's bounded policy.
pub fn generate_smartart_layout_xml(smartart: &SmartArt, limits: &Limits) -> Result<String> {
    limits.validate()?;
    let _ = validate_smartart(smartart, limits)?;
    ensure_output(64 * 1024, limits, LimitKind::XmlBytes)?;
    let xml = generate_smartart_layout_xml_unchecked(smartart);
    ensure_output(xml.len(), limits, LimitKind::XmlBytes)?;
    Ok(xml)
}

fn generate_smartart_layout_xml_unchecked(smartart: &SmartArt) -> String {
    let mut xml = String::with_capacity(4096);

    xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
    // Use a layout type identifier for SmartArt diagrams that matches the
    // loTypeId we emit in the root dgm:prSet so that the data model correctly
    // binds to this embedded layout definition.
    let layout_type_id = match smartart.diagram_type {
        DiagramType::Process => "urn:microsoft.com/office/officeart/2005/8/layout/process1",
        DiagramType::Cycle => "urn:microsoft.com/office/officeart/2005/8/layout/cycle2",
        DiagramType::Hierarchy => "urn:microsoft.com/office/officeart/2005/8/layout/orgChart1",
        DiagramType::Relationship => "urn:microsoft.com/office/officeart/2005/8/layout/venn1",
        DiagramType::Matrix => "urn:microsoft.com/office/officeart/2005/8/layout/matrix3",
        DiagramType::Pyramid => "urn:microsoft.com/office/officeart/2005/8/layout/pyramid1",
        DiagramType::List | DiagramType::Picture | DiagramType::Unknown => {
            "urn:microsoft.com/office/officeart/2005/8/layout/default"
        },
    };
    xml.push_str(&format!(
        "<dgm:layoutDef xmlns:dgm=\"http://schemas.openxmlformats.org/drawingml/2006/diagram\" uniqueId=\"{layout_type_id}\">",
    ));

    // Category type for the layout, aligned with loCatId in the data model.
    let cat_type = match smartart.diagram_type {
        DiagramType::Process => "process",
        DiagramType::Cycle => "cycle",
        DiagramType::Hierarchy => "hierarchy",
        DiagramType::Relationship => "relationship",
        DiagramType::Matrix => "matrix",
        DiagramType::Pyramid => "pyramid",
        DiagramType::List | DiagramType::Picture | DiagramType::Unknown => "list",
    };

    // Per-diagram node geometry used in the layout tree. This controls the
    // visible shape type that the SmartArt engine creates for each logical
    // node (e.g. chevrons for process, circles for cycle, trapezoids for
    // pyramid, etc.).
    let node_shape_type = match smartart.diagram_type {
        DiagramType::Process => "chevron",
        DiagramType::Cycle | DiagramType::Relationship => "ellipse",
        DiagramType::Matrix => "roundRect",
        DiagramType::Pyramid => "trapezoid",
        DiagramType::List
        | DiagramType::Hierarchy
        | DiagramType::Picture
        | DiagramType::Unknown => "rect",
    };

    // Diagram types for which we provide a full SmartArt-style layout tree
    // (diagram/node/sibTrans with forEach). Picture/Unknown use a minimal
    // fallback layout.
    let has_presentation_model = matches!(
        smartart.diagram_type,
        DiagramType::List
            | DiagramType::Process
            | DiagramType::Cycle
            | DiagramType::Hierarchy
            | DiagramType::Relationship
            | DiagramType::Matrix
            | DiagramType::Pyramid
    );

    // Simple title for layout
    xml.push_str("<dgm:title val=\"\"/>");
    xml.push_str("<dgm:desc val=\"\"/>");

    // Category list
    xml.push_str("<dgm:catLst>");
    xml.push_str(&format!("<dgm:cat type=\"{cat_type}\" pri=\"1000\"/>"));
    // Many built-in layouts also advertise themselves as convertible to/from
    // other layouts using a secondary "convert" category. Mirroring this
    // helps Office treat our embedded layouts more like full SmartArt
    // definitions instead of falling back to the default list layout.
    if matches!(
        smartart.diagram_type,
        DiagramType::Process
            | DiagramType::Cycle
            | DiagramType::Hierarchy
            | DiagramType::Relationship
            | DiagramType::Matrix
            | DiagramType::Pyramid
    ) {
        xml.push_str("<dgm:cat type=\"convert\" pri=\"15000\"/>");
    }
    xml.push_str("</dgm:catLst>");

    // Sample data. We provide a small, schema-valid sample data model that
    // exercises a document root plus a few placeholder nodes. This is not
    // used directly for rendering but helps Office recognize the layout as a
    // proper SmartArt definition.
    xml.push_str("<dgm:sampData>");
    xml.push_str("<dgm:dataModel>");
    xml.push_str("<dgm:ptLst>");
    xml.push_str(r#"<dgm:pt modelId="0" type="doc"/>"#);
    for i in 1..=3 {
        xml.push_str(&format!(r#"<dgm:pt modelId="{i}"/>"#));
    }
    xml.push_str("</dgm:ptLst>");
    xml.push_str("<dgm:cxnLst/>");
    xml.push_str("<dgm:bg/><dgm:whole/>");
    xml.push_str("</dgm:dataModel>");
    xml.push_str("</dgm:sampData>");

    // Minimal style and color data models. Built-in PowerPoint layouts
    // include styleData and clrData sections alongside sampData. Emitting
    // lightweight versions here makes our embedded layouts look more like
    // those full definitions while keeping the XML compact.
    if has_presentation_model {
        // Style data: root doc node plus two child placeholders, with a
        // simple parent-of connection graph.
        xml.push_str("<dgm:styleData>");
        xml.push_str("<dgm:dataModel>");
        xml.push_str("<dgm:ptLst>");
        xml.push_str(r#"<dgm:pt modelId="0" type="doc"/>"#);
        xml.push_str(r#"<dgm:pt modelId="1"/><dgm:pt modelId="2"/>"#);
        xml.push_str("</dgm:ptLst>");
        xml.push_str("<dgm:cxnLst>");
        xml.push_str(r#"<dgm:cxn modelId="3" srcId="0" destId="1" srcOrd="0" destOrd="0"/>"#);
        xml.push_str(r#"<dgm:cxn modelId="4" srcId="0" destId="2" srcOrd="1" destOrd="0"/>"#);
        xml.push_str("</dgm:cxnLst>");
        xml.push_str("<dgm:bg/><dgm:whole/>");
        xml.push_str("</dgm:dataModel>");
        xml.push_str("</dgm:styleData>");

        // Color data: similar miniature graph with a few more placeholders so
        // that color engines have something to bind to if needed.
        xml.push_str("<dgm:clrData>");
        xml.push_str("<dgm:dataModel>");
        xml.push_str("<dgm:ptLst>");
        xml.push_str(r#"<dgm:pt modelId="0" type="doc"/>"#);
        xml.push_str(r#"<dgm:pt modelId="1"/><dgm:pt modelId="2"/><dgm:pt modelId="3"/>"#);
        xml.push_str("</dgm:ptLst>");
        xml.push_str("<dgm:cxnLst>");
        xml.push_str(r#"<dgm:cxn modelId="5" srcId="0" destId="1" srcOrd="0" destOrd="0"/>"#);
        xml.push_str(r#"<dgm:cxn modelId="6" srcId="0" destId="2" srcOrd="1" destOrd="0"/>"#);
        xml.push_str(r#"<dgm:cxn modelId="7" srcId="0" destId="3" srcOrd="2" destOrd="0"/>"#);
        xml.push_str("</dgm:cxnLst>");
        xml.push_str("<dgm:bg/><dgm:whole/>");
        xml.push_str("</dgm:dataModel>");
        xml.push_str("</dgm:clrData>");
    }

    // Layout node definition
    if has_presentation_model {
        xml.push_str("<dgm:layoutNode name=\"diagram\">");
        xml.push_str("<dgm:varLst>");
        xml.push_str("<dgm:dir/>");
        xml.push_str(r#"<dgm:resizeHandles val="exact"/>"#);
        xml.push_str("</dgm:varLst>");
        xml.push_str("<dgm:alg type=\"lin\"/>");
        xml.push_str("<dgm:shape/>");
        xml.push_str("<dgm:presOf/>");
        xml.push_str("<dgm:constrLst>");
        xml.push_str(r#"<dgm:constr type="primFontSz" val="65"/>"#);
        xml.push_str(r#"<dgm:constr type="w" for="ch" forName="node" refType="w"/>"#);
        xml.push_str(r#"<dgm:constr type="h" for="ch" forName="node" refType="w" refFor="ch" refForName="node" fact="0.5"/>"#);
        xml.push_str(r#"<dgm:constr type="w" for="ch" forName="sibTrans" refType="w" refFor="ch" refForName="node" fact="0.15"/>"#);
        xml.push_str(r#"<dgm:constr type="sp" refType="w" refFor="ch" refForName="sibTrans"/>"#);
        xml.push_str("</dgm:constrLst>");
        xml.push_str("<dgm:ruleLst/>");

        xml.push_str(r#"<dgm:forEach name="items" axis="ch" ptType="node">"#);
        xml.push_str(r#"<dgm:layoutNode name="node">"#);
        xml.push_str("<dgm:varLst>");
        xml.push_str(r#"<dgm:bulletEnabled val="1"/>"#);
        xml.push_str("</dgm:varLst>");
        xml.push_str(r#"<dgm:alg type="tx"/>"#);
        xml.push_str(&format!("<dgm:shape type=\"{node_shape_type}\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:blip=\"\"><dgm:adjLst/></dgm:shape>"));
        xml.push_str(r#"<dgm:presOf axis="desOrSelf" ptType="node"/>"#);
        xml.push_str("<dgm:constrLst>");
        xml.push_str(r#"<dgm:constr type="lMarg" refType="primFontSz" fact="0.4"/>"#);
        xml.push_str(r#"<dgm:constr type="rMarg" refType="primFontSz" fact="0.4"/>"#);
        xml.push_str(r#"<dgm:constr type="tMarg" refType="primFontSz" fact="0.4"/>"#);
        xml.push_str(r#"<dgm:constr type="bMarg" refType="primFontSz" fact="0.4"/>"#);
        xml.push_str("</dgm:constrLst>");
        xml.push_str("<dgm:ruleLst>");
        xml.push_str(r#"<dgm:rule type="primFontSz" val="5"/>"#);
        xml.push_str("</dgm:ruleLst>");
        xml.push_str("</dgm:layoutNode>");

        xml.push_str(r#"<dgm:forEach name="spacing" axis="followSib" ptType="sibTrans" cnt="1">"#);
        xml.push_str(r#"<dgm:layoutNode name="sibTrans">"#);
        xml.push_str(r#"<dgm:alg type="sp"/>"#);
        xml.push_str(r#"<dgm:shape xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:blip=""><dgm:adjLst/></dgm:shape>"#);
        xml.push_str("<dgm:presOf/>");
        xml.push_str("<dgm:constrLst/>");
        xml.push_str("<dgm:ruleLst/>");
        xml.push_str("</dgm:layoutNode>");
        xml.push_str("</dgm:forEach>");

        xml.push_str("</dgm:forEach>");
        xml.push_str("</dgm:layoutNode>");
    } else {
        xml.push_str("<dgm:layoutNode name=\"root\">");
        xml.push_str("<dgm:varLst>");
        xml.push_str(r#"<dgm:dir val="norm"/>"#);
        xml.push_str("</dgm:varLst>");
        xml.push_str("<dgm:alg type=\"lin\"/>");
        xml.push_str("<dgm:shape/>");
        xml.push_str("<dgm:presOf/>");
        xml.push_str("<dgm:constrLst>");

        // Add constraints based on diagram type
        let spacing = match smartart.diagram_type {
            DiagramType::Process => 150,
            DiagramType::Hierarchy => 200,
            DiagramType::List
            | DiagramType::Cycle
            | DiagramType::Relationship
            | DiagramType::Matrix
            | DiagramType::Pyramid
            | DiagramType::Picture
            | DiagramType::Unknown => 100,
        };
        xml.push_str(r#"<dgm:constr type="primFontSz" val="65"/>"#);
        xml.push_str(&format!(r#"<dgm:constr type="sp" val="{spacing}"/>"#));

        xml.push_str("</dgm:constrLst>");
        xml.push_str("<dgm:ruleLst/>");
        xml.push_str("</dgm:layoutNode>");
    }

    xml.push_str("</dgm:layoutDef>");

    xml
}

/// Generate `SmartArt` drawing XML under the caller's bounded policy.
pub fn generate_smartart_drawing_xml(
    smartart: &SmartArt,
    x: i64,
    y: i64,
    width: i64,
    height: i64,
    limits: &Limits,
) -> Result<String> {
    limits.validate()?;
    let (node_count, text_bytes) = validate_smartart(smartart, limits)?;
    let estimate = estimate_smartart_output(node_count, text_bytes, 4096, 3072)?;
    ensure_output(estimate, limits, LimitKind::DiagramDrawingBytes)?;
    let xml = generate_smartart_drawing_xml_unchecked(smartart, x, y, width, height, estimate);
    ensure_output(xml.len(), limits, LimitKind::DiagramDrawingBytes)?;
    Ok(xml)
}

fn generate_smartart_drawing_xml_unchecked(
    smartart: &SmartArt,
    x: i64,
    y: i64,
    width: i64,
    height: i64,
    capacity: usize,
) -> String {
    let mut xml = String::with_capacity(capacity);

    xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
    // Use the Microsoft extension namespace for diagram drawings (dsp), as used by
    // PowerPoint, Apache POI, and LibreOffice for SmartArt pre-rendered shapes.
    xml.push_str(
        r#"<dsp:drawing xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram" "#,
    );
    xml.push_str(r#"xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" "#);
    xml.push_str(r#"xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">"#);

    xml.push_str("<dsp:spTree>");

    // Group shape container (match typical PowerPoint output: empty grpSpPr)
    xml.push_str("<dsp:nvGrpSpPr>");
    xml.push_str("<dsp:cNvPr id=\"0\" name=\"\"/>");
    xml.push_str("<dsp:cNvGrpSpPr/>");
    xml.push_str("</dsp:nvGrpSpPr>");
    xml.push_str("<dsp:grpSpPr/>");

    // Generate one shape for every flattened node, including descendants.
    // The previous top-level-only loop left nested data points without a
    // corresponding cached shape and made valid nested graphics lose content.
    let mut flat_nodes: Vec<(i32, u32, &DiagramNode)> = Vec::new();
    fn flatten_nodes<'a>(
        output: &mut Vec<(i32, u32, &'a DiagramNode)>,
        next_id: &mut i32,
        node: &'a DiagramNode,
        depth: u32,
    ) {
        let id = *next_id;
        *next_id = (*next_id).saturating_add(1);
        output.push((id, depth, node));
        for child in &node.children {
            flatten_nodes(output, next_id, child, depth.saturating_add(1));
        }
    }
    let mut next_id = 1_i32;
    for node in &smartart.nodes {
        flatten_nodes(&mut flat_nodes, &mut next_id, node, 0);
    }
    let node_count = flat_nodes.len().max(1);
    let node_width = width / index_i64(node_count);
    let max_depth = flat_nodes
        .iter()
        .map(|(_, depth, _)| *depth as i64)
        .max()
        .unwrap_or(0);
    let row_height = height / (max_depth + 1).max(1);
    let node_height = row_height * 8 / 10;

    for (idx, (node_id, depth, node)) in flat_nodes.iter().enumerate() {
        let node_x = x + index_i64(idx).saturating_mul(node_width);
        let node_y =
            y + i64::from(*depth).saturating_mul(row_height) + (row_height - node_height) / 2;

        // Use a numeric modelId per shape. ST_ModelId is a union of xsd:int and
        // GUID; using integers keeps us schema-valid. For diagram types with a
        // full presentation model, we align the shape modelIds with the
        // presentation points (type="pres") defined in the data model (1001,
        // 1002, ...). For minimal types, we fall back to the node ids
        // (1-based).
        let has_presentation_model = matches!(
            smartart.diagram_type,
            DiagramType::List
                | DiagramType::Process
                | DiagramType::Cycle
                | DiagramType::Hierarchy
                | DiagramType::Relationship
                | DiagramType::Matrix
                | DiagramType::Pyramid
        );
        let model_id: i32 = if has_presentation_model {
            let base_pres_id = 1000_i32;
            base_pres_id.saturating_add(*node_id)
        } else {
            *node_id
        };

        xml.push_str(&format!("<dsp:sp modelId=\"{model_id}\">"));
        xml.push_str("<dsp:nvSpPr>");
        xml.push_str(&format!(
            "<dsp:cNvPr id=\"{}\" name=\"Shape {}\"/>",
            idx + 1,
            idx + 1
        ));
        xml.push_str("<dsp:cNvSpPr/>");
        xml.push_str("</dsp:nvSpPr>");

        xml.push_str("<dsp:spPr>");
        xml.push_str("<a:xfrm>");
        xml.push_str(&format!(r#"<a:off x="{node_x}" y="{node_y}"/>"#));
        xml.push_str(&format!(
            r#"<a:ext cx="{}" cy="{}"/>"#,
            node_width * 9 / 10,
            node_height
        ));
        xml.push_str("</a:xfrm>");

        // Shape type based on diagram type
        let shape_type = match smartart.diagram_type {
            DiagramType::Process => "chevron",
            DiagramType::Cycle => "ellipse",
            DiagramType::Hierarchy => "rect",
            DiagramType::List
            | DiagramType::Relationship
            | DiagramType::Matrix
            | DiagramType::Pyramid
            | DiagramType::Picture
            | DiagramType::Unknown => "roundRect",
        };
        xml.push_str(&format!(
            r#"<a:prstGeom prst="{shape_type}"><a:avLst/></a:prstGeom>"#
        ));

        // Default fill and outline
        xml.push_str("<a:solidFill><a:schemeClr val=\"accent1\"/></a:solidFill>");
        xml.push_str("<a:ln><a:solidFill><a:schemeClr val=\"lt1\"/></a:solidFill></a:ln>");
        xml.push_str("</dsp:spPr>");

        // Basic style block, mirroring the structure used in typical
        // PowerPoint-generated SmartArt drawings.
        xml.push_str("<dsp:style>");
        xml.push_str("<a:lnRef idx=\"2\"><a:schemeClr val=\"accent1\"/></a:lnRef>");
        xml.push_str("<a:fillRef idx=\"1\"><a:schemeClr val=\"accent1\"/></a:fillRef>");
        xml.push_str("<a:effectRef idx=\"0\"><a:schemeClr val=\"accent1\"/></a:effectRef>");
        xml.push_str("<a:fontRef idx=\"minor\"><a:schemeClr val=\"tx1\"/></a:fontRef>");
        xml.push_str("</dsp:style>");

        // Text body
        xml.push_str("<dsp:txBody>");
        xml.push_str(r#"<a:bodyPr anchor="ctr"/>"#);
        xml.push_str("<a:lstStyle/>");
        xml.push_str("<a:p>");
        xml.push_str("<a:pPr algn=\"ctr\"/>");
        xml.push_str("<a:r>");
        xml.push_str(r#"<a:rPr lang="en-US"/>"#);
        xml.push_str(&format!("<a:t>{}</a:t>", escape_xml(&node.text)));
        xml.push_str("</a:r>");
        xml.push_str("</a:p>");
        xml.push_str("</dsp:txBody>");

        // Text transform (position/size of text box inside the shape). We
        // mirror the shape's geometry so PowerPoint and POI have clear anchors.
        xml.push_str("<dsp:txXfrm>");
        xml.push_str(&format!(r#"<a:off x="{node_x}" y="{node_y}"/>"#));
        xml.push_str(&format!(
            r#"<a:ext cx="{}" cy="{}"/>"#,
            node_width * 9 / 10,
            node_height
        ));
        xml.push_str("</dsp:txXfrm>");

        xml.push_str("</dsp:sp>");
    }

    xml.push_str("</dsp:spTree>");
    xml.push_str("</dsp:drawing>");

    xml
}

fn index_i32(value: usize) -> i32 {
    i32::try_from(value).unwrap_or(i32::MAX)
}

fn index_i64(value: usize) -> i64 {
    i64::try_from(value).unwrap_or(i64::MAX)
}

fn validate_smartart(smartart: &SmartArt, limits: &Limits) -> Result<(usize, usize)> {
    let max_nodes = limits.nodes_value().min(200_000);
    let max_depth = limits.node_depth_value().min(128);
    let max_text = limits.node_text_bytes_value().min(1024 * 1024);
    let mut count = 0usize;
    let mut text_bytes = 0usize;
    if smartart.nodes.len() > max_nodes {
        return Err(Error::limit(
            LimitKind::DiagramNodes,
            max_nodes + 1,
            max_nodes,
        ));
    }
    let mut stack = Vec::with_capacity(smartart.nodes.len());
    for node in smartart.nodes.iter().rev() {
        stack.push((node, 0));
    }
    while let Some((node, depth)) = stack.pop() {
        count = count.saturating_add(1);
        if count > max_nodes {
            return Err(Error::limit(LimitKind::DiagramNodes, count, max_nodes));
        }
        if depth > max_depth {
            return Err(Error::limit(LimitKind::DiagramNodeDepth, depth, max_depth));
        }
        if node.text.len() > max_text {
            return Err(Error::limit(
                LimitKind::DiagramNodeTextBytes,
                node.text.len(),
                max_text,
            ));
        }
        text_bytes = text_bytes
            .checked_add(node.text.len())
            .ok_or_else(|| Error::Invalid("SmartArt text length overflows".to_owned()))?;
        if node.text.chars().any(|character| {
            !matches!(character, '\t' | '\n' | '\r' | '\u{20}'..='\u{D7FF}' | '\u{E000}'..='\u{FFFD}' | '\u{10000}'..)
        }) {
            return Err(Error::Invalid(
                "SmartArt node text holds a character XML cannot carry".to_owned(),
            ));
        }
        let pending_capacity = max_nodes.saturating_sub(count).saturating_sub(stack.len());
        if node.children.len() > pending_capacity {
            return Err(Error::limit(
                LimitKind::DiagramNodes,
                max_nodes + 1,
                max_nodes,
            ));
        }
        stack.extend(node.children.iter().map(|child| (child, depth + 1)));
    }
    Ok((count, text_bytes))
}

fn estimate_smartart_output(
    count: usize,
    text_bytes: usize,
    base: usize,
    per_node: usize,
) -> Result<usize> {
    let escaped_text = text_bytes
        .checked_mul(6)
        .ok_or_else(|| Error::Invalid("SmartArt output length overflows".to_owned()))?;
    base.checked_add(
        count
            .checked_mul(per_node)
            .ok_or_else(|| Error::Invalid("SmartArt output length overflows".to_owned()))?,
    )
    .and_then(|value| value.checked_add(escaped_text))
    .ok_or_else(|| Error::Invalid("SmartArt output length overflows".to_owned()))
}

fn ensure_output(observed: usize, limits: &Limits, kind: LimitKind) -> Result<()> {
    let maximum = limits.output_bytes_value().min(16 * 1024 * 1024);
    if observed > maximum {
        return Err(Error::limit(kind, observed, maximum));
    }
    Ok(())
}

/// Generate `SmartArt` colors XML under the caller's bounded policy.
pub fn generate_smartart_colors_xml(limits: &Limits) -> Result<String> {
    limits.validate()?;
    ensure_output(16 * 1024, limits, LimitKind::XmlBytes)?;
    let xml = generate_smartart_colors_xml_unchecked();
    ensure_output(xml.len(), limits, LimitKind::XmlBytes)?;
    Ok(xml)
}

fn generate_smartart_colors_xml_unchecked() -> String {
    let mut xml = String::with_capacity(1024);

    xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
    xml.push_str(
        r#"<dgm:colorsDef xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" "#,
    );
    xml.push_str(r#"xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" uniqueId="urn:microsoft.com/office/officeart/2005/8/colors/accent1_1">"#);

    xml.push_str("<dgm:title val=\"\"/>");
    xml.push_str("<dgm:desc val=\"\"/>");
    xml.push_str("<dgm:catLst><dgm:cat type=\"accent1\" pri=\"10100\"/></dgm:catLst>");

    xml.push_str("<dgm:styleLbl name=\"node0\">");
    xml.push_str("<dgm:fillClrLst meth=\"repeat\"><a:schemeClr val=\"accent1\"/></dgm:fillClrLst>");
    xml.push_str("<dgm:linClrLst meth=\"repeat\"><a:schemeClr val=\"lt1\"/></dgm:linClrLst>");
    xml.push_str("<dgm:effectClrLst/>");
    xml.push_str("<dgm:txLinClrLst/>");
    xml.push_str("<dgm:txFillClrLst meth=\"repeat\"><a:schemeClr val=\"lt1\"/></dgm:txFillClrLst>");
    xml.push_str("<dgm:txEffectClrLst/>");
    xml.push_str("</dgm:styleLbl>");

    xml.push_str("</dgm:colorsDef>");

    xml
}

/// Generate `SmartArt` quick style XML under the caller's bounded policy.
pub fn generate_smartart_quickstyle_xml(limits: &Limits) -> Result<String> {
    limits.validate()?;
    ensure_output(16 * 1024, limits, LimitKind::XmlBytes)?;
    let xml = generate_smartart_quickstyle_xml_unchecked();
    ensure_output(xml.len(), limits, LimitKind::XmlBytes)?;
    Ok(xml)
}

fn generate_smartart_quickstyle_xml_unchecked() -> String {
    let mut xml = String::with_capacity(1024);

    xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
    xml.push_str(
        r#"<dgm:styleDef xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" "#,
    );
    xml.push_str(r#"xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" uniqueId="urn:microsoft.com/office/officeart/2005/8/quickstyle/simple1">"#);

    xml.push_str("<dgm:title val=\"\"/>");
    xml.push_str("<dgm:desc val=\"\"/>");
    xml.push_str("<dgm:catLst><dgm:cat type=\"simple\" pri=\"10000\"/></dgm:catLst>");

    xml.push_str("<dgm:styleLbl name=\"node0\">");
    xml.push_str("<dgm:scene3d><a:camera prst=\"orthographicFront\"/><a:lightRig rig=\"threePt\" dir=\"t\"/></dgm:scene3d>");
    xml.push_str("<dgm:txPr/>");
    xml.push_str("<dgm:style><a:lnRef idx=\"1\"><a:schemeClr val=\"accent1\"/></a:lnRef><a:fillRef idx=\"2\"><a:schemeClr val=\"accent1\"/></a:fillRef><a:effectRef idx=\"1\"><a:schemeClr val=\"accent1\"/></a:effectRef><a:fontRef idx=\"minor\"><a:schemeClr val=\"dk1\"/></a:fontRef></dgm:style>");
    xml.push_str("</dgm:styleLbl>");

    xml.push_str("</dgm:styleDef>");

    xml
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn attribute_lookup_checks_large_tags_and_stops_at_duplicate_names() {
        let attributes = (0..80)
            .map(|index| format!(" a{index}=\"{index}\""))
            .collect::<String>();
        let content = format!("pt{attributes} wanted=\"A &amp; B\"");
        let element = BytesStart::from_content(content, 2);
        assert_eq!(
            attribute_value(&element, b"wanted").unwrap().as_deref(),
            Some("A & B")
        );
        assert_eq!(attribute_value(&element, b"missing").unwrap(), None);

        let duplicate = format!("pt{attributes} a79=\"duplicate wanted='wrong'\" wanted=\"late\"");
        let element = BytesStart::from_content(duplicate, 2);
        assert!(matches!(
            attribute_value(&element, b"wanted"),
            Err(Error::Xml(_))
        ));
    }

    #[test]
    fn test_diagram_type_from_uri() {
        assert_eq!(
            DiagramType::from_layout_uri("urn:microsoft.com/office/list"),
            DiagramType::List
        );
        assert_eq!(
            DiagramType::from_layout_uri("urn:microsoft.com/office/orgChart"),
            DiagramType::Hierarchy
        );
    }

    #[test]
    fn test_smartart_builder() {
        let smartart = SmartArtBuilder::new(DiagramType::List)
            .layout_name("Basic List")
            .add_items(vec!["Item 1", "Item 2", "Item 3"])
            .build();

        assert_eq!(smartart.diagram_type, DiagramType::List);
        assert_eq!(smartart.node_count(), 3);
        assert!(smartart.text().contains("Item 1"));
    }

    #[test]
    fn raw_policy_bounds_smartart_reader_and_generators() {
        let defaults = Limits::default();
        let invalid = Limits::new(
            0,
            defaults.output_bytes_value(),
            defaults.xml_nodes_value(),
            defaults.xml_depth_value(),
            defaults.points_value(),
            defaults.connections_value(),
            defaults.data_text_bytes_value(),
            defaults.nodes_value(),
            defaults.node_depth_value(),
            defaults.node_text_bytes_value(),
        );
        let source = concat!(
            "<dgm:dataModel xmlns:dgm=\"http://schemas.openxmlformats.org/drawingml/2006/diagram\">",
            "<dgm:ptLst><dgm:pt><dgm:t>A B</dgm:t></dgm:pt></dgm:ptLst>",
            "</dgm:dataModel>"
        );
        assert!(matches!(
            SmartArt::parse_data_xml(source, &invalid),
            Err(Error::InvalidLimit { .. })
        ));
        assert_eq!(
            SmartArt::parse_data_xml(source, &defaults)
                .unwrap()
                .iter()
                .map(|node| node.text.as_str())
                .collect::<Vec<_>>(),
            ["A B"]
        );
        assert!(matches!(
            SmartArt::parse_data_xml(
                "<dgm:wrongRoot xmlns:dgm=\"http://schemas.openxmlformats.org/drawingml/2006/diagram\"/>",
                &defaults,
            ),
            Err(Error::Invalid(_))
        ));
        assert!(matches!(
            SmartArt::parse_data_xml(
                "<dgm:dataModel xmlns:dgm=\"http://schemas.openxmlformats.org/drawingml/2006/diagram\">text</dgm:dataModel>",
                &defaults,
            ),
            Err(Error::Invalid(_))
        ));
        assert!(matches!(
            SmartArt::parse_data_xml(
                "<dgm:dataModel xmlns:dgm=\"http://schemas.openxmlformats.org/drawingml/2006/diagram\"/>",
                &defaults,
            ),
            Err(Error::Invalid(_))
        ));

        let smartart = SmartArtBuilder::new(DiagramType::Process)
            .add_item("A & B")
            .add_item("C")
            .build();
        let data = generate_smartart_data_xml(&smartart, &defaults).unwrap();
        assert!(data.contains("A &amp; B"));
        assert!(
            !generate_smartart_layout_xml(&smartart, &defaults)
                .unwrap()
                .is_empty()
        );
        assert!(
            !generate_smartart_drawing_xml(&smartart, 0, 0, 100, 100, &defaults)
                .unwrap()
                .is_empty()
        );
        assert!(!generate_smartart_colors_xml(&defaults).unwrap().is_empty());
        assert!(
            !generate_smartart_quickstyle_xml(&defaults)
                .unwrap()
                .is_empty()
        );

        let tiny_output = Limits::new(
            defaults.input_bytes_value(),
            1,
            defaults.xml_nodes_value(),
            defaults.xml_depth_value(),
            defaults.points_value(),
            defaults.connections_value(),
            defaults.data_text_bytes_value(),
            defaults.nodes_value(),
            defaults.node_depth_value(),
            defaults.node_text_bytes_value(),
        );
        assert!(matches!(
            generate_smartart_data_xml(&smartart, &tiny_output),
            Err(Error::Limit {
                kind: LimitKind::DiagramDataModelBytes,
                ..
            })
        ));
        assert!(matches!(
            generate_smartart_layout_xml(&smartart, &tiny_output),
            Err(Error::Limit {
                kind: LimitKind::XmlBytes,
                ..
            })
        ));
        assert!(matches!(
            generate_smartart_drawing_xml(&smartart, 0, 0, 100, 100, &tiny_output),
            Err(Error::Limit {
                kind: LimitKind::DiagramDrawingBytes,
                ..
            })
        ));
        assert!(matches!(
            generate_smartart_colors_xml(&tiny_output),
            Err(Error::Limit {
                kind: LimitKind::XmlBytes,
                ..
            })
        ));
        assert!(matches!(
            generate_smartart_quickstyle_xml(&tiny_output),
            Err(Error::Limit {
                kind: LimitKind::XmlBytes,
                ..
            })
        ));
    }

    #[test]
    fn generated_nonempty_paired_data_model_is_accepted() {
        let limits = Limits::default();
        let smartart = SmartArtBuilder::new(DiagramType::List)
            .add_item("generated root")
            .build();
        let data = generate_smartart_data_xml(&smartart, &limits).unwrap();
        assert!(data.contains("<dgm:dataModel "));
        assert!(data.contains("</dgm:dataModel>"));
        let nodes = SmartArt::parse_data_xml(&data, &limits).unwrap();
        assert_eq!(nodes.len(), 1);
        assert_eq!(nodes[0].text, "generated root");
    }

    #[test]
    fn nested_nodes_round_trip_through_data_and_drawing_generators() {
        let limits = Limits::default();
        let mut root = DiagramNode::new("root &");
        let mut child = DiagramNode::new("child <one>");
        child.add_child(DiagramNode::new("grandchild"));
        root.add_child(child);
        let smartart = SmartArt {
            diagram_type: DiagramType::Hierarchy,
            nodes: vec![root],
            layout_name: None,
            id: None,
        };

        let data = generate_smartart_data_xml(&smartart, &limits).unwrap();
        let parsed = SmartArt::parse_data_xml(&data, &limits).unwrap();
        assert_eq!(parsed.len(), 1);
        assert_eq!(parsed[0].text, "root &");
        assert_eq!(parsed[0].children.len(), 1);
        assert_eq!(parsed[0].children[0].text, "child <one>");
        assert_eq!(parsed[0].children[0].depth, 1);
        assert_eq!(parsed[0].children[0].children[0].text, "grandchild");
        assert_eq!(parsed[0].children[0].children[0].depth, 2);

        let drawing = generate_smartart_drawing_xml(&smartart, 0, 0, 900, 600, &limits).unwrap();
        assert!(drawing.contains("root &amp;"));
        assert!(drawing.contains("child &lt;one&gt;"));
        assert!(drawing.contains("grandchild"));
        assert_eq!(drawing.matches("<dsp:sp modelId=").count(), 3);
        assert!(drawing.contains("modelId=\"1001\""));
        assert!(drawing.contains("modelId=\"1002\""));
        assert!(drawing.contains("modelId=\"1003\""));
    }

    fn limits_with(
        points: usize,
        connections: usize,
        data_text_bytes: usize,
        nodes: usize,
        node_depth: usize,
    ) -> Limits {
        let defaults = Limits::default();
        Limits::new(
            defaults.input_bytes_value(),
            defaults.output_bytes_value(),
            defaults.xml_nodes_value(),
            defaults.xml_depth_value(),
            points,
            connections,
            data_text_bytes,
            nodes,
            node_depth,
            defaults.node_text_bytes_value(),
        )
    }

    #[test]
    fn raw_parser_honors_points_connections_text_and_depth_limits() {
        let xml = concat!(
            "<dgm:dataModel xmlns:dgm=\"http://schemas.openxmlformats.org/drawingml/2006/diagram\">",
            "<dgm:ptLst>",
            "<dgm:pt modelId=\"0\" type=\"doc\"/>",
            "<dgm:pt modelId=\"1\"><dgm:t>parent</dgm:t></dgm:pt>",
            "<dgm:pt modelId=\"2\"><dgm:t>child</dgm:t></dgm:pt>",
            "<dgm:pt modelId=\"3\"><dgm:t>grandchild</dgm:t></dgm:pt>",
            "</dgm:ptLst>",
            "<dgm:cxnLst>",
            "<dgm:cxn modelId=\"10\" srcId=\"0\" destId=\"1\" srcOrd=\"0\"/>",
            "<dgm:cxn modelId=\"11\" srcId=\"1\" destId=\"2\" srcOrd=\"0\"/>",
            "<dgm:cxn modelId=\"12\" srcId=\"2\" destId=\"3\" srcOrd=\"0\"/>",
            "</dgm:cxnLst>",
            "</dgm:dataModel>"
        );
        assert!(matches!(
            SmartArt::parse_data_xml(xml, &limits_with(2, 100, 100, 100, 100)),
            Err(Error::Limit {
                kind: LimitKind::DiagramPoints,
                ..
            })
        ));
        assert!(matches!(
            SmartArt::parse_data_xml(xml, &limits_with(100, 1, 100, 100, 100)),
            Err(Error::Limit {
                kind: LimitKind::DiagramConnections,
                ..
            })
        ));
        assert!(matches!(
            SmartArt::parse_data_xml(xml, &limits_with(100, 100, 3, 100, 100)),
            Err(Error::Limit {
                kind: LimitKind::DiagramDataModelTextBytes,
                ..
            })
        ));
        assert!(matches!(
            SmartArt::parse_data_xml(xml, &limits_with(100, 100, 100, 100, 1)),
            Err(Error::Limit {
                kind: LimitKind::DiagramNodeDepth,
                ..
            })
        ));
    }

    #[test]
    fn raw_parser_preserves_empty_parents_and_source_order() {
        let xml = concat!(
            "<dgm:dataModel xmlns:dgm=\"http://schemas.openxmlformats.org/drawingml/2006/diagram\">",
            "<dgm:ptLst>",
            "<dgm:pt modelId=\"0\" type=\"doc\"/>",
            "<dgm:pt modelId=\"1\"/>",
            "<dgm:pt modelId=\"2\"><dgm:t>second</dgm:t></dgm:pt>",
            "<dgm:pt modelId=\"3\"><dgm:t>first</dgm:t></dgm:pt>",
            "</dgm:ptLst>",
            "<dgm:cxnLst>",
            "<dgm:cxn modelId=\"10\" srcId=\"0\" destId=\"1\" srcOrd=\"0\"/>",
            "<dgm:cxn modelId=\"11\" srcId=\"1\" destId=\"2\" srcOrd=\"1\"/>",
            "<dgm:cxn modelId=\"12\" srcId=\"1\" destId=\"3\" srcOrd=\"0\"/>",
            "</dgm:cxnLst>",
            "</dgm:dataModel>"
        );
        let nodes = SmartArt::parse_data_xml(xml, &Limits::default()).unwrap();
        assert_eq!(nodes.len(), 1);
        assert_eq!(nodes[0].text, "");
        assert_eq!(
            nodes[0]
                .children
                .iter()
                .map(|node| node.text.as_str())
                .collect::<Vec<_>>(),
            ["first", "second"]
        );
    }

    #[test]
    fn raw_parser_stably_orders_equal_source_order_roots() {
        let xml = concat!(
            "<dgm:dataModel xmlns:dgm=\"http://schemas.openxmlformats.org/drawingml/2006/diagram\">",
            "<dgm:ptLst>",
            "<dgm:pt modelId=\"0\" type=\"doc\"/>",
            "<dgm:pt modelId=\"1\"><dgm:t>point order one</dgm:t></dgm:pt>",
            "<dgm:pt modelId=\"2\"><dgm:t>point order two</dgm:t></dgm:pt>",
            "<dgm:pt modelId=\"3\"><dgm:t>point order three</dgm:t></dgm:pt>",
            "</dgm:ptLst>",
            "<dgm:cxnLst>",
            "<dgm:cxn modelId=\"10\" srcId=\"0\" destId=\"3\" srcOrd=\"0\"/>",
            "<dgm:cxn modelId=\"11\" srcId=\"0\" destId=\"1\" srcOrd=\"0\"/>",
            "<dgm:cxn modelId=\"12\" srcId=\"0\" destId=\"2\" srcOrd=\"0\"/>",
            "</dgm:cxnLst>",
            "</dgm:dataModel>"
        );
        let nodes = SmartArt::parse_data_xml(xml, &Limits::default()).unwrap();
        assert_eq!(
            nodes
                .iter()
                .map(|node| node.text.as_str())
                .collect::<Vec<_>>(),
            ["point order one", "point order two", "point order three"]
        );
    }

    #[test]
    fn raw_parser_rejects_misplaced_roots_and_dangling_parent_connections() {
        let dgm = "http://schemas.openxmlformats.org/drawingml/2006/diagram";
        for xml in [
            format!("<dgm:dataModel xmlns:dgm=\"{dgm}\"></dgm:dataModel>"),
            format!("<dgm:dataModel xmlns:dgm=\"{dgm}\"><dgm:pt modelId=\"1\"/></dgm:dataModel>"),
            format!(
                "<dgm:dataModel xmlns:dgm=\"{dgm}\"><dgm:ptLst><dgm:ptLst/></dgm:ptLst></dgm:dataModel>"
            ),
            format!(
                "<dgm:dataModel xmlns:dgm=\"{dgm}\"><dgm:ptLst><dgm:pt modelId=\"1\"/></dgm:ptLst><dgm:cxn modelId=\"2\"/></dgm:dataModel>"
            ),
            format!(
                "<dgm:dataModel xmlns:dgm=\"{dgm}\"><dgm:ptLst><dgm:dataModel/></dgm:ptLst></dgm:dataModel>"
            ),
        ] {
            assert!(
                SmartArt::parse_data_xml(&xml, &Limits::default()).is_err(),
                "{xml}"
            );
        }

        let dangling = format!(
            "<dgm:dataModel xmlns:dgm=\"{dgm}\"><dgm:ptLst><dgm:pt modelId=\"0\" type=\"doc\"/><dgm:pt modelId=\"1\"/></dgm:ptLst><dgm:cxnLst><dgm:cxn modelId=\"2\" srcId=\"0\" destId=\"missing\"/></dgm:cxnLst></dgm:dataModel>"
        );
        assert!(matches!(
            SmartArt::parse_data_xml(&dangling, &Limits::default()),
            Err(Error::Invalid(_))
        ));
    }
}
