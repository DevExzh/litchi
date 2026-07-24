//! DrawingML diagram data-model (`dgm:dataModel`) parsing shared across formats.
//!
//! The data model is the semantic heart of a SmartArt diagram: it declares the
//! logical points (`dgm:pt`) — the document root, content nodes, transition
//! points, and presentation points — and the connection graph (`dgm:cxn`) that
//! links them. This parser reads both the transitional and the ISO Strict
//! `drawingml/diagram` namespaces and treats everything as inert metadata.

use crate::common::mce::process_str;
use crate::diagrams::{DGM_NAMESPACE, DGM_NAMESPACE_STRICT, DiagramNode};
use crate::error::{OoxmlError, Result};
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use std::collections::{HashMap, HashSet};

const MAX_DATA_MODEL_XML: usize = 16 * 1024 * 1024;
const MAX_NODES: usize = 200_000;
const MAX_DEPTH: usize = 128;
const MAX_POINTS: usize = 100_000;
const MAX_CONNECTIONS: usize = 100_000;
const MAX_TEXT_BYTES: usize = 1024 * 1024;
/// Recursion guard for [`DiagramDataModel::node_tree`] on cyclic graphs.
const MAX_TREE_DEPTH: u32 = 64;

/// The type of a diagram data-model point (`dgm:pt@type`).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DiagramPointType {
    /// Content node (`node`, the default when `type` is absent).
    Node,
    /// Document root point (`doc`).
    Document,
    /// Assistant node (`asst`).
    Assistant,
    /// Parent transition point (`parTrans`).
    ParTrans,
    /// Sibling transition point (`sibTrans`).
    SibTrans,
    /// Presentation point (`pres`).
    Pres,
    /// Any other point type.
    Other,
}

impl Default for DiagramPointType {
    fn default() -> Self {
        Self::Node
    }
}

impl DiagramPointType {
    fn parse(value: Option<&str>) -> Self {
        match value {
            None | Some("node") => Self::Node,
            Some("doc") => Self::Document,
            Some("asst") => Self::Assistant,
            Some("parTrans") => Self::ParTrans,
            Some("sibTrans") => Self::SibTrans,
            Some("pres") => Self::Pres,
            Some(_) => Self::Other,
        }
    }
}

/// The type of a diagram data-model connection (`dgm:cxn@type`).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DiagramConnectionType {
    /// Structural parent-of (`parOf`, the default when `type` is absent).
    ParOf,
    /// Presentation mapping (`presOf`).
    PresOf,
    /// Presentation parent-of (`presParOf`).
    PresParOf,
    /// Any other connection type.
    Other,
}

impl DiagramConnectionType {
    fn parse(value: Option<&str>) -> Self {
        match value {
            None | Some("parOf") => Self::ParOf,
            Some("presOf") => Self::PresOf,
            Some("presParOf") => Self::PresParOf,
            Some(_) => Self::Other,
        }
    }
}

/// A single point (`dgm:pt`) in a diagram data model.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DiagramPoint {
    /// Point identifier (`modelId`); an integer or GUID string per `ST_ModelId`.
    pub model_id: String,
    /// Point type.
    pub point_type: DiagramPointType,
    /// Owning connection for transition points (`cxnId`).
    pub cxn_id: Option<String>,
    /// Concatenated text of the point's `dgm:t` body.
    pub text: String,
    /// Layout type identifier from `dgm:prSet@loTypeId` (document point).
    pub layout_type_id: Option<String>,
    /// Quick style type identifier from `dgm:prSet@qsTypeId` (document point).
    pub quick_style_type_id: Option<String>,
    /// Color style type identifier from `dgm:prSet@csTypeId` (document point).
    pub color_style_type_id: Option<String>,
}

/// A single connection (`dgm:cxn`) in a diagram data model.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DiagramConnection {
    /// Connection identifier (`modelId`).
    pub model_id: String,
    /// Connection type.
    pub connection_type: DiagramConnectionType,
    /// Source point identifier (`srcId`).
    pub src_id: String,
    /// Destination point identifier (`destId`).
    pub dest_id: String,
    /// Source order (`srcOrd`).
    pub src_ord: u32,
    /// Destination order (`destOrd`).
    pub dest_ord: u32,
}

/// A parsed DrawingML diagram data model (`dgm:dataModel`).
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct DiagramDataModel {
    /// Data-model points in document order.
    pub points: Vec<DiagramPoint>,
    /// Data-model connections in document order.
    pub connections: Vec<DiagramConnection>,
}

impl DiagramDataModel {
    /// Parse a `dgm:dataModel` document (transitional or Strict namespace).
    ///
    /// The input is first rewritten by markup-compatibility processing so
    /// `mc:AlternateContent` wrappers resolve to their fallback content.
    pub fn parse(xml: &str) -> Result<Self> {
        if xml.len() > MAX_DATA_MODEL_XML {
            return Err(limit("data-model XML bytes"));
        }
        let processed = process_str(xml)?;
        if processed.len() > MAX_DATA_MODEL_XML {
            return Err(limit("processed data-model XML bytes"));
        }
        let mut reader = NsReader::from_reader(processed.as_bytes());
        // Text inside `a:t` is significant: preserve whitespace verbatim and
        // resolve entity references explicitly (`Event::GeneralRef`).
        reader.config_mut().trim_text(false);

        let mut model = DiagramDataModel::default();
        let mut buffer = Vec::new();
        let mut depth = 0usize;
        let mut nodes = 0usize;
        let mut root_seen = false;
        // Currently open point: depth at which it started plus its builder.
        let mut open_point: Option<(usize, PointBuilder)> = None;
        // Depth of the open `dgm:t` text body, if any.
        let mut text_depth: Option<usize> = None;

        loop {
            match reader.read_event_into(&mut buffer).map_err(xml_error)? {
                Event::Start(element) => {
                    depth += 1;
                    nodes += 1;
                    if nodes > MAX_NODES || depth > MAX_DEPTH {
                        return Err(limit("data-model XML structure"));
                    }
                    let namespace = element_namespace(&reader, &element)?;
                    let local = local_name(&element)?;
                    if !root_seen {
                        if !is_dgm(&namespace) || local != "dataModel" {
                            return Err(invalid("invalid data-model root or namespace"));
                        }
                        root_seen = true;
                    } else if is_dgm(&namespace) {
                        match local.as_str() {
                            "pt" => {
                                if open_point.is_some() {
                                    return Err(invalid("nested diagram point"));
                                }
                                if model.points.len() >= MAX_POINTS {
                                    return Err(limit("diagram point count"));
                                }
                                open_point = Some((depth, PointBuilder::from_element(&element)?));
                            },
                            "cxn" => push_connection(&mut model, &element)?,
                            "prSet" => {
                                if let Some((_, builder)) = &mut open_point {
                                    builder.read_pr_set(&element)?;
                                }
                            },
                            "t" if open_point.is_some() => text_depth = Some(depth),
                            _ => {},
                        }
                    }
                },
                Event::Empty(element) => {
                    nodes += 1;
                    if nodes > MAX_NODES || depth + 1 > MAX_DEPTH {
                        return Err(limit("data-model XML structure"));
                    }
                    let namespace = element_namespace(&reader, &element)?;
                    let local = local_name(&element)?;
                    if !root_seen {
                        if !is_dgm(&namespace) || local != "dataModel" {
                            return Err(invalid("invalid data-model root or namespace"));
                        }
                        root_seen = true;
                    } else if is_dgm(&namespace) {
                        match local.as_str() {
                            "pt" => {
                                if open_point.is_some() {
                                    return Err(invalid("nested diagram point"));
                                }
                                if model.points.len() >= MAX_POINTS {
                                    return Err(limit("diagram point count"));
                                }
                                let builder = PointBuilder::from_element(&element)?;
                                model.points.push(builder.finish());
                            },
                            "cxn" => push_connection(&mut model, &element)?,
                            "prSet" => {
                                if let Some((_, builder)) = &mut open_point {
                                    builder.read_pr_set(&element)?;
                                }
                            },
                            _ => {},
                        }
                    }
                },
                Event::Text(event) => {
                    if text_depth.is_some()
                        && let Some((_, builder)) = &mut open_point
                    {
                        let text = std::str::from_utf8(event.as_ref()).map_err(xml_error)?;
                        let text = quick_xml::escape::unescape(text).map_err(xml_error)?;
                        if builder.text.len() + text.len() > MAX_TEXT_BYTES {
                            return Err(limit("diagram text bytes"));
                        }
                        builder.text.push_str(&text);
                    }
                },
                Event::GeneralRef(reference) => {
                    if text_depth.is_some()
                        && let Some((_, builder)) = &mut open_point
                    {
                        let text = crate::common::xml::decode_xml_reference(&reference)?;
                        if builder.text.len() + text.len() > MAX_TEXT_BYTES {
                            return Err(limit("diagram text bytes"));
                        }
                        builder.text.push_str(&text);
                    }
                },
                Event::End(element) => {
                    if text_depth == Some(depth) {
                        text_depth = None;
                    }
                    if open_point.as_ref().is_some_and(|(start, _)| *start == depth) {
                        let local = local_name_end(&element)?;
                        if local != "pt" {
                            return Err(invalid("unbalanced diagram point"));
                        }
                        let (_, builder) = open_point.take().expect("open point checked above");
                        model.points.push(builder.finish());
                    }
                    if depth == 0 {
                        return Err(invalid("unexpected data-model closing element"));
                    }
                    depth -= 1;
                },
                Event::DocType(_) => return Err(invalid("DTDs are rejected")),
                Event::CData(_) => return Err(invalid("CDATA is rejected")),
                Event::Eof => break,
                _ => {},
            }
            buffer.clear();
        }
        if !root_seen || depth != 0 || open_point.is_some() {
            return Err(invalid("missing or unterminated data-model root"));
        }
        Ok(model)
    }

    /// The document root point (`type="doc"`), if present.
    pub fn document_point(&self) -> Option<&DiagramPoint> {
        self.points
            .iter()
            .find(|point| point.point_type == DiagramPointType::Document)
    }

    /// Iterate over content points (nodes and assistants, in document order).
    pub fn content_points(&self) -> impl Iterator<Item = &DiagramPoint> {
        self.points.iter().filter(|point| {
            matches!(
                point.point_type,
                DiagramPointType::Node | DiagramPointType::Assistant
            )
        })
    }

    /// Build the content-node hierarchy implied by the `parOf` connection
    /// graph, ordered by `srcOrd`. Cycles and dangling references are
    /// tolerated: nodes not reachable from the document root are omitted.
    pub fn node_tree(&self) -> Vec<DiagramNode> {
        let Some(root) = self.document_point() else {
            return Vec::new();
        };
        let points: HashMap<&str, &DiagramPoint> = self
            .points
            .iter()
            .map(|point| (point.model_id.as_str(), point))
            .collect();
        let mut children: HashMap<&str, Vec<(u32, &str)>> = HashMap::new();
        for connection in &self.connections {
            if connection.connection_type == DiagramConnectionType::ParOf {
                children
                    .entry(connection.src_id.as_str())
                    .or_default()
                    .push((connection.src_ord, connection.dest_id.as_str()));
            }
        }
        for entries in children.values_mut() {
            entries.sort_by_key(|(ord, _)| *ord);
        }
        let mut visiting = HashSet::new();
        build_children(
            root.model_id.as_str(),
            0,
            &points,
            &children,
            &mut visiting,
        )
    }

    /// All text content of the diagram, one line per content node.
    pub fn text(&self) -> String {
        self.node_tree()
            .iter()
            .map(|node| node.all_text())
            .collect::<Vec<_>>()
            .join("\n")
    }
}

fn build_children(
    parent_id: &str,
    depth: u32,
    points: &HashMap<&str, &DiagramPoint>,
    children: &HashMap<&str, Vec<(u32, &str)>>,
    visiting: &mut HashSet<String>,
) -> Vec<DiagramNode> {
    if depth >= MAX_TREE_DEPTH || !visiting.insert(parent_id.to_owned()) {
        return Vec::new();
    }
    let mut nodes = Vec::new();
    if let Some(entries) = children.get(parent_id) {
        for (_, dest_id) in entries {
            let Some(point) = points.get(dest_id) else {
                continue;
            };
            if !matches!(
                point.point_type,
                DiagramPointType::Node | DiagramPointType::Assistant
            ) || visiting.contains(*dest_id)
            {
                continue;
            }
            let mut node = DiagramNode::new(point.text.clone());
            node.depth = depth;
            node.children = build_children(dest_id, depth + 1, points, children, visiting);
            nodes.push(node);
        }
    }
    visiting.remove(parent_id);
    nodes
}

#[derive(Default)]
struct PointBuilder {
    model_id: Option<String>,
    point_type: DiagramPointType,
    cxn_id: Option<String>,
    text: String,
    layout_type_id: Option<String>,
    quick_style_type_id: Option<String>,
    color_style_type_id: Option<String>,
}

impl PointBuilder {
    fn from_element(element: &BytesStart<'_>) -> Result<Self> {
        let mut builder = PointBuilder::default();
        for (name, value) in attributes(element)? {
            match name.as_str() {
                "modelId" => builder.model_id = Some(value),
                "type" => builder.point_type = DiagramPointType::parse(Some(&value)),
                "cxnId" => builder.cxn_id = Some(value),
                _ => {},
            }
        }
        if builder.model_id.is_none() {
            return Err(invalid("diagram point lacks modelId"));
        }
        Ok(builder)
    }

    fn read_pr_set(&mut self, element: &BytesStart<'_>) -> Result<()> {
        for (name, value) in attributes(element)? {
            match name.as_str() {
                "loTypeId" => self.layout_type_id = Some(value),
                "qsTypeId" => self.quick_style_type_id = Some(value),
                "csTypeId" => self.color_style_type_id = Some(value),
                _ => {},
            }
        }
        Ok(())
    }

    fn finish(self) -> DiagramPoint {
        DiagramPoint {
            model_id: self.model_id.unwrap_or_default(),
            point_type: self.point_type,
            cxn_id: self.cxn_id,
            text: self.text,
            layout_type_id: self.layout_type_id,
            quick_style_type_id: self.quick_style_type_id,
            color_style_type_id: self.color_style_type_id,
        }
    }
}

fn push_connection(model: &mut DiagramDataModel, element: &BytesStart<'_>) -> Result<()> {
    if model.connections.len() >= MAX_CONNECTIONS {
        return Err(limit("diagram connection count"));
    }
    let mut connection = DiagramConnection {
        model_id: String::new(),
        connection_type: DiagramConnectionType::ParOf,
        src_id: String::new(),
        dest_id: String::new(),
        src_ord: 0,
        dest_ord: 0,
    };
    let mut has_model_id = false;
    let mut has_src = false;
    let mut has_dest = false;
    for (name, value) in attributes(element)? {
        match name.as_str() {
            "modelId" => {
                has_model_id = true;
                connection.model_id = value;
            },
            "type" => connection.connection_type = DiagramConnectionType::parse(Some(&value)),
            "srcId" => {
                has_src = true;
                connection.src_id = value;
            },
            "destId" => {
                has_dest = true;
                connection.dest_id = value;
            },
            "srcOrd" => connection.src_ord = parse_order(&value)?,
            "destOrd" => connection.dest_ord = parse_order(&value)?,
            _ => {},
        }
    }
    if !has_model_id || !has_src || !has_dest {
        return Err(invalid("diagram connection lacks required attributes"));
    }
    model.connections.push(connection);
    Ok(())
}

fn parse_order(value: &str) -> Result<u32> {
    value
        .parse()
        .map_err(|_| invalid("invalid diagram connection order"))
}

/// Unnamespaced, unescaped `(local name, value)` attribute pairs.
fn attributes(element: &BytesStart<'_>) -> Result<Vec<(String, String)>> {
    let mut values = Vec::new();
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(xml_error)?;
        let raw = attribute.key.as_ref();
        if raw == b"xmlns" || raw.starts_with(b"xmlns:") {
            continue;
        }
        let name = std::str::from_utf8(attribute.key.local_name().as_ref())
            .map_err(xml_error)?
            .to_owned();
        let value = std::str::from_utf8(attribute.value.as_ref()).map_err(xml_error)?;
        let value = quick_xml::escape::unescape(value)
            .map_err(xml_error)?
            .into_owned();
        values.push((name, value));
    }
    Ok(values)
}

fn element_namespace(reader: &NsReader<&[u8]>, element: &BytesStart<'_>) -> Result<String> {
    match reader.resolver().resolve_element(element.name()).0 {
        ResolveResult::Bound(Namespace(namespace)) => Ok(std::str::from_utf8(namespace)
            .map_err(xml_error)?
            .to_owned()),
        ResolveResult::Unbound => Ok(String::new()),
        ResolveResult::Unknown(prefix) => Err(invalid(format!(
            "unbound XML prefix '{}'",
            String::from_utf8_lossy(prefix.as_ref())
        ))),
    }
}

fn local_name(element: &BytesStart<'_>) -> Result<String> {
    Ok(std::str::from_utf8(element.local_name().as_ref())
        .map_err(xml_error)?
        .to_owned())
}

fn local_name_end(element: &quick_xml::events::BytesEnd<'_>) -> Result<String> {
    Ok(std::str::from_utf8(element.local_name().as_ref())
        .map_err(xml_error)?
        .to_owned())
}

fn is_dgm(namespace: &str) -> bool {
    matches!(namespace, DGM_NAMESPACE | DGM_NAMESPACE_STRICT)
}

fn xml_error(error: impl std::fmt::Display) -> OoxmlError {
    OoxmlError::Xml(error.to_string())
}

fn invalid(message: impl Into<String>) -> OoxmlError {
    OoxmlError::InvalidFormat(message.into())
}

fn limit(label: &str) -> OoxmlError {
    invalid(format!("diagram {label} limit exceeded"))
}

#[cfg(test)]
mod tests {
    use super::*;

    const TRANSITIONAL: &str = concat!(
        "<?xml version=\"1.0\"?>",
        "<dgm:dataModel xmlns:dgm=\"http://schemas.openxmlformats.org/drawingml/2006/diagram\" ",
        "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">",
        "<dgm:ptLst>",
        "<dgm:pt modelId=\"0\" type=\"doc\"><dgm:prSet loTypeId=\"urn:test/layout/process1\" ",
        "qsTypeId=\"urn:test/quickstyle/simple1\" csTypeId=\"urn:test/colors/accent1_1\"/>",
        "<dgm:spPr/><dgm:t><a:p><a:endParaRPr/></a:p></dgm:t></dgm:pt>",
        "<dgm:pt modelId=\"1\"><dgm:prSet/><dgm:t><a:p><a:r><a:t>Alpha &amp; </a:t></a:r>",
        "<a:r><a:t>Beta</a:t></a:r></a:p></dgm:t></dgm:pt>",
        "<dgm:pt modelId=\"2\"><dgm:prSet/><dgm:t><a:p><a:r><a:t>Gamma</a:t></a:r></a:p></dgm:t></dgm:pt>",
        "<dgm:pt modelId=\"3\" type=\"node\"><dgm:t><a:p><a:r><a:t>Child</a:t></a:r></a:p></dgm:t></dgm:pt>",
        "<dgm:pt modelId=\"2000\" type=\"parTrans\" cxnId=\"100\"/>",
        "<dgm:pt modelId=\"1000\" type=\"pres\"/>",
        "</dgm:ptLst>",
        "<dgm:cxnLst>",
        "<dgm:cxn modelId=\"100\" srcId=\"0\" destId=\"1\" srcOrd=\"0\" destOrd=\"0\"/>",
        "<dgm:cxn modelId=\"101\" srcId=\"0\" destId=\"2\" srcOrd=\"1\" destOrd=\"0\"/>",
        "<dgm:cxn modelId=\"102\" srcId=\"2\" destId=\"3\" srcOrd=\"0\" destOrd=\"0\"/>",
        "<dgm:cxn modelId=\"300\" type=\"presOf\" srcId=\"0\" destId=\"1000\" srcOrd=\"0\" destOrd=\"0\"/>",
        "</dgm:cxnLst>",
        "<dgm:bg/><dgm:whole/>",
        "</dgm:dataModel>"
    );

    const STRICT: &str = concat!(
        "<?xml version=\"1.0\"?>",
        "<dgm:dataModel xmlns:dgm=\"http://purl.oclc.org/ooxml/drawingml/diagram\" ",
        "xmlns:a=\"http://purl.oclc.org/ooxml/drawingml/main\">",
        "<dgm:ptLst>",
        "<dgm:pt modelId=\"{DOC}\" type=\"doc\"><dgm:prSet loTypeId=\"urn:test/layout/cycle2\"/></dgm:pt>",
        "<dgm:pt modelId=\"{A}\"><dgm:t><a:p><a:r><a:t>a</a:t></a:r></a:p></dgm:t></dgm:pt>",
        "<dgm:pt modelId=\"{T}\" type=\"sibTrans\" cxnId=\"{C}\"/>",
        "</dgm:ptLst>",
        "<dgm:cxnLst>",
        "<dgm:cxn modelId=\"{C}\" srcId=\"{DOC}\" destId=\"{A}\" srcOrd=\"0\" destOrd=\"0\"/>",
        "</dgm:cxnLst>",
        "</dgm:dataModel>"
    );

    #[test]
    fn parses_transitional_model_with_hierarchy_and_multi_run_text() {
        let model = DiagramDataModel::parse(TRANSITIONAL).unwrap();
        assert_eq!(model.points.len(), 6);
        assert_eq!(model.connections.len(), 4);
        let root = model.document_point().unwrap();
        assert_eq!(
            root.layout_type_id.as_deref(),
            Some("urn:test/layout/process1")
        );
        assert_eq!(
            root.quick_style_type_id.as_deref(),
            Some("urn:test/quickstyle/simple1")
        );
        assert_eq!(model.points[4].point_type, DiagramPointType::ParTrans);
        assert_eq!(model.points[4].cxn_id.as_deref(), Some("100"));
        assert_eq!(model.points[5].point_type, DiagramPointType::Pres);
        assert_eq!(
            model.connections[3].connection_type,
            DiagramConnectionType::PresOf
        );

        let tree = model.node_tree();
        assert_eq!(tree.len(), 2);
        assert_eq!(tree[0].text, "Alpha & Beta");
        assert_eq!(tree[0].depth, 0);
        assert_eq!(tree[1].text, "Gamma");
        assert_eq!(tree[1].children.len(), 1);
        assert_eq!(tree[1].children[0].text, "Child");
        assert_eq!(tree[1].children[0].depth, 1);
        assert_eq!(model.text(), "Alpha & Beta\nGamma\nChild");
    }

    #[test]
    fn parses_strict_namespace_model() {
        let model = DiagramDataModel::parse(STRICT).unwrap();
        assert_eq!(model.points.len(), 3);
        assert_eq!(model.points[2].point_type, DiagramPointType::SibTrans);
        assert_eq!(
            model.document_point().unwrap().layout_type_id.as_deref(),
            Some("urn:test/layout/cycle2")
        );
        let tree = model.node_tree();
        assert_eq!(tree.len(), 1);
        assert_eq!(tree[0].text, "a");
    }

    #[test]
    fn tolerates_cycles_and_dangling_connections() {
        let xml = concat!(
            "<dgm:dataModel xmlns:dgm=\"http://schemas.openxmlformats.org/drawingml/2006/diagram\">",
            "<dgm:ptLst><dgm:pt modelId=\"0\" type=\"doc\"/><dgm:pt modelId=\"1\"/><dgm:pt modelId=\"2\"/></dgm:ptLst>",
            "<dgm:cxnLst>",
            "<dgm:cxn modelId=\"10\" srcId=\"0\" destId=\"1\" srcOrd=\"0\" destOrd=\"0\"/>",
            "<dgm:cxn modelId=\"11\" srcId=\"1\" destId=\"2\" srcOrd=\"0\" destOrd=\"0\"/>",
            "<dgm:cxn modelId=\"12\" srcId=\"2\" destId=\"1\" srcOrd=\"0\" destOrd=\"0\"/>",
            "<dgm:cxn modelId=\"13\" srcId=\"0\" destId=\"9\" srcOrd=\"1\" destOrd=\"0\"/>",
            "</dgm:cxnLst></dgm:dataModel>"
        );
        let model = DiagramDataModel::parse(xml).unwrap();
        let tree = model.node_tree();
        assert_eq!(tree.len(), 1);
        assert_eq!(tree[0].children.len(), 1);
        assert!(tree[0].children[0].children.is_empty());
    }

    #[test]
    fn rejects_wrong_root_and_dtd() {
        assert!(
            DiagramDataModel::parse(
                "<dgm:layoutDef xmlns:dgm=\"http://schemas.openxmlformats.org/drawingml/2006/diagram\"/>"
            )
            .is_err()
        );
        assert!(
            DiagramDataModel::parse(
                "<!DOCTYPE dgm:dataModel><dgm:dataModel xmlns:dgm=\"http://schemas.openxmlformats.org/drawingml/2006/diagram\"/>"
            )
            .is_err()
        );
    }

    #[test]
    fn rejects_missing_ids() {
        assert!(
            DiagramDataModel::parse(
                "<dgm:dataModel xmlns:dgm=\"http://schemas.openxmlformats.org/drawingml/2006/diagram\"><dgm:ptLst><dgm:pt type=\"doc\"/></dgm:ptLst></dgm:dataModel>"
            )
            .is_err()
        );
        assert!(
            DiagramDataModel::parse(
                "<dgm:dataModel xmlns:dgm=\"http://schemas.openxmlformats.org/drawingml/2006/diagram\"><dgm:cxnLst><dgm:cxn modelId=\"1\" srcId=\"0\"/></dgm:cxnLst></dgm:dataModel>"
            )
            .is_err()
        );
    }
}
