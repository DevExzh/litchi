//! Source-preserving reads and edits of a `SmartArt` data-model part
//! (record 0284).
//!
//! [`Outline::read`] projects a `dgm:dataModel` onto its content nodes in
//! document order (the `parOf` connection tree from the document point,
//! siblings by `srcOrd`), with their text, the layout, quick-style and color
//! identifiers, and the relationship of the cached drawing
//! (`dsp:dataModelExt/@relId`). The edits rewrite only the bytes they change:
//!
//! - [`set_text`] replaces one node's text body and keeps its body and list
//!   properties, its first paragraph's properties and its first run's
//!   properties;
//! - [`drawing_text`] makes the same change in the cached drawing when every
//!   drawn shape of the node shows exactly that node's text, and answers
//!   `None` otherwise, so the caller removes the drawing instead of leaving
//!   it inconsistent with the data model;
//! - [`set_drawing_reference`] names, renames or removes the cached-drawing
//!   relationship the data model records.
//!
//! Both the transitional and the ISO Strict namespaces are read; written
//! elements use the namespaces the part already uses.

use std::collections::{HashMap, HashSet};

use litchi_core::xml::escape_xml;
use litchi_ooxml_common::xml::attributes::BytesStartExt as _;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;

pub use super::Limits;
use super::{DGM_NAMESPACE, DGM_NAMESPACE_STRICT};
use crate::{Error, Result};

/// Largest data-model or drawing part these functions accept.
pub const MAX_PART_BYTES: usize = 16 * 1024 * 1024;
/// Most XML element nodes a source-preserving diagram operation inspects.
pub const MAX_XML_NODES: usize = 200_000;
/// Most connection records a source-preserving operation inspects.
pub const MAX_CONNECTIONS: usize = 100_000;
/// Most content nodes an outline holds.
pub const MAX_NODES: usize = 10_000;
/// Deepest node an outline holds (zero-based).
pub const MAX_DEPTH: u32 = 64;

const DML: &[u8] = b"http://schemas.openxmlformats.org/drawingml/2006/main";
const DML_STRICT: &[u8] = b"http://purl.oclc.org/ooxml/drawingml/main";
/// The Microsoft diagram-drawing namespace (`dsp:`), also the `uri` of the
/// extension that holds `dsp:dataModelExt`.
pub const DRAWING_NAMESPACE: &str = "http://schemas.microsoft.com/office/drawing/2008/diagram";

/// One content node of an [`Outline`], in document order.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Entry {
    /// The point's `modelId`, as written.
    pub id: String,
    /// The node's text: its paragraphs joined with `\n`.
    pub text: String,
    /// The index of the parent entry; `None` for a top-level node.
    pub parent: Option<usize>,
    /// Zero for a top-level node.
    pub depth: u32,
}

/// The content of a `SmartArt` data model.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct Outline {
    /// The document point's `modelId`.
    pub document: String,
    /// `loTypeId` of the document point.
    pub layout_type_id: Option<String>,
    /// `qsTypeId` of the document point.
    pub quick_style_type_id: Option<String>,
    /// `csTypeId` of the document point.
    pub color_type_id: Option<String>,
    /// The host relationship of the cached drawing, if the model names one.
    pub drawing_relationship: Option<String>,
    /// The content nodes in document order (depth first).
    pub entries: Vec<Entry>,
}

impl Outline {
    /// Read the outline of a data-model part.
    ///
    /// # Errors
    ///
    /// Returns an error for malformed XML, a part that is not a diagram data
    /// model, a model without a document point, or exceeded bounds.
    pub fn read(xml: &[u8], limits: &Limits) -> Result<Self> {
        limits.validate()?;
        Ok(Parsed::read(xml, limits)?.outline)
    }

    /// The indexes of the children of `parent` (`None`: the top level).
    pub fn children(&self, parent: Option<usize>) -> impl Iterator<Item = usize> + '_ {
        self.entries
            .iter()
            .enumerate()
            .filter(move |(_, entry)| entry.parent == parent)
            .map(|(index, _)| index)
    }

    /// The index of the entry with `modelId` `id`.
    #[must_use]
    pub fn position(&self, id: &str) -> Option<usize> {
        self.entries.iter().position(|entry| entry.id == id)
    }
}

/// A point of the data model, with its byte spans.
#[derive(Debug, Clone, Default)]
struct PointRecord {
    id: String,
    kind: String,
    /// The whole `pt` element.
    span: (usize, usize),
    /// End of the `pt` start tag (equal to `span.1` for an empty element).
    open_end: usize,
    /// The `t` element, if any.
    text: Option<TextBody>,
    pres_assoc: Option<String>,
}

/// A connection of the data model.
#[derive(Debug, Clone, Default)]
struct ConnectionRecord {
    kind: String,
    source: String,
    destination: String,
    order: u32,
}

/// A text body (`dgm:t` or `dsp:txBody`): its span and the reusable pieces.
#[derive(Debug, Clone, Default)]
struct TextBody {
    span: (usize, usize),
    /// End of the start tag.
    open_end: usize,
    /// Start of the end tag (equal to `span.1` for an empty element).
    close_start: usize,
    /// Raw `bodyPr` and `lstStyle` elements.
    preamble: Vec<(usize, usize)>,
    /// Raw `pPr` of the first paragraph.
    paragraph_properties: Option<(usize, usize)>,
    /// Raw `rPr` (or `endParaRPr`, as a fallback) of the first run.
    run_properties: Option<(usize, usize)>,
    run_properties_is_end: bool,
    /// The prefix of the paragraph elements (`a` by default).
    prefix: Option<String>,
    paragraphs: Vec<String>,
}

impl TextBody {
    fn text(&self) -> String {
        self.paragraphs.join("\n")
    }
}

struct Parsed {
    outline: Outline,
    points: Vec<PointRecord>,
    connections: Vec<ConnectionRecord>,
    /// The `a:ext` element holding `dsp:dataModelExt`, and its `relId`
    /// attribute value span.
    drawing_ext: Option<((usize, usize), (usize, usize))>,
    /// The `extLst` element of the data model.
    ext_list: Option<(usize, usize, usize)>,
    /// Start of the root end tag.
    root_close: usize,
    /// The root element's qualified-name prefix (`dgm`), if any.
    root_prefix: Option<String>,
    strict: bool,
}

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}

fn xml_error(error: impl std::fmt::Display) -> Error {
    Error::Xml(error.to_string())
}

/// The namespaces these functions tell apart.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Ns {
    Dgm,
    DgmStrict,
    Dml,
    DmlStrict,
    Drawing,
    Other,
}

fn classify(resolved: &ResolveResult<'_>) -> Ns {
    match resolved {
        ResolveResult::Bound(Namespace(value)) => match *value {
            value if value == DGM_NAMESPACE.as_bytes() => Ns::Dgm,
            value if value == DGM_NAMESPACE_STRICT.as_bytes() => Ns::DgmStrict,
            DML => Ns::Dml,
            DML_STRICT => Ns::DmlStrict,
            value if value == DRAWING_NAMESPACE.as_bytes() => Ns::Drawing,
            _ => Ns::Other,
        },
        _ => Ns::Other,
    }
}

fn is_dgm(ns: Ns) -> bool {
    matches!(ns, Ns::Dgm | Ns::DgmStrict)
}

fn is_dml(ns: Ns) -> bool {
    matches!(ns, Ns::Dml | Ns::DmlStrict)
}

/// Unqualified, unescaped attributes of `element`.
fn attributes(element: &BytesStart<'_>) -> Result<Vec<(String, String)>> {
    let mut values = Vec::new();
    for attribute in element.checked_attributes() {
        let attribute = attribute.map_err(xml_error)?;
        let raw = attribute.key.as_ref();
        if raw.contains(&b':') || raw == b"xmlns" {
            continue;
        }
        let name = std::str::from_utf8(raw).map_err(xml_error)?.to_owned();
        let value = std::str::from_utf8(attribute.value.as_ref()).map_err(xml_error)?;
        let value = quick_xml::escape::unescape(value)
            .map_err(xml_error)?
            .into_owned();
        values.push((name, value));
    }
    Ok(values)
}

fn attribute<'v>(values: &'v [(String, String)], name: &str) -> Option<&'v str> {
    values
        .iter()
        .find(|(key, _)| key == name)
        .map(|(_, value)| value.as_str())
}

fn prefix_of(element: &BytesStart<'_>) -> Option<String> {
    let name = element.name();
    name.prefix()
        .map(|prefix| String::from_utf8_lossy(prefix.as_ref()).into_owned())
}

/// Text-body collection state while a `t`/`txBody` element is open.
struct BodyState {
    body: TextBody,
    depth: usize,
    in_paragraph: bool,
    paragraph: String,
    in_text: bool,
    first_paragraph_done: bool,
    text_bytes: usize,
}

impl BodyState {
    fn new(start: usize, open_end: usize) -> Self {
        Self {
            body: TextBody {
                span: (start, open_end),
                open_end,
                close_start: open_end,
                ..TextBody::default()
            },
            depth: 0,
            in_paragraph: false,
            paragraph: String::new(),
            in_text: false,
            first_paragraph_done: false,
            text_bytes: 0,
        }
    }

    /// One event inside the body (at body-relative `depth` after a start
    /// event, before an end event).
    fn element(&mut self, ns: Ns, element: &BytesStart<'_>, span: (usize, usize), empty: bool) {
        if !is_dml(ns) {
            return;
        }
        let local = element.local_name();
        let local = local.as_ref();
        match (self.depth, local) {
            (1, b"bodyPr" | b"lstStyle") if empty => self.body.preamble.push(span),
            (1, b"p") => {
                if self.body.prefix.is_none() {
                    self.body.prefix = Some(prefix_of(element).unwrap_or_default());
                }
                if !empty {
                    self.in_paragraph = true;
                    self.paragraph.clear();
                } else {
                    self.body.paragraphs.push(String::new());
                    self.first_paragraph_done = true;
                }
            },
            (2, b"pPr") if self.in_paragraph && !self.first_paragraph_done && empty => {
                self.body.paragraph_properties = Some(span);
            },
            (2, b"br") if self.in_paragraph => self.paragraph.push('\n'),
            (2, b"endParaRPr")
                if self.in_paragraph
                    && !self.first_paragraph_done
                    && empty
                    && self.body.run_properties.is_none() =>
            {
                self.body.run_properties = Some(span);
                self.body.run_properties_is_end = true;
            },
            (3, b"rPr")
                if self.in_paragraph
                    && !self.first_paragraph_done
                    && empty
                    && (self.body.run_properties.is_none() || self.body.run_properties_is_end) =>
            {
                self.body.run_properties = Some(span);
                self.body.run_properties_is_end = false;
            },
            (3, b"t") if self.in_paragraph && !empty => self.in_text = true,
            _ => {},
        }
    }

    fn end(&mut self, ns: Ns, local: &[u8]) {
        if !is_dml(ns) {
            return;
        }
        match (self.depth, local) {
            (1, b"p") if self.in_paragraph => {
                self.in_paragraph = false;
                self.first_paragraph_done = true;
                self.body
                    .paragraphs
                    .push(std::mem::take(&mut self.paragraph));
            },
            (3, b"t") => self.in_text = false,
            _ => {},
        }
    }
}

/// Elements whose full span (start to end) a caller needs: `bodyPr`,
/// `lstStyle`, `pPr`, `rPr` may have children; this records their extent.
struct Pending {
    depth: usize,
    start: usize,
    kind: u8,
}

const PENDING_PREAMBLE: u8 = 0;
const PENDING_PPR: u8 = 1;
const PENDING_RPR: u8 = 2;
const PENDING_END_RPR: u8 = 3;

/// Parse a text body element that starts at `start` (its start tag already
/// read, ending at `open_end`) through its end tag.
fn read_body<R: std::io::BufRead>(
    reader: &mut NsReader<R>,
    buffer: &mut Vec<u8>,
    start: usize,
    open_end: usize,
    base_depth: usize,
    xml_nodes: &mut usize,
    limits: &Limits,
) -> Result<TextBody> {
    let mut state = BodyState::new(start, open_end);
    let mut pending: Option<Pending> = None;
    state.depth = 0;
    loop {
        buffer.clear();
        let event_start = position(reader)?;
        let (resolved, event) = reader.read_resolved_event_into(buffer).map_err(xml_error)?;
        let ns = classify(&resolved);
        let event_end = position(reader)?;
        match event {
            Event::Start(element) => {
                account_xml_node(xml_nodes, limits)?;
                state.depth = state
                    .depth
                    .checked_add(1)
                    .ok_or_else(|| invalid("diagram text-body depth overflows"))?;
                let total_depth = base_depth
                    .checked_add(state.depth)
                    .ok_or_else(|| invalid("diagram XML depth overflows"))?;
                let max_depth = limits.xml_depth_value().min(MAX_DEPTH as usize);
                if total_depth > max_depth {
                    return Err(Error::limit(
                        crate::LimitKind::XmlDepth,
                        total_depth,
                        max_depth,
                    ));
                }
                let local = element.local_name();
                if pending.is_none() && is_dml(ns) {
                    let kind = match (state.depth, local.as_ref()) {
                        (1, b"bodyPr" | b"lstStyle") => Some(PENDING_PREAMBLE),
                        (2, b"pPr") if state.in_paragraph && !state.first_paragraph_done => {
                            Some(PENDING_PPR)
                        },
                        (2, b"endParaRPr")
                            if state.in_paragraph
                                && !state.first_paragraph_done
                                && state.body.run_properties.is_none() =>
                        {
                            Some(PENDING_END_RPR)
                        },
                        (3, b"rPr")
                            if state.in_paragraph
                                && !state.first_paragraph_done
                                && (state.body.run_properties.is_none()
                                    || state.body.run_properties_is_end) =>
                        {
                            Some(PENDING_RPR)
                        },
                        _ => None,
                    };
                    if let Some(kind) = kind {
                        pending = Some(Pending {
                            depth: state.depth,
                            start: event_start,
                            kind,
                        });
                    }
                }
                state.element(ns, &element, (event_start, event_end), false);
            },
            Event::Empty(element) => {
                account_xml_node(xml_nodes, limits)?;
                let child_depth = state
                    .depth
                    .checked_add(1)
                    .ok_or_else(|| invalid("diagram text-body depth overflows"))?;
                let total_depth = base_depth
                    .checked_add(child_depth)
                    .ok_or_else(|| invalid("diagram XML depth overflows"))?;
                let max_depth = limits.xml_depth_value().min(MAX_DEPTH as usize);
                if total_depth > max_depth {
                    return Err(Error::limit(
                        crate::LimitKind::XmlDepth,
                        total_depth,
                        max_depth,
                    ));
                }
                state.depth = child_depth;
                state.element(ns, &element, (event_start, event_end), true);
                state.depth -= 1;
            },
            Event::End(element) => {
                if state.depth == 0 {
                    state.body.close_start = event_start;
                    state.body.span.1 = event_end;
                    return Ok(state.body);
                }
                if let Some(open) = pending.as_ref()
                    && open.depth == state.depth
                {
                    let span = (open.start, event_end);
                    match open.kind {
                        PENDING_PREAMBLE => state.body.preamble.push(span),
                        PENDING_PPR => state.body.paragraph_properties = Some(span),
                        PENDING_END_RPR => {
                            state.body.run_properties = Some(span);
                            state.body.run_properties_is_end = true;
                        },
                        _ => {
                            state.body.run_properties = Some(span);
                            state.body.run_properties_is_end = false;
                        },
                    }
                    pending = None;
                }
                let local = element.local_name();
                state.end(ns, local.as_ref());
                state.depth -= 1;
            },
            Event::Text(text) if state.in_text => {
                let text = text.decode().map_err(xml_error)?;
                let text = quick_xml::escape::unescape(&text).map_err(xml_error)?;
                let max_text = limits.node_text_bytes_value().min(1024 * 1024);
                let observed = state.text_bytes.saturating_add(text.len());
                if observed > max_text {
                    return Err(Error::limit(
                        crate::LimitKind::DiagramNodeTextBytes,
                        observed,
                        max_text,
                    ));
                }
                state.text_bytes = observed;
                state.paragraph.push_str(&text);
            },
            Event::GeneralRef(reference) if state.in_text => {
                let name = reference.decode().map_err(xml_error)?;
                let entity = format!("&{name};");
                let text = quick_xml::escape::unescape(&entity).map_err(xml_error)?;
                let max_text = limits.node_text_bytes_value().min(1024 * 1024);
                let observed = state.text_bytes.saturating_add(text.len());
                if observed > max_text {
                    return Err(Error::limit(
                        crate::LimitKind::DiagramNodeTextBytes,
                        observed,
                        max_text,
                    ));
                }
                state.text_bytes = observed;
                state.paragraph.push_str(&text);
            },
            Event::CData(data) if state.in_text => {
                let max_text = limits.node_text_bytes_value().min(1024 * 1024);
                let text = String::from_utf8_lossy(data.as_ref());
                let observed = state.text_bytes.saturating_add(text.len());
                if observed > max_text {
                    return Err(Error::limit(
                        crate::LimitKind::DiagramNodeTextBytes,
                        observed,
                        max_text,
                    ));
                }
                state.text_bytes = observed;
                state.paragraph.push_str(&text);
            },
            Event::Eof => return Err(invalid("unterminated text body")),
            Event::DocType(_) => return Err(invalid("document type in a diagram part")),
            _ => {},
        }
    }
}

fn position<R>(reader: &NsReader<R>) -> Result<usize> {
    usize::try_from(reader.buffer_position()).map_err(|_| invalid("XML offset exceeds usize"))
}

fn account_xml_node(nodes: &mut usize, limits: &Limits) -> Result<()> {
    *nodes = nodes
        .checked_add(1)
        .ok_or_else(|| invalid("diagram XML node count overflows"))?;
    let maximum = limits.xml_nodes_value().min(MAX_XML_NODES);
    if *nodes > maximum {
        return Err(Error::limit(crate::LimitKind::XmlNodes, *nodes, maximum));
    }
    Ok(())
}

impl Parsed {
    fn read(xml: &[u8], limits: &Limits) -> Result<Self> {
        let max_input = limits.input_bytes_value().min(MAX_PART_BYTES);
        if xml.len() > max_input {
            return Err(Error::limit(
                crate::LimitKind::DiagramDataModelBytes,
                xml.len(),
                max_input,
            ));
        }
        let mut reader = NsReader::from_reader(xml);
        reader.config_mut().trim_text(false);
        let mut buffer = Vec::new();
        let mut depth = 0_usize;
        let mut xml_nodes = 0_usize;
        // Names of the open elements, as (dgm local name or empty).
        let mut stack: Vec<Vec<u8>> = Vec::new();
        let mut points = Vec::new();
        let mut connections = Vec::new();
        let mut current: Option<PointRecord> = None;
        let mut document_attributes = None;
        let mut drawing_ext = None;
        let mut ext_start: Option<usize> = None;
        let mut ext_list = None;
        let mut ext_list_start = None;
        let mut root_close = None;
        let mut root_prefix = None;
        let mut strict = false;
        let mut root_seen = false;
        let mut root_closed = false;
        let mut pending_rel: Option<(usize, usize)> = None;
        loop {
            buffer.clear();
            let event_start = position(&reader)?;
            let (resolved, event) = reader
                .read_resolved_event_into(&mut buffer)
                .map_err(xml_error)?;
            let ns = classify(&resolved);
            let event_end = position(&reader)?;
            match event {
                Event::Start(ref element) | Event::Empty(ref element) => {
                    let empty = matches!(event, Event::Empty(_));
                    account_xml_node(&mut xml_nodes, limits)?;
                    let event_depth = depth.saturating_add(1);
                    let max_depth = limits.xml_depth_value().min(MAX_DEPTH as usize);
                    if event_depth > max_depth {
                        return Err(Error::limit(
                            crate::LimitKind::XmlDepth,
                            event_depth,
                            max_depth,
                        ));
                    }
                    let local_name = element.local_name();
                    let local = local_name.as_ref().to_vec();
                    let dgm = is_dgm(ns);
                    if root_closed {
                        return Err(invalid("trailing markup follows the diagram data model"));
                    }
                    if depth == 0 {
                        if root_seen {
                            return Err(invalid("the diagram data model has multiple roots"));
                        }
                        if !dgm || local != b"dataModel" {
                            return Err(invalid("the part is not a diagram data model"));
                        }
                        root_seen = true;
                        strict = ns == Ns::DgmStrict;
                        root_prefix = prefix_of(element);
                        if empty {
                            return Err(invalid("the diagram data model is empty"));
                        }
                    } else if dgm && local == b"dataModel" {
                        return Err(invalid("the diagram data model has a nested root"));
                    }
                    let parent = stack.last().map(Vec::as_slice);
                    if dgm && local == b"pt" && parent == Some(b"ptLst".as_slice()) {
                        let values = attributes(element)?;
                        let record = PointRecord {
                            id: attribute(&values, "modelId").unwrap_or_default().to_owned(),
                            kind: attribute(&values, "type").unwrap_or("node").to_owned(),
                            span: (event_start, event_end),
                            open_end: event_end,
                            ..PointRecord::default()
                        };
                        let max_points = limits.points_value().min(MAX_NODES * 16);
                        if points.len() >= max_points {
                            return Err(Error::limit(
                                crate::LimitKind::DiagramPoints,
                                points.len() + 1,
                                max_points,
                            ));
                        }
                        if empty {
                            points.push(record);
                        } else {
                            current = Some(record);
                        }
                    } else if dgm && local == b"prSet" && parent == Some(b"pt".as_slice()) {
                        let values = attributes(element)?;
                        if let Some(point) = current.as_mut() {
                            point.pres_assoc = attribute(&values, "presAssocID").map(str::to_owned);
                            if point.kind == "doc" {
                                document_attributes = Some(values);
                            }
                        }
                    } else if dgm
                        && local == b"t"
                        && parent == Some(b"pt".as_slice())
                        && current.is_some()
                    {
                        let body = if empty {
                            TextBody {
                                span: (event_start, event_end),
                                open_end: event_end,
                                close_start: event_end,
                                ..TextBody::default()
                            }
                        } else {
                            read_body(
                                &mut reader,
                                &mut buffer,
                                event_start,
                                event_end,
                                event_depth,
                                &mut xml_nodes,
                                limits,
                            )?
                        };
                        if let Some(point) = current.as_mut() {
                            point.text = Some(body);
                        }
                        continue;
                    } else if dgm && local == b"cxn" && parent == Some(b"cxnLst".as_slice()) {
                        let max_connections = limits.connections_value().min(MAX_CONNECTIONS);
                        if connections.len() >= max_connections {
                            return Err(Error::limit(
                                crate::LimitKind::DiagramConnections,
                                connections.len() + 1,
                                max_connections,
                            ));
                        }
                        let values = attributes(element)?;
                        connections.push(ConnectionRecord {
                            kind: attribute(&values, "type").unwrap_or("parOf").to_owned(),
                            source: attribute(&values, "srcId").unwrap_or_default().to_owned(),
                            destination: attribute(&values, "destId")
                                .unwrap_or_default()
                                .to_owned(),
                            order: attribute(&values, "srcOrd")
                                .and_then(|value| value.trim().parse().ok())
                                .unwrap_or(0),
                        });
                    } else if dgm && local == b"extLst" && depth == 1 {
                        ext_list_start = Some((event_start, event_end));
                        if empty {
                            ext_list = Some((event_start, event_end, event_end));
                        }
                    } else if local == b"ext" && depth == 2 && ext_list_start.is_some() {
                        ext_start = Some(event_start);
                    } else if local == b"dataModelExt" && ns == Ns::Drawing && ext_start.is_some() {
                        pending_rel = relationship_value_span(element, event_start);
                    }
                    if !empty {
                        stack.push(if dgm { local } else { Vec::new() });
                        depth += 1;
                    }
                },
                Event::End(ref element) => {
                    if depth == 0 || stack.is_empty() {
                        return Err(invalid("unexpected diagram data model end"));
                    }
                    let local_name = element.local_name();
                    let local = local_name.as_ref().to_vec();
                    if depth == 1 && (!is_dgm(ns) || local != b"dataModel") {
                        return Err(invalid("the diagram data model root is not closed"));
                    }
                    stack.pop();
                    depth -= 1;
                    if depth == 0 {
                        root_closed = true;
                        root_close = Some(event_start);
                    }
                    if is_dgm(ns) && local == b"pt" && depth == 2 {
                        if let Some(mut point) = current.take() {
                            point.span.1 = event_end;
                            points.push(point);
                        }
                    } else if local == b"ext" && depth == 2 {
                        if let (Some(start), Some(rel)) = (ext_start.take(), pending_rel.take()) {
                            drawing_ext = Some(((start, event_end), rel));
                        }
                    } else if is_dgm(ns) && local == b"extLst" && depth == 1 {
                        if let Some((start, open_end)) = ext_list_start {
                            ext_list = Some((start, event_end, open_end));
                        }
                    }
                },
                Event::Eof => break,
                Event::DocType(_) => return Err(invalid("document type in a diagram part")),
                Event::Text(text) => {
                    if text.as_ref().iter().any(|byte| !byte.is_ascii_whitespace()) {
                        return Err(invalid(
                            "non-whitespace text is outside the diagram data model",
                        ));
                    }
                },
                Event::CData(data) => {
                    if data.as_ref().iter().any(|byte| !byte.is_ascii_whitespace()) {
                        return Err(invalid("CDATA is outside the diagram data model"));
                    }
                },
                Event::GeneralRef(_) => {
                    return Err(invalid(
                        "entity reference is outside the diagram data model",
                    ));
                },
                _ => {},
            }
        }
        if !root_seen || !root_closed || depth != 0 || !stack.is_empty() {
            return Err(invalid("unterminated diagram data model"));
        }
        let root_close = root_close.ok_or_else(|| invalid("unterminated diagram data model"))?;
        let document = points
            .iter()
            .find(|point| point.kind == "doc")
            .ok_or_else(|| invalid("the diagram data model has no document point"))?
            .id
            .clone();
        let document_attributes = document_attributes.unwrap_or_default();
        let drawing_relationship = drawing_ext.map(|(_, (start, end))| {
            quick_xml::escape::unescape(&String::from_utf8_lossy(&xml[start..end])).map_or_else(
                |_| String::from_utf8_lossy(&xml[start..end]).into_owned(),
                std::borrow::Cow::into_owned,
            )
        });
        let entries = tree(&document, &points, &connections, limits)?;
        Ok(Self {
            outline: Outline {
                document,
                layout_type_id: attribute(&document_attributes, "loTypeId").map(str::to_owned),
                quick_style_type_id: attribute(&document_attributes, "qsTypeId").map(str::to_owned),
                color_type_id: attribute(&document_attributes, "csTypeId").map(str::to_owned),
                drawing_relationship,
                entries,
            },
            points,
            connections,
            drawing_ext,
            ext_list,
            root_close,
            root_prefix,
            strict,
        })
    }

    fn point(&self, id: &str) -> Option<&PointRecord> {
        self.points.iter().find(|point| point.id == id)
    }

    fn a_namespace(&self) -> &'static str {
        if self.strict {
            "http://purl.oclc.org/ooxml/drawingml/main"
        } else {
            "http://schemas.openxmlformats.org/drawingml/2006/main"
        }
    }
}

/// The span of the `relId` attribute value of a `dsp:dataModelExt` tag that
/// starts at `start` in the part.
fn relationship_value_span(element: &BytesStart<'_>, start: usize) -> Option<(usize, usize)> {
    // The tag text is `<` + name + attributes; locate `relId="` in the raw
    // attribute bytes.
    let raw: &[u8] = element.as_ref();
    let mut index = 0;
    while let Some(found) = find(&raw[index..], b"relId") {
        let at = index + found;
        let before_ok = at == 0 || raw[at - 1].is_ascii_whitespace();
        let mut cursor = at + 5;
        while cursor < raw.len() && raw[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        if before_ok && raw.get(cursor) == Some(&b'=') {
            cursor += 1;
            while cursor < raw.len() && raw[cursor].is_ascii_whitespace() {
                cursor += 1;
            }
            let quote = *raw.get(cursor)?;
            if quote != b'"' && quote != b'\'' {
                return None;
            }
            let value_start = cursor + 1;
            let value_end = value_start + raw[value_start..].iter().position(|&b| b == quote)?;
            // `raw` excludes the leading `<`.
            return Some((start + 1 + value_start, start + 1 + value_end));
        }
        index = at + 5;
    }
    None
}

fn find(haystack: &[u8], needle: &[u8]) -> Option<usize> {
    haystack
        .windows(needle.len())
        .position(|window| window == needle)
}

/// The content-node tree, depth first from the document point.
fn tree(
    document: &str,
    points: &[PointRecord],
    connections: &[ConnectionRecord],
    limits: &Limits,
) -> Result<Vec<Entry>> {
    let content: HashMap<&str, &PointRecord> = points
        .iter()
        .filter(|point| matches!(point.kind.as_str(), "node" | "asst"))
        .map(|point| (point.id.as_str(), point))
        .collect();
    let mut children: HashMap<&str, Vec<(u32, usize, &str)>> = HashMap::new();
    for (index, connection) in connections.iter().enumerate() {
        if connection.kind == "parOf" && content.contains_key(connection.destination.as_str()) {
            children
                .entry(connection.source.as_str())
                .or_default()
                .push((connection.order, index, connection.destination.as_str()));
        }
    }
    for list in children.values_mut() {
        list.sort_unstable();
    }
    let mut entries = Vec::new();
    let mut seen = HashSet::new();
    // (point id, parent entry, depth)
    let max_nodes = limits.nodes_value().min(MAX_NODES);
    let root_children = children.get(document).map_or(&[][..], Vec::as_slice);
    if root_children.len() > max_nodes {
        return Err(Error::limit(
            crate::LimitKind::DiagramNodes,
            max_nodes + 1,
            max_nodes,
        ));
    }
    let mut stack = Vec::with_capacity(root_children.len());
    for (_, _, id) in root_children.iter().rev() {
        stack.push((*id, None, 0));
    }
    while let Some((id, parent, depth)) = stack.pop() {
        if !seen.insert(id) {
            continue;
        }
        let max_depth = limits.node_depth_value().min(MAX_DEPTH as usize);
        if usize::try_from(depth).unwrap_or(usize::MAX) > max_depth {
            let observed =
                usize::try_from(depth).map_err(|_| invalid("diagram depth exceeds usize"))?;
            return Err(Error::limit(
                crate::LimitKind::DiagramNodeDepth,
                observed,
                max_depth,
            ));
        }
        if entries.len() >= max_nodes {
            return Err(Error::limit(
                crate::LimitKind::DiagramNodes,
                entries.len() + 1,
                max_nodes,
            ));
        }
        let point = content[id];
        let index = entries.len();
        entries.push(Entry {
            id: id.to_owned(),
            text: point.text.as_ref().map(TextBody::text).unwrap_or_default(),
            parent,
            depth,
        });
        if let Some(list) = children.get(id) {
            for (_, _, child) in list.iter().rev() {
                if stack.len() >= max_nodes {
                    return Err(Error::limit(
                        crate::LimitKind::DiagramNodes,
                        max_nodes + 1,
                        max_nodes,
                    ));
                }
                stack.push((child, Some(index), depth + 1));
            }
        }
    }
    Ok(entries)
}

/// Check that `text` holds only XML characters.
fn check_text(text: &str, limits: &Limits) -> Result<()> {
    if text.chars().any(|character| {
        !matches!(character, '\t' | '\n' | '\r' | '\u{20}'..='\u{D7FF}' | '\u{E000}'..='\u{FFFD}' | '\u{10000}'..)
    }) {
        return Err(invalid("diagram text holds a character XML cannot carry"));
    }
    let max_text = limits.node_text_bytes_value().min(1024 * 1024);
    if text.len() > max_text {
        return Err(Error::limit(
            crate::LimitKind::DiagramNodeTextBytes,
            text.len(),
            max_text,
        ));
    }
    Ok(())
}

fn ensure_output(
    input_len: usize,
    additional: usize,
    limits: &Limits,
    kind: crate::LimitKind,
) -> Result<usize> {
    let total = input_len
        .checked_add(additional)
        .ok_or_else(|| invalid("edited diagram output length overflows"))?;
    ensure_output_total(total, limits, kind)
}

fn ensure_output_total(total: usize, limits: &Limits, kind: crate::LimitKind) -> Result<usize> {
    let maximum = limits.output_bytes_value().min(MAX_PART_BYTES);
    if total > maximum {
        return Err(Error::limit(kind, total, maximum));
    }
    Ok(total)
}

fn escaped_xml_len(text: &str) -> Result<usize> {
    text.as_bytes().iter().try_fold(0usize, |length, byte| {
        let replacement = match byte {
            b'&' => 5,
            b'<' | b'>' => 4,
            b'"' | b'\'' => 6,
            _ => 1,
        };
        length
            .checked_add(replacement)
            .ok_or_else(|| invalid("escaped diagram text length overflows"))
    })
}

fn span_len(xml: &[u8], span: (usize, usize)) -> Result<usize> {
    let (start, end) = span;
    if start > end || end > xml.len() {
        return Err(invalid("diagram source span is outside the input"));
    }
    Ok(end - start)
}

/// Upper bound for the text body emitted by [`body_content`]. The source
/// spans are counted before the replacement buffer is allocated, so retained
/// properties cannot hide output that would exceed the caller's budget.
fn body_content_bound(xml: &[u8], body: &TextBody, text: &str, a_namespace: &str) -> Result<usize> {
    let prefix = body.prefix.as_deref().unwrap_or("a");
    let qualified = |local: &str| {
        if prefix.is_empty() {
            local.to_owned()
        } else {
            format!("{prefix}:{local}")
        }
    };
    let qualified_len = |local: &str| qualified(local).len();
    let declaration = if prefix.is_empty() {
        format!(" xmlns=\"{a_namespace}\"")
    } else {
        format!(" xmlns:{prefix}=\"{a_namespace}\"")
    };
    let empty_body_properties = 1usize
        .checked_add(qualified_len("bodyPr"))
        .and_then(|length| length.checked_add(declaration.len()))
        .and_then(|length| length.checked_add(2))
        .and_then(|length| {
            1usize
                .checked_add(qualified_len("lstStyle"))
                .and_then(|value| value.checked_add(declaration.len()))
                .and_then(|value| value.checked_add(2))
                .and_then(|value| length.checked_add(value))
        })
        .ok_or_else(|| invalid("diagram text-body length overflows"))?;
    let mut length = if body.preamble.is_empty() {
        empty_body_properties
    } else {
        0
    };
    for span in &body.preamble {
        length = length
            .checked_add(span_len(xml, *span)?)
            .ok_or_else(|| invalid("diagram text-body length overflows"))?;
    }
    let paragraph_properties = body
        .paragraph_properties
        .map(|span| span_len(xml, span))
        .transpose()?;
    let run_properties = match body.run_properties {
        Some(span) => Some(
            span_len(xml, span)?
                .checked_add(if body.run_properties_is_end { 0 } else { 7 })
                .ok_or_else(|| invalid("diagram text-body length overflows"))?,
        ),
        None => None,
    };
    let default_end_para = format!("<{} lang=\"en-US\"/>", qualified("endParaRPr"));
    let default_run = format!("<{} lang=\"en-US\"/>", qualified("rPr"));
    for line in text.split('\n') {
        let line = line.strip_suffix('\r').unwrap_or(line);
        length = length
            .checked_add(1 + qualified_len("p") + declaration.len())
            .and_then(|value| value.checked_add(paragraph_properties.unwrap_or(0)))
            .ok_or_else(|| invalid("diagram text-body length overflows"))?;
        if line.is_empty() {
            length = length
                .checked_add(run_properties.unwrap_or(default_end_para.len()))
                .ok_or_else(|| invalid("diagram text-body length overflows"))?;
        } else {
            let run = 5 + 2 * qualified_len("r");
            let text_element = 1 + qualified_len("t") + 1 + 2 + qualified_len("t") + 1;
            length = length
                .checked_add(run)
                .and_then(|value| value.checked_add(run_properties.unwrap_or(default_run.len())))
                .and_then(|value| value.checked_add(text_element))
                .and_then(|value| value.checked_add(escaped_xml_len(line).ok()?))
                .ok_or_else(|| invalid("diagram text-body length overflows"))?;
        }
        length = length
            .checked_add(3 + qualified_len("p"))
            .ok_or_else(|| invalid("diagram text-body length overflows"))?;
    }
    Ok(length)
}

/// The new content of a text body: its kept preamble, then one paragraph per
/// line of `text` with the kept paragraph and run properties.
fn body_content(xml: &[u8], body: &TextBody, text: &str, a_namespace: &str) -> Result<String> {
    let capacity = body_content_bound(xml, body, text, a_namespace)?;
    let prefix = body.prefix.clone().unwrap_or_else(|| "a".to_owned());
    let qualified = |local: &str| {
        if prefix.is_empty() {
            local.to_owned()
        } else {
            format!("{prefix}:{local}")
        }
    };
    let declaration = if prefix.is_empty() {
        format!(" xmlns=\"{a_namespace}\"")
    } else {
        format!(" xmlns:{prefix}=\"{a_namespace}\"")
    };
    let raw = |(start, end): (usize, usize)| String::from_utf8_lossy(&xml[start..end]).into_owned();
    let mut out = String::with_capacity(capacity);
    if body.preamble.is_empty() {
        out.push_str(&format!(
            "<{}{declaration}/><{}{declaration}/>",
            qualified("bodyPr"),
            qualified("lstStyle")
        ));
    } else {
        for span in &body.preamble {
            out.push_str(&raw(*span));
        }
    }
    let paragraph_properties = body.paragraph_properties.map(raw);
    let run_properties = body.run_properties.map(|span| {
        let raw = raw(span);
        if body.run_properties_is_end {
            // `endParaRPr` and `rPr` share a type; rename the element.
            rename_element(&raw, "endParaRPr", "rPr")
        } else {
            raw
        }
    });
    for line in text.split('\n') {
        let line = line.strip_suffix('\r').unwrap_or(line);
        out.push_str(&format!("<{}{declaration}>", qualified("p")));
        if let Some(properties) = &paragraph_properties {
            out.push_str(properties);
        }
        if line.is_empty() {
            match &run_properties {
                Some(properties) => out.push_str(&rename_element(properties, "rPr", "endParaRPr")),
                None => out.push_str(&format!("<{} lang=\"en-US\"/>", qualified("endParaRPr"))),
            }
        } else {
            out.push_str(&format!("<{}>", qualified("r")));
            match &run_properties {
                Some(properties) => out.push_str(properties),
                None => out.push_str(&format!("<{} lang=\"en-US\"/>", qualified("rPr"))),
            }
            out.push_str(&format!(
                "<{t}>{}</{t}></{r}>",
                escape_xml(line),
                t = qualified("t"),
                r = qualified("r")
            ));
        }
        out.push_str(&format!("</{}>", qualified("p")));
    }
    Ok(out)
}

/// Rename the outermost element of `raw` from local name `from` to `to`.
fn rename_element(raw: &str, from: &str, to: &str) -> String {
    let mut out = raw.to_owned();
    // The start tag: `<prefix:from` or `<from`.
    if let Some(position) = out
        .find(&format!(":{from}"))
        .filter(|&p| p < out.find(['>', ' ', '/']).unwrap_or(usize::MAX).max(p + 1))
    {
        out.replace_range(position + 1..position + 1 + from.len(), to);
    } else if out.starts_with(&format!("<{from}")) {
        out.replace_range(1..1 + from.len(), to);
    }
    // The end tag, if any.
    if let Some(position) = out.rfind(&format!("{from}>")) {
        if out[..position].ends_with("</") || out[..position].ends_with(':') {
            out.replace_range(position..position + from.len(), to);
        }
    }
    out
}

/// Replace the text of the content node `id`.
///
/// The node keeps its body and list properties, the properties of its first
/// paragraph and those of its first run, applied to every new paragraph and
/// run; each line of `text` becomes a paragraph. Identical text returns the
/// input bytes.
///
/// # Errors
///
/// Returns an error for malformed XML, an unknown or non-content node, or
/// text XML cannot carry.
pub fn set_text(xml: &[u8], id: &str, text: &str, limits: &Limits) -> Result<Vec<u8>> {
    limits.validate()?;
    check_text(text, limits)?;
    let parsed = Parsed::read(xml, limits)?;
    let index = parsed
        .outline
        .position(id)
        .ok_or_else(|| invalid(format!("diagram node {id} was not found")))?;
    if parsed.outline.entries[index].text == text.replace("\r\n", "\n") {
        ensure_output(
            xml.len(),
            0,
            limits,
            crate::LimitKind::DiagramDataModelBytes,
        )?;
        return Ok(xml.to_vec());
    }
    let point = parsed
        .point(id)
        .ok_or_else(|| invalid(format!("diagram node {id} was not found")))?;
    let a_namespace = parsed.a_namespace();
    let dgm_prefix = parsed.root_prefix.clone();
    let output_capacity = match &point.text {
        Some(body) if body.open_end != body.span.1 => {
            let base = body
                .open_end
                .checked_add(
                    xml.len()
                        .checked_sub(body.close_start)
                        .ok_or_else(|| invalid("diagram text-body span is outside the input"))?,
                )
                .ok_or_else(|| invalid("edited diagram output length overflows"))?;
            ensure_output(
                base,
                body_content_bound(xml, body, text, a_namespace)?,
                limits,
                crate::LimitKind::DiagramDataModelBytes,
            )?
        },
        Some(body) => {
            let span = span_len(xml, body.span)?;
            let content_bound = body_content_bound(xml, &TextBody::default(), text, a_namespace)?;
            let replacement = span
                .checked_mul(2)
                .and_then(|length| length.checked_add(1))
                .and_then(|length| length.checked_add(content_bound))
                .ok_or_else(|| invalid("edited diagram output length overflows"))?;
            ensure_output(
                xml.len()
                    .checked_sub(span)
                    .ok_or_else(|| invalid("diagram text-body span is outside the input"))?,
                replacement,
                limits,
                crate::LimitKind::DiagramDataModelBytes,
            )?
        },
        None => {
            let text_body = body_content_bound(xml, &TextBody::default(), text, a_namespace)?;
            let qualified_t_len = dgm_prefix.as_deref().map_or(1, |prefix| prefix.len() + 2);
            let element = text_body
                .checked_add(
                    qualified_t_len
                        .checked_mul(2)
                        .ok_or_else(|| invalid("edited diagram output length overflows"))?,
                )
                .and_then(|length| length.checked_add(5))
                .ok_or_else(|| invalid("edited diagram output length overflows"))?;
            if point.open_end == point.span.1 {
                let span = span_len(xml, point.span)?;
                let replacement = span
                    .checked_mul(2)
                    .and_then(|length| length.checked_add(1))
                    .and_then(|length| length.checked_add(element))
                    .ok_or_else(|| invalid("edited diagram output length overflows"))?;
                ensure_output(
                    xml.len()
                        .checked_sub(span)
                        .ok_or_else(|| invalid("diagram point span is outside the input"))?,
                    replacement,
                    limits,
                    crate::LimitKind::DiagramDataModelBytes,
                )?
            } else {
                ensure_output(
                    xml.len(),
                    element,
                    limits,
                    crate::LimitKind::DiagramDataModelBytes,
                )?
            }
        },
    };
    let mut out = Vec::with_capacity(output_capacity);
    match &point.text {
        Some(body) if body.open_end != body.span.1 => {
            out.extend_from_slice(&xml[..body.open_end]);
            let content = body_content(xml, body, text, a_namespace)?;
            out.extend_from_slice(content.as_bytes());
            out.extend_from_slice(&xml[body.close_start..]);
        },
        Some(body) => {
            // `<dgm:t/>`: expand it.
            let raw = String::from_utf8_lossy(&xml[body.span.0..body.span.1]).into_owned();
            let open = raw
                .trim_end_matches('>')
                .trim_end_matches('/')
                .trim_end()
                .to_owned();
            let name = open
                .trim_start_matches('<')
                .split(|c: char| c.is_whitespace())
                .next()
                .unwrap_or("t")
                .to_owned();
            out.extend_from_slice(&xml[..body.span.0]);
            out.extend_from_slice(open.as_bytes());
            out.push(b'>');
            let content = body_content(xml, &TextBody::default(), text, a_namespace)?;
            out.extend_from_slice(content.as_bytes());
            out.extend_from_slice(format!("</{name}>").as_bytes());
            out.extend_from_slice(&xml[body.span.1..]);
        },
        None => {
            let t = match &dgm_prefix {
                Some(prefix) => format!("{prefix}:t"),
                None => "t".to_owned(),
            };
            let content = body_content(xml, &TextBody::default(), text, a_namespace)?;
            let element = format!("<{t}>{}</{t}>", content);
            if point.open_end == point.span.1 {
                // An empty `pt`: expand it.
                let raw = String::from_utf8_lossy(&xml[point.span.0..point.span.1]).into_owned();
                let open = raw
                    .trim_end_matches('>')
                    .trim_end_matches('/')
                    .trim_end()
                    .to_owned();
                let name = open
                    .trim_start_matches('<')
                    .split(|c: char| c.is_whitespace())
                    .next()
                    .unwrap_or("pt")
                    .to_owned();
                out.extend_from_slice(&xml[..point.span.0]);
                out.extend_from_slice(open.as_bytes());
                out.push(b'>');
                out.extend_from_slice(element.as_bytes());
                out.extend_from_slice(format!("</{name}>").as_bytes());
                out.extend_from_slice(&xml[point.span.1..]);
            } else {
                // `t` follows `prSet` and `spPr`; insert it before the end tag
                // or a trailing `extLst`.
                let inner = &xml[point.open_end..point.span.1];
                let close = point.span.0
                    + rfind(&xml[point.span.0..point.span.1], b"</").unwrap_or(inner.len());
                let insert = find(inner, b"extLst")
                    .and_then(|at| rfind(&inner[..at], b"<").map(|lt| point.open_end + lt))
                    .unwrap_or(close);
                out.extend_from_slice(&xml[..insert]);
                out.extend_from_slice(element.as_bytes());
                out.extend_from_slice(&xml[insert..]);
            }
        },
    }
    // Verify the edit reads back.
    ensure_output(
        out.len(),
        0,
        limits,
        crate::LimitKind::DiagramDataModelBytes,
    )?;
    let check = Outline::read(&out, limits)?;
    let expected = text.replace("\r\n", "\n");
    if check.entries.get(index).map(|entry| entry.text.as_str()) != Some(expected.as_str()) {
        return Err(invalid("the edited diagram text did not read back"));
    }
    Ok(out)
}

fn rfind(haystack: &[u8], needle: &[u8]) -> Option<usize> {
    haystack
        .windows(needle.len())
        .rposition(|window| window == needle)
}

/// A drawing shape (`dsp:sp`) with its text body.
struct DrawnShape {
    model: String,
    body: TextBody,
}

fn read_drawing(xml: &[u8], limits: &Limits) -> Result<(Vec<DrawnShape>, bool)> {
    let max_input = limits.input_bytes_value().min(MAX_PART_BYTES);
    if xml.len() > max_input {
        return Err(Error::limit(
            crate::LimitKind::DiagramDrawingBytes,
            xml.len(),
            max_input,
        ));
    }
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut shapes = Vec::new();
    let mut models: Vec<Option<String>> = Vec::new();
    let mut xml_nodes = 0_usize;
    let mut depth = 0_usize;
    let mut strict = false;
    let mut root_seen = false;
    let mut root_closed = false;
    let mut sp_tree_seen = false;
    let mut shape_count = 0_usize;
    loop {
        buffer.clear();
        let event_start = position(&reader)?;
        let (resolved, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(xml_error)?;
        let ns = classify(&resolved);
        let event_end = position(&reader)?;
        match event {
            Event::Start(ref element) => {
                if root_closed {
                    return Err(invalid("trailing markup follows the diagram drawing"));
                }
                account_xml_node(&mut xml_nodes, limits)?;
                let event_depth = depth
                    .checked_add(1)
                    .ok_or_else(|| invalid("diagram drawing XML depth overflows"))?;
                let max_depth = limits.xml_depth_value().min(MAX_DEPTH as usize);
                if event_depth > max_depth {
                    return Err(Error::limit(
                        crate::LimitKind::XmlDepth,
                        event_depth,
                        max_depth,
                    ));
                }
                let drawing = ns == Ns::Drawing;
                let local = element.local_name();
                if depth == 0 {
                    if root_seen {
                        return Err(invalid("the diagram drawing has multiple roots"));
                    }
                    if !drawing || local.as_ref() != b"drawing" {
                        return Err(invalid("the part is not a diagram drawing"));
                    }
                    root_seen = true;
                } else if drawing && local.as_ref() == b"drawing" {
                    return Err(invalid("the diagram drawing has a nested root"));
                }
                if drawing && local.as_ref() == b"spTree" && depth == 1 {
                    if sp_tree_seen {
                        return Err(invalid("the diagram drawing has duplicate shape trees"));
                    }
                    sp_tree_seen = true;
                }
                if drawing && local.as_ref() == b"txBody" {
                    let body = read_body(
                        &mut reader,
                        &mut buffer,
                        event_start,
                        event_end,
                        event_depth,
                        &mut xml_nodes,
                        limits,
                    )?;
                    if let Some(Some(model)) = models.last() {
                        shapes.push(DrawnShape {
                            model: model.clone(),
                            body,
                        });
                    }
                    continue;
                }
                if drawing && local.as_ref() == b"sp" {
                    let max_shapes = limits.nodes_value().min(MAX_NODES);
                    shape_count = shape_count
                        .checked_add(1)
                        .ok_or_else(|| invalid("diagram shape count overflows"))?;
                    if shape_count > max_shapes {
                        return Err(Error::limit(
                            crate::LimitKind::DiagramNodes,
                            shape_count,
                            max_shapes,
                        ));
                    }
                    let values = attributes(element)?;
                    models.push(attribute(&values, "modelId").map(str::to_owned));
                } else {
                    models.push(models.last().cloned().flatten());
                }
                depth = event_depth;
            },
            Event::Empty(ref element) => {
                if root_closed {
                    return Err(invalid("trailing markup follows the diagram drawing"));
                }
                account_xml_node(&mut xml_nodes, limits)?;
                let event_depth = depth
                    .checked_add(1)
                    .ok_or_else(|| invalid("diagram drawing XML depth overflows"))?;
                let max_depth = limits.xml_depth_value().min(MAX_DEPTH as usize);
                if event_depth > max_depth {
                    return Err(Error::limit(
                        crate::LimitKind::XmlDepth,
                        event_depth,
                        max_depth,
                    ));
                }
                let drawing = ns == Ns::Drawing;
                let local = element.local_name();
                if depth == 0 {
                    if root_seen {
                        return Err(invalid("the diagram drawing has multiple roots"));
                    }
                    if !drawing || local.as_ref() != b"drawing" {
                        return Err(invalid("the part is not a diagram drawing"));
                    }
                    return Err(invalid("the diagram drawing root is empty"));
                }
                if drawing && local.as_ref() == b"drawing" {
                    return Err(invalid("the diagram drawing has a nested root"));
                }
                if drawing && local.as_ref() == b"spTree" && depth == 1 {
                    if sp_tree_seen {
                        return Err(invalid("the diagram drawing has duplicate shape trees"));
                    }
                    sp_tree_seen = true;
                }
                if drawing && local.as_ref() == b"sp" {
                    let max_shapes = limits.nodes_value().min(MAX_NODES);
                    shape_count = shape_count
                        .checked_add(1)
                        .ok_or_else(|| invalid("diagram shape count overflows"))?;
                    if shape_count > max_shapes {
                        return Err(Error::limit(
                            crate::LimitKind::DiagramNodes,
                            shape_count,
                            max_shapes,
                        ));
                    }
                }
                if is_dml(ns) && !strict {
                    strict = ns == Ns::DmlStrict;
                }
            },
            Event::End(ref element) => {
                if depth == 0 || models.is_empty() {
                    return Err(invalid("unexpected diagram drawing end"));
                }
                if depth == 1 && (ns != Ns::Drawing || element.local_name().as_ref() != b"drawing")
                {
                    return Err(invalid("the diagram drawing root is not closed"));
                }
                models.pop();
                depth -= 1;
                if depth == 0 {
                    root_closed = true;
                }
            },
            Event::Eof => break,
            Event::DocType(_) => return Err(invalid("document type in a diagram part")),
            Event::Text(text) => {
                if text.as_ref().iter().any(|byte| !byte.is_ascii_whitespace()) {
                    return Err(invalid(
                        "non-whitespace text is outside the diagram drawing",
                    ));
                }
            },
            Event::CData(data) => {
                if data.as_ref().iter().any(|byte| !byte.is_ascii_whitespace()) {
                    return Err(invalid("CDATA is outside the diagram drawing"));
                }
            },
            Event::GeneralRef(_) => {
                return Err(invalid("entity reference is outside the diagram drawing"));
            },
            _ => {},
        }
    }
    if !root_seen || !root_closed || !sp_tree_seen || depth != 0 || !models.is_empty() {
        return Err(invalid("unterminated diagram drawing"));
    }
    Ok((shapes, strict))
}

/// Make the change [`set_text`] makes to the data model in the cached
/// drawing `drawing` of the data model `data` (the model *before* the
/// change).
///
/// The node's shapes are the drawing shapes whose `modelId` is a
/// presentation point the node is presented by (`presOf`). The drawing is
/// changed only when the node has at least one such shape, no such
/// presentation point also presents another node, and every such shape
/// shows exactly the node's current text; otherwise this returns `None`,
/// and the caller removes the drawing so consumers lay the diagram out
/// again.
///
/// # Errors
///
/// Returns an error for malformed XML or text XML cannot carry.
pub fn drawing_text(
    drawing: &[u8],
    data: &[u8],
    id: &str,
    text: &str,
    limits: &Limits,
) -> Result<Option<Vec<u8>>> {
    limits.validate()?;
    check_text(text, limits)?;
    let parsed = Parsed::read(data, limits)?;
    let Some(index) = parsed.outline.position(id) else {
        return Ok(None);
    };
    let current = parsed.outline.entries[index].text.clone();
    let presented: HashSet<&str> = parsed
        .connections
        .iter()
        .filter(|connection| connection.kind == "presOf" && connection.source == id)
        .map(|connection| connection.destination.as_str())
        .collect();
    if presented.is_empty() {
        return Ok(None);
    }
    let shared = parsed.connections.iter().any(|connection| {
        connection.kind == "presOf"
            && connection.source != id
            && presented.contains(connection.destination.as_str())
    });
    if shared {
        return Ok(None);
    }
    let (shapes, strict) = read_drawing(drawing, limits)?;
    let mine: Vec<&DrawnShape> = shapes
        .iter()
        .filter(|shape| presented.contains(shape.model.as_str()))
        .collect();
    if mine.is_empty() || mine.iter().any(|shape| shape.body.text() != current) {
        return Ok(None);
    }
    let a_namespace = if strict || parsed.strict {
        "http://purl.oclc.org/ooxml/drawingml/main"
    } else {
        "http://schemas.openxmlformats.org/drawingml/2006/main"
    };
    let mut output_capacity = drawing.len();
    for shape in &mine {
        let body = &shape.body;
        if body.open_end == body.span.1 {
            return Ok(None);
        }
        let retained = body
            .close_start
            .checked_sub(body.open_end)
            .ok_or_else(|| invalid("diagram drawing text span is inverted"))?;
        let replacement = body_content_bound(drawing, body, text, a_namespace)?;
        output_capacity = output_capacity
            .checked_sub(retained)
            .and_then(|length| length.checked_add(replacement))
            .ok_or_else(|| invalid("edited drawing output length overflows"))?;
    }
    ensure_output_total(
        output_capacity,
        limits,
        crate::LimitKind::DiagramDrawingBytes,
    )?;
    let mut out = Vec::with_capacity(output_capacity);
    let mut cursor = 0;
    for shape in mine {
        let body = &shape.body;
        out.extend_from_slice(&drawing[cursor..body.open_end]);
        let content = body_content(drawing, body, text, a_namespace)?;
        out.extend_from_slice(content.as_bytes());
        cursor = body.close_start;
    }
    out.extend_from_slice(&drawing[cursor..]);
    ensure_output(out.len(), 0, limits, crate::LimitKind::DiagramDrawingBytes)?;
    Ok(Some(out))
}

fn drawing_extension_len(
    a_namespace: &str,
    dgm_namespace: &str,
    relationship: &str,
) -> Result<usize> {
    let relationship = escaped_xml_len(relationship)?;
    let prefix = format!(
        "<a:ext xmlns:a=\"{a_namespace}\" uri=\"{DRAWING_NAMESPACE}\"><dsp:dataModelExt xmlns:dsp=\"{DRAWING_NAMESPACE}\" relId=\""
    );
    let suffix = format!("\" minVer=\"{dgm_namespace}\"/></a:ext>");
    prefix
        .len()
        .checked_add(relationship)
        .and_then(|length| length.checked_add(suffix.len()))
        .ok_or_else(|| invalid("diagram drawing extension length overflows"))
}

/// Name `relationship` as the data model's cached-drawing relationship, or
/// with `None` remove the reference (`dsp:dataModelExt` and its extension;
/// an extension list left empty is removed too). Returns the input bytes
/// when nothing changes.
///
/// # Errors
///
/// Returns an error for malformed XML or a relationship ID XML cannot carry.
pub fn set_drawing_reference(
    xml: &[u8],
    relationship: Option<&str>,
    limits: &Limits,
) -> Result<Vec<u8>> {
    limits.validate()?;
    if let Some(relationship) = relationship {
        check_text(relationship, limits)?;
    }
    let parsed = Parsed::read(xml, limits)?;
    if parsed.outline.drawing_relationship.as_deref() == relationship {
        ensure_output(
            xml.len(),
            0,
            limits,
            crate::LimitKind::DiagramDataModelBytes,
        )?;
        return Ok(xml.to_vec());
    }
    let escaped_relationship = relationship.map(escaped_xml_len).transpose()?;
    let output_capacity = match (relationship, parsed.drawing_ext) {
        (Some(_), Some((_, (start, end)))) => xml
            .len()
            .checked_sub(
                end.checked_sub(start)
                    .ok_or_else(|| invalid("diagram drawing relationship span is inverted"))?,
            )
            .and_then(|length| length.checked_add(escaped_relationship.unwrap_or(0)))
            .ok_or_else(|| invalid("edited diagram output length overflows"))?,
        (Some(relationship), None) => {
            let a_namespace = parsed.a_namespace();
            let dgm_namespace = if parsed.strict {
                DGM_NAMESPACE_STRICT
            } else {
                DGM_NAMESPACE
            };
            let ext = drawing_extension_len(a_namespace, dgm_namespace, relationship)?;
            let list_len = parsed
                .root_prefix
                .as_deref()
                .map_or(6, |prefix| prefix.len() + 1 + 6);
            let element = ext
                .checked_add(5 + 2 * list_len)
                .ok_or_else(|| invalid("edited diagram output length overflows"))?;
            match parsed.ext_list {
                Some((_, end, open_end)) if open_end < end => xml
                    .len()
                    .checked_add(ext)
                    .ok_or_else(|| invalid("edited diagram output length overflows"))?,
                Some((start, end, _)) => xml
                    .len()
                    .checked_sub(
                        end.checked_sub(start)
                            .ok_or_else(|| invalid("diagram extension-list span is inverted"))?,
                    )
                    .and_then(|length| length.checked_add(element))
                    .ok_or_else(|| invalid("edited diagram output length overflows"))?,
                None => xml
                    .len()
                    .checked_add(element)
                    .ok_or_else(|| invalid("edited diagram output length overflows"))?,
            }
        },
        (None, Some(((start, end), _))) => xml
            .len()
            .checked_sub(
                end.checked_sub(start)
                    .ok_or_else(|| invalid("diagram drawing extension span is inverted"))?,
            )
            .ok_or_else(|| invalid("edited diagram output length underflows"))?,
        (None, None) => xml.len(),
    };
    ensure_output_total(
        output_capacity,
        limits,
        crate::LimitKind::DiagramDataModelBytes,
    )?;
    let mut out = Vec::with_capacity(output_capacity);
    match (relationship, parsed.drawing_ext) {
        (Some(relationship), Some((_, (start, end)))) => {
            out.extend_from_slice(&xml[..start]);
            out.extend_from_slice(escape_xml(relationship).as_bytes());
            out.extend_from_slice(&xml[end..]);
        },
        (Some(relationship), None) => {
            let a_namespace = parsed.a_namespace();
            let dgm_namespace = if parsed.strict {
                DGM_NAMESPACE_STRICT
            } else {
                DGM_NAMESPACE
            };
            let ext = format!(
                "<a:ext xmlns:a=\"{a_namespace}\" uri=\"{DRAWING_NAMESPACE}\"><dsp:dataModelExt xmlns:dsp=\"{DRAWING_NAMESPACE}\" relId=\"{}\" minVer=\"{dgm_namespace}\"/></a:ext>",
                escape_xml(relationship)
            );
            match parsed.ext_list {
                Some((start, end, open_end)) if open_end < end => {
                    // A non-empty list: append before its end tag.
                    let close = start + rfind(&xml[start..end], b"</").unwrap_or(end - start);
                    out.extend_from_slice(&xml[..close]);
                    out.extend_from_slice(ext.as_bytes());
                    out.extend_from_slice(&xml[close..]);
                },
                existing => {
                    let list = match &parsed.root_prefix {
                        Some(prefix) => format!("{prefix}:extLst"),
                        None => "extLst".to_owned(),
                    };
                    let element = format!("<{list}>{ext}</{list}>");
                    let (start, end) = existing
                        .map_or((parsed.root_close, parsed.root_close), |(start, end, _)| {
                            (start, end)
                        });
                    out.extend_from_slice(&xml[..start]);
                    out.extend_from_slice(element.as_bytes());
                    out.extend_from_slice(&xml[end..]);
                },
            }
        },
        (None, Some(((start, end), _))) => {
            out.extend_from_slice(&xml[..start]);
            out.extend_from_slice(&xml[end..]);
            // Remove an extension list left without elements.
            let reparsed = Parsed::read(&out, limits)?;
            if let Some((list_start, list_end, open_end)) = reparsed.ext_list {
                let inner = &out[open_end..list_end];
                let close = rfind(inner, b"</").unwrap_or(inner.len());
                if inner[..close].iter().all(u8::is_ascii_whitespace) {
                    out.drain(list_start..list_end);
                }
            }
        },
        (None, None) => return Ok(xml.to_vec()),
    }
    ensure_output(
        out.len(),
        0,
        limits,
        crate::LimitKind::DiagramDataModelBytes,
    )?;
    let check = Outline::read(&out, limits)?;
    if check.drawing_relationship.as_deref() != relationship {
        return Err(invalid("the drawing reference did not read back"));
    }
    Ok(out)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn limits_with_nodes(nodes: usize) -> Limits {
        let defaults = Limits::default();
        Limits::new(
            defaults.input_bytes_value(),
            defaults.output_bytes_value(),
            defaults.xml_nodes_value(),
            defaults.xml_depth_value(),
            defaults.points_value(),
            defaults.connections_value(),
            defaults.data_text_bytes_value(),
            nodes,
            defaults.node_depth_value(),
            defaults.node_text_bytes_value(),
        )
    }

    #[test]
    fn drawing_shape_limit_counts_sibling_shapes() {
        let drawing = br#"<dsp:drawing xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram"><dsp:spTree><dsp:sp/><dsp:sp/></dsp:spTree></dsp:drawing>"#;
        assert!(matches!(
            read_drawing(drawing, &limits_with_nodes(1)),
            Err(Error::Limit {
                kind: crate::LimitKind::DiagramNodes,
                ..
            })
        ));
    }

    #[test]
    fn drawing_rejects_nested_empty_root() {
        let drawing = br#"<dsp:drawing xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram"><dsp:drawing/></dsp:drawing>"#;
        assert!(matches!(
            read_drawing(drawing, &Limits::default()),
            Err(Error::Invalid(_))
        ));
    }

    #[test]
    fn drawing_requires_a_shape_tree() {
        let drawing = br#"<dsp:drawing xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram"></dsp:drawing>"#;
        assert!(matches!(
            read_drawing(drawing, &Limits::default()),
            Err(Error::Invalid(_))
        ));
    }
}
