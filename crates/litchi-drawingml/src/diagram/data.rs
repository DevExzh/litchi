//! DrawingML diagram data-model (`dgm:dataModel`) CRUD shared across formats.
//!
//! The data model is the semantic heart of a SmartArt diagram: it declares the
//! logical points (`dgm:pt`) — the document root, content nodes, transition
//! points, and presentation points — and the connection graph (`dgm:cxn`) that
//! links them. The semantic-subset codec reads and writes both the transitional
//! and ISO Strict `drawingml/diagram` namespaces. Schema-defined identifiers
//! and relation kinds are represented by closed Rust types, so invalid wire
//! strings cannot enter an authored model.
//!
//! Parsing extracts graph semantics and literal text but deliberately does not
//! retain rich-text formatting, shape properties, backgrounds, whole-diagram
//! formatting, or extension lists. Consequently, serialization is intended for
//! fresh authoring or explicit canonicalization, not lossless publication of an
//! arbitrary parsed part.

use crate::diagram::{DGM_NAMESPACE, DGM_NAMESPACE_STRICT, DiagramNode};
use crate::{Error, Result};
use litchi_ooxml_common::mce::process_str;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use std::collections::{HashMap, HashSet};
use std::fmt;
use std::str::FromStr;

const MAX_DATA_MODEL_XML: usize = 16 * 1024 * 1024;
const MAX_NODES: usize = 200_000;
const MAX_DEPTH: usize = 128;
const MAX_POINTS: usize = 100_000;
const MAX_CONNECTIONS: usize = 100_000;
const MAX_TEXT_BYTES: usize = 1024 * 1024;
const DML_NAMESPACE: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const DML_NAMESPACE_STRICT: &str = "http://purl.oclc.org/ooxml/drawingml/main";
/// Recursion guard for [`DiagramDataModel::node_tree`] on cyclic graphs.
const MAX_TREE_DEPTH: u32 = 64;

/// A diagram model identifier (`ST_ModelId`).
///
/// ECMA-376 defines this domain as the union of a signed 32-bit integer and an
/// uppercase, braced GUID. Keeping either representation inline avoids a heap
/// allocation for every point and connection edge.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Id {
    /// Signed 32-bit numeric identifier.
    Number(i32),
    /// The 16 bytes of an uppercase, braced GUID.
    Guid([u8; 16]),
}

impl Id {
    /// Creates a numeric model identifier.
    #[inline]
    pub const fn number(value: i32) -> Self {
        Self::Number(value)
    }

    /// Creates a GUID model identifier from its 16 bytes.
    #[inline]
    pub const fn guid(value: [u8; 16]) -> Self {
        Self::Guid(value)
    }

    /// Returns the numeric value, if this is a numeric identifier.
    #[inline]
    pub const fn as_number(self) -> Option<i32> {
        match self {
            Self::Number(value) => Some(value),
            Self::Guid(_) => None,
        }
    }

    /// Returns the GUID bytes, if this is a GUID identifier.
    #[inline]
    pub const fn as_guid(self) -> Option<[u8; 16]> {
        match self {
            Self::Number(_) => None,
            Self::Guid(value) => Some(value),
        }
    }
}

impl From<i32> for Id {
    #[inline]
    fn from(value: i32) -> Self {
        Self::Number(value)
    }
}

impl FromStr for Id {
    type Err = IdError;

    fn from_str(value: &str) -> std::result::Result<Self, Self::Err> {
        let value = value.trim_matches([' ', '\t', '\r', '\n']);
        if let Ok(number) = value.parse::<i32>() {
            return Ok(Self::Number(number));
        }
        parse_guid(value).map(Self::Guid)
    }
}

impl TryFrom<&str> for Id {
    type Error = IdError;

    #[inline]
    fn try_from(value: &str) -> std::result::Result<Self, Self::Error> {
        value.parse()
    }
}

impl fmt::Display for Id {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Number(value) => value.fmt(formatter),
            Self::Guid(value) => write!(
                formatter,
                "{{{:02X}{:02X}{:02X}{:02X}-{:02X}{:02X}-{:02X}{:02X}-{:02X}{:02X}-{:02X}{:02X}{:02X}{:02X}{:02X}{:02X}}}",
                value[0],
                value[1],
                value[2],
                value[3],
                value[4],
                value[5],
                value[6],
                value[7],
                value[8],
                value[9],
                value[10],
                value[11],
                value[12],
                value[13],
                value[14],
                value[15],
            ),
        }
    }
}

/// A lexical value is outside the ECMA-376 `ST_ModelId` domain.
#[derive(Debug, Clone, Copy, PartialEq, Eq, thiserror::Error)]
#[error("expected a signed 32-bit integer or uppercase braced GUID")]
pub struct IdError;

fn parse_guid(value: &str) -> std::result::Result<[u8; 16], IdError> {
    if value.len() != 38 || !value.starts_with('{') || !value.ends_with('}') {
        return Err(IdError);
    }
    let bytes = value.as_bytes();
    if [9, 14, 19, 24]
        .into_iter()
        .any(|position| bytes[position] != b'-')
    {
        return Err(IdError);
    }
    let mut guid = [0_u8; 16];
    let mut source = 1usize;
    for byte in &mut guid {
        if matches!(source, 9 | 14 | 19 | 24) {
            source += 1;
        }
        let high = hex(bytes[source]).ok_or(IdError)?;
        let low = hex(bytes[source + 1]).ok_or(IdError)?;
        *byte = (high << 4) | low;
        source += 2;
    }
    Ok(guid)
}

#[inline]
const fn hex(value: u8) -> Option<u8> {
    match value {
        b'0'..=b'9' => Some(value - b'0'),
        b'A'..=b'F' => Some(value - b'A' + 10),
        _ => None,
    }
}

/// The type and required relation of a diagram data-model point (`dgm:pt`).
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub enum PointType {
    /// Content node (`node`, the default when `type` is absent).
    #[default]
    Node,
    /// Document root point (`doc`).
    Document,
    /// Assistant node (`asst`).
    Assistant,
    /// Parent transition point (`parTrans`) owned by a connection.
    ///
    /// A missing `cxnId` parses as numeric zero, the schema default. The writer
    /// always emits the identifier because Microsoft Office requires the
    /// attribute on transition points.
    ParentTransition(Id),
    /// Sibling transition point (`sibTrans`) owned by a connection.
    ///
    /// A missing `cxnId` parses as numeric zero, the schema default. The writer
    /// always emits the identifier because Microsoft Office requires the
    /// attribute on transition points.
    SiblingTransition(Id),
    /// Presentation point (`pres`).
    Presentation,
}

impl PointType {
    /// Returns the owning connection for a transition point.
    #[inline]
    pub const fn connection(self) -> Option<Id> {
        match self {
            Self::ParentTransition(connection) | Self::SiblingTransition(connection) => {
                Some(connection)
            },
            _ => None,
        }
    }
}

/// The type and Office-required metadata of a diagram connection (`dgm:cxn`).
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum ConnectionType {
    /// Structural parent-of (`parOf`, the default when `type` is absent).
    ///
    /// Missing `parTransId` and `sibTransId` attributes parse as numeric zero,
    /// their schema defaults. The writer emits both explicitly because Office
    /// requires them for `parOf` relations.
    Parent {
        /// Parent transition point identifier.
        parent_transition: Id,
        /// Sibling transition point identifier.
        sibling_transition: Id,
    },
    /// Presentation mapping (`presOf`) and its required presentation ID.
    ///
    /// A missing `presId` parses as the schema-default empty string. The writer
    /// emits it explicitly because Office requires the attribute for `presOf`.
    Presentation(String),
    /// Presentation parent-of (`presParOf`).
    PresentationParent,
    /// The schema-defined `unknownRelationship` value.
    Unknown,
}

impl ConnectionType {
    /// Creates a structural parent relation with both Office-required
    /// transition identifiers.
    #[inline]
    pub const fn parent(parent_transition: Id, sibling_transition: Id) -> Self {
        Self::Parent {
            parent_transition,
            sibling_transition,
        }
    }

    /// Returns whether Office treats this relation as a parent edge for its
    /// one-parent-per-destination constraint.
    #[inline]
    pub const fn is_parent(&self) -> bool {
        matches!(self, Self::Parent { .. } | Self::PresentationParent)
    }
}

/// A single point (`dgm:pt`) in a diagram data model.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Point {
    /// Point identifier (`modelId`).
    pub id: Id,
    /// Point type and any type-required identifier.
    pub kind: PointType,
    /// Concatenated text of the point's `dgm:t` body.
    pub text: String,
    /// Layout type identifier from `dgm:prSet@loTypeId` (document point).
    pub layout_type_id: Option<String>,
    /// Quick style type identifier from `dgm:prSet@qsTypeId` (document point).
    pub quick_style_type_id: Option<String>,
    /// Color style type identifier from `dgm:prSet@csTypeId` (document point).
    pub color_style_type_id: Option<String>,
}

impl Point {
    /// Creates a point with no text or style-definition references.
    #[inline]
    pub fn new(id: Id, kind: PointType) -> Self {
        Self {
            id,
            kind,
            text: String::new(),
            layout_type_id: None,
            quick_style_type_id: None,
            color_style_type_id: None,
        }
    }

    /// Creates a content node with literal text.
    #[inline]
    pub fn node(id: Id, text: impl Into<String>) -> Self {
        Self::new(id, PointType::Node).with_text(text)
    }

    /// Replaces this point's literal text.
    #[inline]
    pub fn with_text(mut self, text: impl Into<String>) -> Self {
        self.text = text.into();
        self
    }
}

/// A single connection (`dgm:cxn`) in a diagram data model.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Connection {
    /// Connection identifier (`modelId`).
    pub id: Id,
    /// Connection type and any type-required metadata.
    pub kind: ConnectionType,
    /// Source point identifier (`srcId`).
    pub source: Id,
    /// Destination point identifier (`destId`).
    pub destination: Id,
    /// Source order (`srcOrd`).
    pub src_ord: u32,
    /// Destination order (`destOrd`).
    pub dest_ord: u32,
}

impl Connection {
    /// Creates a connection edge.
    #[inline]
    pub const fn new(
        id: Id,
        kind: ConnectionType,
        source: Id,
        destination: Id,
        src_ord: u32,
        dest_ord: u32,
    ) -> Self {
        Self {
            id,
            kind,
            source,
            destination,
            src_ord,
            dest_ord,
        }
    }
}

/// XML namespace conformance used when serializing a data model.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub enum Conformance {
    /// ECMA-376 transitional namespaces.
    #[default]
    Transitional,
    /// ISO/IEC 29500 Strict namespaces.
    Strict,
}

/// The modeled semantic subset of a DrawingML diagram data model.
///
/// This type does not retain unmodeled XML such as rich-text formatting, shape
/// properties, backgrounds, whole-diagram formatting, or extension lists.
/// Use [`Self::to_xml`] only for fresh authoring or when that canonicalization
/// is explicitly intended.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct DiagramDataModel {
    /// Data-model points in document order.
    pub points: Vec<Point>,
    /// Data-model connections in document order.
    pub connections: Vec<Connection>,
}

impl DiagramDataModel {
    /// Parse a `dgm:dataModel` document (transitional or Strict namespace).
    ///
    /// The input is first rewritten by markup-compatibility processing so
    /// `mc:AlternateContent` wrappers resolve to their fallback content.
    /// Unmodeled formatting and extension content is validated as XML structure
    /// but is not retained; see the publication warning on [`Self::to_xml`].
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
        let mut root_closed = false;
        let mut point_list_seen = false;
        let mut point_list_depth: Option<usize> = None;
        let mut connection_list_seen = false;
        let mut connection_list_depth: Option<usize> = None;
        // Currently open point: depth at which it started plus its builder.
        let mut open_point: Option<(usize, PointBuilder)> = None;
        // Depths of the open `dgm:t` text body and `a:t` text leaf, if any.
        let mut text_body_depth: Option<usize> = None;
        let mut run_text_depth: Option<usize> = None;

        loop {
            match reader.read_event_into(&mut buffer).map_err(xml_error)? {
                Event::Start(element) => {
                    depth = depth
                        .checked_add(1)
                        .ok_or_else(|| limit("data-model XML depth"))?;
                    nodes = nodes
                        .checked_add(1)
                        .ok_or_else(|| limit("data-model XML node count"))?;
                    if nodes > MAX_NODES || depth > MAX_DEPTH {
                        return Err(limit("data-model XML structure"));
                    }
                    let namespace = element_namespace(&reader, &element)?;
                    let local = local_name(&element)?;
                    if depth == 1 {
                        if root_seen || root_closed {
                            return Err(invalid("multiple data-model roots"));
                        }
                        if !is_dgm(&namespace) || local != "dataModel" {
                            return Err(invalid("invalid data-model root or namespace"));
                        }
                        root_seen = true;
                    } else if is_dgm(&namespace) {
                        match local.as_str() {
                            "ptLst" => {
                                if depth != 2 || point_list_seen {
                                    return Err(invalid("invalid or duplicate diagram point list"));
                                }
                                point_list_seen = true;
                                point_list_depth = Some(depth);
                            },
                            "cxnLst" => {
                                if depth != 2
                                    || connection_list_seen
                                    || !point_list_seen
                                    || point_list_depth.is_some()
                                {
                                    return Err(invalid(
                                        "invalid, duplicate, or out-of-order diagram connection list",
                                    ));
                                }
                                connection_list_seen = true;
                                connection_list_depth = Some(depth);
                            },
                            "pt" => {
                                if point_list_depth.is_none_or(|list| depth != list + 1) {
                                    return Err(invalid("diagram point is outside its point list"));
                                }
                                if open_point.is_some() {
                                    return Err(invalid("nested diagram point"));
                                }
                                if model.points.len() >= MAX_POINTS {
                                    return Err(limit("diagram point count"));
                                }
                                open_point = Some((depth, PointBuilder::from_element(&element)?));
                            },
                            "cxn" => {
                                if connection_list_depth.is_none_or(|list| depth != list + 1) {
                                    return Err(invalid(
                                        "diagram connection is outside its connection list",
                                    ));
                                }
                                push_connection(&mut model, &element)?;
                            },
                            "prSet" => {
                                let Some((point_depth, builder)) = &mut open_point else {
                                    return Err(invalid("diagram property set is outside a point"));
                                };
                                if depth != *point_depth + 1 {
                                    return Err(invalid(
                                        "diagram property set is not a direct point child",
                                    ));
                                }
                                builder.read_pr_set(&element)?;
                            },
                            "t" => {
                                let Some((point_depth, builder)) = &mut open_point else {
                                    return Err(invalid("diagram text body is outside a point"));
                                };
                                if depth != *point_depth + 1
                                    || text_body_depth.is_some()
                                    || builder.text_body_seen
                                {
                                    return Err(invalid(
                                        "invalid or duplicate diagram point text body",
                                    ));
                                }
                                builder.text_body_seen = true;
                                text_body_depth = Some(depth);
                            },
                            _ => {},
                        }
                    }
                    if is_dml(&namespace)
                        && local == "t"
                        && text_body_depth.is_some_and(|body| depth > body)
                    {
                        if run_text_depth.is_some() {
                            return Err(invalid("nested DrawingML text leaf"));
                        }
                        run_text_depth = Some(depth);
                    }
                },
                Event::Empty(element) => {
                    nodes = nodes
                        .checked_add(1)
                        .ok_or_else(|| limit("data-model XML node count"))?;
                    let child_depth = depth
                        .checked_add(1)
                        .ok_or_else(|| limit("data-model XML depth"))?;
                    if nodes > MAX_NODES || child_depth > MAX_DEPTH {
                        return Err(limit("data-model XML structure"));
                    }
                    let namespace = element_namespace(&reader, &element)?;
                    let local = local_name(&element)?;
                    if child_depth == 1 {
                        if root_seen || root_closed {
                            return Err(invalid("multiple data-model roots"));
                        }
                        if !is_dgm(&namespace) || local != "dataModel" {
                            return Err(invalid("invalid data-model root or namespace"));
                        }
                        return Err(invalid("diagram data model lacks its required point list"));
                    } else if is_dgm(&namespace) {
                        match local.as_str() {
                            "ptLst" => {
                                if child_depth != 2 || point_list_seen {
                                    return Err(invalid("invalid or duplicate diagram point list"));
                                }
                                point_list_seen = true;
                            },
                            "cxnLst" => {
                                if child_depth != 2
                                    || connection_list_seen
                                    || !point_list_seen
                                    || point_list_depth.is_some()
                                {
                                    return Err(invalid(
                                        "invalid, duplicate, or out-of-order diagram connection list",
                                    ));
                                }
                                connection_list_seen = true;
                            },
                            "pt" => {
                                if point_list_depth.is_none_or(|list| child_depth != list + 1) {
                                    return Err(invalid("diagram point is outside its point list"));
                                }
                                if open_point.is_some() {
                                    return Err(invalid("nested diagram point"));
                                }
                                if model.points.len() >= MAX_POINTS {
                                    return Err(limit("diagram point count"));
                                }
                                let builder = PointBuilder::from_element(&element)?;
                                model.points.push(builder.finish()?);
                            },
                            "cxn" => {
                                if connection_list_depth.is_none_or(|list| child_depth != list + 1)
                                {
                                    return Err(invalid(
                                        "diagram connection is outside its connection list",
                                    ));
                                }
                                push_connection(&mut model, &element)?;
                            },
                            "prSet" => {
                                let Some((point_depth, builder)) = &mut open_point else {
                                    return Err(invalid("diagram property set is outside a point"));
                                };
                                if child_depth != *point_depth + 1 {
                                    return Err(invalid(
                                        "diagram property set is not a direct point child",
                                    ));
                                }
                                builder.read_pr_set(&element)?;
                            },
                            "t" => {
                                let Some((point_depth, builder)) = &mut open_point else {
                                    return Err(invalid("diagram text body is outside a point"));
                                };
                                if child_depth != *point_depth + 1 || builder.text_body_seen {
                                    return Err(invalid(
                                        "invalid or duplicate diagram point text body",
                                    ));
                                }
                                builder.text_body_seen = true;
                            },
                            _ => {},
                        }
                    }
                },
                Event::Text(event) => {
                    if run_text_depth.is_some()
                        && let Some((_, builder)) = &mut open_point
                    {
                        let text = std::str::from_utf8(event.as_ref()).map_err(xml_error)?;
                        let text = quick_xml::escape::unescape(text).map_err(xml_error)?;
                        if builder
                            .text
                            .len()
                            .checked_add(text.len())
                            .is_none_or(|length| length > MAX_TEXT_BYTES)
                        {
                            return Err(limit("diagram text bytes"));
                        }
                        builder.text.push_str(&text);
                    } else if depth == 0
                        && event
                            .as_ref()
                            .iter()
                            .any(|byte| !matches!(byte, b' ' | b'\t' | b'\r' | b'\n'))
                    {
                        return Err(invalid("text is outside the data-model root"));
                    }
                },
                Event::GeneralRef(reference) => {
                    if run_text_depth.is_some()
                        && let Some((_, builder)) = &mut open_point
                    {
                        let text = litchi_ooxml_common::xml::decode_xml_reference(&reference)?;
                        if builder
                            .text
                            .len()
                            .checked_add(text.len())
                            .is_none_or(|length| length > MAX_TEXT_BYTES)
                        {
                            return Err(limit("diagram text bytes"));
                        }
                        builder.text.push_str(&text);
                    } else if depth == 0 {
                        return Err(invalid("entity reference is outside the data-model root"));
                    }
                },
                Event::End(element) => {
                    let local = local_name_end(&element)?;
                    if run_text_depth == Some(depth) {
                        if local != "t" {
                            return Err(invalid("unbalanced DrawingML text leaf"));
                        }
                        run_text_depth = None;
                    }
                    if text_body_depth == Some(depth) {
                        if local != "t" {
                            return Err(invalid("unbalanced diagram text body"));
                        }
                        if run_text_depth.is_some() {
                            return Err(invalid("unterminated DrawingML text leaf"));
                        }
                        text_body_depth = None;
                    }
                    if open_point
                        .as_ref()
                        .is_some_and(|(start, _)| *start == depth)
                    {
                        if local != "pt" {
                            return Err(invalid("unbalanced diagram point"));
                        }
                        let Some((_, builder)) = open_point.take() else {
                            return Err(invalid("diagram point state is inconsistent"));
                        };
                        model.points.push(builder.finish()?);
                    }
                    if point_list_depth == Some(depth) {
                        if local != "ptLst" {
                            return Err(invalid("unbalanced diagram point list"));
                        }
                        point_list_depth = None;
                    }
                    if connection_list_depth == Some(depth) {
                        if local != "cxnLst" {
                            return Err(invalid("unbalanced diagram connection list"));
                        }
                        connection_list_depth = None;
                    }
                    if depth == 0 {
                        return Err(invalid("unexpected data-model closing element"));
                    }
                    if depth == 1 {
                        if local != "dataModel" || !root_seen {
                            return Err(invalid("unbalanced data-model root"));
                        }
                        if !point_list_seen {
                            return Err(invalid(
                                "diagram data model lacks its required point list",
                            ));
                        }
                        root_closed = true;
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
        if !root_seen
            || !root_closed
            || !point_list_seen
            || depth != 0
            || point_list_depth.is_some()
            || connection_list_depth.is_some()
            || open_point.is_some()
            || text_body_depth.is_some()
            || run_text_depth.is_some()
        {
            return Err(invalid("missing or unterminated data-model root"));
        }
        Ok(model)
    }

    /// Creates an empty data model.
    #[inline]
    pub const fn new() -> Self {
        Self {
            points: Vec::new(),
            connections: Vec::new(),
        }
    }

    /// Looks up a point by its typed model identifier.
    #[inline]
    pub fn point(&self, id: Id) -> Option<&Point> {
        self.points.iter().find(|point| point.id == id)
    }

    /// Looks up a point mutably by its typed model identifier.
    #[inline]
    pub fn point_mut(&mut self, id: Id) -> Option<&mut Point> {
        self.points.iter_mut().find(|point| point.id == id)
    }

    /// Looks up a connection by its typed model identifier.
    #[inline]
    pub fn connection(&self, id: Id) -> Option<&Connection> {
        self.connections
            .iter()
            .find(|connection| connection.id == id)
    }

    /// Looks up a connection mutably by its typed model identifier.
    #[inline]
    pub fn connection_mut(&mut self, id: Id) -> Option<&mut Connection> {
        self.connections
            .iter_mut()
            .find(|connection| connection.id == id)
    }

    /// Adds a point, rejecting duplicate identifiers and a second document
    /// root before changing the model.
    pub fn add_point(&mut self, point: Point) -> Result<()> {
        if self.points.len() >= MAX_POINTS {
            return Err(limit("diagram point count"));
        }
        if self.point(point.id).is_some() || self.connection(point.id).is_some() {
            return Err(invalid(format!(
                "duplicate diagram model identifier {}",
                point.id
            )));
        }
        if point.kind == PointType::Document && self.document_point().is_some() {
            return Err(invalid("diagram data model already has a document point"));
        }
        if point.text.len() > MAX_TEXT_BYTES {
            return Err(limit("diagram text bytes"));
        }
        if let Some(connection_id) = point.kind.connection()
            && let Some(connection) = self.connection(connection_id)
        {
            validate_transition_point(&point, connection)?;
        }
        self.points.push(point);
        Ok(())
    }

    /// Adds a connection after checking its identifier and endpoint references.
    pub fn add_connection(&mut self, connection: Connection) -> Result<()> {
        if self.connections.len() >= MAX_CONNECTIONS {
            return Err(limit("diagram connection count"));
        }
        if self.connection(connection.id).is_some() || self.point(connection.id).is_some() {
            return Err(invalid(format!(
                "duplicate diagram model identifier {}",
                connection.id
            )));
        }
        if self.point(connection.source).is_none() {
            return Err(invalid(format!(
                "diagram connection {} has no source point {}",
                connection.id, connection.source
            )));
        }
        if self.point(connection.destination).is_none() {
            return Err(invalid(format!(
                "diagram connection {} has no destination point {}",
                connection.id, connection.destination
            )));
        }
        if connection.kind.is_parent()
            && self.connections.iter().any(|existing| {
                existing.kind.is_parent() && existing.destination == connection.destination
            })
        {
            return Err(invalid(format!(
                "diagram point {} already has a parent",
                connection.destination
            )));
        }
        validate_connection_transitions(&connection, |id| self.point(id))?;
        if let ConnectionType::Presentation(presentation) = &connection.kind
            && self.connections.iter().any(|existing| {
                matches!(
                    &existing.kind,
                    ConnectionType::Presentation(existing) if existing != presentation
                )
            })
        {
            return Err(invalid(
                "diagram presentation connections use different presentation identifiers",
            ));
        }
        self.connections.push(connection);
        Ok(())
    }

    /// Removes a point and all connections or transition points that depend on
    /// it. The returned value is the removed point.
    pub fn remove_point(&mut self, id: Id) -> Option<Point> {
        let position = self.points.iter().position(|point| point.id == id)?;
        let point = self.points.remove(position);
        let mut removed_connections = HashSet::new();
        self.connections.retain(|connection| {
            let depends_on_point = connection.source == id
                || connection.destination == id
                || match connection.kind {
                    ConnectionType::Parent {
                        parent_transition,
                        sibling_transition,
                    } => parent_transition == id || sibling_transition == id,
                    _ => false,
                };
            if depends_on_point {
                removed_connections.insert(connection.id);
            }
            !depends_on_point
        });
        self.points.retain(|candidate| {
            candidate
                .kind
                .connection()
                .is_none_or(|connection| !removed_connections.contains(&connection))
        });
        Some(point)
    }

    /// Removes a connection and transition points owned by it.
    pub fn remove_connection(&mut self, id: Id) -> Option<Connection> {
        let position = self
            .connections
            .iter()
            .position(|connection| connection.id == id)?;
        let connection = self.connections.remove(position);
        self.points
            .retain(|point| point.kind.connection() != Some(id));
        Some(connection)
    }

    /// Validates identifiers, references, Office's single-parent rule, and
    /// configured resource limits without changing the model.
    pub fn validate(&self) -> Result<()> {
        self.serialized_xml_len(Conformance::Transitional)
            .map(|_| ())
    }

    fn validated_xml_capacity(&self) -> Result<usize> {
        if self.points.len() > MAX_POINTS {
            return Err(limit("diagram point count"));
        }
        if self.connections.len() > MAX_CONNECTIONS {
            return Err(limit("diagram connection count"));
        }
        let mut xml_capacity = 256usize;
        let mut model_ids = HashSet::with_capacity(self.points.len() + self.connections.len());
        let mut points = HashMap::with_capacity(self.points.len());
        let mut document_seen = false;
        for point in &self.points {
            if !model_ids.insert(point.id) {
                return Err(invalid(format!(
                    "duplicate diagram model identifier {}",
                    point.id
                )));
            }
            points.insert(point.id, point);
            if point.kind == PointType::Document {
                if document_seen {
                    return Err(invalid("diagram data model has multiple document points"));
                }
                document_seen = true;
            }
            if point.text.len() > MAX_TEXT_BYTES {
                return Err(limit("diagram text bytes"));
            }
            let has_modeled_children = !point.text.is_empty()
                || point.layout_type_id.is_some()
                || point.quick_style_type_id.is_some()
                || point.color_style_type_id.is_some();
            xml_capacity = xml_capacity
                .checked_add(if has_modeled_children { 320 } else { 128 })
                .ok_or_else(|| limit("serialized data-model bytes"))?;
            xml_capacity = add_xml_value(xml_capacity, &point.text, "diagram point text")?;
            for (value, description) in [
                (
                    point.layout_type_id.as_deref(),
                    "diagram layout type identifier",
                ),
                (
                    point.quick_style_type_id.as_deref(),
                    "diagram quick-style type identifier",
                ),
                (
                    point.color_style_type_id.as_deref(),
                    "diagram color-style type identifier",
                ),
            ] {
                if let Some(value) = value {
                    xml_capacity = add_xml_value(xml_capacity, value, description)?;
                }
            }
        }

        let mut connections = HashMap::with_capacity(self.connections.len());
        let mut parent_destinations = HashSet::new();
        let mut presentation_id: Option<&str> = None;
        for connection in &self.connections {
            xml_capacity = xml_capacity
                .checked_add(512)
                .ok_or_else(|| limit("serialized data-model bytes"))?;
            if !model_ids.insert(connection.id) {
                return Err(invalid(format!(
                    "duplicate diagram model identifier {}",
                    connection.id
                )));
            }
            connections.insert(connection.id, connection);
            if !points.contains_key(&connection.source) {
                return Err(invalid(format!(
                    "diagram connection {} has no source point {}",
                    connection.id, connection.source
                )));
            }
            if !points.contains_key(&connection.destination) {
                return Err(invalid(format!(
                    "diagram connection {} has no destination point {}",
                    connection.id, connection.destination
                )));
            }
            if connection.kind.is_parent() && !parent_destinations.insert(connection.destination) {
                return Err(invalid(format!(
                    "diagram point {} has multiple parents",
                    connection.destination
                )));
            }
            if let ConnectionType::Presentation(presentation) = &connection.kind {
                if presentation_id.is_some_and(|expected| expected != presentation) {
                    return Err(invalid(
                        "diagram presentation connections use different presentation identifiers",
                    ));
                }
                presentation_id.get_or_insert(presentation);
                xml_capacity = add_xml_value(
                    xml_capacity,
                    presentation,
                    "diagram presentation identifier",
                )?;
            }
            validate_connection_transitions(connection, |id| points.get(&id).copied())?;
        }
        for point in &self.points {
            if let Some(connection_id) = point.kind.connection() {
                let connection = connections.get(&connection_id).copied().ok_or_else(|| {
                    invalid(format!(
                        "diagram transition point {} refers to missing connection {}",
                        point.id, connection_id
                    ))
                })?;
                validate_transition_point(point, connection)?;
            }
        }
        Ok(xml_capacity.min(MAX_DATA_MODEL_XML))
    }

    fn serialized_xml_len(&self, conformance: Conformance) -> Result<usize> {
        self.validated_xml_capacity()?;
        let mut count = XmlByteCount::default();
        self.write_validated_xml(&mut count, conformance)?;
        if count.0 > MAX_DATA_MODEL_XML {
            return Err(limit("serialized data-model bytes"));
        }
        Ok(count.0)
    }

    /// Serializes a canonical, validated document from the modeled semantics.
    ///
    /// # Publication safety
    ///
    /// This is not a lossless rewrite API for arbitrary XML passed to
    /// [`Self::parse`]. Rich-text formatting, shape properties, backgrounds,
    /// whole-diagram formatting, and extension lists are outside this model and
    /// will not be emitted. Use this method for fresh authoring or only after
    /// deliberately accepting that canonicalization.
    pub fn to_xml(&self, conformance: Conformance) -> Result<String> {
        let capacity = self.serialized_xml_len(conformance)?;
        let mut xml = String::new();
        xml.try_reserve_exact(capacity).map_err(reserve_error)?;
        self.write_validated_xml(&mut xml, conformance)?;
        Ok(xml)
    }

    /// Serializes modeled semantics into a caller-owned sink for allocation
    /// reuse.
    ///
    /// This has the same non-lossless publication contract as [`Self::to_xml`].
    /// Validation and allocation complete before the destination is changed, so
    /// an error leaves its previous contents intact.
    pub fn write_xml(&self, xml: &mut String, conformance: Conformance) -> Result<()> {
        let additional = self.serialized_xml_len(conformance)?;
        xml.try_reserve_exact(additional).map_err(reserve_error)?;
        let initial_len = xml.len();
        if let Err(error) = self.write_validated_xml(xml, conformance) {
            xml.truncate(initial_len);
            return Err(error);
        }
        Ok(())
    }

    fn write_validated_xml(
        &self,
        xml: &mut impl fmt::Write,
        conformance: Conformance,
    ) -> Result<()> {
        let (diagram_namespace, main_namespace) = match conformance {
            Conformance::Transitional => (DGM_NAMESPACE, DML_NAMESPACE),
            Conformance::Strict => (DGM_NAMESPACE_STRICT, DML_NAMESPACE_STRICT),
        };
        write!(
            xml,
            "<dgm:dataModel xmlns:dgm=\"{diagram_namespace}\" xmlns:a=\"{main_namespace}\"><dgm:ptLst>"
        )
        .map_err(write_error)?;
        for point in &self.points {
            write_point(xml, point)?;
        }
        xml.write_str("</dgm:ptLst><dgm:cxnLst>")
            .map_err(write_error)?;
        for connection in &self.connections {
            write_connection(xml, connection)?;
        }
        xml.write_str("</dgm:cxnLst></dgm:dataModel>")
            .map_err(write_error)
    }

    /// The document root point (`type="doc"`), if present.
    pub fn document_point(&self) -> Option<&Point> {
        self.points
            .iter()
            .find(|point| point.kind == PointType::Document)
    }

    /// Iterate over content points (nodes and assistants, in document order).
    pub fn content_points(&self) -> impl Iterator<Item = &Point> {
        self.points
            .iter()
            .filter(|point| matches!(point.kind, PointType::Node | PointType::Assistant))
    }

    /// Build the content-node hierarchy implied by the `parOf` connection
    /// graph, ordered by `srcOrd`. Cycles and dangling references are
    /// tolerated: nodes not reachable from the document root are omitted.
    pub fn node_tree(&self) -> Vec<DiagramNode> {
        let Some(root) = self.document_point() else {
            return Vec::new();
        };
        let points: HashMap<Id, &Point> =
            self.points.iter().map(|point| (point.id, point)).collect();
        let mut children: HashMap<Id, Vec<(u32, Id)>> = HashMap::new();
        for connection in &self.connections {
            if matches!(connection.kind, ConnectionType::Parent { .. }) {
                children
                    .entry(connection.source)
                    .or_default()
                    .push((connection.src_ord, connection.destination));
            }
        }
        for entries in children.values_mut() {
            entries.sort_by_key(|(ord, _)| *ord);
        }
        let mut visiting = HashSet::new();
        build_children(root.id, 0, &points, &children, &mut visiting)
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

fn validate_connection_transitions<'a>(
    connection: &Connection,
    mut point: impl FnMut(Id) -> Option<&'a Point>,
) -> Result<()> {
    let ConnectionType::Parent {
        parent_transition,
        sibling_transition,
    } = &connection.kind
    else {
        return Ok(());
    };
    let parent = point(*parent_transition).ok_or_else(|| {
        invalid(format!(
            "diagram parent connection {} refers to missing parent transition point {}",
            connection.id, parent_transition
        ))
    })?;
    if parent.kind != PointType::ParentTransition(connection.id) {
        return Err(invalid(format!(
            "diagram point {} is not the parent transition for connection {}",
            parent.id, connection.id
        )));
    }
    let sibling = point(*sibling_transition).ok_or_else(|| {
        invalid(format!(
            "diagram parent connection {} refers to missing sibling transition point {}",
            connection.id, sibling_transition
        ))
    })?;
    if sibling.kind != PointType::SiblingTransition(connection.id) {
        return Err(invalid(format!(
            "diagram point {} is not the sibling transition for connection {}",
            sibling.id, connection.id
        )));
    }
    Ok(())
}

fn validate_transition_point(point: &Point, connection: &Connection) -> Result<()> {
    let matches = match point.kind {
        PointType::ParentTransition(owner) if owner == connection.id => matches!(
            &connection.kind,
            ConnectionType::Parent {
                parent_transition,
                ..
            } if *parent_transition == point.id
        ),
        PointType::SiblingTransition(owner) if owner == connection.id => matches!(
            &connection.kind,
            ConnectionType::Parent {
                sibling_transition,
                ..
            } if *sibling_transition == point.id
        ),
        _ => false,
    };
    if matches {
        Ok(())
    } else {
        Err(invalid(format!(
            "diagram transition point {} and connection {} do not refer to each other",
            point.id, connection.id
        )))
    }
}

fn build_children(
    parent_id: Id,
    depth: u32,
    points: &HashMap<Id, &Point>,
    children: &HashMap<Id, Vec<(u32, Id)>>,
    visiting: &mut HashSet<Id>,
) -> Vec<DiagramNode> {
    if depth >= MAX_TREE_DEPTH || !visiting.insert(parent_id) {
        return Vec::new();
    }
    let mut nodes = Vec::new();
    if let Some(entries) = children.get(&parent_id) {
        for (_, destination) in entries {
            let Some(point) = points.get(destination) else {
                continue;
            };
            if !matches!(point.kind, PointType::Node | PointType::Assistant)
                || visiting.contains(destination)
            {
                continue;
            }
            let mut node = DiagramNode::new(point.text.clone());
            node.depth = depth;
            let Some(child_depth) = depth.checked_add(1) else {
                continue;
            };
            node.children = build_children(*destination, child_depth, points, children, visiting);
            nodes.push(node);
        }
    }
    visiting.remove(&parent_id);
    nodes
}

fn write_point(xml: &mut impl fmt::Write, point: &Point) -> Result<()> {
    write!(xml, "<dgm:pt modelId=\"{}\"", point.id).map_err(write_error)?;
    match point.kind {
        PointType::Node => {},
        PointType::Document => xml.write_str(" type=\"doc\"").map_err(write_error)?,
        PointType::Assistant => xml.write_str(" type=\"asst\"").map_err(write_error)?,
        PointType::ParentTransition(connection) => {
            write!(xml, " type=\"parTrans\" cxnId=\"{connection}\"").map_err(write_error)?
        },
        PointType::SiblingTransition(connection) => {
            write!(xml, " type=\"sibTrans\" cxnId=\"{connection}\"").map_err(write_error)?
        },
        PointType::Presentation => xml.write_str(" type=\"pres\"").map_err(write_error)?,
    }

    let has_property_set = point.layout_type_id.is_some()
        || point.quick_style_type_id.is_some()
        || point.color_style_type_id.is_some();
    if !has_property_set && point.text.is_empty() {
        return xml.write_str("/>").map_err(write_error);
    }
    xml.write_char('>').map_err(write_error)?;
    if has_property_set {
        xml.write_str("<dgm:prSet").map_err(write_error)?;
        write_optional_attribute(xml, "loTypeId", point.layout_type_id.as_deref())?;
        write_optional_attribute(xml, "qsTypeId", point.quick_style_type_id.as_deref())?;
        write_optional_attribute(xml, "csTypeId", point.color_style_type_id.as_deref())?;
        xml.write_str("/>").map_err(write_error)?;
    }
    if !point.text.is_empty() {
        xml.write_str("<dgm:t><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t xml:space=\"preserve\">")
            .map_err(write_error)?;
        write_xml_escaped(xml, &point.text)?;
        xml.write_str("</a:t></a:r></a:p></dgm:t>")
            .map_err(write_error)?;
    }
    xml.write_str("</dgm:pt>").map_err(write_error)
}

fn write_connection(xml: &mut impl fmt::Write, connection: &Connection) -> Result<()> {
    write!(
        xml,
        "<dgm:cxn modelId=\"{}\" srcId=\"{}\" destId=\"{}\" srcOrd=\"{}\" destOrd=\"{}\"",
        connection.id,
        connection.source,
        connection.destination,
        connection.src_ord,
        connection.dest_ord,
    )
    .map_err(write_error)?;
    match &connection.kind {
        ConnectionType::Parent {
            parent_transition,
            sibling_transition,
        } => write!(
            xml,
            " parTransId=\"{parent_transition}\" sibTransId=\"{sibling_transition}\""
        )
        .map_err(write_error)?,
        ConnectionType::Presentation(presentation) => {
            xml.write_str(" type=\"presOf\"").map_err(write_error)?;
            write_optional_attribute(xml, "presId", Some(presentation))?;
        },
        ConnectionType::PresentationParent => {
            xml.write_str(" type=\"presParOf\"").map_err(write_error)?
        },
        ConnectionType::Unknown => xml
            .write_str(" type=\"unknownRelationship\"")
            .map_err(write_error)?,
    }
    xml.write_str("/>").map_err(write_error)
}

fn write_optional_attribute(
    xml: &mut impl fmt::Write,
    name: &str,
    value: Option<&str>,
) -> Result<()> {
    if let Some(value) = value {
        write!(xml, " {name}=\"").map_err(write_error)?;
        write_xml_escaped(xml, value)?;
        xml.write_char('"').map_err(write_error)?;
    }
    Ok(())
}

fn write_error(error: fmt::Error) -> Error {
    Error::Xml(format!("failed to write diagram XML: {error}"))
}

fn reserve_error(error: std::collections::TryReserveError) -> Error {
    invalid(format!("cannot reserve diagram XML output: {error}"))
}

fn write_xml_escaped(xml: &mut impl fmt::Write, value: &str) -> Result<()> {
    let mut start = 0usize;
    for (offset, character) in value.char_indices() {
        let replacement = match character {
            '&' => Some("&amp;"),
            '<' => Some("&lt;"),
            '>' => Some("&gt;"),
            '\'' => Some("&apos;"),
            '"' => Some("&quot;"),
            _ => None,
        };
        if let Some(replacement) = replacement {
            xml.write_str(&value[start..offset]).map_err(write_error)?;
            xml.write_str(replacement).map_err(write_error)?;
            start = offset + character.len_utf8();
        }
    }
    xml.write_str(&value[start..]).map_err(write_error)
}

#[derive(Default)]
struct XmlByteCount(usize);

impl fmt::Write for XmlByteCount {
    fn write_str(&mut self, value: &str) -> fmt::Result {
        self.0 = self.0.saturating_add(value.len());
        Ok(())
    }
}

fn add_xml_value(total: usize, value: &str, description: &str) -> Result<usize> {
    let mut escaped = 0usize;
    for character in value.chars() {
        if !is_xml_character(character) {
            return Err(invalid(format!(
                "{description} contains a character forbidden by XML 1.0"
            )));
        }
        let bytes = match character {
            '&' => 5,
            '<' | '>' => 4,
            '\'' | '"' => 6,
            character => character.len_utf8(),
        };
        escaped = escaped
            .checked_add(bytes)
            .ok_or_else(|| limit("serialized data-model bytes"))?;
    }
    total
        .checked_add(escaped)
        .ok_or_else(|| limit("serialized data-model bytes"))
}

#[inline]
const fn is_xml_character(character: char) -> bool {
    matches!(
        character as u32,
        0x9 | 0xA | 0xD | 0x20..=0xD7FF | 0xE000..=0xFFFD | 0x10000..=0x10FFFF
    )
}

#[derive(Debug, Clone, Copy, Default)]
enum PointTag {
    #[default]
    Node,
    Document,
    Assistant,
    ParentTransition,
    SiblingTransition,
    Presentation,
}

fn parse_point_tag(value: &str) -> Result<PointTag> {
    match xml_token(value) {
        "node" => Ok(PointTag::Node),
        "doc" => Ok(PointTag::Document),
        "asst" => Ok(PointTag::Assistant),
        "parTrans" => Ok(PointTag::ParentTransition),
        "sibTrans" => Ok(PointTag::SiblingTransition),
        "pres" => Ok(PointTag::Presentation),
        value => Err(invalid(format!("invalid diagram point type `{value}`"))),
    }
}

#[derive(Default)]
struct PointBuilder {
    id: Option<Id>,
    tag: PointTag,
    connection: Option<Id>,
    text: String,
    layout_type_id: Option<String>,
    quick_style_type_id: Option<String>,
    color_style_type_id: Option<String>,
    property_set_seen: bool,
    text_body_seen: bool,
}

impl PointBuilder {
    fn from_element(element: &BytesStart<'_>) -> Result<Self> {
        let mut builder = PointBuilder::default();
        for (name, value) in attributes(element)? {
            match name.as_str() {
                "modelId" => builder.id = Some(parse_id(&value, "diagram point modelId")?),
                "type" => builder.tag = parse_point_tag(&value)?,
                "cxnId" => builder.connection = Some(parse_id(&value, "diagram point cxnId")?),
                _ => {},
            }
        }
        if builder.id.is_none() {
            return Err(invalid("diagram point lacks modelId"));
        }
        Ok(builder)
    }

    fn read_pr_set(&mut self, element: &BytesStart<'_>) -> Result<()> {
        if self.property_set_seen {
            return Err(invalid("duplicate diagram point property set"));
        }
        self.property_set_seen = true;
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

    fn finish(self) -> Result<Point> {
        let id = self
            .id
            .ok_or_else(|| invalid("diagram point lacks modelId"))?;
        let kind = match self.tag {
            PointTag::Node => PointType::Node,
            PointTag::Document => PointType::Document,
            PointTag::Assistant => PointType::Assistant,
            PointTag::ParentTransition => {
                PointType::ParentTransition(self.connection.unwrap_or_else(|| Id::number(0)))
            },
            PointTag::SiblingTransition => {
                PointType::SiblingTransition(self.connection.unwrap_or_else(|| Id::number(0)))
            },
            PointTag::Presentation => PointType::Presentation,
        };
        Ok(Point {
            id,
            kind,
            text: self.text,
            layout_type_id: self.layout_type_id,
            quick_style_type_id: self.quick_style_type_id,
            color_style_type_id: self.color_style_type_id,
        })
    }
}

#[derive(Debug, Clone, Copy, Default)]
enum ConnectionTag {
    #[default]
    Parent,
    Presentation,
    PresentationParent,
    Unknown,
}

fn parse_connection_tag(value: &str) -> Result<ConnectionTag> {
    match xml_token(value) {
        "parOf" => Ok(ConnectionTag::Parent),
        "presOf" => Ok(ConnectionTag::Presentation),
        "presParOf" => Ok(ConnectionTag::PresentationParent),
        "unknownRelationship" => Ok(ConnectionTag::Unknown),
        value => Err(invalid(format!(
            "invalid diagram connection type `{value}`"
        ))),
    }
}

fn push_connection(model: &mut DiagramDataModel, element: &BytesStart<'_>) -> Result<()> {
    if model.connections.len() >= MAX_CONNECTIONS {
        return Err(limit("diagram connection count"));
    }
    let mut id = None;
    let mut tag = ConnectionTag::default();
    let mut source = None;
    let mut destination = None;
    let mut source_order = None;
    let mut destination_order = None;
    let mut parent_transition = None;
    let mut sibling_transition = None;
    let mut presentation = None;
    for (name, value) in attributes(element)? {
        match name.as_str() {
            "modelId" => id = Some(parse_id(&value, "diagram connection modelId")?),
            "type" => tag = parse_connection_tag(&value)?,
            "srcId" => source = Some(parse_id(&value, "diagram connection srcId")?),
            "destId" => destination = Some(parse_id(&value, "diagram connection destId")?),
            "srcOrd" => source_order = Some(parse_order(&value)?),
            "destOrd" => destination_order = Some(parse_order(&value)?),
            "parTransId" => {
                parent_transition = Some(parse_id(&value, "diagram connection parTransId")?)
            },
            "sibTransId" => {
                sibling_transition = Some(parse_id(&value, "diagram connection sibTransId")?)
            },
            "presId" => presentation = Some(value),
            _ => {},
        }
    }
    let id = id.ok_or_else(|| invalid("diagram connection lacks modelId"))?;
    let source = source.ok_or_else(|| invalid("diagram connection lacks srcId"))?;
    let destination = destination.ok_or_else(|| invalid("diagram connection lacks destId"))?;
    let source_order = source_order.ok_or_else(|| invalid("diagram connection lacks srcOrd"))?;
    let destination_order =
        destination_order.ok_or_else(|| invalid("diagram connection lacks destOrd"))?;
    let kind = match tag {
        ConnectionTag::Parent => ConnectionType::Parent {
            parent_transition: parent_transition.unwrap_or_else(|| Id::number(0)),
            sibling_transition: sibling_transition.unwrap_or_else(|| Id::number(0)),
        },
        ConnectionTag::Presentation => {
            ConnectionType::Presentation(presentation.unwrap_or_default())
        },
        ConnectionTag::PresentationParent => ConnectionType::PresentationParent,
        ConnectionTag::Unknown => ConnectionType::Unknown,
    };
    model.connections.push(Connection::new(
        id,
        kind,
        source,
        destination,
        source_order,
        destination_order,
    ));
    Ok(())
}

fn parse_order(value: &str) -> Result<u32> {
    xml_token(value)
        .parse()
        .map_err(|_| invalid("invalid diagram connection order"))
}

fn parse_id(value: &str, description: &str) -> Result<Id> {
    value
        .parse()
        .map_err(|_| invalid(format!("invalid {description} `{value}`")))
}

fn xml_token(value: &str) -> &str {
    value.trim_matches([' ', '\t', '\r', '\n'])
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
        // Schema attributes on these elements are unqualified. An extension
        // attribute with the same local name must not impersonate one merely
        // because namespace prefixes are otherwise discarded below.
        if raw.contains(&b':') {
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

fn is_dml(namespace: &str) -> bool {
    matches!(namespace, DML_NAMESPACE | DML_NAMESPACE_STRICT)
}

fn xml_error(error: impl fmt::Display) -> Error {
    Error::Xml(error.to_string())
}

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}

fn limit(label: &str) -> Error {
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
        "<dgm:pt modelId=\"{00000000-0000-0000-0000-000000000001}\" type=\"doc\"><dgm:prSet loTypeId=\"urn:test/layout/cycle2\"/></dgm:pt>",
        "<dgm:pt modelId=\"{00000000-0000-0000-0000-000000000002}\"><dgm:t><a:p><a:r><a:t>a</a:t></a:r></a:p></dgm:t></dgm:pt>",
        "<dgm:pt modelId=\"{00000000-0000-0000-0000-000000000003}\" type=\"sibTrans\" cxnId=\"{00000000-0000-0000-0000-000000000004}\"/>",
        "</dgm:ptLst>",
        "<dgm:cxnLst>",
        "<dgm:cxn modelId=\"{00000000-0000-0000-0000-000000000004}\" srcId=\"{00000000-0000-0000-0000-000000000001}\" destId=\"{00000000-0000-0000-0000-000000000002}\" srcOrd=\"0\" destOrd=\"0\"/>",
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
        assert_eq!(
            model.points[4].kind,
            PointType::ParentTransition(Id::number(100))
        );
        assert_eq!(model.points[5].kind, PointType::Presentation);
        assert_eq!(
            model.connections[3].kind,
            ConnectionType::Presentation(String::new())
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
        let connection_id: Id = "{00000000-0000-0000-0000-000000000004}".parse().unwrap();
        assert_eq!(
            model.points[2].kind,
            PointType::SiblingTransition(connection_id)
        );
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
        assert!(model.validate().is_err());
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
                "<dgm:dataModel xmlns:dgm=\"http://schemas.openxmlformats.org/drawingml/2006/diagram\"><dgm:ptLst/><dgm:cxnLst><dgm:cxn modelId=\"1\" srcId=\"0\"/></dgm:cxnLst></dgm:dataModel>"
            )
            .is_err()
        );
    }

    #[test]
    fn enforces_data_model_structure_and_unqualified_schema_attributes() {
        let namespace = "http://schemas.openxmlformats.org/drawingml/2006/diagram";
        for invalid_xml in [
            format!("<dgm:dataModel xmlns:dgm=\"{namespace}\"/>"),
            format!(
                "<dgm:dataModel xmlns:dgm=\"{namespace}\"><dgm:cxnLst/><dgm:ptLst/></dgm:dataModel>"
            ),
            format!(
                "<dgm:dataModel xmlns:dgm=\"{namespace}\"><dgm:pt modelId=\"1\"/><dgm:ptLst/></dgm:dataModel>"
            ),
            format!(
                "<dgm:dataModel xmlns:dgm=\"{namespace}\"><dgm:ptLst><dgm:pt modelId=\"1\"><dgm:prSet/><dgm:prSet/></dgm:pt></dgm:ptLst></dgm:dataModel>"
            ),
            format!(
                "<dgm:dataModel xmlns:dgm=\"{namespace}\" xmlns:x=\"urn:extension\"><dgm:ptLst><dgm:pt x:modelId=\"1\"/></dgm:ptLst></dgm:dataModel>"
            ),
        ] {
            assert!(
                DiagramDataModel::parse(&invalid_xml).is_err(),
                "accepted structurally invalid XML: {invalid_xml}"
            );
        }
    }

    #[test]
    fn extracts_only_drawingml_text_leaf_content() {
        let xml = concat!(
            "<dgm:dataModel xmlns:dgm=\"http://schemas.openxmlformats.org/drawingml/2006/diagram\" ",
            "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">",
            "<dgm:ptLst><dgm:pt modelId=\"1\"><dgm:t>\n",
            "  <a:p><a:r><a:t>Alpha </a:t></a:r>\n",
            "  <a:r><a:t>Beta</a:t></a:r></a:p>\n",
            "</dgm:t></dgm:pt></dgm:ptLst></dgm:dataModel>"
        );
        assert_eq!(
            DiagramDataModel::parse(xml).unwrap().points[0].text,
            "Alpha Beta"
        );
    }

    #[test]
    fn model_id_is_a_closed_zero_allocation_wire_domain() {
        assert_eq!(" -2147483648 ".parse::<Id>().unwrap(), Id::number(i32::MIN));
        let guid: Id = "{01234567-89AB-CDEF-0123-456789ABCDEF}".parse().unwrap();
        assert_eq!(guid.to_string(), "{01234567-89AB-CDEF-0123-456789ABCDEF}");
        assert_eq!(
            guid.as_guid(),
            Some([
                0x01, 0x23, 0x45, 0x67, 0x89, 0xAB, 0xCD, 0xEF, 0x01, 0x23, 0x45, 0x67, 0x89, 0xAB,
                0xCD, 0xEF,
            ])
        );
        for invalid in [
            "2147483648",
            "id-1",
            "{01234567-89ab-CDEF-0123-456789ABCDEF}",
            "01234567-89AB-CDEF-0123-456789ABCDEF",
            "{01234567-89AB-CDEF-0123-456789ABCDE}",
        ] {
            assert!(invalid.parse::<Id>().is_err(), "accepted {invalid}");
        }
    }

    #[test]
    fn rejects_values_outside_fixed_point_and_connection_domains() {
        let invalid_point = concat!(
            "<dgm:dataModel xmlns:dgm=\"http://schemas.openxmlformats.org/drawingml/2006/diagram\">",
            "<dgm:ptLst><dgm:pt modelId=\"1\" type=\"futureNode\"/></dgm:ptLst>",
            "</dgm:dataModel>"
        );
        assert!(DiagramDataModel::parse(invalid_point).is_err());

        let invalid_connection = concat!(
            "<dgm:dataModel xmlns:dgm=\"http://schemas.openxmlformats.org/drawingml/2006/diagram\">",
            "<dgm:ptLst/>",
            "<dgm:cxnLst><dgm:cxn modelId=\"1\" type=\"futureRelation\" srcId=\"2\" destId=\"3\" srcOrd=\"0\" destOrd=\"0\"/></dgm:cxnLst>",
            "</dgm:dataModel>"
        );
        assert!(DiagramDataModel::parse(invalid_connection).is_err());

        let invalid_identifier = concat!(
            "<dgm:dataModel xmlns:dgm=\"http://schemas.openxmlformats.org/drawingml/2006/diagram\">",
            "<dgm:ptLst><dgm:pt modelId=\"node-one\"/></dgm:ptLst>",
            "</dgm:dataModel>"
        );
        assert!(DiagramDataModel::parse(invalid_identifier).is_err());
    }

    #[test]
    fn accepts_all_schema_connection_types_without_string_fallback() {
        let xml = concat!(
            "<dgm:dataModel xmlns:dgm=\"http://schemas.openxmlformats.org/drawingml/2006/diagram\">",
            "<dgm:ptLst/>",
            "<dgm:cxnLst>",
            "<dgm:cxn modelId=\"1\" type=\"presParOf\" srcId=\"2\" destId=\"3\" srcOrd=\"0\" destOrd=\"0\"/>",
            "<dgm:cxn modelId=\"4\" type=\"unknownRelationship\" srcId=\"2\" destId=\"3\" srcOrd=\"1\" destOrd=\"0\"/>",
            "</dgm:cxnLst></dgm:dataModel>"
        );
        let model = DiagramDataModel::parse(xml).unwrap();
        assert_eq!(
            model.connections[0].kind,
            ConnectionType::PresentationParent
        );
        assert_eq!(model.connections[1].kind, ConnectionType::Unknown);
    }

    #[test]
    fn canonical_writer_round_trips_both_conformance_classes() {
        let mut model = DiagramDataModel::new();
        for point in [
            Point::new(Id::number(0), PointType::Document),
            Point::node(Id::number(1), "A & B"),
            Point::new(Id::number(2), PointType::ParentTransition(Id::number(10))),
            Point::new(Id::number(3), PointType::SiblingTransition(Id::number(10))),
            Point::new(Id::number(4), PointType::Presentation),
        ] {
            model.add_point(point).unwrap();
        }
        model
            .add_connection(Connection::new(
                Id::number(10),
                ConnectionType::parent(Id::number(2), Id::number(3)),
                Id::number(0),
                Id::number(1),
                0,
                0,
            ))
            .unwrap();
        model
            .add_connection(Connection::new(
                Id::number(11),
                ConnectionType::Presentation("urn:test/layout".to_string()),
                Id::number(1),
                Id::number(4),
                0,
                0,
            ))
            .unwrap();

        let xml = model.to_xml(Conformance::Transitional).unwrap();
        assert!(xml.contains("parTransId=\"2\" sibTransId=\"3\""));
        assert!(xml.contains("type=\"presOf\" presId=\"urn:test/layout\""));
        assert_eq!(DiagramDataModel::parse(&xml).unwrap(), model);

        let xml = model.to_xml(Conformance::Strict).unwrap();
        assert!(xml.contains(DGM_NAMESPACE_STRICT));
        assert!(xml.contains("http://purl.oclc.org/ooxml/drawingml/main"));
        assert_eq!(DiagramDataModel::parse(&xml).unwrap(), model);
    }

    #[test]
    fn semantic_crud_guards_duplicates_and_cascades_dependencies() {
        let mut model = DiagramDataModel::new();
        model
            .add_point(Point::new(Id::number(0), PointType::Document))
            .unwrap();
        model.add_point(Point::node(Id::number(1), "one")).unwrap();
        model
            .add_point(Point::new(
                Id::number(2),
                PointType::ParentTransition(Id::number(10)),
            ))
            .unwrap();
        model
            .add_point(Point::new(
                Id::number(3),
                PointType::SiblingTransition(Id::number(10)),
            ))
            .unwrap();
        model
            .add_connection(Connection::new(
                Id::number(10),
                ConnectionType::parent(Id::number(2), Id::number(3)),
                Id::number(0),
                Id::number(1),
                0,
                0,
            ))
            .unwrap();
        assert!(model.validate().is_ok());
        assert!(
            model
                .add_point(Point::node(Id::number(10), "collision"))
                .is_err()
        );

        let mut broken_transition = model.clone();
        broken_transition.point_mut(Id::number(2)).unwrap().kind =
            PointType::ParentTransition(Id::number(99));
        assert!(broken_transition.validate().is_err());
        assert!(broken_transition.to_xml(Conformance::Transitional).is_err());

        let mut cross_domain_collision = model.clone();
        cross_domain_collision.connections[0].id = Id::number(1);
        assert!(cross_domain_collision.validate().is_err());

        let conflicting_parent = Connection::new(
            Id::number(11),
            ConnectionType::PresentationParent,
            Id::number(0),
            Id::number(1),
            1,
            0,
        );
        assert!(model.add_connection(conflicting_parent.clone()).is_err());
        model.connections.push(conflicting_parent);
        assert!(model.validate().is_err());
        model.connections.pop();

        model
            .add_point(Point::new(Id::number(4), PointType::Presentation))
            .unwrap();
        model
            .add_connection(Connection::new(
                Id::number(11),
                ConnectionType::Presentation("urn:layout/a".to_string()),
                Id::number(0),
                Id::number(4),
                0,
                0,
            ))
            .unwrap();
        assert!(
            model
                .add_connection(Connection::new(
                    Id::number(12),
                    ConnectionType::Presentation("urn:layout/b".to_string()),
                    Id::number(1),
                    Id::number(4),
                    0,
                    0,
                ))
                .is_err()
        );
        assert!(
            model
                .add_point(Point::node(Id::number(1), "duplicate"))
                .is_err()
        );
        model.point_mut(Id::number(1)).unwrap().text = "updated".to_string();
        assert_eq!(model.point(Id::number(1)).unwrap().text, "updated");

        let removed = model.remove_connection(Id::number(10)).unwrap();
        assert_eq!(removed.destination, Id::number(1));
        assert!(model.point(Id::number(2)).is_none());
        assert!(model.point(Id::number(3)).is_none());
    }

    #[test]
    fn canonical_writer_enforces_xml_and_aggregate_size_budgets() {
        let mut invalid_text = DiagramDataModel::new();
        invalid_text
            .add_point(Point::node(Id::number(0), "invalid\0text"))
            .unwrap();
        assert!(invalid_text.to_xml(Conformance::Transitional).is_err());
        let mut destination = "unchanged".to_string();
        assert!(
            invalid_text
                .write_xml(&mut destination, Conformance::Transitional)
                .is_err()
        );
        assert_eq!(destination, "unchanged");

        let mut oversized = DiagramDataModel::new();
        oversized
            .add_point(Point::new(Id::number(0), PointType::Document))
            .unwrap();
        oversized
            .add_point(Point::node(Id::number(1), "child"))
            .unwrap();
        oversized
            .add_connection(Connection::new(
                Id::number(2),
                ConnectionType::Presentation("x".repeat(MAX_DATA_MODEL_XML)),
                Id::number(0),
                Id::number(1),
                0,
                0,
            ))
            .unwrap();
        assert!(oversized.validate().is_err());
        assert!(oversized.to_xml(Conformance::Transitional).is_err());
    }
}
