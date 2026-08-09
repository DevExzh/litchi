//! Typed diagram data-model semantics and graph operations.

use crate::Result;
use crate::diagram::DiagramNode;
use std::collections::{HashMap, HashSet};
use std::fmt;
use std::str::FromStr;

use super::validation::{
    invalid, limit, validate_connection_transitions, validate_transition_point,
};
use super::{MAX_CONNECTIONS, MAX_POINTS, MAX_TEXT_BYTES, MAX_TREE_DEPTH};
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
    #[must_use]
    pub const fn number(value: i32) -> Self {
        Self::Number(value)
    }

    /// Creates a GUID model identifier from its 16 bytes.
    #[inline]
    #[must_use]
    pub const fn guid(value: [u8; 16]) -> Self {
        Self::Guid(value)
    }

    /// Returns the numeric value, if this is a numeric identifier.
    #[inline]
    #[must_use]
    pub const fn as_number(self) -> Option<i32> {
        match self {
            Self::Number(value) => Some(value),
            Self::Guid(_) => None,
        }
    }

    /// Returns the GUID bytes, if this is a GUID identifier.
    #[inline]
    #[must_use]
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
    #[must_use]
    pub const fn connection(self) -> Option<Id> {
        match self {
            Self::ParentTransition(connection) | Self::SiblingTransition(connection) => {
                Some(connection)
            },
            Self::Node | Self::Document | Self::Assistant | Self::Presentation => None,
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
    #[must_use]
    pub const fn parent(parent_transition: Id, sibling_transition: Id) -> Self {
        Self::Parent {
            parent_transition,
            sibling_transition,
        }
    }

    /// Returns whether Office treats this relation as a parent edge for its
    /// one-parent-per-destination constraint.
    #[inline]
    #[must_use]
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
    #[must_use]
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
    #[must_use]
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

/// The modeled semantic subset of a `DrawingML` diagram data model.
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
    /// Creates an empty data model.
    #[inline]
    #[must_use]
    pub const fn new() -> Self {
        Self {
            points: Vec::new(),
            connections: Vec::new(),
        }
    }

    /// Looks up a point by its typed model identifier.
    #[inline]
    #[must_use]
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
    #[must_use]
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
    /// # Errors
    ///
    /// Returns an error when input violates DrawingML constraints, exceeds a configured
    /// bound, or an underlying XML, MCE, I/O, or formatting operation fails.
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
    /// # Errors
    ///
    /// Returns an error when input violates DrawingML constraints, exceeds a configured
    /// bound, or an underlying XML, MCE, I/O, or formatting operation fails.
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
                    ConnectionType::Presentation(_)
                    | ConnectionType::PresentationParent
                    | ConnectionType::Unknown => false,
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
    /// The document root point (`type="doc"`), if present.
    #[must_use]
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
    #[must_use]
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
    #[must_use]
    pub fn text(&self) -> String {
        self.node_tree()
            .iter()
            .map(DiagramNode::all_text)
            .collect::<Vec<_>>()
            .join("\n")
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
