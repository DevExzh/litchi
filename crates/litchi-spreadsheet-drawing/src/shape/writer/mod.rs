#![allow(
    clippy::arbitrary_source_item_ordering,
    clippy::missing_errors_doc,
    clippy::ref_option,
    clippy::return_self_not_must_use,
    clippy::shadow_reuse,
    clippy::shadow_unrelated,
    clippy::unreadable_literal,
    reason = "The compact XML authoring implementation preserves its established public API."
)]

//! Shared `SpreadsheetDrawing` authoring model for XLSX and XLSB shapes.
//!
//! [`ShapeSpec`] describes one text-box shape to embed in a worksheet
//! drawing part: a preset geometry, an anchor on the sheet grid, text-body
//! properties, and rich-text runs. [`GroupSpec`] nests shapes, text
//! boxes, and further groups under a `xdr:grpSp` group with an optional
//! coordinate transform, and [`ConnectionShapeSpec`] authors `xdr:cxnSp`
//! connectors whose `a:stCxn`/`a:endCxn` sites reference other shapes by
//! name. All three deliberately reuse the typed read model from
//! the shared shape model (`Anchor`,
//! `Properties`, `GroupTransform`, paragraphs and runs) so
//! anything authored here round-trips through the shape inventory with
//! identical semantics.
//!
//! `State` serializes authored objects as
//! `xdr:twoCellAnchor`/`xdr:oneCellAnchor`/`xdr:absoluteAnchor` elements for
//! the standard worksheet drawing parts used by both XLSX and XLSB,
//! allocating drawing-wide unique object IDs and resolving connection-site
//! name references. Everything is inert: no rendering, no layout
//! computation, and all inputs are bounded and validated at authoring time.

use std::collections::{HashMap, HashSet};
use std::fmt::Write as _;

use litchi_core::xml::escape::escape_xml;
use litchi_drawingml::geom::Preset;
use litchi_drawingml::text::body::writer::write_contents as write_text_body_contents;

use super::{
    Anchor, CellMarker, EditAs, EmuExtent, EmuOffset, Geometry, GroupTransform, Paragraph,
    Properties, Run,
};
use crate::Error;
use litchi_drawingml::geometry::writer::write_custom_geometry;
use litchi_drawingml::geometry::{CustomGeometry, validate_custom_geometry};

macro_rules! write_xml {
    ($xml:expr, $($arguments:tt)*) => {
        if write!($xml, $($arguments)*).is_err() {
            unreachable!("writing to String is infallible");
        }
    };
}

/// Maximum number of authored top-level drawing objects per worksheet.
const MAX_SHAPES_PER_WORKSHEET: usize = 4096;
/// Maximum authored objects including nested group children.
const MAX_OBJECTS_PER_DRAWING: usize = 100_000;
/// Maximum aggregate run text bytes across one authored shape.
const MAX_SHAPE_TEXT_BYTES: usize = 1024 * 1024;
/// Maximum length of the shape name or description.
const MAX_SHAPE_NAME_BYTES: usize = 4096;
/// Maximum paragraphs in one authored shape.
const MAX_SHAPE_PARAGRAPHS: usize = 16_384;
/// Maximum runs in one authored paragraph.
const MAX_RUNS_PER_PARAGRAPH: usize = 4096;
/// Maximum children in one authored group.
const MAX_GROUP_CHILDREN: usize = 1024;
/// Maximum group nesting depth, mirroring the reader's limit.
const MAX_GROUP_DEPTH: usize = 32;

/// One authored `DrawingML` text-box shape for a worksheet drawing part.
///
/// Construct with [`ShapeSpec::text_box`] (or [`ShapeSpec::shape`]
/// for a non-text-box shape), then adjust `properties`, `paragraphs`,
/// or flags before handing it to `MutableWorksheet::add_shape`.
#[derive(Debug, Clone, PartialEq)]
pub struct ShapeSpec {
    /// Shape name written to `xdr:cNvPr@name`.
    pub name: String,
    /// Alternative text (`xdr:cNvPr@descr`), when set.
    pub description: Option<String>,
    /// Whether the shape is hidden (`xdr:cNvPr@hidden`).
    pub hidden: bool,
    /// Whether the shape is marked as a text box (`xdr:cNvSpPr@txBox`).
    pub is_text_box: bool,
    /// How the shape is anchored on the worksheet grid.
    pub anchor: Anchor,
    /// The shape's mutually exclusive preset or custom geometry.
    pub geometry: Geometry,
    /// Text-body properties (`a:bodyPr`).
    pub properties: Properties,
    /// The text story as paragraphs with runs.
    pub paragraphs: Vec<Paragraph>,
}

impl ShapeSpec {
    /// A text-box shape with the given preset, anchor, and plain-text story.
    ///
    /// `text` is split into paragraphs on `\n`; each paragraph becomes one
    /// unformatted run. Body properties start at the ECMA-376 defaults.
    pub fn text_box(name: impl Into<String>, anchor: Anchor, preset: Preset, text: &str) -> Self {
        Self {
            is_text_box: true,
            ..Self::shape(name, anchor, preset, text)
        }
    }

    /// A plain (non-text-box) shape with the given preset, anchor, and text.
    pub fn shape(name: impl Into<String>, anchor: Anchor, preset: Preset, text: &str) -> Self {
        Self::with_geometry(name, anchor, Geometry::Preset(preset), text)
    }

    fn with_geometry(
        name: impl Into<String>,
        anchor: Anchor,
        geometry: Geometry,
        text: &str,
    ) -> Self {
        let paragraphs = text
            .split('\n')
            .map(|line| Paragraph {
                runs: vec![Run {
                    text: line.to_string(),
                    ..Run::default()
                }],
            })
            .collect();
        Self {
            name: name.into(),
            description: None,
            hidden: false,
            is_text_box: false,
            anchor,
            geometry,
            properties: Properties::default(),
            paragraphs,
        }
    }

    /// A shape drawn with a custom `DrawingML` geometry (`a:custGeom`) instead
    /// of a preset.
    ///
    /// `text` is split into paragraphs on `\n` as in [`ShapeSpec::shape`].
    pub fn custom(
        name: impl Into<String>,
        anchor: Anchor,
        geometry: CustomGeometry,
        text: &str,
    ) -> Self {
        Self::with_geometry(name, anchor, geometry.into(), text)
    }

    /// Validate the spec against worksheet bounds and the module limits.
    fn validate_inner(&self, existing: usize) -> Result<(), String> {
        validate_object_slot(existing)?;
        validate_identity(&self.name, &self.description)?;
        if self.paragraphs.len() > MAX_SHAPE_PARAGRAPHS {
            return Err("shape paragraph count limit exceeded".to_string());
        }
        let mut text_bytes = 0usize;
        for paragraph in &self.paragraphs {
            if paragraph.runs.len() > MAX_RUNS_PER_PARAGRAPH {
                return Err("shape run count limit exceeded".to_string());
            }
            for run in &paragraph.runs {
                validate_xml_text(&run.text, "shape run text")?;
                text_bytes = text_bytes
                    .checked_add(run.text.len())
                    .ok_or_else(|| "shape text byte count overflow".to_string())?;
            }
        }
        if text_bytes > MAX_SHAPE_TEXT_BYTES {
            return Err("shape text bytes limit exceeded".to_string());
        }
        if let Geometry::Custom(geometry) = &self.geometry {
            validate_custom_geometry(geometry)?;
        }
        validate_anchor(&self.anchor)
    }
}

/// One end of an authored connection shape, referencing a target shape by name.
///
/// The name resolves to the target's drawing object ID at serialization
/// time; the target may be any authored shape, group child, group, or other
/// connector in the same worksheet drawing part.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ConnectionEndSpec {
    /// Name of the referenced shape (`xdr:cNvPr@name`).
    pub shape_name: String,
    /// Connection site index on the referenced shape (`@idx`).
    pub site: u32,
}

/// One authored connection shape (`xdr:cxnSp`) linking two shapes.
#[derive(Debug, Clone, PartialEq)]
pub struct ConnectionShapeSpec {
    /// Shape name written to `xdr:cNvPr@name`.
    pub name: String,
    /// Alternative text (`xdr:cNvPr@descr`), when set.
    pub description: Option<String>,
    /// Whether the shape is hidden (`xdr:cNvPr@hidden`).
    pub hidden: bool,
    /// How the connector is anchored on the worksheet grid.
    pub anchor: Anchor,
    /// Connector geometry, for example [`Preset::StraightConnector1`],
    /// [`Preset::BentConnector2`], or [`Preset::CurvedConnector3`].
    pub geometry: Geometry,
    /// Start connection site.
    pub start: ConnectionEndSpec,
    /// End connection site.
    pub end: ConnectionEndSpec,
}

impl ConnectionShapeSpec {
    /// A connector between two named shapes.
    pub fn new(
        name: impl Into<String>,
        anchor: Anchor,
        geometry: impl Into<Geometry>,
        start: ConnectionEndSpec,
        end: ConnectionEndSpec,
    ) -> Self {
        Self {
            name: name.into(),
            description: None,
            hidden: false,
            anchor,
            geometry: geometry.into(),
            start,
            end,
        }
    }

    /// Validate the spec against worksheet bounds and the module limits.
    ///
    /// Referenced shapes are resolved by name at serialization time, so
    /// existence is checked when the drawing XML is generated.
    fn validate_inner(&self, existing: usize) -> Result<(), String> {
        validate_object_slot(existing)?;
        validate_identity(&self.name, &self.description)?;
        if self.start.shape_name.is_empty() || self.end.shape_name.is_empty() {
            return Err("connection shape references cannot be empty".to_string());
        }
        validate_xml_text(&self.start.shape_name, "connection start shape name")?;
        validate_xml_text(&self.end.shape_name, "connection end shape name")?;
        if self.start.shape_name.len() > MAX_SHAPE_NAME_BYTES
            || self.end.shape_name.len() > MAX_SHAPE_NAME_BYTES
        {
            return Err("connection shape reference is too long".to_string());
        }
        if let Geometry::Custom(geometry) = &self.geometry {
            validate_custom_geometry(geometry)?;
        }
        validate_anchor(&self.anchor)
    }
}

/// One authored shape group (`xdr:grpSp`) with nested objects.
#[derive(Debug, Clone, PartialEq)]
pub struct GroupSpec {
    /// Group name written to `xdr:cNvPr@name`.
    pub name: String,
    /// Alternative text (`xdr:cNvPr@descr`), when set.
    pub description: Option<String>,
    /// Whether the group is hidden (`xdr:cNvPr@hidden`).
    pub hidden: bool,
    /// How the group is anchored on the worksheet grid.
    pub anchor: Anchor,
    /// Group coordinate transform (`a:off`/`a:ext`/`a:chOff`/`a:chExt`), when set.
    pub transform: Option<GroupTransform>,
    /// Nested objects in document order; groups may nest. Anchor fields of
    /// children are ignored — children position themselves through the
    /// group's coordinate transform.
    pub children: Vec<ObjectSpec>,
}

impl GroupSpec {
    /// An empty group with the given name and anchor.
    pub fn new(name: impl Into<String>, anchor: Anchor) -> Self {
        Self {
            name: name.into(),
            description: None,
            hidden: false,
            anchor,
            transform: None,
            children: Vec::new(),
        }
    }

    /// Append a child object to the group.
    #[must_use]
    pub fn with_child(mut self, child: ObjectSpec) -> Self {
        self.children.push(child);
        self
    }

    /// Validate the spec against worksheet bounds and the module limits.
    fn validate_inner(&self, existing: usize) -> Result<(), String> {
        validate_object_slot(existing)?;
        validate_identity(&self.name, &self.description)?;
        validate_anchor(&self.anchor)?;
        validate_group_transform(&self.transform)?;
        let mut object_count = 1usize;
        validate_children(&self.children, 1, &mut object_count)
    }
}

/// One authored drawing object: a shape, a group, or a connection shape.
#[derive(Debug, Clone, PartialEq)]
pub enum ObjectSpec {
    /// A text-box or plain shape.
    Shape(ShapeSpec),
    /// A shape group.
    Group(GroupSpec),
    /// A connection shape.
    Connection(ConnectionShapeSpec),
}

impl From<ShapeSpec> for ObjectSpec {
    fn from(spec: ShapeSpec) -> Self {
        Self::Shape(spec)
    }
}

impl From<GroupSpec> for ObjectSpec {
    fn from(spec: GroupSpec) -> Self {
        Self::Group(spec)
    }
}

impl From<ConnectionShapeSpec> for ObjectSpec {
    fn from(spec: ConnectionShapeSpec) -> Self {
        Self::Connection(spec)
    }
}

impl ShapeSpec {
    /// Validate this shape against worksheet bounds and authoring limits.
    pub fn validate(&self, existing: usize) -> crate::Result<()> {
        self.validate_inner(existing).map_err(Error::Invalid)
    }
}

impl GroupSpec {
    /// Validate this group and its descendants against authoring limits.
    pub fn validate(&self, existing: usize) -> crate::Result<()> {
        self.validate_inner(existing).map_err(Error::Invalid)
    }
}

impl ConnectionShapeSpec {
    /// Validate this connector against worksheet bounds and authoring limits.
    pub fn validate(&self, existing: usize) -> crate::Result<()> {
        self.validate_inner(existing).map_err(Error::Invalid)
    }
}

fn validate_object_slot(existing: usize) -> Result<(), String> {
    if existing >= MAX_SHAPES_PER_WORKSHEET {
        return Err("worksheet drawing object count limit exceeded".to_string());
    }
    Ok(())
}

fn validate_identity(name: &str, description: &Option<String>) -> Result<(), String> {
    if name.is_empty() {
        return Err("shape name cannot be empty".to_string());
    }
    if name.len() > MAX_SHAPE_NAME_BYTES
        || description
            .as_ref()
            .is_some_and(|d| d.len() > MAX_SHAPE_NAME_BYTES)
    {
        return Err("shape name/description is too long".to_string());
    }
    validate_xml_text(name, "shape name")?;
    if let Some(description) = description {
        validate_xml_text(description, "shape description")?;
    }
    Ok(())
}

fn validate_xml_text(value: &str, field: &str) -> Result<(), String> {
    if value.chars().any(|character| {
        !matches!(
            character as u32,
            0x9 | 0xA | 0xD | 0x20..=0xD7FF | 0xE000..=0xFFFD | 0x10000..=0x10FFFF
        )
    }) {
        return Err(format!("{field} contains an invalid XML character"));
    }
    Ok(())
}

/// Validate group children recursively. Children carry anchor fields that
/// the serializer ignores inside a group; they are validated anyway for
/// consistency with top-level objects.
fn validate_children(
    children: &[ObjectSpec],
    depth: usize,
    object_count: &mut usize,
) -> Result<(), String> {
    if depth > MAX_GROUP_DEPTH {
        return Err("shape group depth limit exceeded".to_string());
    }
    if children.len() > MAX_GROUP_CHILDREN {
        return Err("shape group children count limit exceeded".to_string());
    }
    for child in children {
        *object_count = object_count
            .checked_add(1)
            .ok_or_else(|| "shape group object count overflow".to_string())?;
        if *object_count > MAX_OBJECTS_PER_DRAWING {
            return Err("shape group object count limit exceeded".to_string());
        }
        match child {
            ObjectSpec::Shape(shape) => shape.validate_inner(0)?,
            ObjectSpec::Group(group) => {
                validate_identity(&group.name, &group.description)?;
                validate_anchor(&group.anchor)?;
                validate_group_transform(&group.transform)?;
                validate_children(&group.children, depth + 1, object_count)?;
            },
            ObjectSpec::Connection(connection) => connection.validate_inner(0)?,
        }
    }
    Ok(())
}

/// Validate anchor markers against worksheet bounds (and ordering for
/// two-cell anchors), mirroring the checks applied to images and charts.
fn validate_anchor(anchor: &Anchor) -> Result<(), String> {
    const MAX_COLUMNS: u32 = 16_384;
    const MAX_ROWS: u32 = 1_048_576;
    let markers: &[CellMarker] = match anchor {
        Anchor::TwoCell { from, to, .. } => {
            if to.row < from.row || to.column < from.column {
                return Err("shape anchor cannot be descending".to_string());
            }
            &[*from, *to]
        },
        Anchor::OneCell { from, .. } => &[*from],
        Anchor::Absolute { .. } => &[],
    };
    for marker in markers {
        if marker.column >= MAX_COLUMNS || marker.row >= MAX_ROWS {
            return Err("shape anchor exceeds worksheet bounds".to_string());
        }
        if marker.column_offset.emu() < 0 || marker.row_offset.emu() < 0 {
            return Err("shape anchor offsets cannot be negative".to_string());
        }
    }
    match anchor {
        Anchor::OneCell { extent, .. } | Anchor::Absolute { extent, .. } => {
            validate_extent(extent)?;
        },
        Anchor::TwoCell { .. } => {},
    }
    Ok(())
}

fn validate_group_transform(transform: &Option<GroupTransform>) -> Result<(), String> {
    if let Some(transform) = transform {
        if let Some(extent) = &transform.extent {
            validate_extent(extent)?;
        }
        if let Some(child_extent) = &transform.child_extent {
            validate_extent(child_extent)?;
        }
    }
    Ok(())
}

fn validate_extent(extent: &EmuExtent) -> Result<(), String> {
    if extent.width.emu() < 0 || extent.height.emu() < 0 {
        return Err("shape extent cannot be negative".to_string());
    }
    Ok(())
}

/// Drawing-part serializer state: allocates drawing-wide unique object IDs
/// and tracks name-to-ID bindings so connection sites can resolve their
/// `a:stCxn`/`a:endCxn` targets.
///
/// Names must be unique across all authored objects in one drawing part
/// (including group children); duplicate names and unresolvable connection
/// references are serialization errors.
struct State {
    next_id: u32,
    names: HashMap<String, u32>,
    preallocated: bool,
    emitted_names: HashSet<String>,
}

impl State {
    /// Preallocate IDs for a complete worksheet shape graph.
    ///
    /// This makes connection targets independent of XML object order while
    /// preserving deterministic IDs and rejecting duplicate names before any
    /// output is emitted.
    pub(crate) fn for_objects(
        first_id: u32,
        shapes: &[ShapeSpec],
        groups: &[GroupSpec],
        connections: &[ConnectionShapeSpec],
    ) -> Result<Self, String> {
        let mut emitter = Self {
            next_id: first_id,
            names: HashMap::new(),
            preallocated: true,
            emitted_names: HashSet::new(),
        };
        for shape in shapes {
            emitter.reserve(&shape.name)?;
        }
        for group in groups {
            emitter.reserve_group(group)?;
        }
        for connection in connections {
            emitter.reserve(&connection.name)?;
        }
        Ok(emitter)
    }

    fn reserve_group(&mut self, group: &GroupSpec) -> Result<(), String> {
        self.reserve(&group.name)?;
        for child in &group.children {
            match child {
                ObjectSpec::Shape(shape) => self.reserve(&shape.name)?,
                ObjectSpec::Group(group) => self.reserve_group(group)?,
                ObjectSpec::Connection(connection) => self.reserve(&connection.name)?,
            }
        }
        Ok(())
    }

    fn reserve(&mut self, name: &str) -> Result<(), String> {
        if self.names.len() >= MAX_OBJECTS_PER_DRAWING {
            return Err("worksheet drawing object count limit exceeded".to_string());
        }
        let id = self.next_id;
        self.next_id = self
            .next_id
            .checked_add(1)
            .ok_or_else(|| "drawing object ID space exhausted".to_string())?;
        if self.names.insert(name.to_string(), id).is_some() {
            return Err(format!("duplicate shape name '{name}' in drawing"));
        }
        Ok(())
    }

    /// Serialize one top-level authored shape wrapped in its anchor.
    pub(crate) fn write_anchored_shape(
        &mut self,
        xml: &mut String,
        shape: &ShapeSpec,
    ) -> Result<(), String> {
        self.write_anchored(xml, &shape.anchor, |emitter, xml| {
            let id = emitter.allocate(&shape.name)?;
            write_shape_xml(xml, shape, id);
            Ok(())
        })
    }

    /// Serialize one top-level authored group wrapped in its anchor.
    pub(crate) fn write_anchored_group(
        &mut self,
        xml: &mut String,
        group: &GroupSpec,
    ) -> Result<(), String> {
        self.write_anchored(xml, &group.anchor, |emitter, xml| {
            emitter.write_group_xml(xml, group)
        })
    }

    /// Serialize one top-level authored connection shape wrapped in its anchor.
    pub(crate) fn write_anchored_connection(
        &mut self,
        xml: &mut String,
        connection: &ConnectionShapeSpec,
    ) -> Result<(), String> {
        self.write_anchored(xml, &connection.anchor, |emitter, xml| {
            emitter.write_connection_xml(xml, connection)
        })
    }

    /// Wrap an object body in its anchor element.
    fn write_anchored(
        &mut self,
        xml: &mut String,
        anchor: &Anchor,
        inner: impl FnOnce(&mut Self, &mut String) -> Result<(), String>,
    ) -> Result<(), String> {
        write_anchor_open(xml, anchor);
        inner(self, xml)?;
        xml.push_str("<xdr:clientData/>");
        match anchor {
            Anchor::TwoCell { .. } => xml.push_str("</xdr:twoCellAnchor>"),
            Anchor::OneCell { .. } => xml.push_str("</xdr:oneCellAnchor>"),
            Anchor::Absolute { .. } => xml.push_str("</xdr:absoluteAnchor>"),
        }
        Ok(())
    }

    /// Serialize one authored object without an anchor (top-level or nested
    /// in a group).
    fn write_object(&mut self, xml: &mut String, object: &ObjectSpec) -> Result<(), String> {
        match object {
            ObjectSpec::Shape(shape) => {
                let id = self.allocate(&shape.name)?;
                write_shape_xml(xml, shape, id);
                Ok(())
            },
            ObjectSpec::Group(group) => self.write_group_xml(xml, group),
            ObjectSpec::Connection(connection) => self.write_connection_xml(xml, connection),
        }
    }

    fn write_group_xml(&mut self, xml: &mut String, group: &GroupSpec) -> Result<(), String> {
        let id = self.allocate(&group.name)?;
        xml.push_str("<xdr:grpSp><xdr:nvGrpSpPr>");
        write_c_nv_pr(xml, id, &group.name, &group.description, group.hidden);
        xml.push_str("<xdr:cNvGrpSpPr/></xdr:nvGrpSpPr>");
        match &group.transform {
            None => xml.push_str("<xdr:grpSpPr/>"),
            Some(transform) => {
                xml.push_str("<xdr:grpSpPr><a:xfrm>");
                if let Some(offset) = &transform.offset {
                    write_xml!(
                        xml,
                        r#"<a:off x="{}" y="{}"/>"#,
                        offset.x.emu(),
                        offset.y.emu()
                    );
                }
                if let Some(extent) = &transform.extent {
                    write_xml!(
                        xml,
                        r#"<a:ext cx="{}" cy="{}"/>"#,
                        extent.width.emu(),
                        extent.height.emu()
                    );
                }
                if let Some(child_offset) = &transform.child_offset {
                    write_xml!(
                        xml,
                        r#"<a:chOff x="{}" y="{}"/>"#,
                        child_offset.x.emu(),
                        child_offset.y.emu()
                    );
                }
                if let Some(child_extent) = &transform.child_extent {
                    write_xml!(
                        xml,
                        r#"<a:chExt cx="{}" cy="{}"/>"#,
                        child_extent.width.emu(),
                        child_extent.height.emu()
                    );
                }
                xml.push_str("</a:xfrm></xdr:grpSpPr>");
            },
        }
        for child in &group.children {
            self.write_object(xml, child)?;
        }
        xml.push_str("</xdr:grpSp>");
        Ok(())
    }

    fn write_connection_xml(
        &mut self,
        xml: &mut String,
        connection: &ConnectionShapeSpec,
    ) -> Result<(), String> {
        let id = self.allocate(&connection.name)?;
        let start_id = self.resolve(&connection.start.shape_name)?;
        let end_id = self.resolve(&connection.end.shape_name)?;
        xml.push_str("<xdr:cxnSp><xdr:nvCxnSpPr>");
        write_c_nv_pr(
            xml,
            id,
            &connection.name,
            &connection.description,
            connection.hidden,
        );
        write_xml!(
            xml,
            "<xdr:cNvCxnSpPr><a:stCxn id=\"{start_id}\" idx=\"{}\"/>\
             <a:endCxn id=\"{end_id}\" idx=\"{}\"/></xdr:cNvCxnSpPr>",
            connection.start.site,
            connection.end.site
        );
        xml.push_str("</xdr:nvCxnSpPr><xdr:spPr>");
        xml.push_str(r#"<a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm>"#);
        match &connection.geometry {
            Geometry::Preset(preset) => {
                write_xml!(
                    xml,
                    r#"<a:prstGeom prst="{}"><a:avLst/></a:prstGeom>"#,
                    escape_xml(preset.token())
                );
            },
            Geometry::Custom(geometry) => write_custom_geometry(xml, geometry),
        }
        xml.push_str("</xdr:spPr></xdr:cxnSp>");
        Ok(())
    }

    fn allocate(&mut self, name: &str) -> Result<u32, String> {
        if self.preallocated {
            let id = self
                .names
                .get(name)
                .copied()
                .ok_or_else(|| format!("shape name '{name}' was not preallocated"))?;
            if !self.emitted_names.insert(name.to_string()) {
                return Err(format!("duplicate shape name '{name}' in drawing"));
            }
            return Ok(id);
        }
        let id = self.next_id;
        if self.names.len() >= MAX_OBJECTS_PER_DRAWING {
            return Err("worksheet drawing object count limit exceeded".to_string());
        }
        self.next_id = self
            .next_id
            .checked_add(1)
            .ok_or_else(|| "drawing object ID space exhausted".to_string())?;
        if self.names.insert(name.to_string(), id).is_some() {
            return Err(format!("duplicate shape name '{name}' in drawing"));
        }
        Ok(id)
    }

    fn resolve(&self, name: &str) -> Result<u32, String> {
        self.names
            .get(name)
            .copied()
            .ok_or_else(|| format!("connection shape references unknown shape '{name}'"))
    }
}

/// Write the shared non-visual identity element (`xdr:cNvPr`).
fn write_c_nv_pr(
    xml: &mut String,
    id: u32,
    name: &str,
    description: &Option<String>,
    hidden: bool,
) {
    write_xml!(xml, r#"<xdr:cNvPr id="{id}" name="{}""#, escape_xml(name));
    if let Some(description) = description {
        write_xml!(xml, r#" descr="{}""#, escape_xml(description));
    }
    if hidden {
        xml.push_str(r#" hidden="1""#);
    }
    xml.push_str("/>");
}

fn write_anchor_open(xml: &mut String, anchor: &Anchor) {
    match anchor {
        Anchor::TwoCell { from, to, edit_as } => {
            match edit_as {
                // The ECMA-376 default; omitted to keep output canonical.
                EditAs::TwoCell => xml.push_str("<xdr:twoCellAnchor>"),
                EditAs::OneCell => xml.push_str(r#"<xdr:twoCellAnchor editAs="oneCell">"#),
                EditAs::Absolute => xml.push_str(r#"<xdr:twoCellAnchor editAs="absolute">"#),
            }
            write_marker(xml, "from", from);
            write_marker(xml, "to", to);
        },
        Anchor::OneCell { from, extent } => {
            xml.push_str("<xdr:oneCellAnchor>");
            write_marker(xml, "from", from);
            write_extent(xml, extent);
        },
        Anchor::Absolute { position, extent } => {
            xml.push_str("<xdr:absoluteAnchor>");
            write_position(xml, position);
            write_extent(xml, extent);
        },
    }
}

fn write_marker(xml: &mut String, name: &str, marker: &CellMarker) {
    write_xml!(
        xml,
        "<xdr:{name}><xdr:col>{}</xdr:col><xdr:colOff>{}</xdr:colOff>\
         <xdr:row>{}</xdr:row><xdr:rowOff>{}</xdr:rowOff></xdr:{name}>",
        marker.column,
        marker.column_offset.emu(),
        marker.row,
        marker.row_offset.emu()
    );
}

fn write_extent(xml: &mut String, extent: &EmuExtent) {
    write_xml!(
        xml,
        r#"<xdr:ext cx="{}" cy="{}"/>"#,
        extent.width.emu(),
        extent.height.emu()
    );
}

fn write_position(xml: &mut String, position: &EmuOffset) {
    write_xml!(
        xml,
        r#"<xdr:pos x="{}" y="{}"/>"#,
        position.x.emu(),
        position.y.emu()
    );
}

fn write_shape_xml(xml: &mut String, spec: &ShapeSpec, id: u32) {
    xml.push_str(r#"<xdr:sp macro="" textlink=""><xdr:nvSpPr>"#);
    write_c_nv_pr(xml, id, &spec.name, &spec.description, spec.hidden);
    if spec.is_text_box {
        xml.push_str(r#"<xdr:cNvSpPr txBox="1"/>"#);
    } else {
        xml.push_str("<xdr:cNvSpPr/>");
    }
    xml.push_str("</xdr:nvSpPr><xdr:spPr>");
    xml.push_str(r#"<a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm>"#);
    match &spec.geometry {
        Geometry::Custom(geometry) => write_custom_geometry(xml, geometry),
        Geometry::Preset(preset) => {
            write_xml!(
                xml,
                r#"<a:prstGeom prst="{}"><a:avLst/></a:prstGeom>"#,
                escape_xml(preset.token())
            );
        },
    }
    xml.push_str("</xdr:spPr><xdr:txBody>");
    write_text_body_contents(xml, &spec.properties, &spec.paragraphs);
    xml.push_str("</xdr:txBody></xdr:sp>");
}

/// Fresh `SpreadsheetDrawing` XML authoring state.
///
/// The emitter allocates drawing-wide IDs and resolves connector endpoints.
/// It accepts only fresh `*Spec` values, deliberately excluding opaque reader
/// payloads from serialization.
pub struct Emitter {
    inner: State,
}

impl Emitter {
    /// Create an emitter whose first allocated object ID is `first_id`.
    #[must_use]
    pub fn new(first_id: u32) -> Self {
        Self {
            inner: State {
                next_id: first_id,
                names: HashMap::new(),
                preallocated: false,
                emitted_names: HashSet::new(),
            },
        }
    }

    /// Preallocate IDs for a complete fresh drawing graph.
    pub fn for_objects(
        first_id: u32,
        shapes: &[ShapeSpec],
        groups: &[GroupSpec],
        connections: &[ConnectionShapeSpec],
    ) -> crate::Result<Self> {
        State::for_objects(first_id, shapes, groups, connections)
            .map(|inner| Self { inner })
            .map_err(Error::Invalid)
    }

    /// Write a fresh top-level shape enclosed in its anchor.
    pub fn write_anchored_shape(
        &mut self,
        xml: &mut String,
        shape: &ShapeSpec,
    ) -> crate::Result<()> {
        shape.validate(self.inner.names.len())?;
        self.inner
            .write_anchored_shape(xml, shape)
            .map_err(Error::Invalid)
    }

    /// Write a fresh top-level group enclosed in its anchor.
    pub fn write_anchored_group(
        &mut self,
        xml: &mut String,
        group: &GroupSpec,
    ) -> crate::Result<()> {
        group.validate(self.inner.names.len())?;
        self.inner
            .write_anchored_group(xml, group)
            .map_err(Error::Invalid)
    }

    /// Write a fresh top-level connection shape enclosed in its anchor.
    pub fn write_anchored_connection(
        &mut self,
        xml: &mut String,
        connection: &ConnectionShapeSpec,
    ) -> crate::Result<()> {
        connection.validate(self.inner.names.len())?;
        self.inner
            .write_anchored_connection(xml, connection)
            .map_err(Error::Invalid)
    }
}

#[cfg(test)]
mod tests;
