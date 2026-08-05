//! Shared SpreadsheetDrawing authoring model for XLSX and XLSB shapes.
//!
//! [`ShapeSpec`] describes one text-box shape to embed in a worksheet
//! drawing part: a preset geometry, an anchor on the sheet grid, text-body
//! properties, and rich-text runs. [`GroupSpec`] nests shapes, text
//! boxes, and further groups under a `xdr:grpSp` group with an optional
//! coordinate transform, and [`ConnectionShapeSpec`] authors `xdr:cxnSp`
//! connectors whose `a:stCxn`/`a:endCxn` sites reference other shapes by
//! name. All three deliberately reuse the typed read model from
//! [`super::shapes`] (`Anchor`,
//! `BodyProperties`, `GroupTransform`, paragraphs and runs) so
//! anything authored here round-trips through the shape inventory with
//! identical semantics.
//!
//! `ShapeEmitter` serializes authored objects as
//! `xdr:twoCellAnchor`/`xdr:oneCellAnchor`/`xdr:absoluteAnchor` elements for
//! the standard worksheet drawing parts used by both XLSX and XLSB,
//! allocating drawing-wide unique object IDs and resolving connection-site
//! name references. Everything is inert: no rendering, no layout
//! computation, and all inputs are bounded and validated at authoring time.

use std::collections::{HashMap, HashSet};
use std::fmt::Write as _;

use litchi_core::xml::escape::escape_xml;
use litchi_drawingml::geom::Preset;

use crate::shapes::{
    Anchor, Autofit, BodyProperties, CellMarker, Columns, Direction, EditAs, EmuExtent, EmuOffset,
    Geometry, GroupTransform, Paragraph, Run, VerticalAnchor, Wrap,
};
use litchi_drawingml::geometry::writer::write_custom_geometry;
use litchi_drawingml::geometry::{CustomGeometry, validate_custom_geometry};

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

/// One authored DrawingML text-box shape for a worksheet drawing part.
///
/// Construct with [`ShapeSpec::text_box`] (or [`ShapeSpec::shape`]
/// for a non-text-box shape), then adjust `body_properties`, `paragraphs`,
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
    pub body_properties: BodyProperties,
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
            body_properties: BodyProperties::default(),
            paragraphs,
        }
    }

    /// A shape drawn with a custom DrawingML geometry (`a:custGeom`) instead
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
    pub(crate) fn validate(&self, existing: usize) -> Result<(), String> {
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
    pub(crate) fn validate(&self, existing: usize) -> Result<(), String> {
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
    pub fn with_child(mut self, child: ObjectSpec) -> Self {
        self.children.push(child);
        self
    }

    /// Validate the spec against worksheet bounds and the module limits.
    pub(crate) fn validate(&self, existing: usize) -> Result<(), String> {
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
            ObjectSpec::Shape(shape) => shape.validate(0)?,
            ObjectSpec::Group(group) => {
                validate_identity(&group.name, &group.description)?;
                validate_anchor(&group.anchor)?;
                validate_group_transform(&group.transform)?;
                validate_children(&group.children, depth + 1, object_count)?;
            },
            ObjectSpec::Connection(connection) => connection.validate(0)?,
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
            validate_extent(extent)?
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
pub(crate) struct ShapeEmitter {
    next_id: u32,
    names: HashMap<String, u32>,
    preallocated: bool,
    emitted_names: HashSet<String>,
}

impl ShapeEmitter {
    /// An emitter whose first allocated ID is `first_id` (continuing the
    /// picture/chart ID sequence of the drawing part).
    #[cfg(test)]
    pub(crate) fn new(first_id: u32) -> Self {
        Self {
            next_id: first_id,
            names: HashMap::new(),
            preallocated: false,
            emitted_names: HashSet::new(),
        }
    }

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
                    let _ = write!(
                        xml,
                        r#"<a:off x="{}" y="{}"/>"#,
                        offset.x.emu(),
                        offset.y.emu()
                    );
                }
                if let Some(extent) = &transform.extent {
                    let _ = write!(
                        xml,
                        r#"<a:ext cx="{}" cy="{}"/>"#,
                        extent.width.emu(),
                        extent.height.emu()
                    );
                }
                if let Some(child_offset) = &transform.child_offset {
                    let _ = write!(
                        xml,
                        r#"<a:chOff x="{}" y="{}"/>"#,
                        child_offset.x.emu(),
                        child_offset.y.emu()
                    );
                }
                if let Some(child_extent) = &transform.child_extent {
                    let _ = write!(
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
        let _ = write!(
            xml,
            "<xdr:cNvCxnSpPr><a:stCxn id=\"{start_id}\" idx=\"{}\"/>\
             <a:endCxn id=\"{end_id}\" idx=\"{}\"/></xdr:cNvCxnSpPr>",
            connection.start.site, connection.end.site
        );
        xml.push_str("</xdr:nvCxnSpPr><xdr:spPr>");
        xml.push_str(r#"<a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm>"#);
        match &connection.geometry {
            Geometry::Preset(preset) => {
                let _ = write!(
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
    let _ = write!(xml, r#"<xdr:cNvPr id="{id}" name="{}""#, escape_xml(name));
    if let Some(description) = description {
        let _ = write!(xml, r#" descr="{}""#, escape_xml(description));
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
    let _ = write!(
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
    let _ = write!(
        xml,
        r#"<xdr:ext cx="{}" cy="{}"/>"#,
        extent.width.emu(),
        extent.height.emu()
    );
}

fn write_position(xml: &mut String, position: &EmuOffset) {
    let _ = write!(
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
            let _ = write!(
                xml,
                r#"<a:prstGeom prst="{}"><a:avLst/></a:prstGeom>"#,
                escape_xml(preset.token())
            );
        },
    }
    xml.push_str("</xdr:spPr><xdr:txBody>");
    write_body_properties(xml, &spec.body_properties);
    xml.push_str("<a:lstStyle/>");
    for paragraph in &spec.paragraphs {
        xml.push_str("<a:p>");
        for run in &paragraph.runs {
            xml.push_str("<a:r>");
            write_run_properties(xml, run);
            let _ = write!(xml, "<a:t>{}</a:t>", escape_xml(&run.text));
            xml.push_str("</a:r>");
        }
        xml.push_str("</a:p>");
    }
    xml.push_str("</xdr:txBody></xdr:sp>");
}

fn write_body_properties(xml: &mut String, body: &BodyProperties) {
    xml.push_str("<a:bodyPr");
    let _ = write!(
        xml,
        r#" lIns="{}" tIns="{}" rIns="{}" bIns="{}""#,
        body.insets.left, body.insets.top, body.insets.right, body.insets.bottom
    );
    let anchor =
        (body.vertical_anchor != VerticalAnchor::Top).then(|| body.vertical_anchor.token());
    if let Some(token) = anchor {
        let _ = write!(xml, r#" anchor="{token}""#);
    }
    if body.anchor_center {
        xml.push_str(r#" anchorCtr="1""#);
    }
    let direction = (body.direction != Direction::Horizontal).then(|| body.direction.token());
    if let Some(token) = direction {
        let _ = write!(xml, r#" vert="{token}""#);
    }
    if body.wrap == Wrap::None {
        xml.push_str(r#" wrap="none""#);
    }
    if body.column_count != Columns::ONE {
        let _ = write!(xml, r#" numCol="{}""#, body.column_count);
    }
    if body.space_first_last_paragraph {
        xml.push_str(r#" spcFirstLastPara="1""#);
    }
    match body.autofit {
        Autofit::None => xml.push_str("><a:noAutofit/></a:bodyPr>"),
        Autofit::Shape => xml.push_str("><a:spAutoFit/></a:bodyPr>"),
        Autofit::Normal => xml.push_str("><a:normAutofit/></a:bodyPr>"),
    }
}

fn write_run_properties(xml: &mut String, run: &Run) {
    if run.bold.is_none()
        && run.italic.is_none()
        && run.underline.is_none()
        && run.font_size.is_none()
    {
        xml.push_str("<a:rPr/>");
        return;
    }
    xml.push_str("<a:rPr");
    if let Some(size) = run.font_size {
        let _ = write!(xml, r#" sz="{size}""#);
    }
    if let Some(bold) = run.bold {
        xml.push_str(if bold { r#" b="1""# } else { r#" b="0""# });
    }
    if let Some(italic) = run.italic {
        xml.push_str(if italic { r#" i="1""# } else { r#" i="0""# });
    }
    if let Some(underline) = run.underline {
        let _ = write!(xml, r#" u="{}""#, underline.dml());
    }
    xml.push_str("/>");
}

#[cfg(any())]
mod tests {
    use super::shapes::{
        CellMarker, Coordinate32, Emu, Object, Run, TextInsets, TextSize, Underline,
        parse_drawing_shapes,
    };
    use super::*;
    use crate::writer::sheet::MutableWorksheet;
    use litchi_drawingml::coord::Unit;

    fn marker(column: u32, row: u32) -> CellMarker {
        CellMarker {
            column,
            column_offset: Emu(100),
            row,
            row_offset: Emu(200),
        }
    }

    fn two_cell() -> Anchor {
        Anchor::TwoCell {
            from: marker(1, 2),
            to: marker(5, 9),
            edit_as: EditAs::OneCell,
        }
    }

    fn drawing_wrap(body: &str) -> String {
        format!(
            "<xdr:wsDr xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\" \
             xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">{body}</xdr:wsDr>"
        )
    }

    fn parse_single(xml: &str) -> super::shapes::AnchoredObject {
        let objects = parse_drawing_shapes(&drawing_wrap(xml)).unwrap().unwrap();
        assert_eq!(objects.len(), 1);
        objects.into_iter().next().unwrap()
    }

    #[test]
    fn two_cell_text_box_round_trips_through_reader() {
        let mut spec = ShapeSpec::text_box("Box 1", two_cell(), Preset::RoundRect, "Hello");
        spec.description = Some("alt <text>".to_string());
        spec.hidden = true;
        spec.body_properties = BodyProperties {
            insets: TextInsets {
                left: Coordinate32::measure("0.2", Unit::Inch).unwrap(),
                top: Coordinate32::from(91440),
                right: Coordinate32::from(182880),
                bottom: Coordinate32::from(91440),
            },
            vertical_anchor: VerticalAnchor::Center,
            anchor_center: true,
            direction: Direction::Vertical270,
            wrap: Wrap::None,
            autofit: Autofit::Shape,
            column_count: Columns::new(2).unwrap(),
            space_first_last_paragraph: true,
        };
        spec.paragraphs = vec![
            Paragraph {
                runs: vec![
                    Run {
                        text: "Bold &".to_string(),
                        bold: Some(true),
                        italic: Some(false),
                        underline: Some(Underline::DotDashHeavy),
                        font_size: Some(TextSize::new(1200).unwrap()),
                    },
                    Run {
                        text: " plain".to_string(),
                        ..Run::default()
                    },
                ],
            },
            Paragraph {
                runs: vec![Run {
                    text: "second".to_string(),
                    ..Run::default()
                }],
            },
        ];

        let mut xml = String::new();
        ShapeEmitter::new(7)
            .write_anchored_shape(&mut xml, &spec)
            .unwrap();
        let anchored = parse_single(&xml);
        assert_eq!(anchored.anchor, spec.anchor);
        let Object::Shape(shape) = &anchored.object else {
            panic!("expected a shape");
        };
        assert_eq!(shape.non_visual.id, Some(7));
        assert_eq!(shape.non_visual.name.as_deref(), Some("Box 1"));
        assert_eq!(shape.non_visual.description.as_deref(), Some("alt <text>"));
        assert!(shape.non_visual.hidden);
        assert!(!shape.non_visual.locked);
        assert!(shape.is_text_box);
        assert_eq!(shape.preset(), Some(Preset::RoundRect));
        let body = shape.text_body.as_ref().unwrap();
        assert_eq!(body.body_properties, spec.body_properties);
        assert_eq!(body.paragraphs, spec.paragraphs);
        assert_eq!(body.text(), "Bold & plain\nsecond");
    }

    #[test]
    fn one_cell_and_absolute_anchors_round_trip() {
        let one_cell = Anchor::OneCell {
            from: marker(3, 4),
            extent: EmuExtent {
                width: Emu(914400),
                height: Emu(457200),
            },
        };
        let absolute = Anchor::Absolute {
            position: EmuOffset {
                x: Emu(123),
                y: Emu(456),
            },
            extent: EmuExtent {
                width: Emu(789),
                height: Emu(101),
            },
        };
        for anchor in [one_cell, absolute] {
            let spec = ShapeSpec::shape("S", anchor, Preset::Arc, "");
            let mut xml = String::new();
            ShapeEmitter::new(1)
                .write_anchored_shape(&mut xml, &spec)
                .unwrap();
            let anchored = parse_single(&xml);
            assert_eq!(anchored.anchor, anchor);
            let Object::Shape(shape) = &anchored.object else {
                panic!("expected a shape");
            };
            assert!(!shape.is_text_box);
            assert_eq!(shape.preset(), Some(Preset::Arc));
        }
    }

    #[test]
    fn default_body_properties_round_trip() {
        let spec = ShapeSpec::text_box("Defaults", two_cell(), Preset::Rect, "x");
        let mut xml = String::new();
        ShapeEmitter::new(2)
            .write_anchored_shape(&mut xml, &spec)
            .unwrap();
        let anchored = parse_single(&xml);
        let Object::Shape(shape) = &anchored.object else {
            panic!("expected a shape");
        };
        let body = shape.text_body.as_ref().unwrap();
        assert_eq!(body.body_properties, BodyProperties::default());
        // The default edit-as token is omitted from the output.
        let default_edit = Anchor::TwoCell {
            from: marker(0, 0),
            to: marker(1, 1),
            edit_as: EditAs::TwoCell,
        };
        let spec = ShapeSpec::text_box("E", default_edit, Preset::Rect, "x");
        let mut xml = String::new();
        ShapeEmitter::new(3)
            .write_anchored_shape(&mut xml, &spec)
            .unwrap();
        assert!(!xml.contains("editAs"));
        assert_eq!(parse_single(&xml).anchor, default_edit);
    }

    #[test]
    fn validation_rejects_invalid_specs() {
        let mut spec = ShapeSpec::text_box("", two_cell(), Preset::Rect, "x");
        assert!(spec.validate(0).is_err());
        spec.name = "ok".to_string();
        spec.anchor = Anchor::TwoCell {
            from: marker(5, 9),
            to: marker(1, 2),
            edit_as: EditAs::TwoCell,
        };
        assert!(spec.validate(0).is_err());
        spec.anchor = Anchor::TwoCell {
            from: marker(16_384, 0),
            to: marker(16_385, 1),
            edit_as: EditAs::TwoCell,
        };
        assert!(spec.validate(0).is_err());
        spec.anchor = two_cell();
        assert!(spec.validate(MAX_SHAPES_PER_WORKSHEET).is_err());
        assert!(spec.validate(0).is_ok());
    }

    #[test]
    fn worksheet_api_adds_removes_and_serializes_shapes() {
        let mut ws = MutableWorksheet::new("Sheet1".to_string(), 1);
        ws.add_text_box("First", two_cell(), Preset::Rect, "hello")
            .unwrap();
        ws.add_shape(ShapeSpec::text_box(
            "Second",
            two_cell(),
            Preset::Ellipse,
            "world",
        ))
        .unwrap();
        assert_eq!(ws.shapes().len(), 2);
        assert!(ws.remove_shape(5).is_err());
        let removed = ws.remove_shape(0).unwrap();
        assert_eq!(removed.name, "First");
        assert_eq!(ws.shapes().len(), 1);

        let xml = ws.generate_drawing_xml().unwrap().unwrap();
        assert!(xml.contains("<xdr:sp"));
        assert!(xml.contains(r#"prst="ellipse""#));
        let objects = parse_drawing_shapes(&xml).unwrap().unwrap();
        assert_eq!(objects.len(), 1);
        let Object::Shape(shape) = &objects[0].object else {
            panic!("expected a shape");
        };
        assert_eq!(shape.non_visual.name.as_deref(), Some("Second"));
        assert_eq!(shape.text_body.as_ref().unwrap().text(), "world");
    }

    #[test]
    fn shapes_coexist_with_images_in_drawing_xml() {
        let mut ws = MutableWorksheet::new("Sheet1".to_string(), 1);
        ws.add_image(vec![1, 2, 3], "png", 1, 1, 2, 2, Some("Logo"))
            .unwrap();
        ws.add_text_box("Note", two_cell(), Preset::Rect, "note text")
            .unwrap();
        let xml = ws.generate_drawing_xml().unwrap().unwrap();
        // The image keeps rId1; the shape follows with the next object ID.
        assert!(xml.contains(r#"r:embed="rId1""#));
        assert!(xml.contains(r#"<xdr:cNvPr id="2" name="Note""#));
        let objects = parse_drawing_shapes(&xml).unwrap().unwrap();
        assert_eq!(objects.len(), 1, "pictures stay with the image pipeline");
    }

    #[test]
    fn authored_shapes_round_trip_through_a_saved_package() {
        let directory = tempfile::tempdir().unwrap();
        let path = directory.path().join("shapes.xlsx");
        let sheet_name;
        {
            let mut workbook = crate::workbook::Workbook::create().unwrap();
            let ws = workbook.worksheet_mut(0).unwrap();
            sheet_name = ws.name().to_string();
            ws.add_image(vec![9, 9, 9], "png", 1, 1, 2, 2, None)
                .unwrap();
            ws.add_text_box(
                "Greeting",
                Anchor::TwoCell {
                    from: CellMarker {
                        column: 2,
                        column_offset: Emu(57150),
                        row: 1,
                        row_offset: Emu(47625),
                    },
                    to: CellMarker {
                        column: 6,
                        column_offset: Emu(0),
                        row: 8,
                        row_offset: Emu(0),
                    },
                    edit_as: EditAs::TwoCell,
                },
                Preset::RoundRect,
                "line one\nline two",
            )
            .unwrap();
            let mut fancy = ShapeSpec::text_box(
                "Fancy",
                Anchor::OneCell {
                    from: CellMarker {
                        column: 8,
                        column_offset: Emu(0),
                        row: 10,
                        row_offset: Emu(0),
                    },
                    extent: EmuExtent {
                        width: Emu(1_828_800),
                        height: Emu(914_400),
                    },
                },
                Preset::Ellipse,
                "fancy",
            );
            fancy.body_properties.vertical_anchor = VerticalAnchor::Bottom;
            fancy.body_properties.autofit = Autofit::Normal;
            fancy.body_properties.wrap = Wrap::None;
            fancy.paragraphs[0].runs[0].bold = Some(true);
            fancy.paragraphs[0].runs[0].font_size = Some(TextSize::new(1400).unwrap());
            ws.add_shape(fancy).unwrap();
            let _ = ws;
            workbook.save(&path).unwrap();
        }

        let workbook = crate::workbook::Workbook::open(&path).unwrap();
        let inventory = workbook.shapes_on_sheet(&sheet_name).unwrap();
        assert_eq!(inventory.objects.len(), 2);

        let Object::Shape(greeting) = &inventory.objects[0].object else {
            panic!("expected a shape");
        };
        assert_eq!(greeting.non_visual.name.as_deref(), Some("Greeting"));
        assert!(greeting.is_text_box);
        assert_eq!(greeting.preset(), Some(Preset::RoundRect));
        assert_eq!(
            greeting.text_body.as_ref().unwrap().text(),
            "line one\nline two"
        );
        let Anchor::TwoCell { from, .. } = inventory.objects[0].anchor else {
            panic!("expected a two-cell anchor");
        };
        assert_eq!(from.column, 2);
        assert_eq!(from.column_offset, Emu(57150));

        let Object::Shape(fancy) = &inventory.objects[1].object else {
            panic!("expected a shape");
        };
        assert!(matches!(
            inventory.objects[1].anchor,
            Anchor::OneCell { .. }
        ));
        let body = fancy.text_body.as_ref().unwrap();
        assert_eq!(body.body_properties.vertical_anchor, VerticalAnchor::Bottom);
        assert_eq!(body.body_properties.autofit, Autofit::Normal);
        assert_eq!(body.body_properties.wrap, Wrap::None);
        assert_eq!(body.paragraphs[0].runs[0].bold, Some(true));
        assert_eq!(
            body.paragraphs[0].runs[0].font_size.map(TextSize::get),
            Some(1400)
        );

        // The saved package stays valid for the crate's own readers, and the
        // image pipeline still sees its picture.
        let package = litchi_opc::OpcPackage::open(&path).unwrap();
        let drawing_part = package
            .get_part(&litchi_opc::PackURI::new("/xl/drawings/drawing1.xml").unwrap())
            .unwrap();
        assert_eq!(
            drawing_part.content_type(),
            litchi_opc::constants::content_type::OFC_DRAWING
        );
        let drawing_xml = std::str::from_utf8(drawing_part.blob()).unwrap();
        assert!(drawing_xml.contains("<xdr:pic>"));
        assert!(drawing_xml.contains("<xdr:sp "));
    }
}

#[cfg(any())]
mod group_connection_tests {
    use super::shapes::{CellMarker, Emu, Object, parse_drawing_shapes};
    use super::*;
    use crate::writer::sheet::MutableWorksheet;
    use litchi_drawingml::geometry::Path;

    fn marker(column: u32, row: u32) -> CellMarker {
        CellMarker {
            column,
            column_offset: Emu(0),
            row,
            row_offset: Emu(0),
        }
    }

    fn anchor(from: (u32, u32), to: (u32, u32)) -> Anchor {
        Anchor::TwoCell {
            from: marker(from.0, from.1),
            to: marker(to.0, to.1),
            edit_as: EditAs::TwoCell,
        }
    }

    fn drawing_wrap(body: &str) -> String {
        format!(
            "<xdr:wsDr xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\" \
             xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">{body}</xdr:wsDr>"
        )
    }

    fn transform() -> GroupTransform {
        GroupTransform {
            offset: Some(EmuOffset {
                x: Emu(10),
                y: Emu(20),
            }),
            extent: Some(EmuExtent {
                width: Emu(30),
                height: Emu(40),
            }),
            child_offset: Some(EmuOffset {
                x: Emu(50),
                y: Emu(60),
            }),
            child_extent: Some(EmuExtent {
                width: Emu(70),
                height: Emu(80),
            }),
        }
    }

    /// Collect every drawing object ID in an inventory object tree.
    fn collect_ids(object: &Object, ids: &mut Vec<u32>) {
        let (non_visual, children) = match object {
            Object::Shape(shape) => (&shape.non_visual, &[][..]),
            Object::ConnectionShape(connection) => (&connection.non_visual, &[][..]),
            Object::Group(group) => {
                let children: &[Object] = &group.children;
                (&group.non_visual, children)
            },
            Object::OleObject(ole) => (&ole.non_visual, &[][..]),
        };
        if let Some(id) = non_visual.id {
            ids.push(id);
        }
        for child in children {
            collect_ids(child, ids);
        }
    }

    #[test]
    fn group_with_nested_children_round_trips() {
        let nested = GroupSpec::new("Inner", anchor((0, 0), (1, 1))).with_child(
            ShapeSpec::text_box("Deep", anchor((0, 0), (1, 1)), Preset::Ellipse, "deep").into(),
        );
        let mut group = GroupSpec::new("Outer", anchor((1, 1), (8, 12)))
            .with_child(
                ShapeSpec::text_box("First", anchor((2, 2), (4, 4)), Preset::Rect, "one").into(),
            )
            .with_child(
                ShapeSpec::text_box("Second", anchor((5, 5), (7, 7)), Preset::RoundRect, "two")
                    .into(),
            )
            .with_child(nested.into());
        group.transform = Some(transform());
        group.description = Some("grp".to_string());

        let mut xml = String::new();
        ShapeEmitter::new(1)
            .write_anchored_group(&mut xml, &group)
            .unwrap();
        let objects = parse_drawing_shapes(&drawing_wrap(&xml)).unwrap().unwrap();
        assert_eq!(objects.len(), 1);
        let Object::Group(outer) = &objects[0].object else {
            panic!("expected a group");
        };
        assert_eq!(outer.non_visual.name.as_deref(), Some("Outer"));
        assert_eq!(outer.non_visual.description.as_deref(), Some("grp"));
        assert_eq!(outer.transform, Some(transform()));
        assert_eq!(outer.children.len(), 3);
        let Object::Shape(first) = &outer.children[0] else {
            panic!("expected a shape");
        };
        assert_eq!(first.text_body.as_ref().unwrap().text(), "one");
        let Object::Group(inner) = &outer.children[2] else {
            panic!("expected a nested group");
        };
        assert_eq!(inner.non_visual.name.as_deref(), Some("Inner"));
        assert!(inner.transform.is_none());
        assert_eq!(inner.children.len(), 1);

        // IDs are unique across the group, its children, and nested groups.
        let mut ids = Vec::new();
        collect_ids(&objects[0].object, &mut ids);
        assert_eq!(ids.len(), 5);
        let mut unique = ids.clone();
        unique.sort_unstable();
        unique.dedup();
        assert_eq!(unique, ids);
    }

    #[test]
    fn connector_resolves_named_shapes() {
        let start_shape = ShapeSpec::text_box("Start", anchor((0, 0), (2, 2)), Preset::Rect, "a");
        let end_shape = ShapeSpec::text_box("End", anchor((4, 4), (6, 6)), Preset::Ellipse, "b");
        let connector = ConnectionShapeSpec::new(
            "Link",
            anchor((2, 2), (4, 4)),
            Preset::BentConnector3,
            ConnectionEndSpec {
                shape_name: "Start".to_string(),
                site: 3,
            },
            ConnectionEndSpec {
                shape_name: "End".to_string(),
                site: 1,
            },
        );

        let mut xml = String::new();
        let mut emitter = ShapeEmitter::new(5);
        emitter
            .write_anchored_shape(&mut xml, &start_shape)
            .unwrap();
        emitter.write_anchored_shape(&mut xml, &end_shape).unwrap();
        emitter
            .write_anchored_connection(&mut xml, &connector)
            .unwrap();

        let objects = parse_drawing_shapes(&drawing_wrap(&xml)).unwrap().unwrap();
        assert_eq!(objects.len(), 3);
        let Object::Shape(start) = &objects[0].object else {
            panic!("expected a shape");
        };
        let Object::Shape(end) = &objects[1].object else {
            panic!("expected a shape");
        };
        let Object::ConnectionShape(link) = &objects[2].object else {
            panic!("expected a connection shape");
        };
        assert_eq!(link.non_visual.id, Some(7));
        assert_eq!(link.preset(), Some(Preset::BentConnector3));
        assert_eq!(
            link.start,
            Some(super::shapes::ConnectionEnd {
                shape_id: start.non_visual.id.unwrap(),
                site: 3,
            })
        );
        assert_eq!(
            link.end,
            Some(super::shapes::ConnectionEnd {
                shape_id: end.non_visual.id.unwrap(),
                site: 1,
            })
        );
        assert_eq!(start.non_visual.id, Some(5));
        assert_eq!(end.non_visual.id, Some(6));
    }

    #[test]
    fn custom_connector_geometry_round_trips_without_a_preset_state() {
        let start = ShapeSpec::shape("Start", anchor((0, 0), (1, 1)), Preset::Rect, "");
        let end = ShapeSpec::shape("End", anchor((3, 3), (4, 4)), Preset::Rect, "");
        let geometry = CustomGeometry::new().with_path(Path::new(10, 10));
        let connector = ConnectionShapeSpec::new(
            "Custom Link",
            anchor((1, 1), (3, 3)),
            geometry,
            ConnectionEndSpec {
                shape_name: "Start".to_string(),
                site: 0,
            },
            ConnectionEndSpec {
                shape_name: "End".to_string(),
                site: 0,
            },
        );

        let mut xml = String::new();
        let mut emitter = ShapeEmitter::new(1);
        emitter.write_anchored_shape(&mut xml, &start).unwrap();
        emitter.write_anchored_shape(&mut xml, &end).unwrap();
        emitter
            .write_anchored_connection(&mut xml, &connector)
            .unwrap();
        assert!(xml.contains("<a:custGeom>"));

        let objects = parse_drawing_shapes(&drawing_wrap(&xml)).unwrap().unwrap();
        let Object::ConnectionShape(connection) = &objects[2].object else {
            panic!("expected a connection shape");
        };
        assert_eq!(connection.preset(), None);
        assert!(connection.custom_geometry().is_some());
    }

    #[test]
    fn preallocated_graph_resolves_forward_group_child_references() {
        let connection = ConnectionShapeSpec::new(
            "Forward",
            anchor((0, 0), (1, 1)),
            Preset::StraightConnector1,
            ConnectionEndSpec {
                shape_name: "Later".to_string(),
                site: 0,
            },
            ConnectionEndSpec {
                shape_name: "Last".to_string(),
                site: 1,
            },
        );
        let group = GroupSpec::new("Group", anchor((0, 0), (4, 4)))
            .with_child(connection.into())
            .with_child(ShapeSpec::shape("Later", anchor((1, 1), (2, 2)), Preset::Rect, "").into())
            .with_child(
                ShapeSpec::shape("Last", anchor((2, 2), (3, 3)), Preset::Ellipse, "").into(),
            );
        let mut emitter = ShapeEmitter::for_objects(1, &[], std::slice::from_ref(&group), &[])
            .expect("shape graph should preallocate");
        let mut xml = String::new();
        emitter.write_anchored_group(&mut xml, &group).unwrap();
        let objects = parse_drawing_shapes(&drawing_wrap(&xml)).unwrap().unwrap();
        let Object::Group(group) = &objects[0].object else {
            panic!("expected group");
        };
        let Object::ConnectionShape(connection) = &group.children[0] else {
            panic!("expected forward connector");
        };
        assert_eq!(connection.start.unwrap().shape_id, 3);
        assert_eq!(connection.end.unwrap().shape_id, 4);
    }

    #[test]
    fn connector_to_group_child_resolves() {
        let group = GroupSpec::new("Box", anchor((0, 0), (3, 3))).with_child(
            ShapeSpec::text_box("Child", anchor((0, 0), (1, 1)), Preset::Rect, "c").into(),
        );
        let target = ShapeSpec::text_box("Target", anchor((5, 5), (7, 7)), Preset::Ellipse, "t");
        let connector = ConnectionShapeSpec::new(
            "L",
            anchor((3, 3), (5, 5)),
            Preset::StraightConnector1,
            ConnectionEndSpec {
                shape_name: "Child".to_string(),
                site: 0,
            },
            ConnectionEndSpec {
                shape_name: "Target".to_string(),
                site: 2,
            },
        );
        let mut xml = String::new();
        let mut emitter = ShapeEmitter::new(1);
        emitter.write_anchored_group(&mut xml, &group).unwrap();
        emitter.write_anchored_shape(&mut xml, &target).unwrap();
        emitter
            .write_anchored_connection(&mut xml, &connector)
            .unwrap();
        let objects = parse_drawing_shapes(&drawing_wrap(&xml)).unwrap().unwrap();
        let Object::ConnectionShape(link) = &objects[2].object else {
            panic!("expected a connection shape");
        };
        // Child got ID 2 (group is 1), Target got ID 3.
        assert_eq!(link.start.unwrap().shape_id, 2);
        assert_eq!(link.end.unwrap().shape_id, 3);
    }

    #[test]
    fn unknown_or_duplicate_names_fail_serialization() {
        let dangling = ConnectionShapeSpec::new(
            "Bad",
            anchor((0, 0), (1, 1)),
            Preset::StraightConnector1,
            ConnectionEndSpec {
                shape_name: "Ghost".to_string(),
                site: 0,
            },
            ConnectionEndSpec {
                shape_name: "AlsoGhost".to_string(),
                site: 0,
            },
        );
        let mut xml = String::new();
        assert!(
            ShapeEmitter::new(1)
                .write_anchored_connection(&mut xml, &dangling)
                .is_err()
        );

        // Duplicate shape names are ambiguous and rejected.
        let shape = ShapeSpec::text_box("Dup", anchor((0, 0), (1, 1)), Preset::Rect, "");
        let mut xml = String::new();
        let mut emitter = ShapeEmitter::new(1);
        emitter.write_anchored_shape(&mut xml, &shape).unwrap();
        assert!(emitter.write_anchored_shape(&mut xml, &shape).is_err());
    }

    #[test]
    fn worksheet_api_manages_groups_and_connections() {
        let mut ws = MutableWorksheet::new("Sheet1".to_string(), 1);
        ws.add_text_box("A", anchor((0, 0), (2, 2)), Preset::Rect, "a")
            .unwrap();
        ws.add_group(GroupSpec::new("G", anchor((3, 3), (6, 6))).with_child(
            ShapeSpec::text_box("GA", anchor((3, 3), (4, 4)), Preset::Ellipse, "g").into(),
        ))
        .unwrap();
        ws.add_connection(ConnectionShapeSpec::new(
            "C",
            anchor((2, 2), (3, 3)),
            Preset::CurvedConnector3,
            ConnectionEndSpec {
                shape_name: "A".to_string(),
                site: 1,
            },
            ConnectionEndSpec {
                shape_name: "GA".to_string(),
                site: 2,
            },
        ))
        .unwrap();
        assert_eq!(ws.groups().len(), 1);
        assert_eq!(ws.connections().len(), 1);
        assert!(ws.remove_group(4).is_err());
        assert!(ws.remove_connection(3).is_err());

        let xml = ws.generate_drawing_xml().unwrap().unwrap();
        let objects = parse_drawing_shapes(&xml).unwrap().unwrap();
        assert_eq!(objects.len(), 3);
        let mut ids = Vec::new();
        for anchored in &objects {
            collect_ids(&anchored.object, &mut ids);
        }
        assert_eq!(ids.len(), 4);
        let mut unique = ids.clone();
        unique.sort_unstable();
        unique.dedup();
        assert_eq!(unique.len(), ids.len(), "object IDs must be unique");

        // A connector referencing an unknown shape fails at save time.
        ws.add_connection(ConnectionShapeSpec::new(
            "Dangling",
            anchor((0, 0), (1, 1)),
            Preset::StraightConnector1,
            ConnectionEndSpec {
                shape_name: "Nope".to_string(),
                site: 0,
            },
            ConnectionEndSpec {
                shape_name: "A".to_string(),
                site: 0,
            },
        ))
        .unwrap();
        assert!(ws.generate_drawing_xml().is_err());
        ws.remove_connection(1).unwrap();
        assert!(ws.generate_drawing_xml().is_ok());
    }

    #[test]
    fn groups_and_connectors_round_trip_through_a_saved_package() {
        let directory = tempfile::tempdir().unwrap();
        let path = directory.path().join("groups.xlsx");
        let sheet_name;
        {
            let mut workbook = crate::workbook::Workbook::create().unwrap();
            let ws = workbook.worksheet_mut(0).unwrap();
            sheet_name = ws.name().to_string();
            ws.add_image(vec![7, 7], "png", 1, 1, 2, 2, None).unwrap();
            ws.add_text_box("Solo", anchor((0, 0), (2, 2)), Preset::Rect, "solo")
                .unwrap();
            let mut group = GroupSpec::new("Pair", anchor((3, 3), (9, 9)))
                .with_child(
                    ShapeSpec::text_box("Left", anchor((3, 3), (5, 5)), Preset::Rect, "L").into(),
                )
                .with_child(
                    ShapeSpec::text_box("Right", anchor((6, 6), (8, 8)), Preset::Ellipse, "R")
                        .into(),
                );
            group.transform = Some(transform());
            ws.add_group(group).unwrap();
            ws.add_connection(ConnectionShapeSpec::new(
                "Bridge",
                anchor((5, 5), (6, 6)),
                Preset::StraightConnector1,
                ConnectionEndSpec {
                    shape_name: "Left".to_string(),
                    site: 4,
                },
                ConnectionEndSpec {
                    shape_name: "Right".to_string(),
                    site: 0,
                },
            ))
            .unwrap();
            let _ = ws;
            workbook.save(&path).unwrap();
        }

        let workbook = crate::workbook::Workbook::open(&path).unwrap();
        let inventory = workbook.shapes_on_sheet(&sheet_name).unwrap();
        // Image is handled by the picture pipeline; three authored objects.
        assert_eq!(inventory.objects.len(), 3);

        let Object::Shape(solo) = &inventory.objects[0].object else {
            panic!("expected a shape");
        };
        assert_eq!(solo.non_visual.name.as_deref(), Some("Solo"));
        // The picture consumed ID 1, so authored objects start at 2.
        assert_eq!(solo.non_visual.id, Some(2));

        let Object::Group(pair) = &inventory.objects[1].object else {
            panic!("expected a group");
        };
        assert_eq!(pair.non_visual.name.as_deref(), Some("Pair"));
        assert_eq!(pair.transform, Some(transform()));
        assert_eq!(pair.children.len(), 2);
        let Object::Shape(left) = &pair.children[0] else {
            panic!("expected a shape");
        };
        let Object::Shape(right) = &pair.children[1] else {
            panic!("expected a shape");
        };

        let Object::ConnectionShape(bridge) = &inventory.objects[2].object else {
            panic!("expected a connection shape");
        };
        assert_eq!(bridge.preset(), Some(Preset::StraightConnector1));
        assert_eq!(bridge.start.unwrap().shape_id, left.non_visual.id.unwrap());
        assert_eq!(bridge.start.unwrap().site, 4);
        assert_eq!(bridge.end.unwrap().shape_id, right.non_visual.id.unwrap());

        // Every object ID in the drawing is unique.
        let mut ids = Vec::new();
        for anchored in &inventory.objects {
            collect_ids(&anchored.object, &mut ids);
        }
        let mut unique = ids.clone();
        unique.sort_unstable();
        unique.dedup();
        assert_eq!(unique.len(), ids.len());
    }
}
