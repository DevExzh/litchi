//! High-performance Shape enum for representing all shape types.
//!
//! Idiomatic Rust implementation using enum variants instead of trait objects.

use super::picture::PictureShape;
use super::shape::{Shape, ShapeType, SourceIdentity};
use super::{AutoShape, Placeholder, TextBox};
use crate::package::Result;
use crate::slide_extension::ShapeMetadata;

/// Represents any shape on a slide using an enum for zero-cost abstraction.
///
/// Uses lifetime parameter `'a` for zero-copy parsing of shape data.
///
/// # Performance
///
/// - Enum dispatch (no vtable overhead)
/// - Stack-allocated (no heap allocation for shape variants)
/// - Pattern matching compiles to efficient jump tables
/// - Zero-copy parsing when possible
#[derive(Debug, Clone)]
pub enum ShapeEnum<'a> {
    /// Text box shape containing editable text
    TextBox(TextBox<'a>),
    /// Placeholder shape (title, body, footer, etc.)
    Placeholder(Placeholder<'a>),
    /// Auto shape (rectangle, ellipse, arrow, etc.)
    AutoShape(AutoShape<'a>),
    /// Picture frame, including OLE-object and media previews
    Picture(PictureShape),
    /// Table shape
    Table(TableShape),
    /// Group shape containing other shapes
    Group(GroupShape<'a>),
    /// Line/connector shape
    Line(LineShape),
}

impl<'a> ShapeEnum<'a>
where
    'a: 'static,
{
    pub(crate) fn mark_source_bound_recursive(&mut self) {
        match self {
            ShapeEnum::TextBox(shape) => shape.mark_source_bound(),
            ShapeEnum::Placeholder(shape) => shape.mark_source_bound(),
            ShapeEnum::AutoShape(shape) => shape.mark_source_bound(),
            ShapeEnum::Picture(shape) => shape.mark_source_bound(),
            ShapeEnum::Table(shape) => shape.mark_source_bound(),
            ShapeEnum::Group(shape) => {
                shape.mark_source_bound();
                for child in &mut shape.children {
                    child.mark_source_bound_recursive();
                }
            },
            ShapeEnum::Line(shape) => shape.mark_source_bound(),
        }
    }

    pub(crate) fn apply_source_identity(&mut self, identity: SourceIdentity) {
        match self {
            ShapeEnum::TextBox(shape) => shape.apply_source_identity(identity),
            ShapeEnum::Placeholder(shape) => shape.apply_source_identity(identity),
            ShapeEnum::AutoShape(shape) => shape.apply_source_identity(identity),
            ShapeEnum::Picture(shape) => shape.apply_source_identity(identity),
            ShapeEnum::Table(shape) => shape.apply_source_identity(identity),
            ShapeEnum::Group(shape) => shape.apply_source_identity(identity),
            ShapeEnum::Line(shape) => shape.apply_source_identity(identity),
        }
    }

    /// The `OfficeArt` shape identifier (`spid`), unique within its drawing
    /// (MS-ODRAW 2.2.40); the stable identity a shape keeps across reads
    /// (L-PPT-R4).
    #[must_use]
    pub fn id(&self) -> u32 {
        match self {
            ShapeEnum::TextBox(shape) => shape.properties().id,
            ShapeEnum::Placeholder(shape) => shape.properties().id,
            ShapeEnum::AutoShape(shape) => shape.properties().id,
            ShapeEnum::Picture(shape) => shape.properties().id,
            ShapeEnum::Table(shape) => shape.id(),
            ShapeEnum::Group(shape) => shape.id(),
            ShapeEnum::Line(shape) => shape.id(),
        }
    }

    /// The shape's name (`wzName`, MS-ODRAW 2.3.4.1), such as
    /// `"Title 1"` or `"Rectangle 3"`, when it carries one (L-PPT-R3).
    #[must_use]
    pub fn name(&self) -> Option<&str> {
        match self {
            ShapeEnum::TextBox(shape) => shape.properties().name.as_deref(),
            ShapeEnum::Placeholder(shape) => shape.properties().name.as_deref(),
            ShapeEnum::AutoShape(shape) => shape.properties().name.as_deref(),
            ShapeEnum::Picture(shape) => shape.properties().name.as_deref(),
            ShapeEnum::Table(shape) => shape.name(),
            ShapeEnum::Group(shape) => shape.name(),
            ShapeEnum::Line(shape) => shape.name(),
        }
    }

    /// The shape's alternative text (`wzDescription`, MS-ODRAW 2.3.4.3),
    /// when it carries one (L-PPT-R3).
    #[must_use]
    pub fn description(&self) -> Option<&str> {
        match self {
            ShapeEnum::TextBox(shape) => shape.properties().description.as_deref(),
            ShapeEnum::Placeholder(shape) => shape.properties().description.as_deref(),
            ShapeEnum::AutoShape(shape) => shape.properties().description.as_deref(),
            ShapeEnum::Picture(shape) => shape.properties().description.as_deref(),
            ShapeEnum::Table(shape) => shape.description(),
            ShapeEnum::Group(shape) => shape.description(),
            ShapeEnum::Line(shape) => shape.description(),
        }
    }

    /// The shape's stable position in its source drawing (L-PPT-R4): its
    /// zero-based index among its siblings — the drawing's top-level shapes,
    /// or its group's children — in drawing order.
    ///
    /// The index counts every shape of the drawing's shape tree, including
    /// shapes the reader does not return (an unreadable shape, a line
    /// without an anchor, an unknown kind), so it does not shift when a
    /// sibling is dropped, and it is the same on every read of the same
    /// drawing; the background shape and a group's own header are not
    /// siblings. The position in [`crate::Slide::shapes`] counts returned
    /// shapes only. `None` for a shape that was not read from a drawing.
    #[must_use]
    pub fn source_index(&self) -> Option<usize> {
        match self {
            ShapeEnum::TextBox(shape) => shape.properties().source_index,
            ShapeEnum::Placeholder(shape) => shape.properties().source_index,
            ShapeEnum::AutoShape(shape) => shape.properties().source_index,
            ShapeEnum::Picture(shape) => shape.properties().source_index,
            ShapeEnum::Table(shape) => shape.source_index(),
            ShapeEnum::Group(shape) => shape.source_index(),
            ShapeEnum::Line(shape) => shape.source_index(),
        }
    }

    /// Get the shape type.
    #[must_use]
    pub fn shape_type(&self) -> ShapeType {
        match self {
            ShapeEnum::TextBox(_) => ShapeType::TextBox,
            ShapeEnum::Placeholder(_) => ShapeType::Placeholder,
            ShapeEnum::AutoShape(_) => ShapeType::AutoShape,
            ShapeEnum::Picture(picture) => picture.properties.shape_type,
            ShapeEnum::Table(_) => ShapeType::Table,
            ShapeEnum::Group(_) => ShapeType::Group,
            ShapeEnum::Line(line) => line.shape_type(),
        }
    }

    /// Extract text from the shape if it contains text.
    ///
    /// # Performance
    ///
    /// - Pattern matching compiles to jump table
    /// - No heap allocation for empty text
    /// - Recursive for group shapes
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn text(&self) -> Result<String> {
        match self {
            ShapeEnum::TextBox(tb) => Shape::text(tb),
            ShapeEnum::Placeholder(ph) => Shape::text(ph),
            ShapeEnum::AutoShape(as_) => Shape::text(as_),
            ShapeEnum::Table(table) => {
                // Extract text from all table cells
                let mut text_parts = Vec::new();
                for row in 0..table.rows() {
                    for col in 0..table.columns() {
                        if let Some(cell_text) = table.cell(row, col)
                            && !cell_text.is_empty()
                        {
                            text_parts.push(cell_text.to_string());
                        }
                    }
                }
                Ok(text_parts.join(" "))
            },
            ShapeEnum::Group(group) => {
                let mut text_parts = Vec::new();
                let mut shapes: Vec<_> = group.children().iter().rev().collect();
                let mut visited = 0u32;
                while let Some(shape) = shapes.pop() {
                    visited = visited.checked_add(1).ok_or_else(|| {
                        crate::package::Error::Corrupted(
                            "PowerPoint shape count overflowed".to_string(),
                        )
                    })?;
                    if visited > 1_000_000 {
                        return Err(crate::package::Error::Corrupted(
                            "PowerPoint group contains more than 1000000 shapes".to_string(),
                        ));
                    }

                    if let ShapeEnum::Group(child_group) = shape {
                        shapes.extend(child_group.children().iter().rev());
                    } else {
                        let child_text = shape.text()?;
                        if !child_text.is_empty() {
                            text_parts.push(child_text);
                        }
                    }
                }
                Ok(text_parts.join("\n"))
            },
            ShapeEnum::Picture(_) | ShapeEnum::Line(_) => Ok(String::new()),
        }
    }

    /// Get shape as `TextBox` if it is one.
    #[inline]
    #[must_use]
    pub fn as_textbox(&self) -> Option<&TextBox<'a>> {
        match self {
            ShapeEnum::TextBox(tb) => Some(tb),
            ShapeEnum::Placeholder(_)
            | ShapeEnum::AutoShape(_)
            | ShapeEnum::Picture(_)
            | ShapeEnum::Table(_)
            | ShapeEnum::Group(_)
            | ShapeEnum::Line(_) => None,
        }
    }

    /// Get shape as Placeholder if it is one.
    #[inline]
    #[must_use]
    pub fn as_placeholder(&self) -> Option<&Placeholder<'a>> {
        match self {
            ShapeEnum::Placeholder(ph) => Some(ph),
            ShapeEnum::TextBox(_)
            | ShapeEnum::AutoShape(_)
            | ShapeEnum::Picture(_)
            | ShapeEnum::Table(_)
            | ShapeEnum::Group(_)
            | ShapeEnum::Line(_) => None,
        }
    }

    /// Get shape as `AutoShape` if it is one.
    #[inline]
    #[must_use]
    pub fn as_autoshape(&self) -> Option<&AutoShape<'a>> {
        match self {
            ShapeEnum::AutoShape(as_) => Some(as_),
            ShapeEnum::TextBox(_)
            | ShapeEnum::Placeholder(_)
            | ShapeEnum::Picture(_)
            | ShapeEnum::Table(_)
            | ShapeEnum::Group(_)
            | ShapeEnum::Line(_) => None,
        }
    }

    /// Get shape as `PictureShape` if it is one.
    #[inline]
    #[must_use]
    pub fn as_picture(&self) -> Option<&PictureShape> {
        match self {
            ShapeEnum::Picture(pic) => Some(pic),
            ShapeEnum::TextBox(_)
            | ShapeEnum::Placeholder(_)
            | ShapeEnum::AutoShape(_)
            | ShapeEnum::Table(_)
            | ShapeEnum::Group(_)
            | ShapeEnum::Line(_) => None,
        }
    }

    /// Get shape as mutable `PictureShape` if it is one.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    #[inline]
    pub fn as_picture_mut(
        &mut self,
    ) -> std::result::Result<Option<&mut PictureShape>, crate::shapes::shape::MutationError> {
        match self {
            ShapeEnum::Picture(picture) => {
                picture.ensure_mutable(crate::shapes::shape::Mutation::Picture)?;
                Ok(Some(picture))
            },
            ShapeEnum::TextBox(_)
            | ShapeEnum::Placeholder(_)
            | ShapeEnum::AutoShape(_)
            | ShapeEnum::Table(_)
            | ShapeEnum::Group(_)
            | ShapeEnum::Line(_) => Ok(None),
        }
    }

    /// Get an embedded or linked OLE object frame.
    #[inline]
    #[must_use]
    pub fn as_object_frame(&self) -> Option<&PictureShape> {
        self.as_picture()
            .filter(|picture| picture.frame_kind() == super::picture::PictureFrameKind::OleObject)
    }

    /// Get an audio or video media frame.
    #[inline]
    #[must_use]
    pub fn as_media_frame(&self) -> Option<&PictureShape> {
        self.as_picture()
            .filter(|picture| picture.frame_kind() == super::picture::PictureFrameKind::Media)
    }

    /// Get shape as `TableShape` if it is one.
    #[inline]
    #[must_use]
    pub fn as_table(&self) -> Option<&TableShape> {
        match self {
            ShapeEnum::Table(table) => Some(table),
            ShapeEnum::TextBox(_)
            | ShapeEnum::Placeholder(_)
            | ShapeEnum::AutoShape(_)
            | ShapeEnum::Picture(_)
            | ShapeEnum::Group(_)
            | ShapeEnum::Line(_) => None,
        }
    }

    /// Get shape as `GroupShape` if it is one.
    #[inline]
    #[must_use]
    pub fn as_group(&self) -> Option<&GroupShape<'a>> {
        match self {
            ShapeEnum::Group(group) => Some(group),
            ShapeEnum::TextBox(_)
            | ShapeEnum::Placeholder(_)
            | ShapeEnum::AutoShape(_)
            | ShapeEnum::Picture(_)
            | ShapeEnum::Table(_)
            | ShapeEnum::Line(_) => None,
        }
    }

    /// Get shape as `LineShape` if it is one.
    #[inline]
    #[must_use]
    pub fn as_line(&self) -> Option<&LineShape> {
        match self {
            ShapeEnum::Line(line) => Some(line),
            ShapeEnum::TextBox(_)
            | ShapeEnum::Placeholder(_)
            | ShapeEnum::AutoShape(_)
            | ShapeEnum::Picture(_)
            | ShapeEnum::Table(_)
            | ShapeEnum::Group(_) => None,
        }
    }

    /// Return inert `PowerPoint` 12 placeholder metadata retained for round trips.
    #[must_use]
    pub fn powerpoint12_shape_metadata(&self) -> Option<&ShapeMetadata> {
        match self {
            ShapeEnum::TextBox(shape) => shape.properties().powerpoint12_shape_metadata.as_ref(),
            ShapeEnum::Placeholder(shape) => {
                shape.properties().powerpoint12_shape_metadata.as_ref()
            },
            ShapeEnum::AutoShape(shape) => shape.properties().powerpoint12_shape_metadata.as_ref(),
            ShapeEnum::Picture(shape) => shape.properties().powerpoint12_shape_metadata.as_ref(),
            ShapeEnum::Table(shape) => shape.powerpoint12_shape_metadata.as_ref(),
            ShapeEnum::Group(shape) => shape.powerpoint12_shape_metadata.as_ref(),
            ShapeEnum::Line(shape) => shape.powerpoint12_shape_metadata.as_ref(),
        }
    }
}

/// The identity of a table, group, or line read from a drawing; the other
/// shape kinds keep it in their [`ShapeProperties`](super::shape::ShapeProperties).
#[derive(Debug, Clone, Default, PartialEq, Eq)]
struct ShapeIdentity {
    name: Option<String>,
    description: Option<String>,
    source_index: Option<usize>,
}

impl From<SourceIdentity> for ShapeIdentity {
    fn from(identity: SourceIdentity) -> Self {
        Self {
            name: identity.name,
            description: identity.description,
            source_index: Some(identity.index),
        }
    }
}

// PictureShape is now defined in picture.rs and re-exported

/// Table shape.
///
/// Represents a table with rows and columns.
#[derive(Debug, Clone)]
pub struct TableShape {
    /// Shape ID
    id: u32,
    /// Number of rows
    rows: usize,
    /// Number of columns
    columns: usize,
    /// Table cells (row-major order)
    cells: Vec<Vec<String>>,
    /// The `OfficeArt` shape ID of each cell's shape, when read from a
    /// drawing (record 0323).
    cell_ids: Vec<Vec<Option<u32>>>,
    /// The top of each row and the left of each column, in master units,
    /// when read from a drawing.
    row_tops: Vec<i32>,
    column_lefts: Vec<i32>,
    /// Left coordinate
    left: i32,
    /// Top coordinate
    top: i32,
    /// Width
    width: i32,
    /// Height
    height: i32,
    powerpoint12_shape_metadata: Option<ShapeMetadata>,
    identity: ShapeIdentity,
    source_bound: bool,
}

impl TableShape {
    /// Create a new table shape.
    #[must_use]
    pub fn new(id: u32, rows: usize, columns: usize) -> Self {
        let cells = vec![vec![String::new(); columns]; rows];
        Self {
            id,
            rows,
            columns,
            cells,
            cell_ids: Vec::new(),
            row_tops: Vec::new(),
            column_lefts: Vec::new(),
            left: 0,
            top: 0,
            width: 0,
            height: 0,
            powerpoint12_shape_metadata: None,
            identity: ShapeIdentity::default(),
            source_bound: false,
        }
    }

    pub(crate) fn set_cell_text(&mut self, row: usize, col: usize, text: String) {
        if let Some(cell) = self.cells.get_mut(row).and_then(|cells| cells.get_mut(col)) {
            *cell = text;
        }
    }

    pub(crate) fn set_cell_id(&mut self, row: usize, col: usize, id: u32) {
        if self.cell_ids.is_empty() {
            self.cell_ids = vec![vec![None; self.columns]; self.rows];
        }
        if let Some(cell) = self
            .cell_ids
            .get_mut(row)
            .and_then(|cells| cells.get_mut(col))
        {
            *cell = Some(id);
        }
    }

    pub(crate) fn set_grid(&mut self, row_tops: Vec<i32>, column_lefts: Vec<i32>) {
        self.row_tops = row_tops;
        self.column_lefts = column_lefts;
    }

    /// The `OfficeArt` shape ID of the cell shape at `row`, `col` (record
    /// 0323); `None` for a covered cell or a table not read from a drawing.
    #[must_use]
    pub fn cell_shape_id(&self, row: usize, col: usize) -> Option<u32> {
        self.cell_ids.get(row)?.get(col).copied().flatten()
    }

    /// The height of each row in master units (1/576 inch): the distance to
    /// the next row's top, the last to the table's bottom (record 0323).
    /// Empty for a table not read from a drawing.
    #[must_use]
    pub fn row_heights(&self) -> Vec<i32> {
        extents(&self.row_tops, self.top.saturating_add(self.height))
    }

    /// The width of each column in master units, as [`Self::row_heights`].
    #[must_use]
    pub fn column_widths(&self) -> Vec<i32> {
        extents(&self.column_lefts, self.left.saturating_add(self.width))
    }

    pub(crate) fn set_bounds(&mut self, left: i32, top: i32, width: i32, height: i32) {
        self.left = left;
        self.top = top;
        self.width = width;
        self.height = height;
    }

    /// Get shape ID.
    #[must_use]
    pub fn id(&self) -> u32 {
        self.id
    }

    /// Get number of rows.
    #[must_use]
    pub fn rows(&self) -> usize {
        self.rows
    }

    /// Get number of columns.
    #[must_use]
    pub fn columns(&self) -> usize {
        self.columns
    }

    /// Get cell text.
    #[must_use]
    pub fn cell(&self, row: usize, col: usize) -> Option<&str> {
        self.cells
            .get(row)
            .and_then(|r| r.get(col))
            .map(String::as_str)
    }

    /// Get the left coordinate.
    #[must_use]
    pub fn left(&self) -> i32 {
        self.left
    }

    /// Get the top coordinate.
    #[must_use]
    pub fn top(&self) -> i32 {
        self.top
    }

    /// Get the table width.
    #[must_use]
    pub fn width(&self) -> i32 {
        self.width
    }

    /// Get the table height.
    #[must_use]
    pub fn height(&self) -> i32 {
        self.height
    }

    pub(crate) fn set_powerpoint12_shape_metadata(&mut self, metadata: Option<ShapeMetadata>) {
        self.powerpoint12_shape_metadata = metadata;
    }

    /// The shape's name (`wzName`, MS-ODRAW 2.3.4.1), when it carries one
    /// (L-PPT-R3).
    #[must_use]
    pub fn name(&self) -> Option<&str> {
        self.identity.name.as_deref()
    }

    /// The shape's alternative text (`wzDescription`, MS-ODRAW 2.3.4.3),
    /// when it carries one (L-PPT-R3).
    #[must_use]
    pub fn description(&self) -> Option<&str> {
        self.identity.description.as_deref()
    }

    /// The shape's stable zero-based position among its siblings in its
    /// source drawing (L-PPT-R4); `None` when it was not read from one.
    #[must_use]
    pub fn source_index(&self) -> Option<usize> {
        self.identity.source_index
    }

    pub(crate) fn apply_source_identity(&mut self, identity: SourceIdentity) {
        self.identity = ShapeIdentity::from(identity);
    }

    pub(crate) fn mark_source_bound(&mut self) {
        self.source_bound = true;
    }
}

/// The extent of each grid band from its start to the next start, the last
/// to `end`.
fn extents(starts: &[i32], end: i32) -> Vec<i32> {
    starts
        .iter()
        .enumerate()
        .map(|(index, start)| {
            starts
                .get(index + 1)
                .copied()
                .unwrap_or(end)
                .saturating_sub(*start)
                .max(0)
        })
        .collect()
}

/// Group shape containing other shapes.
///
/// Groups allow hierarchical organization of shapes.
#[derive(Debug, Clone)]
pub struct GroupShape<'a> {
    /// Shape ID
    id: u32,
    /// Child shapes
    children: Vec<ShapeEnum<'a>>,
    /// Left coordinate
    left: i32,
    /// Top coordinate
    top: i32,
    /// Width
    width: i32,
    /// Height
    height: i32,
    powerpoint12_shape_metadata: Option<ShapeMetadata>,
    identity: ShapeIdentity,
    source_bound: bool,
}

impl<'a> GroupShape<'a> {
    /// Create a new group shape.
    #[must_use]
    pub fn new(id: u32) -> Self {
        Self {
            id,
            children: Vec::new(),
            left: 0,
            top: 0,
            width: 0,
            height: 0,
            powerpoint12_shape_metadata: None,
            identity: ShapeIdentity::default(),
            source_bound: false,
        }
    }

    /// Add a child shape.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn add_child(
        &mut self,
        shape: ShapeEnum<'a>,
    ) -> std::result::Result<(), crate::shapes::shape::MutationError> {
        if self.source_bound {
            return Err(crate::shapes::shape::MutationError::source_bound(
                crate::shapes::shape::Mutation::Structure,
            ));
        }
        self.children.push(shape);
        Ok(())
    }

    pub(crate) fn add_decoded_child(&mut self, shape: ShapeEnum<'a>) {
        self.children.push(shape);
    }

    /// Get child shapes.
    #[must_use]
    pub fn children(&self) -> &[ShapeEnum<'a>] {
        &self.children
    }

    /// Set group bounds.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn set_bounds(
        &mut self,
        left: i32,
        top: i32,
        width: i32,
        height: i32,
    ) -> std::result::Result<(), crate::shapes::shape::MutationError> {
        if self.source_bound {
            return Err(crate::shapes::shape::MutationError::source_bound(
                crate::shapes::shape::Mutation::Geometry,
            ));
        }
        self.set_decoded_bounds(left, top, width, height);
        Ok(())
    }

    pub(crate) fn set_decoded_bounds(&mut self, left: i32, top: i32, width: i32, height: i32) {
        self.left = left;
        self.top = top;
        self.width = width;
        self.height = height;
    }

    /// Get shape ID.
    #[must_use]
    pub fn id(&self) -> u32 {
        self.id
    }

    /// Get the left coordinate.
    #[must_use]
    pub fn left(&self) -> i32 {
        self.left
    }

    /// Get the top coordinate.
    #[must_use]
    pub fn top(&self) -> i32 {
        self.top
    }

    /// Get the group width.
    #[must_use]
    pub fn width(&self) -> i32 {
        self.width
    }

    /// Get the group height.
    #[must_use]
    pub fn height(&self) -> i32 {
        self.height
    }

    pub(crate) fn set_powerpoint12_shape_metadata(&mut self, metadata: Option<ShapeMetadata>) {
        self.powerpoint12_shape_metadata = metadata;
    }

    /// The shape's name (`wzName`, MS-ODRAW 2.3.4.1), when it carries one
    /// (L-PPT-R3).
    #[must_use]
    pub fn name(&self) -> Option<&str> {
        self.identity.name.as_deref()
    }

    /// The shape's alternative text (`wzDescription`, MS-ODRAW 2.3.4.3),
    /// when it carries one (L-PPT-R3).
    #[must_use]
    pub fn description(&self) -> Option<&str> {
        self.identity.description.as_deref()
    }

    /// The shape's stable zero-based position among its siblings in its
    /// source drawing (L-PPT-R4); `None` when it was not read from one.
    #[must_use]
    pub fn source_index(&self) -> Option<usize> {
        self.identity.source_index
    }

    pub(crate) fn apply_source_identity(&mut self, identity: SourceIdentity) {
        self.identity = ShapeIdentity::from(identity);
    }

    pub(crate) fn mark_source_bound(&mut self) {
        self.source_bound = true;
    }
}

/// Line/connector shape.
///
/// Represents a line or connector between two points.
#[derive(Debug, Clone)]
pub struct LineShape {
    kind: LineKind,
    /// Shape ID
    id: u32,
    /// Start X coordinate
    x1: i32,
    /// Start Y coordinate
    y1: i32,
    /// End X coordinate
    x2: i32,
    /// End Y coordinate
    y2: i32,
    /// Line width
    width: i32,
    /// Line color
    color: Option<u32>,
    powerpoint12_shape_metadata: Option<ShapeMetadata>,
    identity: ShapeIdentity,
    source_bound: bool,
}

#[derive(Debug, Clone, Copy)]
enum LineKind {
    Line,
    Connector,
}

impl LineShape {
    /// Create a new line shape.
    #[must_use]
    pub fn new(id: u32, x1: i32, y1: i32, x2: i32, y2: i32) -> Self {
        Self {
            kind: LineKind::Line,
            id,
            x1,
            y1,
            x2,
            y2,
            width: 1,
            color: None,
            powerpoint12_shape_metadata: None,
            identity: ShapeIdentity::default(),
            source_bound: false,
        }
    }

    /// Create a connector between two points.
    #[must_use]
    pub fn connector(id: u32, x1: i32, y1: i32, x2: i32, y2: i32) -> Self {
        Self {
            kind: LineKind::Connector,
            ..Self::new(id, x1, y1, x2, y2)
        }
    }

    /// Return whether this shape is a plain line or a connector.
    #[must_use]
    pub fn shape_type(&self) -> ShapeType {
        match self.kind {
            LineKind::Line => ShapeType::Line,
            LineKind::Connector => ShapeType::Connector,
        }
    }

    /// Set line width.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn set_width(
        &mut self,
        width: i32,
    ) -> std::result::Result<(), crate::shapes::shape::MutationError> {
        if self.source_bound {
            return Err(crate::shapes::shape::MutationError::source_bound(
                crate::shapes::shape::Mutation::Line,
            ));
        }
        self.width = width;
        Ok(())
    }

    pub(crate) fn set_decoded_width(&mut self, width: i32) {
        self.width = width;
    }

    /// Set line color.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn set_color(
        &mut self,
        color: u32,
    ) -> std::result::Result<(), crate::shapes::shape::MutationError> {
        if self.source_bound {
            return Err(crate::shapes::shape::MutationError::source_bound(
                crate::shapes::shape::Mutation::Line,
            ));
        }
        self.color = Some(color);
        Ok(())
    }

    pub(crate) fn set_decoded_color(&mut self, color: u32) {
        self.color = Some(color);
    }

    /// Get shape ID.
    #[must_use]
    pub fn id(&self) -> u32 {
        self.id
    }

    /// Get line length.
    #[must_use]
    pub fn length(&self) -> f64 {
        let dx = f64::from(self.x2 - self.x1);
        let dy = f64::from(self.y2 - self.y1);
        (dx * dx + dy * dy).sqrt()
    }

    pub(crate) fn set_powerpoint12_shape_metadata(&mut self, metadata: Option<ShapeMetadata>) {
        self.powerpoint12_shape_metadata = metadata;
    }

    /// The shape's name (`wzName`, MS-ODRAW 2.3.4.1), when it carries one
    /// (L-PPT-R3).
    #[must_use]
    pub fn name(&self) -> Option<&str> {
        self.identity.name.as_deref()
    }

    /// The shape's alternative text (`wzDescription`, MS-ODRAW 2.3.4.3),
    /// when it carries one (L-PPT-R3).
    #[must_use]
    pub fn description(&self) -> Option<&str> {
        self.identity.description.as_deref()
    }

    /// The shape's stable zero-based position among its siblings in its
    /// source drawing (L-PPT-R4); `None` when it was not read from one.
    #[must_use]
    pub fn source_index(&self) -> Option<usize> {
        self.identity.source_index
    }

    pub(crate) fn apply_source_identity(&mut self, identity: SourceIdentity) {
        self.identity = ShapeIdentity::from(identity);
    }

    pub(crate) fn mark_source_bound(&mut self) {
        self.source_bound = true;
    }
}

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]
mod tests {
    use super::*;

    #[test]
    fn line_family_preserves_connector_semantics() {
        let line: ShapeEnum<'static> = ShapeEnum::Line(LineShape::new(1, 0, 0, 10, 10));
        let connector: ShapeEnum<'static> = ShapeEnum::Line(LineShape::connector(2, 0, 0, 10, 10));

        assert_eq!(line.shape_type(), ShapeType::Line);
        assert_eq!(connector.shape_type(), ShapeType::Connector);
    }

    #[test]
    fn source_bound_line_refuses_mutation() {
        let mut line = LineShape::new(1, 0, 0, 10, 10);
        let before_width = line.width;
        let before_color = line.color;
        line.mark_source_bound();

        assert_eq!(
            line.set_width(2),
            Err(crate::shapes::shape::MutationError::SourceBound {
                mutation: crate::shapes::shape::Mutation::Line,
            })
        );
        assert_eq!(
            line.set_color(0x0000_ff00),
            Err(crate::shapes::shape::MutationError::SourceBound {
                mutation: crate::shapes::shape::Mutation::Line,
            })
        );
        assert_eq!(line.width, before_width);
        assert_eq!(line.color, before_color);
    }

    #[test]
    fn source_bound_group_setters_are_atomic() {
        let mut group = GroupShape::new(1);
        let before_bounds = (group.left, group.top, group.width, group.height);
        let before_children = group.children.len();
        group.mark_source_bound();

        assert_eq!(
            group.set_bounds(1, 2, 3, 4),
            Err(crate::shapes::shape::MutationError::SourceBound {
                mutation: crate::shapes::shape::Mutation::Geometry,
            })
        );
        assert_eq!(
            group.add_child(ShapeEnum::Picture(PictureShape::new(2))),
            Err(crate::shapes::shape::MutationError::SourceBound {
                mutation: crate::shapes::shape::Mutation::Structure,
            })
        );
        assert_eq!(
            (group.left, group.top, group.width, group.height),
            before_bounds
        );
        assert_eq!(group.children.len(), before_children);
    }

    #[test]
    fn source_bound_shape_enum_picture_mut_refuses_typed_access() {
        let mut picture = PictureShape::new(1);
        picture.mark_source_bound();
        let before_name = picture.name.clone();
        let mut shape: ShapeEnum<'static> = ShapeEnum::Picture(picture);

        assert_eq!(
            shape
                .as_picture_mut()
                .map(|picture_ref| picture_ref.is_some()),
            Err(crate::shapes::shape::MutationError::SourceBound {
                mutation: crate::shapes::shape::Mutation::Picture,
            })
        );
        assert_eq!(
            shape.as_picture().and_then(PictureShape::name),
            before_name.as_deref()
        );
    }
}
