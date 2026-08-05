//! High-performance Shape enum for representing all shape types.
//!
//! Idiomatic Rust implementation using enum variants instead of trait objects.

use super::picture::PictureShape;
use super::shape::{Shape, ShapeType};
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
    /// Get the shape type.
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

    /// Get shape as TextBox if it is one.
    #[inline]
    pub fn as_textbox(&self) -> Option<&TextBox<'a>> {
        match self {
            ShapeEnum::TextBox(tb) => Some(tb),
            _ => None,
        }
    }

    /// Get shape as Placeholder if it is one.
    #[inline]
    pub fn as_placeholder(&self) -> Option<&Placeholder<'a>> {
        match self {
            ShapeEnum::Placeholder(ph) => Some(ph),
            _ => None,
        }
    }

    /// Get shape as AutoShape if it is one.
    #[inline]
    pub fn as_autoshape(&self) -> Option<&AutoShape<'a>> {
        match self {
            ShapeEnum::AutoShape(as_) => Some(as_),
            _ => None,
        }
    }

    /// Get shape as PictureShape if it is one.
    #[inline]
    pub fn as_picture(&self) -> Option<&PictureShape> {
        match self {
            ShapeEnum::Picture(pic) => Some(pic),
            _ => None,
        }
    }

    /// Get shape as mutable PictureShape if it is one.
    #[inline]
    pub fn as_picture_mut(&mut self) -> Option<&mut PictureShape> {
        match self {
            ShapeEnum::Picture(pic) => Some(pic),
            _ => None,
        }
    }

    /// Get an embedded or linked OLE object frame.
    #[inline]
    pub fn as_object_frame(&self) -> Option<&PictureShape> {
        self.as_picture()
            .filter(|picture| picture.frame_kind() == super::picture::PictureFrameKind::OleObject)
    }

    /// Get an audio or video media frame.
    #[inline]
    pub fn as_media_frame(&self) -> Option<&PictureShape> {
        self.as_picture()
            .filter(|picture| picture.frame_kind() == super::picture::PictureFrameKind::Media)
    }

    /// Get shape as TableShape if it is one.
    #[inline]
    pub fn as_table(&self) -> Option<&TableShape> {
        match self {
            ShapeEnum::Table(table) => Some(table),
            _ => None,
        }
    }

    /// Get shape as GroupShape if it is one.
    #[inline]
    pub fn as_group(&self) -> Option<&GroupShape<'a>> {
        match self {
            ShapeEnum::Group(group) => Some(group),
            _ => None,
        }
    }

    /// Get shape as LineShape if it is one.
    #[inline]
    pub fn as_line(&self) -> Option<&LineShape> {
        match self {
            ShapeEnum::Line(line) => Some(line),
            _ => None,
        }
    }

    /// Return inert PowerPoint 12 placeholder metadata retained for round trips.
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
    /// Left coordinate
    left: i32,
    /// Top coordinate
    top: i32,
    /// Width
    width: i32,
    /// Height
    height: i32,
    powerpoint12_shape_metadata: Option<ShapeMetadata>,
}

impl TableShape {
    /// Create a new table shape.
    pub fn new(id: u32, rows: usize, columns: usize) -> Self {
        let cells = vec![vec![String::new(); columns]; rows];
        Self {
            id,
            rows,
            columns,
            cells,
            left: 0,
            top: 0,
            width: 0,
            height: 0,
            powerpoint12_shape_metadata: None,
        }
    }

    pub(crate) fn set_cell_text(&mut self, row: usize, col: usize, text: String) {
        if let Some(cell) = self.cells.get_mut(row).and_then(|cells| cells.get_mut(col)) {
            *cell = text;
        }
    }

    pub(crate) fn set_bounds(&mut self, left: i32, top: i32, width: i32, height: i32) {
        self.left = left;
        self.top = top;
        self.width = width;
        self.height = height;
    }

    /// Get shape ID.
    pub fn id(&self) -> u32 {
        self.id
    }

    /// Get number of rows.
    pub fn rows(&self) -> usize {
        self.rows
    }

    /// Get number of columns.
    pub fn columns(&self) -> usize {
        self.columns
    }

    /// Get cell text.
    pub fn cell(&self, row: usize, col: usize) -> Option<&str> {
        self.cells
            .get(row)
            .and_then(|r| r.get(col))
            .map(|s| s.as_str())
    }

    /// Get the left coordinate.
    pub fn left(&self) -> i32 {
        self.left
    }

    /// Get the top coordinate.
    pub fn top(&self) -> i32 {
        self.top
    }

    /// Get the table width.
    pub fn width(&self) -> i32 {
        self.width
    }

    /// Get the table height.
    pub fn height(&self) -> i32 {
        self.height
    }

    pub(crate) fn set_powerpoint12_shape_metadata(&mut self, metadata: Option<ShapeMetadata>) {
        self.powerpoint12_shape_metadata = metadata;
    }
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
}

impl<'a> GroupShape<'a> {
    /// Create a new group shape.
    pub fn new(id: u32) -> Self {
        Self {
            id,
            children: Vec::new(),
            left: 0,
            top: 0,
            width: 0,
            height: 0,
            powerpoint12_shape_metadata: None,
        }
    }

    /// Add a child shape.
    pub fn add_child(&mut self, shape: ShapeEnum<'a>) {
        self.children.push(shape);
    }

    /// Get child shapes.
    pub fn children(&self) -> &[ShapeEnum<'a>] {
        &self.children
    }

    /// Set group bounds.
    pub fn set_bounds(&mut self, left: i32, top: i32, width: i32, height: i32) {
        self.left = left;
        self.top = top;
        self.width = width;
        self.height = height;
    }

    /// Get shape ID.
    pub fn id(&self) -> u32 {
        self.id
    }

    /// Get the left coordinate.
    pub fn left(&self) -> i32 {
        self.left
    }

    /// Get the top coordinate.
    pub fn top(&self) -> i32 {
        self.top
    }

    /// Get the group width.
    pub fn width(&self) -> i32 {
        self.width
    }

    /// Get the group height.
    pub fn height(&self) -> i32 {
        self.height
    }

    pub(crate) fn set_powerpoint12_shape_metadata(&mut self, metadata: Option<ShapeMetadata>) {
        self.powerpoint12_shape_metadata = metadata;
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
}

#[derive(Debug, Clone, Copy)]
enum LineKind {
    Line,
    Connector,
}

impl LineShape {
    /// Create a new line shape.
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
        }
    }

    /// Create a connector between two points.
    pub fn connector(id: u32, x1: i32, y1: i32, x2: i32, y2: i32) -> Self {
        Self {
            kind: LineKind::Connector,
            ..Self::new(id, x1, y1, x2, y2)
        }
    }

    /// Return whether this shape is a plain line or a connector.
    pub fn shape_type(&self) -> ShapeType {
        match self.kind {
            LineKind::Line => ShapeType::Line,
            LineKind::Connector => ShapeType::Connector,
        }
    }

    /// Set line width.
    pub fn set_width(&mut self, width: i32) {
        self.width = width;
    }

    /// Set line color.
    pub fn set_color(&mut self, color: u32) {
        self.color = Some(color);
    }

    /// Get shape ID.
    pub fn id(&self) -> u32 {
        self.id
    }

    /// Get line length.
    pub fn length(&self) -> f64 {
        let dx = (self.x2 - self.x1) as f64;
        let dy = (self.y2 - self.y1) as f64;
        (dx * dx + dy * dy).sqrt()
    }

    pub(crate) fn set_powerpoint12_shape_metadata(&mut self, metadata: Option<ShapeMetadata>) {
        self.powerpoint12_shape_metadata = metadata;
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn line_family_preserves_connector_semantics() {
        let line: ShapeEnum<'static> = ShapeEnum::Line(LineShape::new(1, 0, 0, 10, 10));
        let connector: ShapeEnum<'static> = ShapeEnum::Line(LineShape::connector(2, 0, 0, 10, 10));

        assert_eq!(line.shape_type(), ShapeType::Line);
        assert_eq!(connector.shape_type(), ShapeType::Connector);
    }
}
