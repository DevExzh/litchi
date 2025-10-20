//! High-performance Shape enum for representing all shape types.
//!
//! Idiomatic Rust implementation using enum variants instead of trait objects.

use super::{TextBox, Placeholder, AutoShape};
use super::shape::{ShapeType, Shape};
use crate::ole::ppt::package::Result;

/// Represents any shape on a slide using an enum for zero-cost abstraction.
///
/// # Performance
///
/// - Enum dispatch (no vtable overhead)
/// - Stack-allocated (no heap allocation for shape variants)
/// - Pattern matching compiles to efficient jump tables
#[derive(Debug, Clone)]
pub enum ShapeEnum {
    /// Text box shape containing editable text
    TextBox(TextBox),
    /// Placeholder shape (title, body, footer, etc.)
    Placeholder(Placeholder),
    /// Auto shape (rectangle, ellipse, arrow, etc.)
    AutoShape(AutoShape),
    /// Picture/image shape (not yet implemented)
    Picture(PictureShape),
    /// Table shape (not yet implemented)
    Table(TableShape),
    /// Group shape containing other shapes (not yet implemented)
    Group(GroupShape),
    /// Line/connector shape (not yet implemented)
    Line(LineShape),
}

impl ShapeEnum {
    /// Get the shape type.
    pub fn shape_type(&self) -> ShapeType {
        match self {
            ShapeEnum::TextBox(_) => ShapeType::TextBox,
            ShapeEnum::Placeholder(_) => ShapeType::Placeholder,
            ShapeEnum::AutoShape(_) => ShapeType::AutoShape,
            ShapeEnum::Picture(_) => ShapeType::Picture,
            ShapeEnum::Table(_) => ShapeType::Table,
            ShapeEnum::Group(_) => ShapeType::Group,
            ShapeEnum::Line(_) => ShapeType::Line,
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
                            && !cell_text.is_empty() {
                                text_parts.push(cell_text.to_string());
                            }
                    }
                }
                Ok(text_parts.join(" "))
            }
            ShapeEnum::Group(group) => {
                // Recursively extract text from all child shapes
                let mut text_parts = Vec::new();
                for child in group.children() {
                    if let Ok(child_text) = child.text()
                        && !child_text.is_empty() {
                            text_parts.push(child_text);
                        }
                }
                Ok(text_parts.join("\n"))
            }
            ShapeEnum::Picture(_) | ShapeEnum::Line(_) => Ok(String::new()),
        }
    }

    /// Get shape as TextBox if it is one.
    #[inline]
    pub fn as_textbox(&self) -> Option<&TextBox> {
        match self {
            ShapeEnum::TextBox(tb) => Some(tb),
            _ => None,
        }
    }

    /// Get shape as Placeholder if it is one.
    #[inline]
    pub fn as_placeholder(&self) -> Option<&Placeholder> {
        match self {
            ShapeEnum::Placeholder(ph) => Some(ph),
            _ => None,
        }
    }

    /// Get shape as AutoShape if it is one.
    #[inline]
    pub fn as_autoshape(&self) -> Option<&AutoShape> {
        match self {
            ShapeEnum::AutoShape(as_) => Some(as_),
            _ => None,
        }
    }
}

/// Picture/image shape.
///
/// Represents an embedded image (JPEG, PNG, EMF, WMF, etc.)
#[derive(Debug, Clone)]
pub struct PictureShape {
    /// Shape ID
    id: u32,
    /// Picture blip ID (reference to image data)
    blip_id: Option<u32>,
    /// Picture name/filename
    name: Option<String>,
    /// Left coordinate
    left: i32,
    /// Top coordinate
    top: i32,
    /// Width
    width: i32,
    /// Height
    height: i32,
}

impl PictureShape {
    /// Create a new picture shape.
    pub fn new(id: u32) -> Self {
        Self {
            id,
            blip_id: None,
            name: None,
            left: 0,
            top: 0,
            width: 0,
            height: 0,
        }
    }
    
    /// Set picture coordinates and size.
    pub fn set_bounds(&mut self, left: i32, top: i32, width: i32, height: i32) {
        self.left = left;
        self.top = top;
        self.width = width;
        self.height = height;
    }
    
    /// Set blip ID.
    pub fn set_blip_id(&mut self, blip_id: u32) {
        self.blip_id = Some(blip_id);
    }
    
    /// Set picture name.
    pub fn set_name(&mut self, name: String) {
        self.name = Some(name);
    }
    
    /// Get shape ID.
    pub fn id(&self) -> u32 {
        self.id
    }
    
    /// Get blip ID.
    pub fn blip_id(&self) -> Option<u32> {
        self.blip_id
    }
}

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
}

impl TableShape {
    /// Create a new table shape.
    pub fn new(id: u32, rows: usize, columns: usize) -> Self {
        let cells = vec![vec![String::new(); columns]; rows];
        Self { id, rows, columns, cells }
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
        self.cells.get(row).and_then(|r| r.get(col)).map(|s| s.as_str())
    }
}

/// Group shape containing other shapes.
///
/// Groups allow hierarchical organization of shapes.
#[derive(Debug, Clone)]
pub struct GroupShape {
    /// Shape ID
    id: u32,
    /// Child shapes
    children: Vec<ShapeEnum>,
    /// Left coordinate
    left: i32,
    /// Top coordinate
    top: i32,
    /// Width
    width: i32,
    /// Height
    height: i32,
}

impl GroupShape {
    /// Create a new group shape.
    pub fn new(id: u32) -> Self {
        Self {
            id,
            children: Vec::new(),
            left: 0,
            top: 0,
            width: 0,
            height: 0,
        }
    }
    
    /// Add a child shape.
    pub fn add_child(&mut self, shape: ShapeEnum) {
        self.children.push(shape);
    }
    
    /// Get child shapes.
    pub fn children(&self) -> &[ShapeEnum] {
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
}

/// Line/connector shape.
///
/// Represents a line or connector between two points.
#[derive(Debug, Clone)]
pub struct LineShape {
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
}

impl LineShape {
    /// Create a new line shape.
    pub fn new(id: u32, x1: i32, y1: i32, x2: i32, y2: i32) -> Self {
        Self {
            id,
            x1,
            y1,
            x2,
            y2,
            width: 1,
            color: None,
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
}

