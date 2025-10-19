/// High-performance Shape enum for representing all shape types.
///
/// Idiomatic Rust implementation using enum variants instead of trait objects.

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
    pub fn text(&self) -> Result<String> {
        match self {
            ShapeEnum::TextBox(tb) => Shape::text(tb),
            ShapeEnum::Placeholder(ph) => Shape::text(ph),
            ShapeEnum::AutoShape(as_) => Shape::text(as_),
            _ => Ok(String::new()),
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

/// Picture/image shape (placeholder for future implementation).
#[derive(Debug, Clone)]
pub struct PictureShape {
    // TODO: Implement picture shape parsing
}

/// Table shape (placeholder for future implementation).
#[derive(Debug, Clone)]
pub struct TableShape {
    // TODO: Implement table shape parsing
}

/// Group shape containing other shapes (placeholder for future implementation).
#[derive(Debug, Clone)]
pub struct GroupShape {
    // TODO: Implement group shape parsing
}

/// Line/connector shape (placeholder for future implementation).
#[derive(Debug, Clone)]
pub struct LineShape {
    // TODO: Implement line shape parsing
}

