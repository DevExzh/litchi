/// Shapes module for PowerPoint presentations.
///
/// This module provides types for working with shapes on slides, including:
/// - Text shapes with text frames
/// - Pictures (images)
/// - Tables
/// - Placeholders
///
/// The design follows the python-pptx library structure.
pub mod base;
pub mod picture;
pub mod table;
pub mod textframe;

pub use base::{BaseShape, Shape, ShapeType};
pub use picture::Picture;
pub use table::{Table, TableCell, TableRow};
pub use textframe::TextFrame;
