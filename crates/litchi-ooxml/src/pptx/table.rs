//! DrawingML table content in PowerPoint graphic frames.
//!
//! Shape discovery lives in [`crate::pptx::shape`]. After matching
//! [`crate::pptx::shape::Shape::Table`], this module parses the borrowed frame
//! XML into its focused table model.

pub use super::shapes::table::{
    Table, TableCell as Cell, TableProperties as Props, TableRow as Row,
};
