//! Drawing-page semantics.

use super::shape::Shape;

/// A semantic drawing page.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Page {
    name: Option<String>,
    shapes: Vec<Shape>,
}

impl Page {
    pub fn new(name: impl Into<String>) -> Self {
        Self {
            name: Some(name.into()),
            shapes: Vec::new(),
        }
    }

    pub(crate) fn parsed(name: Option<String>) -> Self {
        Self {
            name,
            shapes: Vec::new(),
        }
    }

    pub(crate) fn push_shape(&mut self, shape: Shape) {
        self.shapes.push(shape);
    }

    pub(crate) fn shape_mut(&mut self, index: usize) -> Option<&mut Shape> {
        self.shapes.get_mut(index)
    }

    /// Returns the optional `draw:name`.
    #[must_use]
    pub fn name(&self) -> Option<&str> {
        self.name.as_deref()
    }

    /// Bounded shapes in source order.
    #[must_use]
    pub fn shapes(&self) -> &[Shape] {
        &self.shapes
    }
}
