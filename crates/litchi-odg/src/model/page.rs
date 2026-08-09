//! Drawing-page semantics.

use super::{layer::Layer, shape::Shape};
use litchi_core::Position;
use std::borrow::Cow;

/// Selector for a page in one immutable drawing snapshot.
#[derive(Clone, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum Selector<'a> {
    /// The unique page with this exact `draw:name`.
    Name(Cow<'a, str>),
    /// A checked zero-based position in source order.
    Position(Position),
}

impl<'a> From<&'a str> for Selector<'a> {
    fn from(value: &'a str) -> Self {
        Self::Name(Cow::Borrowed(value))
    }
}

impl From<Position> for Selector<'_> {
    fn from(value: Position) -> Self {
        Self::Position(value)
    }
}

impl From<usize> for Selector<'_> {
    fn from(value: usize) -> Self {
        Self::Position(Position::new(value))
    }
}

/// A semantic drawing page.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Page {
    name: Option<String>,
    xml_id: Option<String>,
    style_name: Option<String>,
    master_page_name: Option<String>,
    layer_set: bool,
    layers: Vec<Layer>,
    shapes: Vec<Shape>,
}

impl Page {
    pub fn new(name: impl Into<String>) -> Self {
        Self {
            name: Some(name.into()),
            xml_id: None,
            style_name: None,
            master_page_name: None,
            layer_set: false,
            layers: Vec::new(),
            shapes: Vec::new(),
        }
    }

    pub(crate) fn parsed(
        name: Option<String>,
        xml_id: Option<String>,
        style_name: Option<String>,
        master_page_name: Option<String>,
    ) -> Self {
        Self {
            name,
            xml_id,
            style_name,
            master_page_name,
            layer_set: false,
            layers: Vec::new(),
            shapes: Vec::new(),
        }
    }

    pub(crate) fn push_layer(&mut self, layer: Layer) {
        self.layers.push(layer);
    }

    pub(crate) fn mark_layer_set(&mut self) {
        self.layer_set = true;
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

    /// Returns the optional XML identity.
    #[must_use]
    pub fn xml_id(&self) -> Option<&str> {
        self.xml_id.as_deref()
    }

    /// Returns the optional drawing-page style name.
    #[must_use]
    pub fn style_name(&self) -> Option<&str> {
        self.style_name.as_deref()
    }

    /// Returns the optional master-page name.
    #[must_use]
    pub fn master_page_name(&self) -> Option<&str> {
        self.master_page_name.as_deref()
    }

    /// Returns page-local layer declarations in source order.
    #[must_use]
    pub fn layers(&self) -> &[Layer] {
        &self.layers
    }

    /// Returns whether this page explicitly owns a local layer set.
    #[must_use]
    pub const fn has_layer_set(&self) -> bool {
        self.layer_set
    }

    /// Bounded shapes in source order.
    #[must_use]
    pub fn shapes(&self) -> &[Shape] {
        &self.shapes
    }

    /// Selects one shape by exact name or checked position.
    ///
    /// # Errors
    ///
    /// Returns an error when an exact name is ambiguous.
    pub fn shape<'selector>(
        &self,
        selector: impl Into<super::shape::Selector<'selector>>,
    ) -> litchi_core::Result<Option<&Shape>> {
        let resolved_selector = selector.into();
        match resolved_selector {
            super::shape::Selector::Position(position) => Ok(self.shapes.get(position.get())),
            super::shape::Selector::Name(name) => select_unique_shape(&self.shapes, &name),
        }
    }
}

fn select_unique_shape<'a>(
    shapes: &'a [Shape],
    name: &str,
) -> litchi_core::Result<Option<&'a Shape>> {
    let mut matches = shapes.iter().filter(|shape| shape.name() == Some(name));
    let selected = matches.next();
    if selected.is_some() && matches.next().is_some() {
        return Err(litchi_core::Error::InvalidFormat(
            "ODG shape name selector is ambiguous".to_string(),
        ));
    }
    Ok(selected)
}
