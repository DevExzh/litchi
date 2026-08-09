//! Semantic drawing-shape views.

use litchi_odf_common::drawing::Frame;

/// The recognized ODF drawing shape family.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum ShapeKind {
    Caption,
    Circle,
    Connector,
    /// Retained inertly; controls are never activated.
    Control,
    Custom,
    Ellipse,
    Frame,
    Group,
    Line,
    Measure,
    Path,
    Polygon,
    Polyline,
    Rectangle,
    RegularPolygon,
}

/// One bounded, inert shape view from `content.xml`.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Shape {
    name: Option<String>,
    layer: Option<String>,
    kind: ShapeKind,
    text: String,
    frame: Option<Frame>,
}

impl Shape {
    pub(crate) fn new(
        name: Option<String>,
        layer: Option<String>,
        kind: ShapeKind,
        frame: Option<Frame>,
    ) -> Self {
        Self {
            name,
            layer,
            kind,
            text: String::new(),
            frame,
        }
    }

    pub(crate) fn push_text(&mut self, text: &str) {
        self.text.push_str(text);
    }

    /// The optional `draw:name` selector.
    #[must_use]
    pub fn name(&self) -> Option<&str> {
        self.name.as_deref()
    }

    /// The optional `draw:layer` selector.
    #[must_use]
    pub fn layer(&self) -> Option<&str> {
        self.layer.as_deref()
    }

    /// The recognized shape family.
    #[must_use]
    pub const fn kind(&self) -> ShapeKind {
        self.kind
    }

    /// Paragraph character data without source rewriting.
    #[must_use]
    pub fn text(&self) -> &str {
        &self.text
    }

    /// Shared inert drawing occurrence context for `draw:frame`.
    #[must_use]
    pub fn frame(&self) -> Option<&Frame> {
        self.frame.as_ref()
    }
}
