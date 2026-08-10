//! Semantic drawing-shape views.

use litchi_core::Position;
use litchi_odf_common::drawing::Frame;
use std::borrow::Cow;

/// Selector for a shape on one drawing page.
#[derive(Clone, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum Selector<'a> {
    /// The unique shape with this exact `draw:name`.
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

/// The recognized ODF drawing shape family.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
#[non_exhaustive]
#[allow(
    clippy::module_name_repetitions,
    reason = "public API compatibility uses ShapeKind"
)]
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
    PageThumbnail,
    Polygon,
    Polyline,
    Rectangle,
    RegularPolygon,
}

impl ShapeKind {
    pub(crate) const fn element_name(self) -> &'static str {
        match self {
            Self::Caption => "caption",
            Self::Circle => "circle",
            Self::Connector => "connector",
            Self::Control => "control",
            Self::Custom => "custom-shape",
            Self::Ellipse => "ellipse",
            Self::Frame => "frame",
            Self::Group => "g",
            Self::Line => "line",
            Self::Measure => "measure",
            Self::Path => "path",
            Self::PageThumbnail => "page-thumbnail",
            Self::Polygon => "polygon",
            Self::Polyline => "polyline",
            Self::Rectangle => "rect",
            Self::RegularPolygon => "regular-polygon",
        }
    }
}

/// One bounded, inert shape view from `content.xml`.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Shape {
    control_reference: Option<String>,
    name: Option<String>,
    layer: Option<String>,
    style_name: Option<String>,
    text_style_name: Option<String>,
    z_index: Option<u32>,
    x: Option<String>,
    y: Option<String>,
    width: Option<String>,
    height: Option<String>,
    path_data: Option<String>,
    transform: Option<String>,
    points: Option<String>,
    view_box: Option<String>,
    line_geometry: [Option<String>; 4],
    title: Option<String>,
    description: Option<String>,
    kind: ShapeKind,
    text: String,
    frame: Option<Frame>,
}

pub(crate) struct Properties {
    pub(crate) control_reference: Option<String>,
    pub(crate) geometry: [Option<String>; 4],
    pub(crate) layer: Option<String>,
    pub(crate) name: Option<String>,
    pub(crate) path_data: Option<String>,
    pub(crate) transform: Option<String>,
    pub(crate) points: Option<String>,
    pub(crate) view_box: Option<String>,
    pub(crate) line_geometry: [Option<String>; 4],
    pub(crate) style_name: Option<String>,
    pub(crate) text_style_name: Option<String>,
    pub(crate) z_index: Option<u32>,
}

impl Shape {
    /// Creates a detached shape value for structural insertion.
    #[must_use]
    pub fn new(kind: ShapeKind) -> Self {
        Self::parsed(
            Properties {
                control_reference: None,
                geometry: [None, None, None, None],
                layer: None,
                name: None,
                path_data: None,
                transform: None,
                points: None,
                view_box: None,
                line_geometry: [None, None, None, None],
                style_name: None,
                text_style_name: None,
                z_index: None,
            },
            kind,
            None,
        )
    }

    pub(crate) fn parsed(properties: Properties, kind: ShapeKind, frame: Option<Frame>) -> Self {
        let Properties {
            control_reference,
            geometry,
            layer,
            name,
            path_data,
            transform,
            points,
            view_box,
            line_geometry,
            style_name,
            text_style_name,
            z_index,
        } = properties;
        let [x, y, width, height] = geometry;
        Self {
            control_reference,
            name,
            layer,
            style_name,
            text_style_name,
            z_index,
            x,
            y,
            width,
            height,
            path_data,
            transform,
            points,
            view_box,
            line_geometry,
            title: None,
            description: None,
            kind,
            text: String::new(),
            frame,
        }
    }

    /// Sets the optional reference name on a detached shape.
    #[must_use]
    pub fn with_name(mut self, name: impl Into<String>) -> Self {
        self.name = Some(name.into());
        self
    }

    /// Assigns a declared drawing layer on a detached shape.
    #[must_use]
    pub fn with_layer(mut self, layer: impl Into<String>) -> Self {
        self.layer = Some(layer.into());
        self
    }

    /// Sets the graphic style reference on a detached shape.
    #[must_use]
    pub fn with_style_name(mut self, style_name: impl Into<String>) -> Self {
        self.style_name = Some(style_name.into());
        self
    }

    /// Sets the paragraph style reference on a detached shape.
    #[must_use]
    pub fn with_text_style_name(mut self, style_name: impl Into<String>) -> Self {
        self.text_style_name = Some(style_name.into());
        self
    }

    /// Sets the explicit stacking position on a detached shape.
    #[must_use]
    pub const fn with_z_index(mut self, z_index: u32) -> Self {
        self.z_index = Some(z_index);
        self
    }

    /// Sets all four lexical SVG geometry values on a detached shape.
    #[must_use]
    pub fn with_geometry(
        mut self,
        x: impl Into<String>,
        y: impl Into<String>,
        width: impl Into<String>,
        height: impl Into<String>,
    ) -> Self {
        self.x = Some(x.into());
        self.y = Some(y.into());
        self.width = Some(width.into());
        self.height = Some(height.into());
        self
    }

    /// Sets lexical SVG path data on a detached path shape.
    #[must_use]
    pub fn with_path_data(mut self, path_data: impl Into<String>) -> Self {
        self.path_data = Some(path_data.into());
        self
    }

    /// Sets a lexical `draw:transform` on a detached shape.
    #[must_use]
    pub fn with_transform(mut self, transform: impl Into<String>) -> Self {
        self.transform = Some(transform.into());
        self
    }

    /// Sets a lexical polygon/polyline view box and point list.
    #[must_use]
    pub fn with_points(mut self, view_box: impl Into<String>, points: impl Into<String>) -> Self {
        self.view_box = Some(view_box.into());
        self.points = Some(points.into());
        self
    }

    /// Sets lexical line endpoints (`svg:x1`, `y1`, `x2`, `y2`).
    #[must_use]
    pub fn with_line_geometry(
        mut self,
        x1: impl Into<String>,
        y1: impl Into<String>,
        x2: impl Into<String>,
        y2: impl Into<String>,
    ) -> Self {
        self.line_geometry = [
            Some(x1.into()),
            Some(y1.into()),
            Some(x2.into()),
            Some(y2.into()),
        ];
        self
    }

    /// Sets the inert form-control reference on a detached control shape.
    #[must_use]
    pub fn with_control_reference(mut self, reference: impl Into<String>) -> Self {
        self.control_reference = Some(reference.into());
        self
    }

    /// Sets plain paragraph text on a detached shape.
    #[must_use]
    pub fn with_text(mut self, text: impl Into<String>) -> Self {
        self.text = text.into();
        self
    }

    /// Sets an accessibility title on a detached shape.
    #[must_use]
    pub fn with_title(mut self, title: impl Into<String>) -> Self {
        self.title = Some(title.into());
        self
    }

    /// Sets an accessibility description on a detached shape.
    #[must_use]
    pub fn with_description(mut self, description: impl Into<String>) -> Self {
        self.description = Some(description.into());
        self
    }

    pub(crate) fn push_text(&mut self, text: &str) {
        self.text.push_str(text);
    }

    pub(crate) fn push_title(&mut self, text: &str) {
        self.title.get_or_insert_with(String::new).push_str(text);
        if let Some(frame) = &mut self.frame {
            frame.title.get_or_insert_with(String::new).push_str(text);
        }
    }

    pub(crate) fn push_description(&mut self, text: &str) {
        self.description
            .get_or_insert_with(String::new)
            .push_str(text);
        if let Some(frame) = &mut self.frame {
            frame
                .description
                .get_or_insert_with(String::new)
                .push_str(text);
        }
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

    /// The optional graphic style reference.
    #[must_use]
    pub fn style_name(&self) -> Option<&str> {
        self.style_name.as_deref()
    }

    /// The optional paragraph style used by shape text.
    #[must_use]
    pub fn text_style_name(&self) -> Option<&str> {
        self.text_style_name.as_deref()
    }

    /// The explicit drawing stacking position.
    #[must_use]
    pub const fn z_index(&self) -> Option<u32> {
        self.z_index
    }

    /// The optional lexical horizontal position.
    #[must_use]
    pub fn x(&self) -> Option<&str> {
        self.x.as_deref()
    }

    /// The optional lexical vertical position.
    #[must_use]
    pub fn y(&self) -> Option<&str> {
        self.y.as_deref()
    }

    /// The optional lexical width.
    #[must_use]
    pub fn width(&self) -> Option<&str> {
        self.width.as_deref()
    }

    /// The optional lexical height.
    #[must_use]
    pub fn height(&self) -> Option<&str> {
        self.height.as_deref()
    }

    /// The optional lexical `svg:d` path data.
    #[must_use]
    pub fn path_data(&self) -> Option<&str> {
        self.path_data.as_deref()
    }

    /// Optional lexical `draw:transform`.
    #[must_use]
    pub fn transform(&self) -> Option<&str> {
        self.transform.as_deref()
    }

    /// Optional lexical `draw:points` polygon/polyline point list.
    #[must_use]
    pub fn points(&self) -> Option<&str> {
        self.points.as_deref()
    }

    /// Optional lexical `svg:viewBox` paired with polygon/polyline points.
    #[must_use]
    pub fn view_box(&self) -> Option<&str> {
        self.view_box.as_deref()
    }

    /// Optional lexical line endpoints in `x1`, `y1`, `x2`, `y2` order.
    #[must_use]
    pub fn line_geometry(&self) -> [Option<&str>; 4] {
        self.line_geometry.each_ref().map(|value| value.as_deref())
    }

    /// The optional inert `draw:control` form-control reference.
    #[must_use]
    pub fn control_reference(&self) -> Option<&str> {
        self.control_reference.as_deref()
    }

    /// The direct accessibility title.
    #[must_use]
    pub fn title(&self) -> Option<&str> {
        self.title.as_deref()
    }

    /// The direct accessibility description.
    #[must_use]
    pub fn description(&self) -> Option<&str> {
        self.description.as_deref()
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
