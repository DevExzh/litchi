//! Slide and shape structures for ODP presentations.

use super::legacy_animation::Node as AnimationNode;
use super::{DrawingHyperlink, Node, Reference, ShapeEventListener, Transition};
use crate::action::validate_event_listeners;
use litchi_core::Result;

/// A slide in an ODP presentation.
///
/// Slides contain text content, optional titles, and shape elements.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Slide {
    /// Optional slide title
    pub title: Option<String>,
    /// Text content of the slide
    pub text: String,
    /// Slide index (0-based)
    pub index: usize,
    /// Optional notes for the slide
    pub notes: Option<String>,
    /// Optional slide transition and automatic-advance properties.
    pub transition: Option<Transition>,
    /// Inert ODF animation and timing trees attached to the slide.
    pub animations: Vec<Node>,
    /// Optional legacy `presentation:animations` effect tree.
    pub legacy_animation: Option<AnimationNode>,
    /// Shapes on the slide
    pub shapes: Vec<Shape>,
}

impl Slide {
    /// Get the title of the slide.
    ///
    /// Returns the slide title if present, None otherwise.
    ///
    /// # Errors
    /// Returns an error when the input is malformed or a configured limit is exceeded.
    pub fn title(&self) -> Result<Option<&str>> {
        Ok(self.title.as_deref())
    }

    /// Get the slide's primary body text.
    ///
    /// Use [`Self::all_text`] to include the title and labeled drawing shapes.
    ///
    /// # Errors
    /// Returns an error when the input is malformed or a configured limit is exceeded.
    pub fn text(&self) -> Result<&str> {
        Ok(&self.text)
    }

    /// Compose all visible text from the title, body, and labeled shapes.
    pub fn all_text(&self) -> String {
        let mut parts = Vec::with_capacity(self.shapes.len() + 2);
        if let Some(title) = self.title.as_deref().map(str::trim)
            && !title.is_empty()
        {
            parts.push(title);
        }
        let body = self.text.trim();
        if !body.is_empty() {
            parts.push(body);
        }
        let mut shapes: Vec<&Shape> = self.shapes.iter().rev().collect();
        while let Some(shape) = shapes.pop() {
            let text = shape.text.trim();
            if !text.is_empty() {
                parts.push(text);
            }
            shapes.extend(shape.children.iter().rev());
        }
        parts.join("\n")
    }

    /// Get all shapes on the slide.
    ///
    /// Returns a slice of shapes contained in this slide.
    ///
    /// # Errors
    /// Returns an error when the input is malformed or a configured limit is exceeded.
    pub fn shapes(&self) -> Result<&[Shape]> {
        Ok(&self.shapes)
    }

    /// Get the slide index.
    ///
    /// Returns the 0-based index of this slide in the presentation.
    #[must_use]
    pub fn index(&self) -> usize {
        self.index
    }

    /// Get the slide notes.
    ///
    /// Returns speaker notes for this slide if present, None otherwise.
    ///
    /// # Errors
    /// Returns an error when the input is malformed or a configured limit is exceeded.
    pub fn notes(&self) -> Result<Option<&str>> {
        Ok(self.notes.as_deref())
    }

    /// Get the slide transition configuration.
    #[must_use]
    pub fn transition(&self) -> Option<&Transition> {
        self.transition.as_ref()
    }

    /// Get or create the slide transition configuration.
    pub fn transition_mut(&mut self) -> &mut Transition {
        self.transition.get_or_insert_with(Transition::new)
    }

    /// Remove the slide transition configuration.
    pub fn clear_transition(&mut self) {
        self.transition = None;
    }

    /// Return the slide's inert animation and timing trees.
    #[must_use]
    pub fn animations(&self) -> &[Node] {
        &self.animations
    }

    /// Return mutable animation and timing trees.
    pub fn animations_mut(&mut self) -> &mut Vec<Node> {
        &mut self.animations
    }

    /// Add a schema-defined animation root to the slide.
    ///
    /// # Errors
    /// Returns an error when the input is malformed or a configured limit is exceeded.
    pub fn add_animation(&mut self, animation: Node) -> Result<()> {
        if !animation.kind().allowed_at_page_root() {
            return Err(litchi_core::Error::InvalidFormat(
                "anim:param is only valid below anim:command".to_string(),
            ));
        }
        self.animations.push(animation);
        Ok(())
    }

    /// Remove all animation and timing trees from the slide.
    pub fn clear_animations(&mut self) {
        self.animations.clear();
    }

    /// Return the optional legacy presentation effect tree.
    #[must_use]
    pub fn legacy_animation(&self) -> Option<&AnimationNode> {
        self.legacy_animation.as_ref()
    }

    /// Set or remove the legacy presentation effect tree.
    pub fn set_legacy_animation(&mut self, animation: Option<AnimationNode>) {
        self.legacy_animation = animation;
    }
}

/// A shape (element) on a slide.
///
/// Shapes represent visual elements like text boxes, images, and drawings.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DrawingShapeKind {
    /// `draw:frame`.
    Frame,
    /// `draw:rect`.
    Rectangle,
    /// `draw:line`.
    Line,
    /// `draw:polyline`.
    Polyline,
    /// `draw:polygon`.
    Polygon,
    /// `draw:regular-polygon`.
    RegularPolygon,
    /// `draw:path`.
    Path,
    /// `draw:circle`.
    Circle,
    /// `draw:ellipse`.
    Ellipse,
    /// `draw:g`.
    Group,
    /// `draw:page-thumbnail`.
    PageThumbnail,
    /// `draw:measure`.
    Measure,
    /// `draw:caption`.
    Caption,
    /// `draw:connector`.
    Connector,
    /// `draw:control`.
    Control,
    /// `draw:custom-shape`.
    CustomShape,
    /// `dr3d:scene`.
    ThreeDimensionalScene,
    /// `dr3d:light`.
    ThreeDimensionalLight,
    /// `dr3d:cube`.
    ThreeDimensionalCube,
    /// `dr3d:sphere`.
    ThreeDimensionalSphere,
    /// `dr3d:extrude`.
    ThreeDimensionalExtrude,
    /// `dr3d:rotate`.
    ThreeDimensionalRotate,
}

impl DrawingShapeKind {
    pub(crate) fn element_name(self) -> &'static str {
        match self {
            Self::Frame => "draw:frame",
            Self::Rectangle => "draw:rect",
            Self::Line => "draw:line",
            Self::Polyline => "draw:polyline",
            Self::Polygon => "draw:polygon",
            Self::RegularPolygon => "draw:regular-polygon",
            Self::Path => "draw:path",
            Self::Circle => "draw:circle",
            Self::Ellipse => "draw:ellipse",
            Self::Group => "draw:g",
            Self::PageThumbnail => "draw:page-thumbnail",
            Self::Measure => "draw:measure",
            Self::Caption => "draw:caption",
            Self::Connector => "draw:connector",
            Self::Control => "draw:control",
            Self::CustomShape => "draw:custom-shape",
            Self::ThreeDimensionalScene => "dr3d:scene",
            Self::ThreeDimensionalLight => "dr3d:light",
            Self::ThreeDimensionalCube => "dr3d:cube",
            Self::ThreeDimensionalSphere => "dr3d:sphere",
            Self::ThreeDimensionalExtrude => "dr3d:extrude",
            Self::ThreeDimensionalRotate => "dr3d:rotate",
        }
    }

    pub(crate) fn is_three_dimensional(self) -> bool {
        matches!(
            self,
            Self::ThreeDimensionalScene
                | Self::ThreeDimensionalLight
                | Self::ThreeDimensionalCube
                | Self::ThreeDimensionalSphere
                | Self::ThreeDimensionalExtrude
                | Self::ThreeDimensionalRotate
        )
    }
}

/// Namespace of an unmodeled drawing-shape attribute.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DrawingAttributeNamespace {
    /// `OpenDocument` drawing namespace (`draw:*`).
    Drawing,
    /// ODF SVG-compatible namespace (`svg:*`).
    Svg,
    /// `OpenDocument` 3D namespace (`dr3d:*`).
    Dr3d,
    /// `OpenDocument` table namespace (`table:*`), used by spreadsheet
    /// shape anchoring attributes such as `table:end-cell-address`.
    Table,
}

impl DrawingAttributeNamespace {
    pub(crate) fn prefix(self) -> &'static str {
        match self {
            Self::Drawing => "draw",
            Self::Svg => "svg",
            Self::Dr3d => "dr3d",
            Self::Table => "table",
        }
    }
}

/// An exact drawing or SVG attribute not represented by a dedicated shape field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DrawingAttribute {
    pub(crate) namespace: DrawingAttributeNamespace,
    pub(crate) local_name: String,
    pub(crate) value: String,
}

impl DrawingAttribute {
    /// Create a drawing attribute after validating its XML local name.
    ///
    /// # Errors
    /// Returns an error when the input is malformed or a configured limit is exceeded.
    pub fn new(
        namespace: DrawingAttributeNamespace,
        local_name: impl Into<String>,
        value: impl Into<String>,
    ) -> Result<Self> {
        let name = local_name.into();
        if !is_xml_local_name(&name) {
            return Err(litchi_core::Error::InvalidFormat(format!(
                "invalid drawing attribute local name '{name}'"
            )));
        }
        Ok(Self {
            namespace,
            local_name: name,
            value: value.into(),
        })
    }

    /// Return the attribute namespace.
    #[must_use]
    pub fn namespace(&self) -> DrawingAttributeNamespace {
        self.namespace
    }

    /// Return the attribute local name.
    #[must_use]
    pub fn local_name(&self) -> &str {
        &self.local_name
    }

    /// Return the decoded attribute value.
    #[must_use]
    pub fn value(&self) -> &str {
        &self.value
    }
}

/// Child kind within `draw:enhanced-geometry`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum EnhancedGeometryChildKind {
    /// Inert `draw:equation` formula declaration.
    Equation,
    /// Inert `draw:handle` adjustment handle.
    Handle,
}

impl EnhancedGeometryChildKind {
    pub(crate) fn element_name(self) -> &'static str {
        match self {
            Self::Equation => "draw:equation",
            Self::Handle => "draw:handle",
        }
    }
}

/// An inert equation or handle declaration in custom-shape geometry.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct EnhancedGeometryChild {
    pub(crate) kind: EnhancedGeometryChildKind,
    pub(crate) attributes: Vec<DrawingAttribute>,
}

impl EnhancedGeometryChild {
    /// Create an empty equation or handle declaration.
    #[must_use]
    pub fn new(kind: EnhancedGeometryChildKind) -> Self {
        Self {
            kind,
            attributes: Vec::new(),
        }
    }

    /// Return the child kind.
    #[must_use]
    pub fn kind(&self) -> EnhancedGeometryChildKind {
        self.kind
    }

    /// Return exact attributes in document order.
    #[must_use]
    pub fn attributes(&self) -> &[DrawingAttribute] {
        &self.attributes
    }

    /// Return mutable exact attributes.
    pub fn attributes_mut(&mut self) -> &mut Vec<DrawingAttribute> {
        &mut self.attributes
    }
}

/// Inert `draw:enhanced-geometry` data for a custom shape.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct EnhancedGeometry {
    pub(crate) attributes: Vec<DrawingAttribute>,
    pub(crate) children: Vec<EnhancedGeometryChild>,
}

impl EnhancedGeometry {
    /// Create an empty enhanced-geometry declaration.
    #[must_use]
    pub fn new() -> Self {
        Self::default()
    }

    /// Return exact geometry attributes in document order.
    #[must_use]
    pub fn attributes(&self) -> &[DrawingAttribute] {
        &self.attributes
    }

    /// Return mutable exact geometry attributes.
    pub fn attributes_mut(&mut self) -> &mut Vec<DrawingAttribute> {
        &mut self.attributes
    }

    /// Return inert equations and handles in document order.
    #[must_use]
    pub fn children(&self) -> &[EnhancedGeometryChild] {
        &self.children
    }

    /// Return mutable inert equations and handles.
    pub fn children_mut(&mut self) -> &mut Vec<EnhancedGeometryChild> {
        &mut self.children
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Shape {
    /// Shape type (text box, image, frame, etc.)
    pub shape_type: litchi_core::ShapeType,
    /// Exact ODF drawing element kind for parsed shapes.
    pub drawing_kind: Option<DrawingShapeKind>,
    /// Exact unmodeled `draw:*` and `svg:*` attributes.
    pub drawing_attributes: Vec<DrawingAttribute>,
    /// Nested shapes when this is a `draw:g` group.
    pub children: Vec<Shape>,
    /// Inert enhanced geometry for `draw:custom-shape`.
    pub enhanced_geometry: Option<EnhancedGeometry>,
    /// Text content if the shape contains text
    pub text: String,
    /// Shape name/ID
    pub name: Option<String>,
    /// X position (in presentation units)
    pub x: Option<String>,
    /// Y position (in presentation units)
    pub y: Option<String>,
    /// Width (in presentation units)
    pub width: Option<String>,
    /// Height (in presentation units)
    pub height: Option<String>,
    /// Style name reference
    pub style_name: Option<String>,
    /// Drawing layer containing this shape.
    pub layer: Option<String>,
    /// Exact `draw:z-index` lexical value.
    ///
    /// ODF uses the unbounded XML Schema `nonNegativeInteger` type here, so a
    /// string avoids truncating valid values that exceed Rust integer widths.
    pub z_index: Option<String>,
    /// Optional ODF drawing transformation list.
    pub transform: Option<String>,
    /// Exact presentation role, such as `object`, `subtitle`, or `chart`.
    pub presentation_class: Option<String>,
    /// Whether this shape is a presentation placeholder.
    pub presentation_placeholder: Option<bool>,
    /// Whether a presentation placeholder was transformed by the user.
    pub presentation_user_transformed: Option<bool>,
    /// Image source referenced by `draw:image`, when this is a picture shape.
    pub image_href: Option<String>,
    /// Inert audio/video plugin referenced by this frame.
    pub media: Option<Reference>,
    /// Optional hyperlink wrapping this shape.
    pub hyperlink: Option<DrawingHyperlink>,
    /// Inert event listeners attached directly to this shape.
    pub event_listeners: Vec<ShapeEventListener>,
}

impl Shape {
    /// Create a new empty shape
    #[must_use]
    pub fn new() -> Self {
        Self {
            shape_type: litchi_core::ShapeType::AutoShape,
            drawing_kind: None,
            drawing_attributes: Vec::new(),
            children: Vec::new(),
            enhanced_geometry: None,
            text: String::new(),
            name: None,
            x: None,
            y: None,
            width: None,
            height: None,
            style_name: None,
            layer: None,
            z_index: None,
            transform: None,
            presentation_class: None,
            presentation_placeholder: None,
            presentation_user_transformed: None,
            image_href: None,
            media: None,
            hyperlink: None,
            event_listeners: Vec::new(),
        }
    }

    /// Get the text content of the shape.
    ///
    /// # Errors
    /// Returns an error when the input is malformed or a configured limit is exceeded.
    pub fn text(&self) -> Result<&str> {
        Ok(&self.text)
    }

    /// Get the shape type.
    #[must_use]
    pub fn shape_type(&self) -> litchi_core::ShapeType {
        self.shape_type
    }

    /// Get the exact ODF drawing element kind, when parsed or explicitly set.
    #[must_use]
    pub fn drawing_kind(&self) -> Option<DrawingShapeKind> {
        self.drawing_kind
    }

    /// Return unmodeled drawing and SVG attributes in source order.
    #[must_use]
    pub fn drawing_attributes(&self) -> &[DrawingAttribute] {
        &self.drawing_attributes
    }

    /// Return nested group shapes in document order.
    #[must_use]
    pub fn children(&self) -> &[Shape] {
        &self.children
    }

    /// Return mutable nested group shapes.
    pub fn children_mut(&mut self) -> &mut Vec<Shape> {
        &mut self.children
    }

    /// Return inert custom-shape enhanced geometry.
    #[must_use]
    pub fn enhanced_geometry(&self) -> Option<&EnhancedGeometry> {
        self.enhanced_geometry.as_ref()
    }

    /// Return mutable inert custom-shape enhanced geometry.
    pub fn enhanced_geometry_mut(&mut self) -> Option<&mut EnhancedGeometry> {
        self.enhanced_geometry.as_mut()
    }

    /// Check if this is a text shape.
    #[must_use]
    pub fn has_text(&self) -> bool {
        !self.text.trim().is_empty()
    }

    /// Get the shape name/ID.
    #[must_use]
    pub fn name(&self) -> Option<&str> {
        self.name.as_deref()
    }

    /// Get the shape position as (x, y).
    #[must_use]
    pub fn position(&self) -> (Option<&str>, Option<&str>) {
        (self.x.as_deref(), self.y.as_deref())
    }

    /// Get the shape dimensions as (width, height).
    #[must_use]
    pub fn dimensions(&self) -> (Option<&str>, Option<&str>) {
        (self.width.as_deref(), self.height.as_deref())
    }

    /// Get the drawing layer containing this shape.
    #[must_use]
    pub fn layer(&self) -> Option<&str> {
        self.layer.as_deref()
    }

    /// Get the exact non-negative stacking index.
    #[must_use]
    pub fn z_index(&self) -> Option<&str> {
        self.z_index.as_deref()
    }

    /// Set or remove the exact non-negative stacking index.
    ///
    /// # Errors
    /// Returns an error when the input is malformed or a configured limit is exceeded.
    pub fn set_z_index(&mut self, z_index: Option<String>) -> Result<()> {
        if let Some(value) = &z_index {
            validate_z_index(value)?;
        }
        self.z_index = z_index;
        Ok(())
    }

    /// Get the ODF drawing transformation list.
    #[must_use]
    pub fn transform(&self) -> Option<&str> {
        self.transform.as_deref()
    }

    /// Get the exact presentation role for this shape.
    #[must_use]
    pub fn presentation_class(&self) -> Option<&str> {
        self.presentation_class.as_deref()
    }

    /// Get the image source referenced by this shape.
    #[must_use]
    pub fn image_href(&self) -> Option<&str> {
        self.image_href.as_deref()
    }

    /// Return the inert audio/video plugin referenced by this shape.
    #[must_use]
    pub fn media(&self) -> Option<&Reference> {
        self.media.as_ref()
    }

    /// Attach an inert audio/video plugin and mark this shape as a graphic frame.
    #[must_use]
    pub fn with_media(mut self, media: Reference) -> Self {
        self.shape_type = litchi_core::ShapeType::GraphicFrame;
        self.image_href = None;
        self.media = Some(media);
        self
    }

    /// Return the optional hyperlink wrapping this shape.
    #[must_use]
    pub fn hyperlink(&self) -> Option<&DrawingHyperlink> {
        self.hyperlink.as_ref()
    }

    /// Attach or remove a hyperlink wrapper.
    pub fn set_hyperlink(&mut self, hyperlink: Option<DrawingHyperlink>) {
        self.hyperlink = hyperlink;
    }

    /// Return inert event bindings attached to the shape.
    #[must_use]
    pub fn event_listeners(&self) -> &[ShapeEventListener] {
        &self.event_listeners
    }

    /// Return mutable inert event bindings.
    pub fn event_listeners_mut(&mut self) -> &mut Vec<ShapeEventListener> {
        &mut self.event_listeners
    }

    /// Add an inert event binding, subject to the per-shape safety limit.
    ///
    /// # Errors
    /// Returns an error when the input is malformed or a configured limit is exceeded.
    pub fn add_event_listener(&mut self, listener: ShapeEventListener) -> Result<()> {
        if self.event_listeners.len() >= 4096 {
            return Err(litchi_core::Error::InvalidFormat(
                "ODP shape exceeds 4096 event listeners".to_string(),
            ));
        }
        validate_event_listeners(std::slice::from_ref(&listener))?;
        self.event_listeners.push(listener);
        Ok(())
    }

    /// Set the image source and mark this shape as a picture.
    #[must_use]
    pub fn with_image_href(mut self, href: impl Into<String>) -> Self {
        self.shape_type = litchi_core::ShapeType::Picture;
        self.image_href = Some(href.into());
        self.media = None;
        self
    }
}

impl Default for Shape {
    fn default() -> Self {
        Self::new()
    }
}

fn is_xml_local_name(value: &str) -> bool {
    let mut characters = value.chars();
    characters
        .next()
        .is_some_and(|character| character == '_' || character.is_ascii_alphabetic())
        && characters.all(|character| {
            character == '_'
                || character == '-'
                || character == '.'
                || character.is_ascii_alphanumeric()
        })
}

pub(crate) fn validate_z_index(value: &str) -> Result<()> {
    let digits = value
        .strip_prefix('+')
        .or_else(|| value.strip_prefix('-'))
        .unwrap_or(value);
    let negative_nonzero = value.starts_with('-') && digits.bytes().any(|byte| byte != b'0');
    if digits.is_empty() || !digits.bytes().all(|byte| byte.is_ascii_digit()) || negative_nonzero {
        return Err(litchi_core::Error::InvalidFormat(format!(
            "invalid draw:z-index value '{value}'"
        )));
    }
    Ok(())
}

#[cfg(test)]
#[allow(
    clippy::default_trait_access,
    clippy::unwrap_used,
    reason = "test assertions panic on failure by design"
)]
mod tests {
    use super::*;

    #[test]
    fn test_slide_new() {
        let slide = Slide {
            title: None,
            text: String::new(),
            index: 0,
            notes: None,
            transition: None,
            animations: vec![],
            legacy_animation: None,
            shapes: vec![],
        };
        assert!(slide.title.is_none());
        assert!(slide.text.is_empty());
    }

    #[test]
    fn test_slide_with_content() {
        let slide = Slide {
            title: Some("Test Slide".to_string()),
            text: "Slide content".to_string(),
            index: 5,
            notes: Some("Speaker notes".to_string()),
            transition: None,
            animations: vec![],
            legacy_animation: None,
            shapes: vec![],
        };
        assert_eq!(slide.title().unwrap(), Some("Test Slide"));
        assert_eq!(slide.text().unwrap(), "Slide content");
        assert_eq!(slide.index(), 5);
        assert_eq!(slide.notes().unwrap(), Some("Speaker notes"));
    }

    #[test]
    fn test_slide_title_method() {
        let slide = Slide {
            title: Some("Title".to_string()),
            text: String::new(),
            index: 0,
            notes: None,
            transition: None,
            animations: vec![],
            legacy_animation: None,
            shapes: vec![],
        };
        assert_eq!(slide.title().unwrap(), Some("Title"));
    }

    #[test]
    fn test_slide_title_none() {
        let slide = Slide {
            title: None,
            text: String::new(),
            index: 0,
            notes: None,
            transition: None,
            animations: vec![],
            legacy_animation: None,
            shapes: vec![],
        };
        assert_eq!(slide.title().unwrap(), None);
    }

    #[test]
    fn test_slide_text_method() {
        let slide = Slide {
            title: None,
            text: "Hello World".to_string(),
            index: 0,
            notes: None,
            transition: None,
            animations: vec![],
            legacy_animation: None,
            shapes: vec![],
        };
        assert_eq!(slide.text().unwrap(), "Hello World");
    }

    #[test]
    fn test_slide_shapes_method() {
        let shapes = vec![Shape {
            shape_type: litchi_core::ShapeType::TextBox,
            text: "Shape 1".to_string(),
            name: Some("Shape1".to_string()),
            x: Some("0cm".to_string()),
            y: Some("0cm".to_string()),
            width: Some("5cm".to_string()),
            height: Some("3cm".to_string()),
            style_name: None,
            ..Shape::new()
        }];
        let slide = Slide {
            title: None,
            text: String::new(),
            index: 0,
            notes: None,
            transition: None,
            animations: vec![],
            legacy_animation: None,
            shapes,
        };
        assert_eq!(slide.shapes().unwrap().len(), 1);
    }

    #[test]
    fn test_slide_index_method() {
        let slide = Slide {
            title: None,
            text: String::new(),
            index: 42,
            notes: None,
            transition: None,
            animations: vec![],
            legacy_animation: None,
            shapes: vec![],
        };
        assert_eq!(slide.index(), 42);
    }

    #[test]
    fn test_slide_notes_method() {
        let slide = Slide {
            title: None,
            text: String::new(),
            index: 0,
            notes: Some("Notes".to_string()),
            transition: None,
            animations: vec![],
            legacy_animation: None,
            shapes: vec![],
        };
        assert_eq!(slide.notes().unwrap(), Some("Notes"));
    }

    #[test]
    fn test_slide_notes_none() {
        let slide = Slide {
            title: None,
            text: String::new(),
            index: 0,
            notes: None,
            transition: None,
            animations: vec![],
            legacy_animation: None,
            shapes: vec![],
        };
        assert_eq!(slide.notes().unwrap(), None);
    }

    #[test]
    fn test_slide_clone() {
        let slide = Slide {
            title: Some("Title".to_string()),
            text: "Content".to_string(),
            index: 1,
            notes: Some("Notes".to_string()),
            transition: None,
            animations: vec![],
            legacy_animation: None,
            shapes: vec![],
        };
        let cloned = slide.clone();
        assert_eq!(slide.title, cloned.title);
        assert_eq!(slide.text, cloned.text);
        assert_eq!(slide.index, cloned.index);
    }

    #[test]
    fn test_shape_new() {
        let shape = Shape::new();
        assert_eq!(shape.shape_type, litchi_core::ShapeType::AutoShape);
        assert!(shape.text.is_empty());
        assert!(shape.name.is_none());
        assert!(shape.x.is_none());
        assert!(shape.y.is_none());
        assert!(shape.width.is_none());
        assert!(shape.height.is_none());
        assert!(shape.style_name.is_none());
    }

    #[test]
    fn test_shape_default() {
        let shape: Shape = Default::default();
        assert_eq!(shape.shape_type, litchi_core::ShapeType::AutoShape);
        assert!(shape.text.is_empty());
    }

    #[test]
    fn test_shape_text_method() {
        let shape = Shape {
            shape_type: litchi_core::ShapeType::TextBox,
            text: "Hello".to_string(),
            name: None,
            x: None,
            y: None,
            width: None,
            height: None,
            style_name: None,
            ..Shape::new()
        };
        assert_eq!(shape.text().unwrap(), "Hello");
    }

    #[test]
    fn test_shape_shape_type_method() {
        let shape = Shape {
            shape_type: litchi_core::ShapeType::Picture,
            text: String::new(),
            name: None,
            x: None,
            y: None,
            width: None,
            height: None,
            style_name: None,
            ..Shape::new()
        };
        assert_eq!(shape.shape_type(), litchi_core::ShapeType::Picture);
    }

    #[test]
    fn test_shape_has_text() {
        let mut shape = Shape::new();
        assert!(!shape.has_text());

        shape.text = "   ".to_string();
        assert!(!shape.has_text());

        shape.text = "Hello".to_string();
        assert!(shape.has_text());
    }

    #[test]
    fn test_shape_name_method() {
        let shape = Shape {
            shape_type: litchi_core::ShapeType::TextBox,
            text: String::new(),
            name: Some("MyShape".to_string()),
            x: None,
            y: None,
            width: None,
            height: None,
            style_name: None,
            ..Shape::new()
        };
        assert_eq!(shape.name(), Some("MyShape"));
    }

    #[test]
    fn test_shape_name_none() {
        let shape = Shape {
            shape_type: litchi_core::ShapeType::TextBox,
            text: String::new(),
            name: None,
            x: None,
            y: None,
            width: None,
            height: None,
            style_name: None,
            ..Shape::new()
        };
        assert_eq!(shape.name(), None);
    }

    #[test]
    fn test_shape_position() {
        let shape = Shape {
            shape_type: litchi_core::ShapeType::TextBox,
            text: String::new(),
            name: None,
            x: Some("10cm".to_string()),
            y: Some("5cm".to_string()),
            width: None,
            height: None,
            style_name: None,
            ..Shape::new()
        };
        let (x, y) = shape.position();
        assert_eq!(x, Some("10cm"));
        assert_eq!(y, Some("5cm"));
    }

    #[test]
    fn test_shape_dimensions() {
        let shape = Shape {
            shape_type: litchi_core::ShapeType::TextBox,
            text: String::new(),
            name: None,
            x: None,
            y: None,
            width: Some("20cm".to_string()),
            height: Some("15cm".to_string()),
            style_name: None,
            ..Shape::new()
        };
        let (w, h) = shape.dimensions();
        assert_eq!(w, Some("20cm"));
        assert_eq!(h, Some("15cm"));
    }

    #[test]
    fn test_shape_clone() {
        let shape = Shape {
            shape_type: litchi_core::ShapeType::Placeholder,
            text: "Content".to_string(),
            name: Some("Shape1".to_string()),
            x: Some("1cm".to_string()),
            y: Some("2cm".to_string()),
            width: Some("10cm".to_string()),
            height: Some("5cm".to_string()),
            style_name: Some("Style1".to_string()),
            ..Shape::new()
        };
        let cloned = shape.clone();
        assert_eq!(shape.shape_type, cloned.shape_type);
        assert_eq!(shape.text, cloned.text);
        assert_eq!(shape.name, cloned.name);
    }

    #[test]
    fn test_shape_debug() {
        let shape = Shape::new();
        let debug_str = format!("{shape:?}");
        assert!(debug_str.contains("Shape"));
    }
}
