//! Slide and shape structures for ODP presentations.

use super::{AnimationNode, LegacyAnimationNode, MediaReference, SlideTransition};
use litchi_core::Result;

/// A slide in an ODP presentation.
///
/// Slides contain text content, optional titles, and shape elements.
#[derive(Debug, Clone)]
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
    pub transition: Option<SlideTransition>,
    /// Inert ODF animation and timing trees attached to the slide.
    pub animations: Vec<AnimationNode>,
    /// Optional legacy `presentation:animations` effect tree.
    pub legacy_animation: Option<LegacyAnimationNode>,
    /// Shapes on the slide
    pub shapes: Vec<Shape>,
}

impl Slide {
    /// Get the title of the slide.
    ///
    /// Returns the slide title if present, None otherwise.
    pub fn title(&self) -> Result<Option<&str>> {
        Ok(self.title.as_deref())
    }

    /// Get the slide's primary body text.
    ///
    /// Use [`Self::all_text`] to include the title and labeled drawing shapes.
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
        parts.extend(
            self.shapes
                .iter()
                .map(|shape| shape.text.trim())
                .filter(|text| !text.is_empty()),
        );
        parts.join("\n")
    }

    /// Get all shapes on the slide.
    ///
    /// Returns a slice of shapes contained in this slide.
    pub fn shapes(&self) -> Result<&[Shape]> {
        Ok(&self.shapes)
    }

    /// Get the slide index.
    ///
    /// Returns the 0-based index of this slide in the presentation.
    pub fn index(&self) -> usize {
        self.index
    }

    /// Get the slide notes.
    ///
    /// Returns speaker notes for this slide if present, None otherwise.
    pub fn notes(&self) -> Result<Option<&str>> {
        Ok(self.notes.as_deref())
    }

    /// Get the slide transition configuration.
    pub fn transition(&self) -> Option<&SlideTransition> {
        self.transition.as_ref()
    }

    /// Get or create the slide transition configuration.
    pub fn transition_mut(&mut self) -> &mut SlideTransition {
        self.transition.get_or_insert_with(SlideTransition::new)
    }

    /// Remove the slide transition configuration.
    pub fn clear_transition(&mut self) {
        self.transition = None;
    }

    /// Return the slide's inert animation and timing trees.
    pub fn animations(&self) -> &[AnimationNode] {
        &self.animations
    }

    /// Return mutable animation and timing trees.
    pub fn animations_mut(&mut self) -> &mut Vec<AnimationNode> {
        &mut self.animations
    }

    /// Add a schema-defined animation root to the slide.
    pub fn add_animation(&mut self, animation: AnimationNode) -> Result<()> {
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
    pub fn legacy_animation(&self) -> Option<&LegacyAnimationNode> {
        self.legacy_animation.as_ref()
    }

    /// Set or remove the legacy presentation effect tree.
    pub fn set_legacy_animation(&mut self, animation: Option<LegacyAnimationNode>) {
        self.legacy_animation = animation;
    }
}

/// A shape (element) on a slide.
///
/// Shapes represent visual elements like text boxes, images, and drawings.
#[derive(Debug, Clone)]
pub struct Shape {
    /// Shape type (text box, image, frame, etc.)
    pub shape_type: litchi_core::ShapeType,
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
    /// Image source referenced by `draw:image`, when this is a picture shape.
    pub image_href: Option<String>,
    /// Inert audio/video plugin referenced by this frame.
    pub media: Option<MediaReference>,
}

impl Shape {
    /// Create a new empty shape
    pub fn new() -> Self {
        Self {
            shape_type: litchi_core::ShapeType::AutoShape,
            text: String::new(),
            name: None,
            x: None,
            y: None,
            width: None,
            height: None,
            style_name: None,
            image_href: None,
            media: None,
        }
    }

    /// Get the text content of the shape.
    pub fn text(&self) -> Result<&str> {
        Ok(&self.text)
    }

    /// Get the shape type.
    pub fn shape_type(&self) -> litchi_core::ShapeType {
        self.shape_type
    }

    /// Check if this is a text shape.
    pub fn has_text(&self) -> bool {
        !self.text.trim().is_empty()
    }

    /// Get the shape name/ID.
    pub fn name(&self) -> Option<&str> {
        self.name.as_deref()
    }

    /// Get the shape position as (x, y).
    pub fn position(&self) -> (Option<&str>, Option<&str>) {
        (self.x.as_deref(), self.y.as_deref())
    }

    /// Get the shape dimensions as (width, height).
    pub fn dimensions(&self) -> (Option<&str>, Option<&str>) {
        (self.width.as_deref(), self.height.as_deref())
    }

    /// Get the image source referenced by this shape.
    pub fn image_href(&self) -> Option<&str> {
        self.image_href.as_deref()
    }

    /// Return the inert audio/video plugin referenced by this shape.
    pub fn media(&self) -> Option<&MediaReference> {
        self.media.as_ref()
    }

    /// Attach an inert audio/video plugin and mark this shape as a graphic frame.
    pub fn with_media(mut self, media: MediaReference) -> Self {
        self.shape_type = litchi_core::ShapeType::GraphicFrame;
        self.image_href = None;
        self.media = Some(media);
        self
    }

    /// Set the image source and mark this shape as a picture.
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

#[cfg(test)]
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
            image_href: None,
            media: None,
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
            image_href: None,
            media: None,
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
            image_href: None,
            media: None,
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
            image_href: None,
            media: None,
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
            image_href: None,
            media: None,
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
            image_href: None,
            media: None,
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
            image_href: None,
            media: None,
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
            image_href: None,
            media: None,
        };
        let cloned = shape.clone();
        assert_eq!(shape.shape_type, cloned.shape_type);
        assert_eq!(shape.text, cloned.text);
        assert_eq!(shape.name, cloned.name);
    }

    #[test]
    fn test_shape_debug() {
        let shape = Shape::new();
        let debug_str = format!("{:?}", shape);
        assert!(debug_str.contains("Shape"));
    }
}
