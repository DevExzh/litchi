//! Slide and shape structures for ODP presentations.

use crate::common::Result;

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

    /// Extract all text content from the slide.
    ///
    /// Returns the combined text from all text elements on the slide.
    pub fn text(&self) -> Result<&str> {
        Ok(&self.text)
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
}

/// A shape (element) on a slide.
///
/// Shapes represent visual elements like text boxes, images, and drawings.
#[derive(Debug, Clone)]
pub struct Shape {
    /// Shape type (text box, image, frame, etc.)
    pub shape_type: crate::common::ShapeType,
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
}

impl Shape {
    /// Create a new empty shape
    pub fn new() -> Self {
        Self {
            shape_type: crate::common::ShapeType::AutoShape,
            text: String::new(),
            name: None,
            x: None,
            y: None,
            width: None,
            height: None,
            style_name: None,
        }
    }

    /// Get the text content of the shape.
    pub fn text(&self) -> Result<&str> {
        Ok(&self.text)
    }

    /// Get the shape type.
    pub fn shape_type(&self) -> crate::common::ShapeType {
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
            shapes: vec![],
        };
        assert_eq!(slide.text().unwrap(), "Hello World");
    }

    #[test]
    fn test_slide_shapes_method() {
        let shapes = vec![Shape {
            shape_type: crate::common::ShapeType::TextBox,
            text: "Shape 1".to_string(),
            name: Some("Shape1".to_string()),
            x: Some("0cm".to_string()),
            y: Some("0cm".to_string()),
            width: Some("5cm".to_string()),
            height: Some("3cm".to_string()),
            style_name: None,
        }];
        let slide = Slide {
            title: None,
            text: String::new(),
            index: 0,
            notes: None,
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
        assert_eq!(shape.shape_type, crate::common::ShapeType::AutoShape);
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
        assert_eq!(shape.shape_type, crate::common::ShapeType::AutoShape);
        assert!(shape.text.is_empty());
    }

    #[test]
    fn test_shape_text_method() {
        let shape = Shape {
            shape_type: crate::common::ShapeType::TextBox,
            text: "Hello".to_string(),
            name: None,
            x: None,
            y: None,
            width: None,
            height: None,
            style_name: None,
        };
        assert_eq!(shape.text().unwrap(), "Hello");
    }

    #[test]
    fn test_shape_shape_type_method() {
        let shape = Shape {
            shape_type: crate::common::ShapeType::Picture,
            text: String::new(),
            name: None,
            x: None,
            y: None,
            width: None,
            height: None,
            style_name: None,
        };
        assert_eq!(shape.shape_type(), crate::common::ShapeType::Picture);
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
            shape_type: crate::common::ShapeType::TextBox,
            text: String::new(),
            name: Some("MyShape".to_string()),
            x: None,
            y: None,
            width: None,
            height: None,
            style_name: None,
        };
        assert_eq!(shape.name(), Some("MyShape"));
    }

    #[test]
    fn test_shape_name_none() {
        let shape = Shape {
            shape_type: crate::common::ShapeType::TextBox,
            text: String::new(),
            name: None,
            x: None,
            y: None,
            width: None,
            height: None,
            style_name: None,
        };
        assert_eq!(shape.name(), None);
    }

    #[test]
    fn test_shape_position() {
        let shape = Shape {
            shape_type: crate::common::ShapeType::TextBox,
            text: String::new(),
            name: None,
            x: Some("10cm".to_string()),
            y: Some("5cm".to_string()),
            width: None,
            height: None,
            style_name: None,
        };
        let (x, y) = shape.position();
        assert_eq!(x, Some("10cm"));
        assert_eq!(y, Some("5cm"));
    }

    #[test]
    fn test_shape_dimensions() {
        let shape = Shape {
            shape_type: crate::common::ShapeType::TextBox,
            text: String::new(),
            name: None,
            x: None,
            y: None,
            width: Some("20cm".to_string()),
            height: Some("15cm".to_string()),
            style_name: None,
        };
        let (w, h) = shape.dimensions();
        assert_eq!(w, Some("20cm"));
        assert_eq!(h, Some("15cm"));
    }

    #[test]
    fn test_shape_clone() {
        let shape = Shape {
            shape_type: crate::common::ShapeType::Placeholder,
            text: "Content".to_string(),
            name: Some("Shape1".to_string()),
            x: Some("1cm".to_string()),
            y: Some("2cm".to_string()),
            width: Some("10cm".to_string()),
            height: Some("5cm".to_string()),
            style_name: Some("Style1".to_string()),
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
