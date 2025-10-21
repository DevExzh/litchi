//! Slide and shape structures for ODP presentations.

use crate::common::Result;

/// A slide in an ODP presentation.
///
/// Slides contain text content, optional titles, and shape elements.
#[derive(Clone)]
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
