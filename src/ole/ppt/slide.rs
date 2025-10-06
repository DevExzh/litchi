/// Slide - represents an individual slide in a PowerPoint presentation.
use super::package::Result;
use super::shapes::{Shape, escher::EscherParser};

/// A slide in a PowerPoint presentation (.ppt).
///
/// This represents an individual slide with its content, text, and formatting.
/// Slides contain shapes that hold the actual content like text boxes, placeholders, etc.
///
/// # Examples
///
/// ```rust,no_run
/// use litchi::ppt::Package;
///
/// let mut pkg = Package::open("presentation.ppt")?;
/// let pres = pkg.presentation()?;
///
/// for slide in pres.slides()? {
///     println!("Slide text: {}", slide.text()?);
///     println!("Slide has {} shapes", slide.shape_count()?);
///
///     // Access individual shapes
///     for shape in slide.shapes()? {
///         match shape {
///             Shape::TextBox(textbox) => {
///                 println!("Text box: {}", textbox.text()?);
///             }
///             Shape::Placeholder(placeholder) => {
///                 println!("Placeholder type: {:?}", placeholder.placeholder_type());
///             }
///             _ => {}
///         }
///     }
/// }
/// # Ok::<(), Box<dyn std::error::Error>>(())
/// ```
pub struct Slide {
    /// Slide content data (Escher binary format)
    data: Vec<u8>,
    /// Slide index (0-based)
    index: usize,
    /// Parsed shapes on this slide
    shapes: Vec<Box<dyn Shape>>,
}

impl Slide {
    /// Create a new slide with shape parsing.
    ///
    /// This is typically called internally by the presentation parser.
    pub(crate) fn new(data: Vec<u8>, index: usize) -> Result<Self> {
        let mut slide = Self {
            data,
            index,
            shapes: Vec::new(),
        };

        // Parse shapes from the slide data
        slide.parse_shapes()?;

        Ok(slide)
    }

    /// Parse shapes from the slide's binary data.
    fn parse_shapes(&mut self) -> Result<()> {
        if self.data.is_empty() {
            return Ok(());
        }

        // Use Escher parser to extract shape data
        let mut parser = EscherParser::new();
        parser.parse_data(&self.data)?;

        // Extract shape properties and create shape objects
        let shape_properties = parser.extract_shapes()?;

        for props in shape_properties {
            // Create shape objects based on shape type
            let shape: Box<dyn Shape> = match props.shape_type {
                super::shapes::shape::ShapeType::TextBox => {
                    // Find the corresponding Escher record for this shape
                    if let Some(shape_record) = parser.find_record_by_shape_id(props.id) {
                        match super::shapes::textbox::TextBox::from_escher_record(shape_record) {
                            Ok(textbox) => Box::new(textbox),
                            Err(_) => Box::new(super::shapes::textbox::TextBox::new(props, Vec::new())),
                        }
                    } else {
                        Box::new(super::shapes::textbox::TextBox::new(props, Vec::new()))
                    }
                }
                super::shapes::shape::ShapeType::Placeholder => {
                    // For placeholders, try to find associated PlaceholderData record
                    if let Some(placeholder_record) = self.find_associated_placeholder_record(&parser, &props.id) {
                        match super::shapes::placeholder::Placeholder::from_escher_record(placeholder_record) {
                            Ok(placeholder) => Box::new(placeholder),
                            Err(_) => Box::new(super::shapes::placeholder::Placeholder::new(props, Vec::new())),
                        }
                    } else {
                        Box::new(super::shapes::placeholder::Placeholder::new(props, Vec::new()))
                    }
                }
                super::shapes::shape::ShapeType::AutoShape => {
                    Box::new(super::shapes::autoshape::AutoShape::new(props, Vec::new()))
                }
                super::shapes::shape::ShapeType::Group => {
                    // Group shapes contain other shapes - parse children if available
                    Box::new(super::shapes::shape::ShapeContainer::new(props, Vec::new()))
                }
                super::shapes::shape::ShapeType::Picture => {
                    // Picture shapes contain image data
                    Box::new(super::shapes::shape::ShapeContainer::new(props, Vec::new()))
                }
                super::shapes::shape::ShapeType::Line => {
                    // Line shapes are simple geometric shapes
                    Box::new(super::shapes::shape::ShapeContainer::new(props, Vec::new()))
                }
                super::shapes::shape::ShapeType::Connector => {
                    // Connector shapes link other shapes
                    Box::new(super::shapes::shape::ShapeContainer::new(props, Vec::new()))
                }
                super::shapes::shape::ShapeType::Object => {
                    // Object shapes contain embedded objects
                    Box::new(super::shapes::shape::ShapeContainer::new(props, Vec::new()))
                }
                super::shapes::shape::ShapeType::Table => {
                    // Table shapes contain tabular data
                    Box::new(super::shapes::shape::ShapeContainer::new(props, Vec::new()))
                }
                _ => {
                    // For unknown shape types, create a basic shape container
                    Box::new(super::shapes::shape::ShapeContainer::new(props, Vec::new()))
                }
            };

            self.shapes.push(shape);
        }

        Ok(())
    }

    /// Find the Escher record for a placeholder by shape ID.
    /// This method attempts to find a PlaceholderData record that might be associated
    /// with the given shape ID. In POI, this association is typically made through
    /// the Escher record hierarchy.
    fn find_associated_placeholder_record<'a>(&self, parser: &'a super::shapes::escher::EscherParser, _shape_id: &u32) -> Option<&'a super::shapes::escher::EscherRecord> {
        // For now, return the first PlaceholderData record found
        // In a full implementation, this would match by shape hierarchy or other criteria
        parser.placeholder_records().first()
    }

    /// Get all text content from the slide.
    ///
    /// This extracts all text from the slide's text boxes, titles, etc.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ppt::Package;
    ///
    /// let mut pkg = Package::open("presentation.ppt")?;
    /// let pres = pkg.presentation()?;
    ///
    /// for slide in pres.slides()? {
    ///     println!("Slide: {}", slide.text()?);
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn text(&self) -> super::package::Result<String> {
        let mut text_parts = Vec::new();

        for shape in &self.shapes {
            let shape_text = shape.text()?;
            if !shape_text.is_empty() {
                text_parts.push(shape_text);
            }
        }

        if text_parts.is_empty() {
            Ok(format!("Slide {} (no text content)", self.index + 1))
        } else {
            Ok(text_parts.join("\n"))
        }
    }

    /// Get the number of shapes on the slide.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ppt::Package;
    ///
    /// let mut pkg = Package::open("presentation.ppt")?;
    /// let pres = pkg.presentation()?;
    ///
    /// for slide in pres.slides()? {
    ///     println!("Shapes: {}", slide.shape_count()?);
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn shape_count(&self) -> super::package::Result<usize> {
        Ok(self.shapes.len())
    }

    /// Get all shapes on the slide.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ppt::Package;
    /// use litchi::ppt::shapes::Shape;
    ///
    /// let mut pkg = Package::open("presentation.ppt")?;
    /// let pres = pkg.presentation()?;
    ///
    /// for slide in pres.slides()? {
    ///     for shape in slide.shapes()? {
    ///         match shape {
    ///             Shape::TextBox(textbox) => {
    ///                 println!("Text box: {}", textbox.text()?);
    ///             }
    ///             Shape::Placeholder(placeholder) => {
    ///                 println!("Placeholder type: {:?}", placeholder.placeholder_type());
    ///             }
    ///             _ => {}
    ///         }
    ///     }
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn shapes(&self) -> super::package::Result<Vec<&dyn Shape>> {
        Ok(self.shapes.iter().map(|s| s.as_ref() as &dyn Shape).collect())
    }

    /// Get text boxes on the slide.
    ///
    /// Based on POI's approach of filtering shapes by type.
    pub fn text_boxes(&self) -> Vec<&super::shapes::textbox::TextBox> {
        // Since we store shapes as trait objects, we need to check the shape type
        // and return references appropriately. For now, this is a simplified implementation.
        // In a full implementation, we'd use downcasting or store typed collections.
        Vec::new()
    }

    /// Get placeholders on the slide.
    ///
    /// Based on POI's HSLFSheet.getPlaceholder() logic which filters shapes
    /// to find those with placeholder information.
    pub fn placeholders(&self) -> Vec<&super::shapes::placeholder::Placeholder> {
        // Extract placeholders by checking shape properties
        // This is a simplified implementation - full version would use downcasting
        Vec::new()
    }

    /// Get a placeholder by index.
    ///
    /// # Arguments
    ///
    /// * `idx` - The index of the placeholder (0-based)
    ///
    /// # Returns
    ///
    /// Returns the placeholder if found, None otherwise
    pub fn get_placeholder(&self, idx: usize) -> Option<&super::shapes::placeholder::Placeholder> {
        self.placeholders().get(idx).copied()
    }

    /// Get placeholders of a specific type.
    ///
    /// # Arguments
    ///
    /// * `placeholder_type` - The type of placeholder to find
    ///
    /// # Returns
    ///
    /// Returns a vector of placeholders of the specified type
    pub fn get_placeholders_by_type(&self, placeholder_type: super::shapes::placeholder::PlaceholderType) -> Vec<&super::shapes::placeholder::Placeholder> {
        self.placeholders()
            .into_iter()
            .filter(|p| p.placeholder_type() == placeholder_type)
            .collect()
    }

    /// Get the slide index (0-based).
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ppt::Package;
    ///
    /// let mut pkg = Package::open("presentation.ppt")?;
    /// let pres = pkg.presentation()?;
    ///
    /// for (i, slide) in pres.slides()?.iter().enumerate() {
    ///     println!("Slide {} index: {}", i + 1, slide.index());
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn index(&self) -> usize {
        self.index
    }

    /// Get the slide's raw data.
    ///
    /// This provides access to the underlying binary data of the slide.
    #[inline]
    pub fn data(&self) -> &[u8] {
        &self.data
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_slide_creation() {
        let data = vec![1, 2, 3, 4, 5];
        let slide = Slide::new(data, 0).unwrap();
        assert_eq!(slide.index(), 0);
        assert_eq!(slide.data().len(), 5);
    }

    #[test]
    fn test_slide_text() {
        let data = vec![];
        let slide = Slide::new(data, 0).unwrap();
        let text = slide.text().unwrap();
        assert!(text.contains("Slide 1"));
    }
}
