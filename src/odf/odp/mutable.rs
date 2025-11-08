//! Mutable presentation structure for in-place modifications.
//!
//! This module provides a mutable wrapper around ODP presentations that allows
//! for in-place modification of slides, shapes, and content.

use crate::common::{Metadata, Result};
use crate::odf::core::{OdfStructure, PackageWriter};
use crate::odf::odp::{Presentation, Shape, Slide};
use std::path::Path;

/// A mutable ODP presentation that supports in-place modifications.
///
/// This struct wraps an ODP presentation and provides methods to modify its content,
/// including adding, updating, and removing slides and shapes.
///
/// # Examples
///
/// ```no_run
/// use litchi::odf::{Presentation, MutablePresentation};
///
/// # fn main() -> litchi::Result<()> {
/// // Open an existing presentation
/// let presentation = Presentation::open("input.odp")?;
/// let mut mutable = MutablePresentation::from_presentation(presentation)?;
///
/// // Modify the presentation
/// mutable.add_slide("New Slide", "Slide content")?;
/// mutable.remove_slide(0)?;
///
/// // Save the modified presentation
/// mutable.save("output.odp")?;
/// # Ok(())
/// # }
/// ```
pub struct MutablePresentation {
    /// Mutable slides
    slides: Vec<Slide>,
    /// Document metadata
    metadata: Metadata,
    /// Original MIME type
    mimetype: String,
    /// Original styles XML (preserved as-is)
    styles_xml: Option<String>,
}

impl MutablePresentation {
    /// Create a mutable presentation from an existing Presentation.
    ///
    /// This parses the presentation structure into mutable elements.
    ///
    /// # Arguments
    ///
    /// * `presentation` - The presentation to make mutable
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi::odf::{Presentation, MutablePresentation};
    ///
    /// # fn main() -> litchi::Result<()> {
    /// let presentation = Presentation::open("slides.odp")?;
    /// let mut mutable = MutablePresentation::from_presentation(presentation)?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn from_presentation(presentation: Presentation) -> Result<Self> {
        let slides = presentation.slides()?;
        let metadata = presentation.metadata()?;
        let mimetype = "application/vnd.oasis.opendocument.presentation".to_string();

        // Extract styles XML from the presentation's package
        // TODO: Add method to Presentation to expose get_file for extracting styles.xml
        let styles_xml = None;

        Ok(Self {
            slides,
            metadata,
            mimetype,
            styles_xml,
        })
    }

    /// Create a new empty mutable presentation.
    ///
    /// # Examples
    ///
    /// ```
    /// use litchi::odf::MutablePresentation;
    ///
    /// let presentation = MutablePresentation::new();
    /// ```
    pub fn new() -> Self {
        Self {
            slides: Vec::new(),
            metadata: Metadata::default(),
            mimetype: "application/vnd.oasis.opendocument.presentation".to_string(),
            styles_xml: None,
        }
    }

    /// Get all slides in the presentation.
    pub fn slides(&self) -> &[Slide] {
        &self.slides
    }

    /// Get a mutable reference to all slides.
    pub fn slides_mut(&mut self) -> &mut Vec<Slide> {
        &mut self.slides
    }

    /// Get the presentation metadata.
    pub fn metadata(&self) -> &Metadata {
        &self.metadata
    }

    /// Get a mutable reference to the presentation metadata.
    pub fn metadata_mut(&mut self) -> &mut Metadata {
        &mut self.metadata
    }

    /// Add a new slide to the end of the presentation.
    ///
    /// # Arguments
    ///
    /// * `title` - Optional title for the slide
    /// * `text` - Text content for the slide
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi::odf::MutablePresentation;
    ///
    /// # fn main() -> litchi::Result<()> {
    /// let mut presentation = MutablePresentation::new();
    /// presentation.add_slide("Slide 1", "Content for slide 1")?;
    /// presentation.add_slide("Slide 2", "Content for slide 2")?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn add_slide(&mut self, title: &str, text: &str) -> Result<()> {
        let slide = Slide {
            title: Some(title.to_string()),
            text: text.to_string(),
            index: self.slides.len(),
            notes: None,
            shapes: Vec::new(),
        };
        self.slides.push(slide);
        Ok(())
    }

    /// Insert a slide at a specific index.
    ///
    /// # Arguments
    ///
    /// * `index` - Position to insert at (0-based)
    /// * `title` - Optional title for the slide
    /// * `text` - Text content for the slide
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi::odf::MutablePresentation;
    ///
    /// # fn main() -> litchi::Result<()> {
    /// let mut presentation = MutablePresentation::new();
    /// presentation.add_slide("First", "Content 1")?;
    /// presentation.add_slide("Third", "Content 3")?;
    /// presentation.insert_slide(1, "Second", "Content 2")?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn insert_slide(&mut self, index: usize, title: &str, text: &str) -> Result<()> {
        if index <= self.slides.len() {
            let slide = Slide {
                title: Some(title.to_string()),
                text: text.to_string(),
                index,
                notes: None,
                shapes: Vec::new(),
            };
            self.slides.insert(index, slide);

            // Update indices of subsequent slides
            for i in (index + 1)..self.slides.len() {
                self.slides[i].index = i;
            }

            Ok(())
        } else {
            Err(crate::common::Error::InvalidFormat(format!(
                "Index {} out of bounds (length: {})",
                index,
                self.slides.len()
            )))
        }
    }

    /// Remove a slide at a specific index.
    ///
    /// # Arguments
    ///
    /// * `index` - Index of the slide to remove (0-based)
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi::odf::MutablePresentation;
    ///
    /// # fn main() -> litchi::Result<()> {
    /// let mut presentation = MutablePresentation::new();
    /// presentation.add_slide("Slide 1", "Content 1")?;
    /// presentation.add_slide("Slide 2", "Content 2")?;
    /// presentation.remove_slide(0)?; // Remove first slide
    /// # Ok(())
    /// # }
    /// ```
    pub fn remove_slide(&mut self, index: usize) -> Result<Slide> {
        if index < self.slides.len() {
            let slide = self.slides.remove(index);

            // Update indices of subsequent slides
            for i in index..self.slides.len() {
                self.slides[i].index = i;
            }

            Ok(slide)
        } else {
            Err(crate::common::Error::InvalidFormat(format!(
                "Index {} out of bounds (length: {})",
                index,
                self.slides.len()
            )))
        }
    }

    /// Update a slide at a specific index.
    ///
    /// # Arguments
    ///
    /// * `index` - Index of the slide to update (0-based)
    /// * `title` - New title for the slide
    /// * `text` - New text content
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi::odf::MutablePresentation;
    ///
    /// # fn main() -> litchi::Result<()> {
    /// let mut presentation = MutablePresentation::new();
    /// presentation.add_slide("Old Title", "Old content")?;
    /// presentation.update_slide(0, "New Title", "New content")?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn update_slide(&mut self, index: usize, title: &str, text: &str) -> Result<()> {
        if index < self.slides.len() {
            self.slides[index].title = Some(title.to_string());
            self.slides[index].text = text.to_string();
            Ok(())
        } else {
            Err(crate::common::Error::InvalidFormat(format!(
                "Index {} out of bounds (length: {})",
                index,
                self.slides.len()
            )))
        }
    }

    /// Clear all slides from the presentation.
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi::odf::MutablePresentation;
    ///
    /// # fn main() -> litchi::Result<()> {
    /// let mut presentation = MutablePresentation::new();
    /// presentation.add_slide("Slide 1", "Content 1")?;
    /// presentation.add_slide("Slide 2", "Content 2")?;
    /// presentation.clear_slides();
    /// assert_eq!(presentation.slides().len(), 0);
    /// # Ok(())
    /// # }
    /// ```
    pub fn clear_slides(&mut self) {
        self.slides.clear();
    }

    /// Add a shape to a slide.
    ///
    /// # Arguments
    ///
    /// * `slide_index` - Index of the slide to add the shape to
    /// * `shape` - Shape to add
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi::odf::{MutablePresentation, Shape};
    ///
    /// # fn main() -> litchi::Result<()> {
    /// let mut presentation = MutablePresentation::new();
    /// presentation.add_slide("Slide 1", "Content")?;
    /// let mut shape = Shape::new();
    /// shape.text = "Shape text".to_string();
    /// presentation.add_shape(0, shape)?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn add_shape(&mut self, slide_index: usize, shape: Shape) -> Result<()> {
        if slide_index < self.slides.len() {
            self.slides[slide_index].shapes.push(shape);
            Ok(())
        } else {
            Err(crate::common::Error::InvalidFormat(format!(
                "Slide index {} out of bounds",
                slide_index
            )))
        }
    }

    /// Remove a shape from a slide.
    ///
    /// # Arguments
    ///
    /// * `slide_index` - Index of the slide
    /// * `shape_index` - Index of the shape to remove
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi::odf::MutablePresentation;
    ///
    /// # fn main() -> litchi::Result<()> {
    /// let mut presentation = MutablePresentation::new();
    /// presentation.add_slide("Slide 1", "Content")?;
    /// // Add shape first, then remove it
    /// presentation.remove_shape(0, 0)?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn remove_shape(&mut self, slide_index: usize, shape_index: usize) -> Result<Shape> {
        if slide_index < self.slides.len() {
            let slide = &mut self.slides[slide_index];
            if shape_index < slide.shapes.len() {
                Ok(slide.shapes.remove(shape_index))
            } else {
                Err(crate::common::Error::InvalidFormat(format!(
                    "Shape index {} out of bounds",
                    shape_index
                )))
            }
        } else {
            Err(crate::common::Error::InvalidFormat(format!(
                "Slide index {} out of bounds",
                slide_index
            )))
        }
    }

    /// Clear all content (text and shapes) from a slide.
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi::odf::MutablePresentation;
    ///
    /// # fn main() -> litchi::Result<()> {
    /// let mut presentation = MutablePresentation::new();
    /// presentation.add_slide("Slide 1", "Content")?;
    /// presentation.clear_slide(0)?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn clear_slide(&mut self, slide_index: usize) -> Result<()> {
        if slide_index < self.slides.len() {
            self.slides[slide_index].text.clear();
            self.slides[slide_index].shapes.clear();
            Ok(())
        } else {
            Err(crate::common::Error::InvalidFormat(format!(
                "Slide index {} out of bounds",
                slide_index
            )))
        }
    }

    /// Generate content.xml from the current mutable state.
    fn generate_content_xml(&self) -> String {
        let mut body = String::new();

        for (i, slide) in self.slides.iter().enumerate() {
            let page_num = i + 1;
            body.push_str(&xml_minifier::minified_xml_format!(
                r#"<draw:page draw:name="page{}" draw:style-name="dp1" draw:master-page-name="Default">"#,
                page_num
            ));

            // Add title frame if title exists
            if let Some(ref title) = slide.title {
                let escaped_title = Self::escape_xml(title);
                body.push_str(&xml_minifier::minified_xml_format!(
                    r#"<draw:frame draw:style-name="gr1" draw:text-style-name="P1" draw:layer="layout" svg:width="25.199cm" svg:height="3.506cm" svg:x="1.4cm" svg:y="0.962cm"><draw:text-box><text:p text:style-name="P1">{}</text:p></draw:text-box></draw:frame>"#,
                    escaped_title
                ));
            }

            // Add text frame
            if !slide.text.is_empty() {
                let y_position = if slide.title.is_some() {
                    "5.0cm"
                } else {
                    "2.0cm"
                };
                let escaped_text = Self::escape_xml(&slide.text);
                body.push_str(&xml_minifier::minified_xml_format!(
                    r#"<draw:frame draw:style-name="gr2" draw:text-style-name="P2" draw:layer="layout" svg:width="25.199cm" svg:height="10cm" svg:x="1.4cm" svg:y="{}"><draw:text-box><text:p text:style-name="P2">{}</text:p></draw:text-box></draw:frame>"#,
                    y_position,
                    escaped_text
                ));
            }

            // Add shapes
            for (shape_idx, shape) in slide.shapes.iter().enumerate() {
                use crate::common::ShapeType;

                let x = shape.x.as_deref().unwrap_or("2cm");
                let y = shape.y.as_deref().unwrap_or("8cm");
                let width = shape.width.as_deref().unwrap_or("10cm");
                let height = shape.height.as_deref().unwrap_or("5cm");
                let default_name = format!("Shape{}", shape_idx + 1);
                let name = shape.name.as_deref().unwrap_or(&default_name);
                let style_name = shape.style_name.as_deref().unwrap_or("gr3");

                match shape.shape_type {
                    ShapeType::TextBox | ShapeType::AutoShape | ShapeType::Placeholder => {
                        if shape.has_text() {
                            let escaped_name = Self::escape_xml(name);
                            let escaped_shape_text = Self::escape_xml(&shape.text);
                            body.push_str(&xml_minifier::minified_xml_format!(
                                r#"<draw:frame draw:name="{}" draw:style-name="{}" draw:layer="layout" svg:x="{}" svg:y="{}" svg:width="{}" svg:height="{}"><draw:text-box><text:p text:style-name="P2">{}</text:p></draw:text-box></draw:frame>"#,
                                escaped_name,
                                style_name,
                                x,
                                y,
                                width,
                                height,
                                escaped_shape_text
                            ));
                        }
                    },
                    _ => {},
                }
            }

            body.push_str("</draw:page>");
        }

        xml_minifier::minified_xml_format!(
            r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" office:version="1.3"><office:scripts/><office:font-face-decls/><office:automatic-styles/><office:body><office:presentation>{}</office:presentation></office:body></office:document-content>"#,
            body
        )
    }

    /// Generate meta.xml with current metadata.
    fn generate_meta_xml(&self) -> String {
        let now = chrono::Utc::now().to_rfc3339();
        let mut meta_fields = String::new();

        // Add optional metadata fields
        if let Some(ref title) = self.metadata.title {
            let escaped_title = Self::escape_xml(title);
            meta_fields.push_str(&xml_minifier::minified_xml_format!(
                r#"<dc:title>{}</dc:title>"#,
                escaped_title
            ));
        }

        if let Some(ref author) = self.metadata.author {
            let escaped_author = Self::escape_xml(author);
            meta_fields.push_str(&xml_minifier::minified_xml_format!(
                r#"<dc:creator>{}</dc:creator>"#,
                escaped_author
            ));
        }

        xml_minifier::minified_xml_format!(
            r#"<?xml version="1.0" encoding="UTF-8"?><office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" office:version="1.3"><office:meta><meta:generator>Litchi/0.0.1</meta:generator><dc:date>{}</dc:date>{}</office:meta></office:document-meta>"#,
            now,
            meta_fields
        )
    }

    /// Escape XML special characters.
    fn escape_xml(text: &str) -> String {
        text.replace('&', "&amp;")
            .replace('<', "&lt;")
            .replace('>', "&gt;")
            .replace('"', "&quot;")
            .replace('\'', "&apos;")
    }

    /// Save the modified presentation to a file.
    ///
    /// # Arguments
    ///
    /// * `path` - Path where the ODP file should be saved
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi::odf::MutablePresentation;
    ///
    /// # fn main() -> litchi::Result<()> {
    /// let mut presentation = MutablePresentation::new();
    /// presentation.add_slide("Slide 1", "Content")?;
    /// presentation.save("output.odp")?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn save<P: AsRef<Path>>(&self, path: P) -> Result<()> {
        let bytes = self.to_bytes()?;
        std::fs::write(path, bytes)?;
        Ok(())
    }

    /// Convert the presentation to bytes.
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi::odf::MutablePresentation;
    ///
    /// # fn main() -> litchi::Result<()> {
    /// let mut presentation = MutablePresentation::new();
    /// presentation.add_slide("Slide 1", "Content")?;
    /// let bytes = presentation.to_bytes()?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        let mut writer = PackageWriter::new();

        // Set MIME type
        writer.set_mimetype(&self.mimetype)?;

        // Add content.xml (regenerated from mutable state)
        let content_xml = self.generate_content_xml();
        writer.add_file("content.xml", content_xml.as_bytes())?;

        // Add styles.xml (preserved or default)
        let default_styles = OdfStructure::default_styles_xml();
        let styles_xml = self.styles_xml.as_deref().unwrap_or(&default_styles);
        writer.add_file("styles.xml", styles_xml.as_bytes())?;

        // Add meta.xml (regenerated with current metadata)
        let meta_xml = self.generate_meta_xml();
        writer.add_file("meta.xml", meta_xml.as_bytes())?;

        writer.finish_to_bytes()
    }
}

impl Default for MutablePresentation {
    fn default() -> Self {
        Self::new()
    }
}
