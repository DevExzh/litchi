//! OpenDocument Presentation builder.
//!
//! This module provides a builder pattern for creating new ODP presentations from scratch.

use crate::common::{Metadata, Result, xml::escape_xml};
use crate::odf::core::{OdfStructure, PackageWriter};
use crate::odf::odp::Slide;
use std::path::Path;

/// Builder for creating new ODP presentations.
///
/// This builder allows you to create ODP presentations programmatically by adding
/// slides with text and shapes, then saving them to a file or bytes.
///
/// # Examples
///
/// ```no_run
/// use litchi::odf::PresentationBuilder;
///
/// # fn main() -> litchi::Result<()> {
/// let mut builder = PresentationBuilder::new();
/// builder.add_slide_with_title("Welcome", "This is my presentation")?;
/// builder.add_slide_with_title("Slide 2", "More content here")?;
/// builder.save("presentation.odp")?;
/// # Ok(())
/// # }
/// ```
pub struct PresentationBuilder {
    slides: Vec<Slide>,
    metadata: Metadata,
}

impl Default for PresentationBuilder {
    fn default() -> Self {
        Self::new()
    }
}

impl PresentationBuilder {
    /// Create a new presentation builder
    ///
    /// # Examples
    ///
    /// ```
    /// use litchi::odf::PresentationBuilder;
    ///
    /// let builder = PresentationBuilder::new();
    /// ```
    pub fn new() -> Self {
        Self {
            slides: Vec::new(),
            metadata: Metadata::default(),
        }
    }

    /// Set document metadata
    ///
    /// # Arguments
    ///
    /// * `metadata` - Document metadata (title, author, etc.)
    pub fn set_metadata(&mut self, metadata: Metadata) {
        self.metadata = metadata;
    }

    /// Add a slide with title and text content
    ///
    /// # Arguments
    ///
    /// * `title` - Title for the slide
    /// * `text` - Text content for the slide
    ///
    /// # Examples
    ///
    /// ```
    /// use litchi::odf::PresentationBuilder;
    ///
    /// # fn main() -> litchi::Result<()> {
    /// let mut builder = PresentationBuilder::new();
    /// builder.add_slide_with_title("Introduction", "Welcome to our presentation")?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn add_slide_with_title(&mut self, title: &str, text: &str) -> Result<&mut Self> {
        let slide = Slide {
            title: Some(title.to_string()),
            text: text.to_string(),
            index: self.slides.len(),
            notes: None,
            shapes: Vec::new(),
        };
        self.slides.push(slide);
        Ok(self)
    }

    /// Add a slide with only text content (no title)
    ///
    /// # Arguments
    ///
    /// * `text` - Text content for the slide
    ///
    /// # Examples
    ///
    /// ```
    /// use litchi::odf::PresentationBuilder;
    ///
    /// # fn main() -> litchi::Result<()> {
    /// let mut builder = PresentationBuilder::new();
    /// builder.add_slide("Simple slide content")?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn add_slide(&mut self, text: &str) -> Result<&mut Self> {
        let slide = Slide {
            title: None,
            text: text.to_string(),
            index: self.slides.len(),
            notes: None,
            shapes: Vec::new(),
        };
        self.slides.push(slide);
        Ok(self)
    }

    /// Add a Slide element directly
    ///
    /// # Arguments
    ///
    /// * `slide` - A complete `Slide` element to add
    ///
    /// # Examples
    ///
    /// ```
    /// use litchi::odf::{PresentationBuilder, Slide, Shape};
    ///
    /// # fn main() -> litchi::Result<()> {
    /// let mut builder = PresentationBuilder::new();
    /// let slide = Slide {
    ///     title: Some("Custom Slide".to_string()),
    ///     text: "Custom content".to_string(),
    ///     index: 0,
    ///     notes: Some("Speaker notes".to_string()),
    ///     shapes: vec![],
    /// };
    /// builder.add_slide_element(slide)?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn add_slide_element(&mut self, mut slide: Slide) -> Result<&mut Self> {
        slide.index = self.slides.len();
        self.slides.push(slide);
        Ok(self)
    }

    /// Generate XML for a shape
    fn generate_shape_xml(shape: &crate::odf::odp::Shape, idx: usize) -> String {
        use crate::common::ShapeType;

        // Determine default position and size if not provided
        let x = shape.x.as_deref().unwrap_or("2cm");
        let y = shape.y.as_deref().unwrap_or("8cm");
        let width = shape.width.as_deref().unwrap_or("10cm");
        let height = shape.height.as_deref().unwrap_or("5cm");
        let default_name = format!("Shape{}", idx + 1);
        let name = shape.name.as_deref().unwrap_or(&default_name);
        let style_name = shape.style_name.as_deref().unwrap_or("gr3");

        match shape.shape_type {
            ShapeType::TextBox | ShapeType::AutoShape | ShapeType::Placeholder => {
                // Text box or auto shape with text content
                if shape.has_text() {
                    format!(
                        r#"<draw:frame draw:name="{}" draw:style-name="{}" draw:layer="layout" svg:x="{}" svg:y="{}" svg:width="{}" svg:height="{}"><draw:text-box><text:p text:style-name="P2">{}</text:p></draw:text-box></draw:frame>"#,
                        escape_xml(name),
                        style_name,
                        x,
                        y,
                        width,
                        height,
                        escape_xml(&shape.text)
                    )
                } else {
                    // Empty frame
                    format!(
                        r#"<draw:frame draw:name="{}" draw:style-name="{}" draw:layer="layout" svg:x="{}" svg:y="{}" svg:width="{}" svg:height="{}"/>"#,
                        escape_xml(name),
                        style_name,
                        x,
                        y,
                        width,
                        height
                    )
                }
            },
            ShapeType::Picture => {
                // Image frame (basic support - would need actual image path)
                format!(
                    r#"<draw:frame draw:name="{}" draw:style-name="{}" draw:layer="layout" svg:x="{}" svg:y="{}" svg:width="{}" svg:height="{}"><draw:image/></draw:frame>"#,
                    escape_xml(name),
                    style_name,
                    x,
                    y,
                    width,
                    height
                )
            },
            ShapeType::Line | ShapeType::Connector => {
                // Line shape - use x,y as start and width,height as end offsets
                let x2 = shape.width.as_deref().unwrap_or("12cm");
                let y2 = shape.height.as_deref().unwrap_or("8cm");
                format!(
                    r#"<draw:line draw:name="{}" draw:style-name="{}" draw:layer="layout" svg:x1="{}" svg:y1="{}" svg:x2="{}" svg:y2="{}"/>"#,
                    escape_xml(name),
                    style_name,
                    x,
                    y,
                    x2,
                    y2
                )
            },
            _ => {
                // Generic shape or unsupported - render as text frame if it has text
                if shape.has_text() {
                    format!(
                        r#"<draw:frame draw:name="{}" draw:style-name="{}" draw:layer="layout" svg:x="{}" svg:y="{}" svg:width="{}" svg:height="{}"><draw:text-box><text:p text:style-name="P2">{}</text:p></draw:text-box></draw:frame>"#,
                        escape_xml(name),
                        style_name,
                        x,
                        y,
                        width,
                        height,
                        escape_xml(&shape.text)
                    )
                } else {
                    String::new() // Skip unsupported shapes without text
                }
            },
        }
    }

    /// Generate the content.xml body for presentation
    fn generate_content_body(&self) -> String {
        let shape_count = self.slides.iter().map(|s| s.shapes.len()).sum::<usize>();
        let mut estimated = 256usize;
        estimated += self.slides.len() * 128;
        estimated += shape_count * 192;
        estimated += self
            .slides
            .iter()
            .map(|s| s.text.len() + s.title.as_ref().map(|t| t.len()).unwrap_or(0))
            .sum::<usize>();
        estimated += self
            .slides
            .iter()
            .flat_map(|s| s.shapes.iter())
            .map(|sh| sh.text.len() + sh.name.as_ref().map(|n| n.len()).unwrap_or(0))
            .sum::<usize>();

        let mut body = String::with_capacity(estimated);

        for (i, slide) in self.slides.iter().enumerate() {
            body.push_str(&format!(
                r#"<draw:page draw:name="page{}" draw:style-name="dp1" draw:master-page-name="Default">"#,
                i + 1
            ));

            // Add title frame if title exists
            if let Some(ref title) = slide.title {
                body.push_str(&format!(
                    r#"<draw:frame draw:style-name="gr1" draw:text-style-name="P1" draw:layer="layout" svg:width="25.199cm" svg:height="3.506cm" svg:x="1.4cm" svg:y="0.962cm"><draw:text-box><text:p text:style-name="P1">{}</text:p></draw:text-box></draw:frame>"#,
                    escape_xml(title)
                ));
            }

            // Add text frame
            if !slide.text.is_empty() {
                let y_position = if slide.title.is_some() {
                    "5.0cm"
                } else {
                    "2.0cm"
                };
                body.push_str(&format!(
                    r#"<draw:frame draw:style-name="gr2" draw:text-style-name="P2" draw:layer="layout" svg:width="25.199cm" svg:height="10cm" svg:x="1.4cm" svg:y="{}"><draw:text-box><text:p text:style-name="P2">{}</text:p></draw:text-box></draw:frame>"#,
                    y_position,
                    escape_xml(&slide.text)
                ));
            }

            // Add custom shapes
            for (shape_idx, shape) in slide.shapes.iter().enumerate() {
                body.push_str(&Self::generate_shape_xml(shape, shape_idx));
            }

            body.push_str("</draw:page>");
        }

        body
    }

    /// Generate the complete content.xml for presentation
    fn generate_content_xml(&self) -> String {
        let body = self.generate_content_body();

        format!(
            r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0" xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0" xmlns:dr3d="urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0" xmlns:math="http://www.w3.org/1998/Math/MathML" xmlns:form="urn:oasis:names:tc:opendocument:xmlns:form:1.0" xmlns:script="urn:oasis:names:tc:opendocument:xmlns:script:1.0" xmlns:ooo="http://openoffice.org/2004/office" office:version="1.3"><office:scripts/><office:font-face-decls/><office:automatic-styles/><office:body><office:presentation>{}</office:presentation></office:body></office:document-content>"#,
            body
        )
    }

    /// Generate meta.xml with metadata
    fn generate_meta_xml(&self) -> String {
        let now = chrono::Utc::now().to_rfc3339();

        let mut meta = format!(
            r#"<?xml version="1.0" encoding="UTF-8"?><office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" office:version="1.3"><office:meta><meta:generator>Litchi/0.0.1</meta:generator><meta:creation-date>{}</meta:creation-date><dc:date>{}</dc:date>"#,
            now, now
        );

        // Add optional metadata fields
        if let Some(ref title) = self.metadata.title {
            meta.push_str(&format!("<dc:title>{}</dc:title>", escape_xml(title)));
        }

        if let Some(ref author) = self.metadata.author {
            meta.push_str(&format!("<dc:creator>{}</dc:creator>", escape_xml(author)));
        }

        meta.push_str("</office:meta>");
        meta.push_str("</office:document-meta>");

        meta
    }

    /// Build the presentation and return as bytes
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi::odf::PresentationBuilder;
    ///
    /// # fn main() -> litchi::Result<()> {
    /// let mut builder = PresentationBuilder::new();
    /// builder.add_slide("Slide content")?;
    /// let bytes = builder.build()?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn build(self) -> Result<Vec<u8>> {
        let mut writer = PackageWriter::new();

        // Set MIME type
        writer.set_mimetype("application/vnd.oasis.opendocument.presentation")?;

        // Add content.xml
        let content_xml = self.generate_content_xml();
        writer.add_file("content.xml", content_xml.as_bytes())?;

        // Add styles.xml
        let styles_xml = OdfStructure::default_styles_xml();
        writer.add_file("styles.xml", styles_xml.as_bytes())?;

        // Add meta.xml
        let meta_xml = self.generate_meta_xml();
        writer.add_file("meta.xml", meta_xml.as_bytes())?;

        // Finish and return bytes
        writer.finish_to_bytes()
    }

    /// Build and save the presentation to a file
    ///
    /// # Arguments
    ///
    /// * `path` - Path where the ODP file should be saved
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi::odf::PresentationBuilder;
    ///
    /// # fn main() -> litchi::Result<()> {
    /// let mut builder = PresentationBuilder::new();
    /// builder.add_slide("Slide content")?;
    /// builder.save("output.odp")?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn save<P: AsRef<Path>>(self, path: P) -> Result<()> {
        let bytes = self.build()?;
        std::fs::write(path, bytes)?;
        Ok(())
    }
}
