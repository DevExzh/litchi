//! Mutable presentation structure for in-place modifications.
//!
//! This module provides a mutable wrapper around ODP presentations that allows
//! for in-place modification of slides, shapes, and content.

use crate::core::{OdfStructure, OwnedPackage, PackageWriter};
use crate::odp::animation::validate_animation_roots;
use crate::odp::legacy_animation::validate_legacy_animation_root;
use crate::odp::media::{EmbeddedMedia, embed_media, validate_package_media_path};
use crate::odp::{MediaReference, Presentation, Shape, Slide};
use litchi_core::{Metadata, Result, xml::escape_xml};
use std::collections::{BTreeMap, BTreeSet};
use std::path::Path;

/// A mutable ODP presentation that supports in-place modifications.
///
/// This struct wraps an ODP presentation and provides methods to modify its content,
/// including adding, updating, and removing slides and shapes.
///
/// # Examples
///
/// ```no_run
/// use litchi_odf::{Presentation, MutablePresentation};
///
/// # fn main() -> litchi_core::Result<()> {
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
    /// Original package retained for copying auxiliary package parts.
    source_package: Option<OwnedPackage>,
    /// Newly embedded package media, keyed by package path.
    media_files: BTreeMap<String, EmbeddedMedia>,
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
    /// use litchi_odf::{Presentation, MutablePresentation};
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let presentation = Presentation::open("slides.odp")?;
    /// let mut mutable = MutablePresentation::from_presentation(presentation)?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn from_presentation(presentation: Presentation) -> Result<Self> {
        let slides = presentation.slides()?;
        let metadata = presentation.metadata()?;
        let mimetype = "application/vnd.oasis.opendocument.presentation".to_string();

        let styles_xml = presentation.styles_xml().map(str::to_owned);
        let source_package = Some(presentation.into_package());

        Ok(Self {
            slides,
            metadata,
            mimetype,
            styles_xml,
            source_package,
            media_files: BTreeMap::new(),
        })
    }

    /// Create a new empty mutable presentation.
    ///
    /// # Examples
    ///
    /// ```
    /// use litchi_odf::MutablePresentation;
    ///
    /// let presentation = MutablePresentation::new();
    /// ```
    pub fn new() -> Self {
        Self {
            slides: Vec::new(),
            metadata: Metadata::default(),
            mimetype: "application/vnd.oasis.opendocument.presentation".to_string(),
            styles_xml: None,
            source_package: None,
            media_files: BTreeMap::new(),
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

    /// Add a package-contained audio or video payload.
    ///
    /// Existing source-package paths cannot be replaced implicitly. The
    /// returned inert reference can be attached with [`Shape::with_media`].
    pub fn embed_media(
        &mut self,
        path: impl Into<String>,
        bytes: impl Into<Vec<u8>>,
        media_type: impl Into<String>,
    ) -> Result<MediaReference> {
        let path = path.into();
        validate_package_media_path(&path)?;
        if let Some(package) = &self.source_package
            && package.has_file(&path)?
        {
            return Err(litchi_core::Error::InvalidFormat(format!(
                "cannot replace existing ODP package media path '{path}' implicitly"
            )));
        }
        embed_media(&mut self.media_files, path, bytes, media_type)
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
    /// use litchi_odf::MutablePresentation;
    ///
    /// # fn main() -> litchi_core::Result<()> {
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
            transition: None,
            animations: Vec::new(),
            legacy_animation: None,
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
    /// use litchi_odf::MutablePresentation;
    ///
    /// # fn main() -> litchi_core::Result<()> {
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
                transition: None,
                animations: Vec::new(),
                legacy_animation: None,
                shapes: Vec::new(),
            };
            self.slides.insert(index, slide);

            // Update indices of subsequent slides
            for i in (index + 1)..self.slides.len() {
                self.slides[i].index = i;
            }

            Ok(())
        } else {
            Err(litchi_core::Error::InvalidFormat(format!(
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
    /// use litchi_odf::MutablePresentation;
    ///
    /// # fn main() -> litchi_core::Result<()> {
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
            Err(litchi_core::Error::InvalidFormat(format!(
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
    /// use litchi_odf::MutablePresentation;
    ///
    /// # fn main() -> litchi_core::Result<()> {
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
            Err(litchi_core::Error::InvalidFormat(format!(
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
    /// use litchi_odf::MutablePresentation;
    ///
    /// # fn main() -> litchi_core::Result<()> {
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
    /// use litchi_odf::{MutablePresentation, Shape};
    ///
    /// # fn main() -> litchi_core::Result<()> {
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
            Err(litchi_core::Error::InvalidFormat(format!(
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
    /// use litchi_odf::MutablePresentation;
    ///
    /// # fn main() -> litchi_core::Result<()> {
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
                Err(litchi_core::Error::InvalidFormat(format!(
                    "Shape index {} out of bounds",
                    shape_index
                )))
            }
        } else {
            Err(litchi_core::Error::InvalidFormat(format!(
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
    /// use litchi_odf::MutablePresentation;
    ///
    /// # fn main() -> litchi_core::Result<()> {
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
            Err(litchi_core::Error::InvalidFormat(format!(
                "Slide index {} out of bounds",
                slide_index
            )))
        }
    }

    /// Generate content.xml from the current mutable state.
    fn generate_content_xml(&self) -> Result<String> {
        let mut extension_uris = BTreeSet::new();
        for slide in &self.slides {
            validate_animation_roots(&slide.animations)?;
            for animation in &slide.animations {
                animation.collect_extension_namespaces(&mut extension_uris);
            }
            if let Some(animation) = &slide.legacy_animation {
                validate_legacy_animation_root(animation)?;
                animation.collect_extension_namespaces(&mut extension_uris);
            }
        }
        let extension_namespaces = extension_uris
            .into_iter()
            .enumerate()
            .map(|(index, uri)| (uri, format!("anim-ext{}", index + 1)))
            .collect::<BTreeMap<_, _>>();
        let mut extension_declarations = String::new();
        for (uri, prefix) in &extension_namespaces {
            if uri.is_empty() {
                return Err(litchi_core::Error::InvalidFormat(
                    "animation extension namespace URI cannot be empty".to_string(),
                ));
            }
            extension_declarations.push_str(" xmlns:");
            extension_declarations.push_str(prefix);
            extension_declarations.push_str("=\"");
            extension_declarations.push_str(&escape_xml(uri));
            extension_declarations.push('"');
        }

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
            let page_num = i + 1;
            let slide_style = super::builder::slide_style_name(slide, i);
            body.push_str(&xml_minifier::minified_xml_format!(
                r#"<draw:page draw:name="page{}" draw:style-name="{}" draw:master-page-name="Default">"#,
                page_num,
                slide_style
            ));

            // Add title frame if title exists
            if let Some(ref title) = slide.title {
                let title_paragraphs = super::builder::generate_text_paragraphs(title, Some("P1"));
                body.push_str(&xml_minifier::minified_xml_format!(
                    r#"<draw:frame draw:style-name="gr1" draw:text-style-name="P1" draw:layer="layout" presentation:class="title" svg:width="25.199cm" svg:height="3.506cm" svg:x="1.4cm" svg:y="0.962cm"><draw:text-box>{}</draw:text-box></draw:frame>"#,
                    title_paragraphs
                ));
            }

            // Add text frame
            if !slide.text.is_empty() {
                let y_position = if slide.title.is_some() {
                    "5.0cm"
                } else {
                    "2.0cm"
                };
                let text_paragraphs =
                    super::builder::generate_text_paragraphs(&slide.text, Some("P2"));
                body.push_str(&xml_minifier::minified_xml_format!(
                    r#"<draw:frame draw:style-name="gr2" draw:text-style-name="P2" draw:layer="layout" presentation:class="object" svg:width="25.199cm" svg:height="10cm" svg:x="1.4cm" svg:y="{}"><draw:text-box>{}</draw:text-box></draw:frame>"#,
                    y_position,
                    text_paragraphs
                ));
            }

            // Add shapes
            for (shape_idx, shape) in slide.shapes.iter().enumerate() {
                body.push_str(&super::builder::PresentationBuilder::generate_shape_xml(
                    shape, shape_idx,
                )?);
            }

            for animation in &slide.animations {
                animation.write_xml(&mut body, &extension_namespaces)?;
            }
            if let Some(animation) = &slide.legacy_animation {
                animation.write_xml(&mut body, &extension_namespaces)?;
            }

            body.push_str(&super::builder::PresentationBuilder::generate_notes_xml(
                slide.notes.as_deref(),
            ));

            body.push_str("</draw:page>");
        }

        let transition_styles = super::builder::generate_transition_styles(&self.slides);
        Ok(format!(
            r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0" xmlns:anim="urn:oasis:names:tc:opendocument:xmlns:animation:1.0" xmlns:smil="urn:oasis:names:tc:opendocument:xmlns:smil-compatible:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:xlink="http://www.w3.org/1999/xlink"{} office:version="1.3"><office:scripts/><office:font-face-decls/><office:automatic-styles>{}</office:automatic-styles><office:body><office:presentation>{}</office:presentation></office:body></office:document-content>"#,
            extension_declarations, transition_styles, body
        ))
    }

    /// Generate meta.xml with current metadata.
    fn generate_meta_xml(&self) -> String {
        let now = chrono::Utc::now().to_rfc3339();
        let mut estimated = 64usize;
        estimated += self.metadata.title.as_ref().map(|s| s.len()).unwrap_or(0);
        estimated += self.metadata.author.as_ref().map(|s| s.len()).unwrap_or(0);
        estimated += self.metadata.subject.as_ref().map(|s| s.len()).unwrap_or(0);
        estimated += self
            .metadata
            .description
            .as_ref()
            .map(|s| s.len())
            .unwrap_or(0);
        estimated += self
            .metadata
            .keywords
            .as_ref()
            .map(|s| s.len())
            .unwrap_or(0);
        let mut meta_fields = String::with_capacity(estimated);

        // Add optional metadata fields
        if let Some(ref title) = self.metadata.title {
            let escaped_title = escape_xml(title);
            meta_fields.push_str(&xml_minifier::minified_xml_format!(
                r#"<dc:title>{}</dc:title>"#,
                escaped_title
            ));
        }

        if let Some(ref author) = self.metadata.author {
            let escaped_author = escape_xml(author);
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

    /// Save the modified presentation to a file.
    ///
    /// # Arguments
    ///
    /// * `path` - Path where the ODP file should be saved
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi_odf::MutablePresentation;
    ///
    /// # fn main() -> litchi_core::Result<()> {
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
    /// use litchi_odf::MutablePresentation;
    ///
    /// # fn main() -> litchi_core::Result<()> {
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
        let content_xml = self.generate_content_xml()?;
        writer.add_file("content.xml", content_xml.as_bytes())?;

        // Add styles.xml (preserved or default)
        let default_styles = OdfStructure::default_styles_xml();
        let styles_xml = self.styles_xml.as_deref().unwrap_or(&default_styles);
        writer.add_file("styles.xml", styles_xml.as_bytes())?;

        // Add meta.xml (regenerated with current metadata)
        let meta_xml = self.generate_meta_xml();
        writer.add_file("meta.xml", meta_xml.as_bytes())?;

        for (path, media) in &self.media_files {
            writer.add_file_with_media_type(path, &media.bytes, &media.media_type)?;
        }

        if let Some(package) = &self.source_package {
            writer.copy_auxiliary_files_from(package)?;
        }

        writer.finish_to_bytes()
    }
}

impl Default for MutablePresentation {
    fn default() -> Self {
        Self::new()
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::odp::{
        AnimationAttribute, AnimationAttributeNamespace, AnimationKind, AnimationNode,
        LegacyAnimationKind, LegacyAnimationNode, PresentationBuilder,
    };

    const STYLES: &str = r#"<?xml version="1.0"?><office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><office:styles><office:marker>preserve-me</office:marker></office:styles></office:document-styles>"#;
    const SETTINGS: &[u8] = b"<settings>presentation-settings</settings>";
    const IMAGE: &[u8] = b"\x89PNG\r\n\x1a\nimage-payload";
    const CUSTOM: &[u8] = b"custom-presentation-data";

    fn presentation_bytes_with_image() -> Vec<u8> {
        let content = r#"<?xml version="1.0"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:xlink="http://www.w3.org/1999/xlink"><office:body><office:presentation><draw:page draw:name="Media"><draw:frame presentation:class="title"><draw:text-box><text:p>Visible Title</text:p></draw:text-box></draw:frame><draw:frame presentation:class="object"><draw:text-box><text:p>Body &amp; more</text:p></draw:text-box></draw:frame><draw:frame draw:name="Photo" svg:x="1cm" svg:y="2cm" svg:width="3cm" svg:height="4cm"><draw:image xlink:href="Pictures/a&amp;b.png" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/></draw:frame><draw:rect draw:name="Labeled"><text:p>Shape label</text:p></draw:rect><draw:connector draw:name="Link" svg:x1="0cm" svg:y1="0cm" svg:x2="2cm" svg:y2="2cm"/><draw:line draw:name="Rule" svg:x1="1cm" svg:y1="1cm" svg:x2="5cm" svg:y2="1cm"/><presentation:notes><draw:frame><draw:text-box><text:p>Speaker note</text:p></draw:text-box></draw:frame></presentation:notes></draw:page></office:presentation></office:body></office:document-content>"#;
        let mut writer = PackageWriter::new();
        writer
            .set_mimetype("application/vnd.oasis.opendocument.presentation")
            .unwrap();
        writer.add_file("content.xml", content.as_bytes()).unwrap();
        writer.add_file("styles.xml", STYLES.as_bytes()).unwrap();
        writer.add_file("settings.xml", SETTINGS).unwrap();
        writer.add_file("Pictures/a&b.png", IMAGE).unwrap();
        writer
            .add_manifest_entry("Object 1/", "application/vnd.oasis.opendocument.text")
            .unwrap();
        writer
            .add_file_with_media_type("custom/data.bin", CUSTOM, "application/x-odp-test")
            .unwrap();
        writer.finish_to_bytes().unwrap()
    }

    #[test]
    fn mutable_presentation_round_trips_images_styles_and_settings() {
        let source_bytes = presentation_bytes_with_image();
        let presentation = Presentation::from_bytes(source_bytes.clone()).unwrap();
        assert_eq!(presentation.to_bytes().unwrap(), source_bytes);
        let source_shapes = presentation.slides().unwrap().remove(0).shapes;
        assert_eq!(source_shapes[0].image_href(), Some("Pictures/a&b.png"));

        let mutable = MutablePresentation::from_presentation(presentation).unwrap();
        let bytes = mutable.to_bytes().unwrap();
        let package = OwnedPackage::from_bytes(bytes.clone()).unwrap();

        assert_eq!(package.get_file("Pictures/a&b.png").unwrap(), IMAGE);
        assert_eq!(package.get_file("settings.xml").unwrap(), SETTINGS);
        assert_eq!(package.get_file("styles.xml").unwrap(), STYLES.as_bytes());
        assert_eq!(package.get_file("custom/data.bin").unwrap(), CUSTOM);
        let borrowed = package.package().unwrap();
        assert_eq!(
            borrowed.manifest().get_media_type("Object 1/"),
            Some("application/vnd.oasis.opendocument.text")
        );
        assert_eq!(
            borrowed.manifest().get_media_type("custom/data.bin"),
            Some("application/x-odp-test")
        );

        let content = String::from_utf8(package.get_file("content.xml").unwrap()).unwrap();
        assert!(content.contains("<draw:image"));
        assert!(content.contains(r#"xlink:href="Pictures/a&amp;b.png""#));
        assert!(content.contains(r#"xlink:show="embed""#));
        assert!(content.contains("<draw:line"));
        assert!(content.contains("<draw:rect"));
        assert!(content.contains("<draw:connector"));
        assert_eq!(content.matches("Visible Title").count(), 1);
        assert_eq!(content.matches("Body &amp; more").count(), 1);
        assert_eq!(content.matches("Shape label").count(), 1);
        assert_eq!(content.matches("Speaker note").count(), 1);

        let reparsed = Presentation::from_bytes(bytes).unwrap();
        let slides = reparsed.slides().unwrap();
        assert_eq!(slides[0].title.as_deref(), Some("Visible Title"));
        assert_eq!(slides[0].text, "Body & more");
        assert_eq!(slides[0].notes.as_deref(), Some("Speaker note"));
        assert_eq!(
            slides[0].all_text(),
            "Visible Title\nBody & more\nShape label"
        );
        let picture = slides[0]
            .shapes
            .iter()
            .find(|shape| shape.shape_type == litchi_core::ShapeType::Picture)
            .unwrap();
        assert_eq!(picture.image_href(), Some("Pictures/a&b.png"));
    }

    #[test]
    fn builder_and_mutable_presentation_round_trip_animation_trees() {
        let mut parameter = AnimationNode::new(AnimationKind::Parameter);
        parameter.set_attribute(
            AnimationAttribute::new(
                AnimationAttributeNamespace::Animation,
                "name",
                "destination",
            )
            .unwrap(),
        );
        parameter.set_attribute(
            AnimationAttribute::new(AnimationAttributeNamespace::Animation, "value", "2 & next")
                .unwrap(),
        );
        let mut command = AnimationNode::new(AnimationKind::Command);
        command.set_attribute(
            AnimationAttribute::new(AnimationAttributeNamespace::Animation, "command", "show")
                .unwrap(),
        );
        command.add_child(parameter).unwrap();

        let mut root = AnimationNode::new(AnimationKind::Sequence);
        root.set_attribute(
            AnimationAttribute::new(AnimationAttributeNamespace::Smil, "begin", "slide.begin")
                .unwrap(),
        );
        root.set_attribute(
            AnimationAttribute::new(
                AnimationAttributeNamespace::Other("urn:example:timing".to_string()),
                "mode",
                "author-defined",
            )
            .unwrap(),
        );
        root.add_child(command).unwrap();
        root.add_child(AnimationNode::new(AnimationKind::TransitionFilter))
            .unwrap();

        let slide = Slide {
            title: Some("Animated".to_string()),
            text: String::new(),
            index: 0,
            notes: None,
            transition: None,
            animations: vec![root.clone()],
            legacy_animation: None,
            shapes: Vec::new(),
        };
        let mut builder = PresentationBuilder::new();
        builder.add_slide_element(slide).unwrap();
        let built = builder.build().unwrap();
        let presentation = Presentation::from_bytes(built).unwrap();
        assert_eq!(presentation.slides().unwrap()[0].animations, [root.clone()]);

        let mutable = MutablePresentation::from_presentation(presentation).unwrap();
        let bytes = mutable.to_bytes().unwrap();
        let package = OwnedPackage::from_bytes(bytes.clone()).unwrap();
        let content = String::from_utf8(package.get_file("content.xml").unwrap()).unwrap();
        assert!(content.contains(r#"xmlns:anim-ext1="urn:example:timing""#));
        assert!(content.contains(r#"anim-ext1:mode="author-defined""#));
        assert!(content.contains(r#"anim:value="2 &amp; next""#));

        let reparsed = Presentation::from_bytes(bytes).unwrap();
        assert_eq!(reparsed.slides().unwrap()[0].animations, [root]);
    }

    #[test]
    fn rejects_invalid_mutated_animation_trees_and_xml_characters() {
        let mut leaf = AnimationNode::new(AnimationKind::Animate);
        leaf.children_mut()
            .push(AnimationNode::new(AnimationKind::Set));
        let mut presentation = MutablePresentation::new();
        presentation.add_slide("Invalid", "").unwrap();
        presentation.slides_mut()[0].animations.push(leaf);
        assert!(presentation.to_bytes().is_err());

        assert!(
            AnimationAttribute::new(AnimationAttributeNamespace::Smil, "begin", "bad\0value")
                .is_err()
        );
        assert!(
            AnimationAttribute::new(
                AnimationAttributeNamespace::Other(
                    "http://www.w3.org/XML/1998/namespace".to_string()
                ),
                "id",
                "bad namespace variant"
            )
            .is_err()
        );
    }

    #[test]
    fn mutable_presentation_preserves_and_adds_embedded_media() {
        const ORIGINAL: &[u8] = b"original-video";
        const ADDED: &[u8] = b"added-audio";
        let mut builder = PresentationBuilder::new();
        let original = builder
            .embed_media("Media/original.mp4", ORIGINAL, "video/mp4")
            .unwrap();
        builder
            .add_slide_element(Slide {
                title: None,
                text: String::new(),
                index: 0,
                notes: None,
                transition: None,
                animations: Vec::new(),
                legacy_animation: None,
                shapes: vec![Shape::new().with_media(original)],
            })
            .unwrap();
        let presentation = Presentation::from_bytes(builder.build().unwrap()).unwrap();
        let mut mutable = MutablePresentation::from_presentation(presentation).unwrap();
        assert!(
            mutable
                .embed_media("Media/original.mp4", b"replacement", "video/mp4")
                .is_err()
        );
        let added = mutable
            .embed_media("Media/added.ogg", ADDED, "audio/ogg")
            .unwrap();
        mutable
            .add_shape(0, Shape::new().with_media(added))
            .unwrap();

        let bytes = mutable.to_bytes().unwrap();
        let package = OwnedPackage::from_bytes(bytes.clone()).unwrap();
        assert_eq!(package.get_file("Media/original.mp4").unwrap(), ORIGINAL);
        assert_eq!(package.get_file("Media/added.ogg").unwrap(), ADDED);
        assert_eq!(
            package
                .package()
                .unwrap()
                .manifest()
                .get_media_type("Media/added.ogg"),
            Some("audio/ogg")
        );

        let reparsed = Presentation::from_bytes(bytes).unwrap();
        let slides = reparsed.slides().unwrap();
        assert_eq!(slides[0].shapes.len(), 2);
        assert_eq!(
            slides[0].shapes[0].media().unwrap().href(),
            "Media/original.mp4"
        );
        assert_eq!(
            slides[0].shapes[1].media().unwrap().href(),
            "Media/added.ogg"
        );
    }

    #[test]
    fn builder_and_mutable_round_trip_legacy_presentation_effects() {
        let attr =
            |namespace, name, value| AnimationAttribute::new(namespace, name, value).unwrap();
        let mut sound = LegacyAnimationNode::new(LegacyAnimationKind::Sound);
        sound.set_attribute(attr(
            AnimationAttributeNamespace::Xlink,
            "href",
            "Sounds/chime.ogg",
        ));
        sound.set_attribute(attr(AnimationAttributeNamespace::Xlink, "type", "simple"));
        let mut show = LegacyAnimationNode::new(LegacyAnimationKind::ShowShape);
        show.set_attribute(attr(
            AnimationAttributeNamespace::Draw,
            "shape-id",
            "shape1",
        ));
        show.set_attribute(attr(
            AnimationAttributeNamespace::Presentation,
            "effect",
            "fade",
        ));
        show.add_child(sound).unwrap();
        let mut root = LegacyAnimationNode::new(LegacyAnimationKind::Animations);
        root.set_attribute(attr(
            AnimationAttributeNamespace::Other("urn:example:legacy-effects".to_string()),
            "mode",
            "preserve",
        ));
        root.add_child(show).unwrap();

        let mut builder = PresentationBuilder::new();
        builder
            .add_slide_element(Slide {
                title: None,
                text: String::new(),
                index: 0,
                notes: None,
                transition: None,
                animations: Vec::new(),
                legacy_animation: Some(root.clone()),
                shapes: Vec::new(),
            })
            .unwrap();
        let presentation = Presentation::from_bytes(builder.build().unwrap()).unwrap();
        assert_eq!(
            presentation.slides().unwrap()[0].legacy_animation(),
            Some(&root)
        );

        let mutable = MutablePresentation::from_presentation(presentation).unwrap();
        let reparsed = Presentation::from_bytes(mutable.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reparsed.slides().unwrap()[0].legacy_animation(),
            Some(&root)
        );
    }
}
