//! OpenDocument Presentation (.odp) support.
//!
//! This module provides a unified API for working with OpenDocument presentations,
//! equivalent to Microsoft PowerPoint presentations.

use crate::common::{Error, Result, Metadata};
use crate::odf::core::{Content, Meta, Package, Styles, Manifest};
use std::io::Cursor;
use std::path::Path;

/// An OpenDocument presentation (.odp)
pub struct Presentation {
    _package: Package<Cursor<Vec<u8>>>,
    _content: Content,
    _styles: Option<Styles>,
    _meta: Option<Meta>,
    _manifest: Manifest,
}

impl Presentation {
    /// Open an ODP presentation from a file path
    pub fn open<P: AsRef<Path>>(path: P) -> Result<Self> {
        let path = path.as_ref();

        // Read the entire file into memory
        let bytes = std::fs::read(path)?;
        Self::from_bytes(bytes)
    }

    /// Create a Presentation from a byte buffer
    pub fn from_bytes(bytes: Vec<u8>) -> Result<Self> {
        let cursor = Cursor::new(bytes);
        let mut package = Package::from_reader(cursor)?;

        // Verify this is a presentation
        let mime_type = package.mimetype();
        if !mime_type.contains("opendocument.presentation") {
            return Err(Error::InvalidFormat(format!(
                "Not an ODP file: MIME type is {}", mime_type
            )));
        }

        // Parse core components
        let content_bytes = package.get_file("content.xml")?;
        let content = Content::from_bytes(&content_bytes)?;

        let styles = if package.has_file("styles.xml") {
            let styles_bytes = package.get_file("styles.xml")?;
            Some(Styles::from_bytes(&styles_bytes)?)
        } else {
            None
        };

        let meta = if package.has_file("meta.xml") {
            let meta_bytes = package.get_file("meta.xml")?;
            Some(Meta::from_bytes(&meta_bytes)?)
        } else {
            None
        };

        let manifest = package.manifest().clone();

        Ok(Self {
            _package: package,
            _content: content,
            _styles: styles,
            _meta: meta,
            _manifest: manifest,
        })
    }

    /// Get the number of slides in the presentation
    pub fn slide_count(&mut self) -> Result<usize> {
        let slides = self.slides()?;
        Ok(slides.len())
    }

    /// Get all slides in the presentation
    pub fn slides(&mut self) -> Result<Vec<Slide>> {
        let content_bytes = self._package.get_file("content.xml")?;
        let content = Content::from_bytes(&content_bytes)?;

        use quick_xml::events::Event;
        use quick_xml::Reader;

        let mut reader = Reader::from_str(content.xml_content());
        let mut buf = Vec::new();
        let mut slides = Vec::new();

        // Simple state tracking
        let mut current_slide_text = String::new();
        let mut current_slide_title: Option<String> = None;
        let mut in_slide = false;
        let mut slide_index = 0;

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) => {
                    match e.name().as_ref() {
                        b"draw:page" => {
                            // Start new slide
                            if in_slide && !current_slide_text.trim().is_empty() {
                                slides.push(Slide {
                                    title: current_slide_title.clone(),
                                    text: current_slide_text.trim().to_string(),
                                    index: slide_index,
                                    notes: None,
                                });
                                slide_index += 1;
                            }

                            // Extract slide name
                            current_slide_title = None;
                            for attr_result in e.attributes() {
                                if let Ok(attr) = attr_result
                                    && attr.key.as_ref() == b"draw:name" {
                                        if let Ok(name) = String::from_utf8(attr.value.to_vec()) {
                                            current_slide_title = Some(name);
                                        }
                                        break;
                                    }
                            }
                            if current_slide_title.is_none() {
                                current_slide_title = Some(format!("Slide{}", slide_index + 1));
                            }

                            current_slide_text.clear();
                            in_slide = true;
                        }
                        b"text:p" | b"text:span" => {
                            // Continue - text will be collected in Text event
                        }
                        _ => {}
                    }
                }
                Ok(Event::Text(ref t)) => {
                    if in_slide
                        && let Ok(text) = String::from_utf8(t.to_vec())
                            && !text.trim().is_empty() {
                                if !current_slide_text.is_empty() {
                                    current_slide_text.push(' ');
                                }
                                current_slide_text.push_str(text.trim());
                            }
                }
                Ok(Event::End(ref e)) => {
                    if e.name().as_ref() == b"draw:page" {
                        // Finish current slide
                        if in_slide && !current_slide_text.trim().is_empty() {
                            slides.push(Slide {
                                title: current_slide_title.clone(),
                                text: current_slide_text.trim().to_string(),
                                index: slide_index,
                                notes: None,
                            });
                            slide_index += 1;
                        }
                        in_slide = false;
                    }
                }
                Ok(Event::Eof) => break,
                Err(e) => return Err(crate::common::Error::InvalidFormat(format!("XML parsing error: {}", e))),
                _ => {}
            }
            buf.clear();
        }

        // Handle any remaining slide
        if in_slide && !current_slide_text.trim().is_empty() {
            slides.push(Slide {
                title: current_slide_title,
                text: current_slide_text.trim().to_string(),
                index: slide_index,
                notes: None,
            });
        }

        Ok(slides)
    }

    /// Get a slide by index
    pub fn slide(&mut self, index: usize) -> Result<Option<Slide>> {
        let slides = self.slides()?;
        Ok(slides.into_iter().nth(index))
    }

    /// Extract all text content from the presentation
    pub fn text(&mut self) -> Result<String> {
        let slides = self.slides()?;
        let mut all_text = Vec::new();

        for slide in slides {
            if !slide.text.trim().is_empty() {
                all_text.push(slide.text.trim().to_string());
            }
        }

        Ok(all_text.join("\n\n"))
    }

    /// Get document metadata
    pub fn metadata(&self) -> Result<Metadata> {
        if let Some(meta) = &self._meta {
            Ok(meta.extract_metadata())
        } else {
            Ok(Metadata::default())
        }
    }
}

/// A slide in an ODP presentation
#[derive(Clone)]
pub struct Slide {
    pub title: Option<String>,
    pub text: String,
    pub index: usize,
    pub notes: Option<String>,
}

impl Slide {
    /// Get the title of the slide
    pub fn title(&self) -> Result<Option<String>> {
        Ok(self.title.clone())
    }

    /// Extract all text content from the slide
    pub fn text(&self) -> Result<String> {
        Ok(self.text.clone())
    }

    /// Get all shapes on the slide
    pub fn shapes(&self) -> Result<Vec<Shape>> {
        Ok(Vec::new()) // TODO: Implement shape extraction
    }

    /// Get the slide index
    pub fn index(&self) -> usize {
        self.index
    }

    /// Get the slide name/notes
    pub fn notes(&self) -> Result<Option<String>> {
        Ok(self.notes.clone())
    }
}


/// A shape (element) on a slide
pub struct Shape {
    // Placeholder for shape implementation
}

impl Shape {
    /// Get the text content of the shape
    pub fn text(&self) -> Result<String> {
        Ok(String::new()) // Placeholder
    }

    /// Get the shape type
    pub fn shape_type(&self) -> crate::common::ShapeType {
        crate::common::ShapeType::AutoShape // Placeholder
    }

    /// Check if this is a text shape
    pub fn has_text(&self) -> bool {
        false // Placeholder
    }
}

