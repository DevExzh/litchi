/// Picture (image) shape implementation.
use crate::ooxml::error::{OoxmlError, Result};
use crate::ooxml::pptx::shapes::base::BaseShape;
use quick_xml::events::Event;
use quick_xml::Reader;

/// A picture (image) shape in a presentation.
///
/// Pictures display images on slides and can have various properties
/// like position, size, and relationships to image files.
///
/// # Examples
///
/// ```rust,ignore
/// if let Some(picture) = shape.as_picture() {
///     println!("Picture: {}", picture.name()?);
///     println!("Image rId: {}", picture.image_rId()?);
/// }
/// ```
#[derive(Debug, Clone)]
pub struct Picture {
    /// Base shape properties
    base: BaseShape,
}

impl Picture {
    /// Create a new Picture from XML bytes.
    pub fn new(xml_bytes: Vec<u8>) -> Self {
        Self {
            base: BaseShape::new(xml_bytes, crate::ooxml::pptx::shapes::base::ShapeType::Picture),
        }
    }

    /// Get the base shape.
    #[inline]
    pub fn base(&mut self) -> &mut BaseShape {
        &mut self.base
    }

    /// Get the relationship ID of the embedded image.
    ///
    /// This rId can be used to locate the actual image file in the package.
    ///
    /// # Examples
    ///
    /// ```rust,ignore
    /// let r_id = picture.image_r_id()?;
    /// // Use r_id to get the image from the package
    /// ```
    pub fn image_r_id(&self) -> Result<String> {
        let mut reader = Reader::from_reader(self.base.xml_bytes());
        reader.config_mut().trim_text(true);
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Empty(e)) | Ok(Event::Start(e)) => {
                    // Look for <a:blip r:embed="rId..."/>
                    if e.local_name().as_ref() == b"blip" {
                        for attr in e.attributes().flatten() {
                            let key = attr.key.as_ref();
                            // Check for r:embed attribute
                            if key == b"r:embed" || 
                               (key.starts_with(b"r:") && attr.key.local_name().as_ref() == b"embed") ||
                               attr.key.local_name().as_ref() == b"embed" {
                                let rid = std::str::from_utf8(&attr.value)
                                    .map_err(|e| OoxmlError::Xml(e.to_string()))?;
                                return Ok(rid.to_string());
                            }
                        }
                    }
                }
                Ok(Event::Eof) => break,
                Err(e) => return Err(OoxmlError::Xml(e.to_string())),
                _ => {}
            }
            buf.clear();
        }

        Err(OoxmlError::PartNotFound("Image relationship not found".to_string()))
    }

    /// Get the image filename from the embedded relationship.
    ///
    /// Returns the filename (e.g., "image1.png") if available.
    pub fn image_filename(&self) -> Option<String> {
        // This would need access to the package's relationships
        // For now, we can try to extract from XML if present
        None
    }
}

