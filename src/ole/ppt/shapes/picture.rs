// Picture shape with integrated BLIP extraction support
//
// This module provides Picture shape support with the ability to extract
// embedded images directly from shapes, similar to python-pptx.

use super::shape::{Shape, ShapeProperties, ShapeType};
#[cfg(feature = "imgconv")]
use crate::common::error::Result;
#[cfg(feature = "imgconv")]
use crate::images::{Blip, ExtractedImage};
#[cfg(feature = "imgconv")]
use crate::ole::ppt::escher::EscherContainer;
use crate::ole::ppt::package::PptError;

/// Picture shape containing an embedded image
///
/// Represents an image embedded in a PowerPoint slide, with methods
/// to extract the underlying BLIP data.
///
/// # Example
/// ```no_run
/// use litchi::ole::ppt::Package;
///
/// let mut pkg = Package::open("presentation.ppt")?;
/// let mut pres = pkg.presentation()?;
///
/// for slide in pres.slides()? {
///     for shape in slide.shapes()? {
///         if let Some(picture) = shape.as_picture() {
///             // Extract the image
///             if let Ok(Some(image)) = picture.extract_image(&pres) {
///                 let png_data = image.to_png(None, None)?;
///                 std::fs::write(image.suggested_filename(), png_data)?;
///             }
///         }
///     }
/// }
/// # Ok::<(), Box<dyn std::error::Error>>(())
/// ```
#[derive(Debug, Clone)]
pub struct PictureShape {
    /// Shape properties
    pub properties: ShapeProperties,
    /// BLIP ID reference (index into BLIP store)
    pub blip_id: Option<u32>,
    /// Picture name/filename
    pub name: Option<String>,
    /// Escher container data (for extracting BLIP)
    #[cfg(feature = "imgconv")]
    escher_data: Option<Vec<u8>>,
}

impl PictureShape {
    /// Create a new picture shape
    pub fn new(id: u32) -> Self {
        let properties = ShapeProperties {
            id,
            shape_type: ShapeType::Picture,
            ..Default::default()
        };

        Self {
            properties,
            blip_id: None,
            name: None,
            #[cfg(feature = "imgconv")]
            escher_data: None,
        }
    }

    /// Create from shape properties
    pub fn from_properties(properties: ShapeProperties) -> Self {
        Self {
            properties,
            blip_id: None,
            name: None,
            #[cfg(feature = "imgconv")]
            escher_data: None,
        }
    }

    /// Set BLIP ID reference
    pub fn set_blip_id(&mut self, blip_id: u32) {
        self.blip_id = Some(blip_id);
    }

    /// Set picture name
    pub fn set_name<S: Into<String>>(&mut self, name: S) {
        self.name = Some(name.into());
    }

    /// Set picture bounds (position and size)
    pub fn set_bounds(&mut self, left: i32, top: i32, width: i32, height: i32) {
        self.properties.x = left;
        self.properties.y = top;
        self.properties.width = width;
        self.properties.height = height;
    }

    /// Set Escher container data for BLIP extraction
    #[cfg(feature = "imgconv")]
    pub fn set_escher_data(&mut self, data: Vec<u8>) {
        self.escher_data = Some(data);
    }

    /// Get BLIP ID
    pub const fn blip_id(&self) -> Option<u32> {
        self.blip_id
    }

    /// Get picture name
    pub fn name(&self) -> Option<&str> {
        self.name.as_deref()
    }

    /// Extract the embedded image from this picture shape
    ///
    /// This method attempts to extract the BLIP data in two ways:
    /// 1. From the shape's Escher container (embedded BLIP)
    /// 2. From the presentation's Pictures stream (referenced BLIP)
    ///
    /// # Arguments
    /// * `presentation` - The presentation containing this shape
    ///
    /// # Returns
    /// The extracted image, or None if no image data is found
    #[cfg(feature = "imgconv")]
    pub fn extract_image(
        &self,
        presentation: &crate::ole::ppt::Presentation,
    ) -> Result<Option<ExtractedImage<'static>>> {
        // Try to extract from embedded Escher data first
        if let Some(ref escher_data) = self.escher_data
            && let Ok(images) = crate::images::ImageExtractor::extract_blips(escher_data)
            && let Some(image) = images.into_iter().next()
        {
            return Ok(Some(image));
        }

        // Try to extract from Pictures stream using BLIP ID
        if let Some(blip_id) = self.blip_id {
            return presentation.extract_image_by_blip_id(blip_id).map_err(|e| {
                crate::common::error::Error::ParseError(format!(
                    "Failed to extract image by BLIP ID: {}",
                    e
                ))
            });
        }

        Ok(None)
    }

    /// Extract the BLIP directly from embedded Escher data
    ///
    /// This is a lower-level method that extracts the BLIP without
    /// requiring access to the full presentation.
    #[cfg(feature = "imgconv")]
    pub fn extract_blip_from_escher(&self) -> Result<Option<Blip<'static>>> {
        if let Some(ref escher_data) = self.escher_data {
            let images = crate::images::ImageExtractor::extract_blips(escher_data)?;
            if let Some(image) = images.into_iter().next() {
                return Ok(Some(image.blip));
            }
        }
        Ok(None)
    }

    /// Get the suggested filename for this picture
    pub fn suggested_filename(&self) -> String {
        if let Some(name) = &self.name {
            if name.contains('.') {
                return name.clone();
            }
            // If name doesn't have extension, add one based on type
            format!("{}.png", name)
        } else {
            format!("picture_{}.png", self.properties.id)
        }
    }
}

impl Shape for PictureShape {
    fn properties(&self) -> &ShapeProperties {
        &self.properties
    }

    fn properties_mut(&mut self) -> &mut ShapeProperties {
        &mut self.properties
    }

    fn text(&self) -> std::result::Result<String, PptError> {
        Ok(String::new()) // Pictures don't have text
    }

    fn has_text(&self) -> bool {
        false
    }

    fn clone_box(&self) -> Box<dyn Shape> {
        Box::new(self.clone())
    }

    fn as_any(&self) -> &dyn std::any::Any {
        self
    }
}

/// Helper function to extract BLIP ID from Escher properties
///
/// Searches for the BlipToDisplay property (0x0104) which contains
/// the reference to the BLIP in the BStoreContainer.
#[cfg(feature = "imgconv")]
pub fn extract_blip_id_from_escher(container: &EscherContainer) -> Option<u32> {
    use crate::ole::ppt::escher::EscherRecordType;

    // Look for shape options (Opt record)
    for child in container.children().flatten() {
        if child.record_type == EscherRecordType::Opt {
            // Parse properties from the Opt record
            let prop_count = child.instance;
            if let Ok(properties) =
                crate::ole::ppt::shapes::escher::EscherProperty::parse_properties(
                    child.data, prop_count,
                )
            {
                // Look for BlipToDisplay property (0x0104)
                for prop in properties {
                    if prop.is_blip_id() && prop.property_number() == 0x0104 {
                        return Some(prop.data);
                    }
                }
            }
        }
    }

    None
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_picture_shape_creation() {
        let picture = PictureShape::new(1);
        assert_eq!(picture.properties.id, 1);
        assert_eq!(picture.properties.shape_type, ShapeType::Picture);
        assert!(picture.blip_id.is_none());
        assert!(picture.name.is_none());
    }

    #[test]
    fn test_picture_shape_set_blip_id() {
        let mut picture = PictureShape::new(1);
        picture.set_blip_id(42);
        assert_eq!(picture.blip_id(), Some(42));
    }

    #[test]
    fn test_picture_shape_set_name() {
        let mut picture = PictureShape::new(1);
        picture.set_name("image.jpg");
        assert_eq!(picture.name(), Some("image.jpg"));
        assert_eq!(picture.suggested_filename(), "image.jpg");
    }

    #[test]
    fn test_picture_shape_suggested_filename() {
        let mut picture = PictureShape::new(5);

        // No name set
        assert_eq!(picture.suggested_filename(), "picture_5.png");

        // Name without extension
        picture.set_name("photo");
        assert_eq!(picture.suggested_filename(), "photo.png");

        // Name with extension
        picture.set_name("photo.jpg");
        assert_eq!(picture.suggested_filename(), "photo.jpg");
    }
}
