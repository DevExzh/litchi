// Picture shape with integrated BLIP extraction support
//
// This module provides Picture shape support with the ability to extract
// embedded images directly from shapes, similar to python-pptx.

use super::shape::{Shape, ShapeProperties, ShapeType};
use crate::package::Error;
use litchi_core::error::Result;
use litchi_odraw::Container;
use litchi_odraw::image::File as ImageFile;
use litchi_odraw::image::Id as ImageId;

/// Semantic kind of a PowerPoint picture frame.
///
/// Binary PowerPoint uses the same OfficeArt frame shape for ordinary images,
/// embedded OLE objects, and media previews. The client-data records distinguish
/// the three cases.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum PictureFrameKind {
    /// An ordinary embedded or linked image.
    #[default]
    Picture,
    /// A preview for an embedded or linked OLE object.
    OleObject,
    /// A preview frame for audio or video media.
    Media,
}

/// Picture shape containing an embedded image
///
/// Represents an image embedded in a PowerPoint slide, with methods
/// to extract the underlying BLIP data.
///
/// # Example
/// ```ignore
/// use litchi_imgconv::Convert as _;
/// use litchi_ppt::Package;
///
/// let mut pkg = Package::open("presentation.ppt")?;
/// let mut pres = pkg.presentation()?;
///
/// for slide in pres.slides()? {
///     for shape in slide.shapes()? {
///         if let Some(picture) = shape.as_picture() {
///             // Extract the image
///             if let Ok(Some(image)) = picture.extract_image(&pres) {
///                 let data = image.extract(Default::default())?;
///                 std::fs::write(image.out_name(), data)?;
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
    pub blip_id: Option<ImageId>,
    /// Picture name/filename
    pub name: Option<String>,
    /// Semantic kind of this picture frame.
    frame_kind: PictureFrameKind,
    /// Reference to an external OLE or media object.
    external_object_id: Option<u32>,
    /// Escher container data (for extracting BLIP)
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
            frame_kind: PictureFrameKind::Picture,
            external_object_id: None,
            escher_data: None,
        }
    }

    /// Create from shape properties
    pub fn from_properties(properties: ShapeProperties) -> Self {
        let frame_kind = match properties.shape_type {
            ShapeType::Object => PictureFrameKind::OleObject,
            ShapeType::Media => PictureFrameKind::Media,
            _ => PictureFrameKind::Picture,
        };
        Self {
            properties,
            blip_id: None,
            name: None,
            frame_kind,
            external_object_id: None,
            escher_data: None,
        }
    }

    /// Set BLIP ID reference
    pub fn set_blip_id(&mut self, id: ImageId) {
        self.blip_id = Some(id);
    }

    /// Set a BLIP ID from a raw host property after validating its range.
    pub fn set_blip_index(&mut self, index: u32) -> litchi_odraw::Result<()> {
        self.set_blip_id(ImageId::new(index)?);
        Ok(())
    }

    /// Set picture name
    pub fn set_name<S: Into<String>>(&mut self, name: S) {
        self.name = Some(name.into());
    }

    /// Set the semantic frame kind and keep the common shape type in sync.
    pub fn set_frame_kind(&mut self, frame_kind: PictureFrameKind) {
        self.frame_kind = frame_kind;
        self.properties.shape_type = match frame_kind {
            PictureFrameKind::Picture => ShapeType::Picture,
            PictureFrameKind::OleObject => ShapeType::Object,
            PictureFrameKind::Media => ShapeType::Media,
        };
    }

    /// Set the external OLE or media object reference.
    pub fn set_external_object_id(&mut self, external_object_id: u32) {
        self.external_object_id = Some(external_object_id);
    }

    /// Set picture bounds (position and size)
    pub fn set_bounds(&mut self, left: i32, top: i32, width: i32, height: i32) {
        self.properties.x = left;
        self.properties.y = top;
        self.properties.width = width;
        self.properties.height = height;
    }

    /// Set Escher container data for BLIP extraction
    pub fn set_escher_data(&mut self, data: Vec<u8>) {
        self.escher_data = Some(data);
    }

    /// Get BLIP ID
    pub const fn blip_id(&self) -> Option<ImageId> {
        self.blip_id
    }

    /// Get picture name
    pub fn name(&self) -> Option<&str> {
        self.name.as_deref()
    }

    /// Get the semantic frame kind.
    pub const fn frame_kind(&self) -> PictureFrameKind {
        self.frame_kind
    }

    /// Get the external OLE or media object reference.
    pub const fn external_object_id(&self) -> Option<u32> {
        self.external_object_id
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
    pub fn extract_image<'data>(
        &'data self,
        presentation: &'data crate::Presentation,
    ) -> Result<Option<ImageFile<'data>>> {
        // Try to extract from embedded Escher data first
        if let Some(ref escher_data) = self.escher_data {
            let images = litchi_odraw::image::scan(escher_data)
                .map_err(|error| litchi_core::error::Error::ParseError(error.to_string()))?;
            if let Some(image) = images.into_iter().next() {
                return Ok(Some(image));
            }
        }

        // Try to extract from Pictures stream using BLIP ID
        if let Some(blip_id) = self.blip_id {
            return presentation.image(blip_id).map_err(|e| {
                litchi_core::error::Error::ParseError(format!(
                    "Failed to extract image by BLIP ID: {}",
                    e
                ))
            });
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

    fn text(&self) -> std::result::Result<String, Error> {
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
///
/// Uses zero-copy OfficeArt property parsing.
pub fn extract_blip_id(container: &Container<'_>) -> litchi_odraw::Result<Option<ImageId>> {
    use litchi_odraw::RecordKind;
    use litchi_odraw::prop::{Id, Props};

    // Look for shape options (Opt record)
    for child in container.children() {
        let child = child?;
        if child.kind() == RecordKind::Opt
            && let Some(raw) = Props::parse(&child)?.get_int(Id::BlipToDisplay)
        {
            let raw = u32::try_from(raw).map_err(|_| litchi_odraw::Error::MalformedProperties {
                reason: "BlipToDisplay must be a positive one-based image identifier",
            })?;
            return ImageId::new(raw).map(Some);
        }
    }

    Ok(None)
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
        picture.set_blip_index(42).unwrap();
        assert_eq!(picture.blip_id().map(ImageId::get), Some(42));
        assert!(picture.set_blip_index(0).is_err());
    }

    #[test]
    fn frame_kind_tracks_the_common_shape_type() {
        let mut picture = PictureShape::new(1);
        picture.set_frame_kind(PictureFrameKind::OleObject);
        assert_eq!(picture.frame_kind(), PictureFrameKind::OleObject);
        assert_eq!(picture.properties.shape_type, ShapeType::Object);

        let media = PictureShape::from_properties(ShapeProperties {
            shape_type: ShapeType::Media,
            ..Default::default()
        });
        assert_eq!(media.frame_kind(), PictureFrameKind::Media);
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
