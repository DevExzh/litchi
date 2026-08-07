use litchi_core::Result;

use super::{ImageMap, parse_image_maps};

impl crate::Package {
    /// Every `draw:image-map` in packaged `content.xml`, in document order.
    pub fn image_maps(&self) -> Result<Vec<ImageMap>> {
        parse_image_maps(&self.content_xml()?)
    }
}

impl crate::FlatDocument {
    /// Every `draw:image-map` in a flat `OpenDocument`, in document order.
    pub fn image_maps(&self) -> Result<Vec<ImageMap>> {
        parse_image_maps(self.xml())
    }
}
