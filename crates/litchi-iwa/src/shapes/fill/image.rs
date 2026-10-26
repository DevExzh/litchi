//! Typed image-fill values shared by Pages, Numbers, and Keynote.

use litchi_iwa_common::shape::fill::{Image, ImageData, ImageTechnique};

use crate::{Error, Result};

use super::super::{DrawableSize, RgbaColor};

/// Stable identifier of image data embedded in an iWork package (the shared
/// `litchi_iwa_common::shape::fill::ImageData`, record 0375).
#[derive(Clone, Copy, PartialEq, Eq, Hash)]
pub struct ShapeImageDataIdentifier(ImageData);

impl std::fmt::Debug for ShapeImageDataIdentifier {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_tuple("ShapeImageDataIdentifier")
            .field(&self.0.get())
            .finish()
    }
}

impl ShapeImageDataIdentifier {
    /// Construct a nonzero embedded-data identifier.
    pub fn new(identifier: u64) -> Result<Self> {
        ImageData::new(identifier).map(Self).map_err(|_| {
            Error::ParseError("An embedded shape-fill image identifier must be nonzero".to_owned())
        })
    }

    /// Return the native `TSP.DataReference` identifier.
    pub const fn get(self) -> u64 {
        self.0.get()
    }
}

/// How an image is laid out inside a shape.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum ShapeImageFillTechnique {
    /// Keep the image at its natural size.
    OriginalSize,
    /// Distort the image to match the shape bounds.
    #[default]
    Stretch,
    /// Repeat the image in both dimensions.
    Tile,
    /// Preserve aspect ratio and cover the shape bounds.
    ScaleToFill,
    /// Preserve aspect ratio and fit entirely inside the shape bounds.
    ScaleToFit,
}

impl ShapeImageFillTechnique {
    const fn common(self) -> ImageTechnique {
        match self {
            Self::OriginalSize => ImageTechnique::OriginalSize,
            Self::Stretch => ImageTechnique::Stretch,
            Self::Tile => ImageTechnique::Tile,
            Self::ScaleToFill => ImageTechnique::ScaleToFill,
            Self::ScaleToFit => ImageTechnique::ScaleToFit,
        }
    }

    const fn from_common(technique: ImageTechnique) -> Self {
        match technique {
            ImageTechnique::OriginalSize => Self::OriginalSize,
            ImageTechnique::Stretch => Self::Stretch,
            ImageTechnique::Tile => Self::Tile,
            ImageTechnique::ScaleToFill => Self::ScaleToFill,
            ImageTechnique::ScaleToFit => Self::ScaleToFit,
        }
    }
}

/// A native simple or advanced image fill (the shared
/// `litchi_iwa_common::shape::fill::Image`, record 0375).
#[derive(Clone, PartialEq)]
pub struct ShapeImageFill(Image);

impl std::fmt::Debug for ShapeImageFill {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("ShapeImageFill")
            .field("data_identifier", &self.data_identifier())
            .field("technique", &self.technique())
            .field("fill_size", &self.fill_size())
            .field("tint", &self.tint())
            .field("reference_color", &self.reference_color())
            .field(
                "interprets_untagged_image_as_generic",
                &self.interprets_untagged_image_as_generic(),
            )
            .finish()
    }
}

fn size_error(_: litchi_iwa_common::shape::fill::Error) -> Error {
    Error::ParseError("Shape image-fill dimensions must be finite and positive".to_owned())
}

impl ShapeImageFill {
    /// Construct a fill backed by embedded image data.
    pub fn embedded(
        data_identifier: ShapeImageDataIdentifier,
        technique: ShapeImageFillTechnique,
        fill_size: DrawableSize,
    ) -> Result<Self> {
        Image::embedded(data_identifier.0, technique.common(), fill_size)
            .map(Self)
            .map_err(size_error)
    }

    /// Construct the empty placeholder displayed before an image is chosen.
    pub fn placeholder(
        technique: ShapeImageFillTechnique,
        fill_size: DrawableSize,
    ) -> Result<Self> {
        Image::placeholder(technique.common(), fill_size)
            .map(Self)
            .map_err(size_error)
    }

    /// Apply a tint, converting the application presentation to Advanced Image Fill.
    pub fn with_tint(self, tint: RgbaColor) -> Self {
        Self(self.0.with_tint(tint))
    }

    /// Remove the tint, converting the application presentation to Image Fill.
    pub fn without_tint(self) -> Self {
        Self(self.0.without_tint())
    }

    /// Preserve iWork's representative color sampled from the source image.
    pub fn with_reference_color(self, color: RgbaColor) -> Self {
        Self(self.0.with_reference_color(color))
    }

    /// Remove representative-color metadata.
    pub fn without_reference_color(self) -> Self {
        Self(self.0.without_reference_color())
    }

    /// Replace the image layout technique.
    pub fn with_technique(self, technique: ShapeImageFillTechnique) -> Self {
        Self(self.0.with_technique(technique.common()))
    }

    /// Replace the natural/tile fill dimensions.
    pub fn with_fill_size(self, fill_size: DrawableSize) -> Result<Self> {
        self.0
            .with_fill_size(fill_size)
            .map(Self)
            .map_err(size_error)
    }

    /// Control whether untagged source pixels are interpreted as generic RGB.
    pub fn with_generic_untagged_image(self, enabled: bool) -> Self {
        Self(self.0.with_generic_untagged_image(enabled))
    }

    pub const fn data_identifier(&self) -> Option<ShapeImageDataIdentifier> {
        match self.0.data() {
            Some(data) => Some(ShapeImageDataIdentifier(data)),
            None => None,
        }
    }

    pub const fn technique(&self) -> ShapeImageFillTechnique {
        ShapeImageFillTechnique::from_common(self.0.technique())
    }

    pub const fn fill_size(&self) -> DrawableSize {
        self.0.fill_size()
    }

    pub const fn tint(&self) -> Option<RgbaColor> {
        self.0.tint()
    }

    pub const fn reference_color(&self) -> Option<RgbaColor> {
        self.0.reference_color()
    }

    pub const fn interprets_untagged_image_as_generic(&self) -> bool {
        self.0.interprets_untagged_image_as_generic()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    const FILL_SIZE: DrawableSize = DrawableSize {
        width: 320.0,
        height: 180.0,
    };

    #[test]
    fn rejects_zero_identifiers_and_invalid_dimensions() {
        assert!(ShapeImageDataIdentifier::new(0).is_err());
        assert!(
            ShapeImageFill::placeholder(
                ShapeImageFillTechnique::Stretch,
                DrawableSize {
                    width: 0.0,
                    height: 1.0,
                },
            )
            .is_err()
        );
    }

    #[test]
    fn embedded_fill_preserves_typed_options() {
        let identifier = ShapeImageDataIdentifier::new(42).unwrap();
        let tint = RgbaColor::new(0.1, 0.2, 0.3, 0.5, crate::shapes::RgbColorSpace::Srgb).unwrap();
        let fill =
            ShapeImageFill::embedded(identifier, ShapeImageFillTechnique::ScaleToFit, FILL_SIZE)
                .unwrap()
                .with_tint(tint)
                .with_reference_color(tint)
                .with_generic_untagged_image(true);
        assert_eq!(fill.data_identifier(), Some(identifier));
        assert_eq!(fill.technique(), ShapeImageFillTechnique::ScaleToFit);
        assert_eq!(fill.fill_size(), FILL_SIZE);
        assert_eq!(fill.tint(), Some(tint));
        assert_eq!(fill.reference_color(), Some(tint));
        assert!(fill.interprets_untagged_image_as_generic());
    }
}
