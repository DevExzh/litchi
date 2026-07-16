//! Typed image-fill values shared by Pages, Numbers, and Keynote.

use crate::{Error, Result};

use super::super::{DrawableSize, RgbaColor};

/// Stable identifier of image data embedded in an iWork package.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct ShapeImageDataIdentifier(u64);

impl ShapeImageDataIdentifier {
    /// Construct a nonzero embedded-data identifier.
    pub fn new(identifier: u64) -> Result<Self> {
        if identifier == 0 {
            return Err(Error::ParseError(
                "An embedded shape-fill image identifier must be nonzero".to_owned(),
            ));
        }
        Ok(Self(identifier))
    }

    /// Return the native `TSP.DataReference` identifier.
    pub const fn get(self) -> u64 {
        self.0
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

/// A native simple or advanced image fill.
#[derive(Debug, Clone, PartialEq)]
pub struct ShapeImageFill {
    data_identifier: Option<ShapeImageDataIdentifier>,
    technique: ShapeImageFillTechnique,
    fill_size: DrawableSize,
    tint: Option<RgbaColor>,
    reference_color: Option<RgbaColor>,
    interprets_untagged_image_as_generic: bool,
}

impl ShapeImageFill {
    /// Construct a fill backed by embedded image data.
    pub fn embedded(
        data_identifier: ShapeImageDataIdentifier,
        technique: ShapeImageFillTechnique,
        fill_size: DrawableSize,
    ) -> Result<Self> {
        Self::new(Some(data_identifier), technique, fill_size)
    }

    /// Construct the empty placeholder displayed before an image is chosen.
    pub fn placeholder(
        technique: ShapeImageFillTechnique,
        fill_size: DrawableSize,
    ) -> Result<Self> {
        Self::new(None, technique, fill_size)
    }

    fn new(
        data_identifier: Option<ShapeImageDataIdentifier>,
        technique: ShapeImageFillTechnique,
        fill_size: DrawableSize,
    ) -> Result<Self> {
        validate_fill_size(fill_size)?;
        Ok(Self {
            data_identifier,
            technique,
            fill_size,
            tint: None,
            reference_color: None,
            interprets_untagged_image_as_generic: false,
        })
    }

    /// Apply a tint, converting the application presentation to Advanced Image Fill.
    pub fn with_tint(mut self, tint: RgbaColor) -> Self {
        self.tint = Some(tint);
        self
    }

    /// Remove the tint, converting the application presentation to Image Fill.
    pub fn without_tint(mut self) -> Self {
        self.tint = None;
        self
    }

    /// Preserve iWork's representative color sampled from the source image.
    pub fn with_reference_color(mut self, color: RgbaColor) -> Self {
        self.reference_color = Some(color);
        self
    }

    /// Remove representative-color metadata.
    pub fn without_reference_color(mut self) -> Self {
        self.reference_color = None;
        self
    }

    /// Replace the image layout technique.
    pub fn with_technique(mut self, technique: ShapeImageFillTechnique) -> Self {
        self.technique = technique;
        self
    }

    /// Replace the natural/tile fill dimensions.
    pub fn with_fill_size(mut self, fill_size: DrawableSize) -> Result<Self> {
        validate_fill_size(fill_size)?;
        self.fill_size = fill_size;
        Ok(self)
    }

    /// Control whether untagged source pixels are interpreted as generic RGB.
    pub fn with_generic_untagged_image(mut self, enabled: bool) -> Self {
        self.interprets_untagged_image_as_generic = enabled;
        self
    }

    pub const fn data_identifier(&self) -> Option<ShapeImageDataIdentifier> {
        self.data_identifier
    }

    pub const fn technique(&self) -> ShapeImageFillTechnique {
        self.technique
    }

    pub const fn fill_size(&self) -> DrawableSize {
        self.fill_size
    }

    pub const fn tint(&self) -> Option<RgbaColor> {
        self.tint
    }

    pub const fn reference_color(&self) -> Option<RgbaColor> {
        self.reference_color
    }

    pub const fn interprets_untagged_image_as_generic(&self) -> bool {
        self.interprets_untagged_image_as_generic
    }
}

fn validate_fill_size(size: DrawableSize) -> Result<()> {
    if !size.width.is_finite()
        || !size.height.is_finite()
        || size.width <= 0.0
        || size.height <= 0.0
    {
        return Err(Error::ParseError(
            "Shape image-fill dimensions must be finite and positive".to_owned(),
        ));
    }
    Ok(())
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
