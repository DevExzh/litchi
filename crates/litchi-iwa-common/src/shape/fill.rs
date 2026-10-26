//! Archive-free shape-fill values shared by iWork format owners: gradients,
//! image fills and the [`Fill`] they form with colors (record 0375 adds the
//! image fill and [`Fill`]).

use std::f32::consts::TAU;

use crate::color::Rgba;
use crate::shape::geometry::Size;

#[allow(
    clippy::arbitrary_source_item_ordering,
    reason = "The macro keeps three identical validated scalar declarations concise"
)]
macro_rules! normalized_value {
    ($name:ident, $non_finite:ident, $out_of_range:ident, $doc:literal) => {
        #[doc = $doc]
        #[repr(transparent)]
        #[derive(Debug, Clone, Copy, PartialEq, PartialOrd)]
        pub struct $name(f32);

        impl $name {
            /// Construct a finite normalized value in the inclusive range `0.0..=1.0`.
            ///
            /// # Errors
            ///
            /// Returns a typed error when `value` is non-finite or outside
            /// the normalized interval.
            pub fn new(value: f32) -> Result<Self> {
                if !value.is_finite() {
                    return Err(Error::$non_finite);
                }
                if !(0.0..=1.0).contains(&value) {
                    return Err(Error::$out_of_range);
                }
                Ok(Self(value))
            }

            /// Return the normalized value.
            #[must_use]
            pub const fn get(self) -> f32 {
                self.0
            }
        }
    };
}

const MINIMUM_STOP_COUNT: usize = 2;
const SIMPLE_STOP_COUNT: usize = 2;

/// Validation failures for shape-gradient values.
#[derive(Debug, Clone, Copy, PartialEq, Eq, thiserror::Error)]
#[non_exhaustive]
pub enum Error {
    /// The gradient angle was not finite.
    #[error("shape gradient angle must be finite")]
    AngleNonFinite,
    /// The gradient angle was outside the native `[0, 2π)` interval.
    #[error("shape gradient angle must be in [0, 2π)")]
    AngleOutOfRange,
    /// A stop position was not finite.
    #[error("shape gradient stop position must be finite")]
    StopPositionNonFinite,
    /// A stop position was outside the normalized interval.
    #[error("shape gradient stop position must be in 0.0..=1.0")]
    StopPositionOutOfRange,
    /// A stop midpoint was not finite.
    #[error("shape gradient stop midpoint must be finite")]
    StopMidpointNonFinite,
    /// A stop midpoint was outside the normalized interval.
    #[error("shape gradient stop midpoint must be in 0.0..=1.0")]
    StopMidpointOutOfRange,
    /// Gradient opacity was not finite.
    #[error("shape gradient opacity must be finite")]
    OpacityNonFinite,
    /// Gradient opacity was outside the normalized interval.
    #[error("shape gradient opacity must be in 0.0..=1.0")]
    OpacityOutOfRange,
    /// A gradient contained fewer than two stops.
    #[error("shape gradients require at least {MINIMUM_STOP_COUNT} color stops")]
    TooFewStops,
    /// Stops were not ordered by their normalized position.
    #[error("shape gradient stops must be ordered by position")]
    StopsUnordered,
    /// A non-advanced gradient did not have the native simple shape.
    #[error("a simple shape gradient must be linear with two centered stops")]
    InvalidSimpleShape,
    /// An image-fill size was not finite and positive.
    #[error("shape image-fill dimensions must be finite and positive")]
    ImageSizeInvalid,
    /// An embedded image-data identifier was zero.
    #[error("an embedded shape-fill image identifier must be nonzero")]
    ImageDataZero,
}

/// Result type for shape-gradient value construction.
pub type Result<T> = std::result::Result<T, Error>;

/// Geometry used to paint a shape gradient.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Kind {
    /// Colors progress along a straight axis.
    Linear,
    /// Colors radiate from the native gradient origin.
    Radial,
}

/// Angle of a shape gradient, stored in native radians.
#[repr(transparent)]
#[derive(Debug, Clone, Copy, PartialEq, PartialOrd)]
pub struct Angle(f32);

impl Angle {
    /// Construct an angle from the degree value shown by iWork.
    ///
    /// # Errors
    ///
    /// Returns a typed error when the converted angle is non-finite or
    /// outside the native `[0, 2π)` interval.
    pub fn from_degrees(degrees: f32) -> Result<Self> {
        Self::from_radians(degrees.to_radians())
    }

    /// Construct an angle from the native radian value.
    ///
    /// # Errors
    ///
    /// Returns a typed error when `radians` is non-finite or outside the
    /// native `[0, 2π)` interval.
    pub fn from_radians(radians: f32) -> Result<Self> {
        if !radians.is_finite() {
            return Err(Error::AngleNonFinite);
        }
        if !(0.0..TAU).contains(&radians) {
            return Err(Error::AngleOutOfRange);
        }
        Ok(Self(radians))
    }

    /// Return the exact radian value stored in the native archive.
    #[must_use]
    pub const fn radians(self) -> f32 {
        self.0
    }

    /// Return the angle in the degree unit displayed by iWork.
    #[must_use]
    pub fn degrees(self) -> f32 {
        self.0.to_degrees()
    }
}

normalized_value!(
    StopPosition,
    StopPositionNonFinite,
    StopPositionOutOfRange,
    "Normalized location of a shape-gradient color stop."
);
normalized_value!(
    StopMidpoint,
    StopMidpointNonFinite,
    StopMidpointOutOfRange,
    "Normalized blend midpoint following a shape-gradient color stop."
);
normalized_value!(
    Opacity,
    OpacityNonFinite,
    OpacityOutOfRange,
    "Normalized overall opacity of a shape gradient."
);

impl StopPosition {
    /// The beginning of a gradient axis.
    pub const START: Self = Self(0.0);
    /// The end of a gradient axis.
    pub const END: Self = Self(1.0);
}

impl StopMidpoint {
    /// The midpoint used by iWork's simple two-color gradient editor.
    pub const CENTER: Self = Self(0.5);
}

impl Opacity {
    /// A fully opaque gradient.
    pub const OPAQUE: Self = Self(1.0);
}

/// One immutable, validated color stop in a shape gradient.
#[repr(C)]
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct Stop {
    color: Rgba,
    position: StopPosition,
    midpoint: StopMidpoint,
}

impl Stop {
    /// Construct one color stop from validated components.
    #[must_use]
    pub const fn new(color: Rgba, position: StopPosition, midpoint: StopMidpoint) -> Self {
        Self {
            color,
            position,
            midpoint,
        }
    }

    /// Return the stop color.
    #[must_use]
    pub const fn color(self) -> Rgba {
        self.color
    }

    /// Return the normalized stop position.
    #[must_use]
    pub const fn position(self) -> StopPosition {
        self.position
    }

    /// Return the normalized stop midpoint.
    #[must_use]
    pub const fn midpoint(self) -> StopMidpoint {
        self.midpoint
    }
}

/// Native simple or advanced linear/radial shape gradient.
#[derive(Debug, Clone, PartialEq)]
pub struct Gradient {
    kind: Kind,
    stops: Box<[Stop]>,
    opacity: Opacity,
    advanced: bool,
    angle: Angle,
}

impl Gradient {
    /// Construct iWork's two-stop simple linear gradient.
    #[must_use]
    pub fn linear(start: Rgba, end: Rgba, angle: Angle) -> Self {
        Self {
            kind: Kind::Linear,
            stops: Box::new([
                Stop::new(start, StopPosition::START, StopMidpoint::CENTER),
                Stop::new(end, StopPosition::END, StopMidpoint::CENTER),
            ]),
            opacity: Opacity::OPAQUE,
            advanced: false,
            angle,
        }
    }

    /// Construct an explicitly advanced linear or radial gradient.
    ///
    /// # Errors
    ///
    /// Returns a typed error when fewer than two stops are supplied or their
    /// positions are not ordered.
    pub fn advanced(
        kind: Kind,
        stops: impl Into<Box<[Stop]>>,
        opacity: Opacity,
        angle: Angle,
    ) -> Result<Self> {
        Self::from_parts(kind, stops, opacity, true, angle)
    }

    /// Construct a gradient from all native semantic fields.
    ///
    /// The `advanced` flag is retained because it changes how iWork presents
    /// and edits the same gradient wire shape; it is not inferred from the
    /// stop count so a read/write cycle remains lossless.
    ///
    /// # Errors
    ///
    /// Returns a typed error when the stops are insufficient or unordered, or
    /// when a simple gradient does not have its required native shape.
    pub fn from_parts(
        kind: Kind,
        stops: impl Into<Box<[Stop]>>,
        opacity: Opacity,
        advanced: bool,
        angle: Angle,
    ) -> Result<Self> {
        let boxed_stops = stops.into();
        validate_stops(&boxed_stops)?;
        if !advanced
            && (kind != Kind::Linear
                || boxed_stops.len() != SIMPLE_STOP_COUNT
                || boxed_stops
                    .iter()
                    .any(|stop| stop.midpoint() != StopMidpoint::CENTER))
        {
            return Err(Error::InvalidSimpleShape);
        }
        Ok(Self {
            kind,
            stops: boxed_stops,
            opacity,
            advanced,
            angle,
        })
    }

    /// Return the gradient geometry.
    #[must_use]
    pub const fn kind(&self) -> Kind {
        self.kind
    }

    /// Return the boxed stop slice without copying it.
    #[must_use]
    pub fn stops(&self) -> &[Stop] {
        &self.stops
    }

    /// Return the overall gradient opacity.
    #[must_use]
    pub const fn opacity(&self) -> Opacity {
        self.opacity
    }

    /// Return whether iWork stores this as an advanced gradient.
    #[must_use]
    pub const fn is_advanced(&self) -> bool {
        self.advanced
    }

    /// Return the gradient angle.
    #[must_use]
    pub const fn angle(&self) -> Angle {
        self.angle
    }
}

fn validate_stops(stops: &[Stop]) -> Result<()> {
    if stops.len() < MINIMUM_STOP_COUNT {
        return Err(Error::TooFewStops);
    }
    if stops
        .windows(2)
        .any(|pair| pair[1].position() < pair[0].position())
    {
        return Err(Error::StopsUnordered);
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use std::mem::{align_of, size_of};

    use super::*;
    use crate::color::RgbColorSpace;

    fn color() -> Rgba {
        Rgba::new(0.2, 0.4, 0.8, 1.0, RgbColorSpace::Srgb).unwrap()
    }

    #[test]
    fn scalar_values_are_typed_and_compact() {
        assert_eq!(size_of::<Angle>(), 4);
        assert_eq!(size_of::<StopPosition>(), 4);
        assert_eq!(size_of::<StopMidpoint>(), 4);
        assert_eq!(size_of::<Opacity>(), 4);
        assert_eq!(align_of::<Angle>(), 4);
        assert_eq!(size_of::<Stop>(), size_of::<Rgba>() + 8);
    }

    #[test]
    fn scalar_validation_is_typed() {
        assert_eq!(Angle::from_degrees(f32::NAN), Err(Error::AngleNonFinite));
        assert_eq!(Angle::from_degrees(360.0), Err(Error::AngleOutOfRange));
        assert_eq!(StopPosition::new(-0.1), Err(Error::StopPositionOutOfRange));
        assert_eq!(
            StopMidpoint::new(f32::INFINITY),
            Err(Error::StopMidpointNonFinite)
        );
        assert_eq!(Opacity::new(1.1), Err(Error::OpacityOutOfRange));
    }

    #[test]
    fn gradients_keep_boxed_stops_and_native_mode() {
        let angle = Angle::from_degrees(45.0).unwrap();
        let gradient = Gradient::from_parts(
            Kind::Radial,
            vec![
                Stop::new(color(), StopPosition::START, StopMidpoint::CENTER),
                Stop::new(color(), StopPosition::END, StopMidpoint::CENTER),
            ],
            Opacity::new(0.8).unwrap(),
            true,
            angle,
        )
        .unwrap();
        assert!(gradient.is_advanced());
        assert_eq!(gradient.stops().len(), 2);
        assert_eq!(gradient.angle(), angle);
    }

    #[test]
    fn gradients_reject_invalid_stop_sequences() {
        let color = color();
        let stops = vec![
            Stop::new(color, StopPosition::END, StopMidpoint::CENTER),
            Stop::new(color, StopPosition::START, StopMidpoint::CENTER),
        ];
        assert_eq!(
            Gradient::advanced(
                Kind::Linear,
                stops,
                Opacity::OPAQUE,
                Angle::from_degrees(0.0).unwrap(),
            ),
            Err(Error::StopsUnordered)
        );
    }
}

/// How an image is laid out inside a shape.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum ImageTechnique {
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

impl ImageTechnique {
    /// The technique stored under a native `ImageFillTechnique` value.
    #[must_use]
    pub const fn from_native(value: i32) -> Option<Self> {
        match value {
            0 => Some(Self::OriginalSize),
            1 => Some(Self::Stretch),
            2 => Some(Self::Tile),
            3 => Some(Self::ScaleToFill),
            4 => Some(Self::ScaleToFit),
            _ => None,
        }
    }

    /// The native `ImageFillTechnique` value.
    #[must_use]
    pub const fn native_value(self) -> i32 {
        match self {
            Self::OriginalSize => 0,
            Self::Stretch => 1,
            Self::Tile => 2,
            Self::ScaleToFill => 3,
            Self::ScaleToFit => 4,
        }
    }
}

/// The non-zero identifier of embedded image data (`TSP.DataReference`), the
/// archive-boundary key of an image fill.
#[repr(transparent)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct ImageData(std::num::NonZeroU64);

impl ImageData {
    /// Construct a non-zero identifier.
    ///
    /// # Errors
    ///
    /// [`Error::ImageDataZero`] for 0.
    pub const fn new(identifier: u64) -> Result<Self> {
        match std::num::NonZeroU64::new(identifier) {
            Some(value) => Ok(Self(value)),
            None => Err(Error::ImageDataZero),
        }
    }

    /// The native identifier.
    #[must_use]
    pub const fn get(self) -> u64 {
        self.0.get()
    }
}

fn validate_image_size(size: Size) -> Result<()> {
    if !size.width.is_finite()
        || !size.height.is_finite()
        || size.width <= 0.0
        || size.height <= 0.0
    {
        return Err(Error::ImageSizeInvalid);
    }
    Ok(())
}

/// A simple or advanced (tinted) image fill.
#[derive(Debug, Clone, PartialEq)]
pub struct Image {
    data: Option<ImageData>,
    technique: ImageTechnique,
    fill_size: Size,
    tint: Option<Rgba>,
    reference_color: Option<Rgba>,
    interprets_untagged_image_as_generic: bool,
}

impl Image {
    /// A fill backed by embedded image data.
    ///
    /// # Errors
    ///
    /// [`Error::ImageSizeInvalid`] for a non-positive or non-finite size.
    pub fn embedded(data: ImageData, technique: ImageTechnique, fill_size: Size) -> Result<Self> {
        Self::new(Some(data), technique, fill_size)
    }

    /// The empty placeholder shown before an image is chosen.
    ///
    /// # Errors
    ///
    /// [`Error::ImageSizeInvalid`] for a non-positive or non-finite size.
    pub fn placeholder(technique: ImageTechnique, fill_size: Size) -> Result<Self> {
        Self::new(None, technique, fill_size)
    }

    fn new(data: Option<ImageData>, technique: ImageTechnique, fill_size: Size) -> Result<Self> {
        validate_image_size(fill_size)?;
        Ok(Self {
            data,
            technique,
            fill_size,
            tint: None,
            reference_color: None,
            interprets_untagged_image_as_generic: false,
        })
    }

    /// Apply a tint (the inspector's Advanced Image Fill).
    #[must_use]
    pub const fn with_tint(mut self, tint: Rgba) -> Self {
        self.tint = Some(tint);
        self
    }

    /// Remove the tint.
    #[must_use]
    pub const fn without_tint(mut self) -> Self {
        self.tint = None;
        self
    }

    /// Keep iWork's representative color sampled from the image.
    #[must_use]
    pub const fn with_reference_color(mut self, color: Rgba) -> Self {
        self.reference_color = Some(color);
        self
    }

    /// Remove the representative color.
    #[must_use]
    pub const fn without_reference_color(mut self) -> Self {
        self.reference_color = None;
        self
    }

    /// Replace the layout technique.
    #[must_use]
    pub const fn with_technique(mut self, technique: ImageTechnique) -> Self {
        self.technique = technique;
        self
    }

    /// Replace the natural/tile fill size.
    ///
    /// # Errors
    ///
    /// [`Error::ImageSizeInvalid`].
    pub fn with_fill_size(mut self, fill_size: Size) -> Result<Self> {
        validate_image_size(fill_size)?;
        self.fill_size = fill_size;
        Ok(self)
    }

    /// Whether untagged pixels are interpreted as generic RGB.
    #[must_use]
    pub const fn with_generic_untagged_image(mut self, enabled: bool) -> Self {
        self.interprets_untagged_image_as_generic = enabled;
        self
    }

    /// The embedded image data, or `None` for the placeholder.
    #[must_use]
    pub const fn data(&self) -> Option<ImageData> {
        self.data
    }

    /// The layout technique.
    #[must_use]
    pub const fn technique(&self) -> ImageTechnique {
        self.technique
    }

    /// The natural/tile fill size.
    #[must_use]
    pub const fn fill_size(&self) -> Size {
        self.fill_size
    }

    /// The tint.
    #[must_use]
    pub const fn tint(&self) -> Option<Rgba> {
        self.tint
    }

    /// The representative color.
    #[must_use]
    pub const fn reference_color(&self) -> Option<Rgba> {
        self.reference_color
    }

    /// Whether untagged pixels are interpreted as generic RGB.
    #[must_use]
    pub const fn interprets_untagged_image_as_generic(&self) -> bool {
        self.interprets_untagged_image_as_generic
    }
}

/// A standard shape fill.
#[derive(Debug, Clone, PartialEq, Default)]
pub enum Fill {
    /// An empty native fill ("No Fill").
    #[default]
    None,
    /// One color.
    Solid(Rgba),
    /// A gradient.
    Gradient(Gradient),
    /// An image fill.
    Image(Image),
}

#[cfg(test)]
mod image_tests {
    use super::{Error, Fill, Image, ImageData, ImageTechnique, Size};

    const SIZE: Size = Size {
        width: 320.0,
        height: 180.0,
    };

    #[test]
    fn image_fills_validate_and_build() {
        assert_eq!(ImageData::new(0), Err(Error::ImageDataZero));
        assert_eq!(
            Image::placeholder(
                ImageTechnique::Stretch,
                Size {
                    width: 0.0,
                    height: 1.0
                }
            ),
            Err(Error::ImageSizeInvalid)
        );
        let data = ImageData::new(42).unwrap();
        let tint = crate::color::Rgba::black();
        let image = Image::embedded(data, ImageTechnique::ScaleToFit, SIZE)
            .unwrap()
            .with_tint(tint)
            .with_reference_color(tint)
            .with_generic_untagged_image(true)
            .with_technique(ImageTechnique::Tile);
        assert_eq!(image.data(), Some(data));
        assert_eq!(image.technique(), ImageTechnique::Tile);
        assert_eq!(image.tint(), Some(tint));
        assert!(
            image
                .clone()
                .without_tint()
                .without_reference_color()
                .tint()
                .is_none()
        );
        assert!(image.interprets_untagged_image_as_generic());
        for technique in [
            ImageTechnique::OriginalSize,
            ImageTechnique::Stretch,
            ImageTechnique::Tile,
            ImageTechnique::ScaleToFill,
            ImageTechnique::ScaleToFit,
        ] {
            assert_eq!(
                ImageTechnique::from_native(technique.native_value()),
                Some(technique)
            );
        }
        assert_eq!(Fill::default(), Fill::None);
    }
}
