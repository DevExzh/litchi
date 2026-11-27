//! Package-independent slide background values and XML codec.
//!
//! This module owns the `PresentationML` semantic model for slide backgrounds
//! and the package-free `<p:bg>`/`DrawingML` fill codec. Relationship lookup and
//! image-part loading remain responsibilities of the package host.

use crate::format::ImageFormat;

/// Slide background configuration.
///
/// Represents the background fill for a slide.
#[derive(Debug, Clone, PartialEq)]
pub enum SlideBackground {
    /// No background (transparent)
    None,
    /// Solid color background
    Solid {
        /// RGB color in hexadecimal format (e.g., "FFFFFF" for white)
        color: String,
    },
    /// Gradient background
    Gradient {
        /// Gradient type (linear or radial)
        gradient_type: GradientType,
        /// Gradient angle in degrees (0-360, for linear gradients)
        angle: Option<f64>,
        /// Gradient stops (position 0.0-1.0, color pairs)
        stops: Vec<GradientStop>,
    },
    /// Picture/image background
    Picture {
        /// Image data (will be embedded)
        image_data: Vec<u8>,
        /// Image format
        format: ImageFormat,
        /// How the image should be displayed
        style: PictureStyle,
    },
    /// Pattern background
    Pattern {
        /// Pattern type
        pattern_type: PatternType,
        /// Foreground color
        fg_color: String,
        /// Background color
        bg_color: String,
    },
}

/// Gradient type for gradient backgrounds.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum GradientType {
    /// Linear gradient
    Linear,
    /// Radial gradient (from center)
    Radial,
    /// Rectangular gradient
    Rectangular,
    /// Path gradient
    Path,
}

/// A gradient stop (position and color).
#[derive(Debug, Clone, PartialEq)]
pub struct GradientStop {
    /// Position from 0.0 to 1.0
    pub position: f64,
    /// RGB color in hex format
    pub color: String,
}

/// Picture display style for picture backgrounds.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PictureStyle {
    /// Stretch to fill the slide
    Stretch,
    /// Fit within the slide (maintain aspect ratio)
    Fit,
    /// Tile the image
    Tile,
    /// Center the image
    Center,
}

/// Pattern type for pattern backgrounds.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PatternType {
    /// 5% dotted pattern
    Pct5,
    /// 10% dotted pattern
    Pct10,
    /// 20% dotted pattern
    Pct20,
    /// 25% dotted pattern
    Pct25,
    /// 30% dotted pattern
    Pct30,
    /// 40% dotted pattern
    Pct40,
    /// 50% dotted pattern
    Pct50,
    /// 60% dotted pattern
    Pct60,
    /// 70% dotted pattern
    Pct70,
    /// 75% dotted pattern
    Pct75,
    /// 80% dotted pattern
    Pct80,
    /// 90% dotted pattern
    Pct90,
    /// Horizontal lines
    Horizontal,
    /// Vertical lines
    Vertical,
    /// Light horizontal lines
    LightHorizontal,
    /// Light vertical lines
    LightVertical,
    /// Dark horizontal lines
    DarkHorizontal,
    /// Dark vertical lines
    DarkVertical,
    /// Narrow horizontal lines
    NarrowHorizontal,
    /// Narrow vertical lines
    NarrowVertical,
    /// Dashed horizontal lines
    DashedHorizontal,
    /// Dashed vertical lines
    DashedVertical,
    /// Diagonal lines (top-left to bottom-right)
    DownDiagonal,
    /// Diagonal lines (bottom-left to top-right)
    UpDiagonal,
    /// Light down diagonal
    LightDownDiagonal,
    /// Light up diagonal
    LightUpDiagonal,
    /// Dark down diagonal
    DarkDownDiagonal,
    /// Dark up diagonal
    DarkUpDiagonal,
    /// Wide down diagonal
    WideDownDiagonal,
    /// Wide up diagonal
    WideUpDiagonal,
    /// Dashed down diagonal
    DashedDownDiagonal,
    /// Dashed up diagonal
    DashedUpDiagonal,
    /// Crosshatch
    Cross,
    /// Diagonal crosshatch
    DiagonalCross,
    /// Small check pattern
    SmallCheck,
    /// Large check pattern
    LargeCheck,
    /// Small grid
    SmallGrid,
    /// Large grid
    LargeGrid,
    /// Dotted grid
    DottedGrid,
    /// Small confetti
    SmallConfetti,
    /// Large confetti
    LargeConfetti,
    /// Horizontal brick
    HorizontalBrick,
    /// Diagonal brick
    DiagonalBrick,
    /// Solid diamond
    SolidDiamond,
    /// Open diamond
    OpenDiamond,
    /// Dotted diamond
    DottedDiamond,
    /// Plaid pattern
    Plaid,
    /// Sphere pattern
    Sphere,
    /// Weave pattern
    Weave,
    /// Divot pattern
    Divot,
    /// Shingle pattern
    Shingle,
    /// Wave pattern
    Wave,
    /// Trellis pattern
    Trellis,
    /// Zig-zag pattern
    ZigZag,
}

impl SlideBackground {
    /// Create a solid color background.
    ///
    /// # Arguments
    /// * `color` - RGB color in hex format (e.g., "FF0000" for red)
    ///
    /// # Examples
    ///
    /// ```rust
    /// use litchi_pptx::SlideBackground;
    ///
    /// let bg = SlideBackground::solid("4472C4"); // Blue background
    /// ```
    pub fn solid(color: impl Into<String>) -> Self {
        SlideBackground::Solid {
            color: color.into(),
        }
    }

    /// Create a linear gradient background.
    ///
    /// # Arguments
    /// * `angle` - Gradient angle in degrees (0-360)
    /// * `stops` - Vector of gradient stops
    ///
    /// # Examples
    ///
    /// ```rust
    /// use litchi_pptx::{GradientStop, SlideBackground};
    ///
    /// let stops = vec![
    ///     GradientStop { position: 0.0, color: "4472C4".to_string() },
    ///     GradientStop { position: 1.0, color: "FFFFFF".to_string() },
    /// ];
    /// let bg = SlideBackground::linear_gradient(90.0, stops);
    /// ```
    #[must_use]
    pub fn linear_gradient(angle: f64, stops: Vec<GradientStop>) -> Self {
        SlideBackground::Gradient {
            gradient_type: GradientType::Linear,
            angle: Some(angle),
            stops,
        }
    }

    /// Create a radial gradient background.
    ///
    /// # Arguments
    /// * `stops` - Vector of gradient stops
    #[must_use]
    pub fn radial_gradient(stops: Vec<GradientStop>) -> Self {
        SlideBackground::Gradient {
            gradient_type: GradientType::Radial,
            angle: None,
            stops,
        }
    }

    /// Create a picture background from image data.
    ///
    /// # Arguments
    /// * `image_data` - Image bytes
    /// * `format` - Image format
    /// * `style` - How to display the image
    #[must_use]
    pub fn picture(image_data: Vec<u8>, format: ImageFormat, style: PictureStyle) -> Self {
        SlideBackground::Picture {
            image_data,
            format,
            style,
        }
    }

    /// Create a pattern background.
    ///
    /// # Arguments
    /// * `pattern_type` - Type of pattern
    /// * `fg_color` - Foreground color in hex
    /// * `bg_color` - Background color in hex
    #[must_use]
    pub fn pattern(pattern_type: PatternType, fg_color: String, bg_color: String) -> Self {
        SlideBackground::Pattern {
            pattern_type,
            fg_color,
            bg_color,
        }
    }

    /// Borrow image data and its format when this is a picture background.
    ///
    /// This is the narrow resource seam used by package writers to publish
    /// the bytes and allocate a relationship without copying the image.
    #[must_use]
    pub fn image_data(&self) -> Option<(&[u8], &ImageFormat)> {
        match self {
            SlideBackground::Picture {
                image_data, format, ..
            } => Some((image_data, format)),
            _ => None,
        }
    }
}
