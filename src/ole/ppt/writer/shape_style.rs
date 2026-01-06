//! Shape styling support for PPT files
//!
//! This module handles fill colors, line styles, and other visual properties
//! for shapes in PowerPoint presentations.
//!
//! Reference: [MS-ODRAW] Section 2.3 - Property Tables

// Shape styling structures don't need zerocopy for now

use crate::common::unit::{EMUS_PER_PT, pt_f32_to_emu_u32};

// =============================================================================
// Escher Property IDs for Shape Styling (MS-ODRAW 2.3)
// =============================================================================

/// Fill properties
pub mod fill_prop {
    /// Fill type (0=solid, 1=pattern, 2=texture, etc.)
    pub const FILL_TYPE: u16 = 0x0180;
    /// Fill color (RGBX)
    pub const FILL_COLOR: u16 = 0x0181;
    /// Fill opacity (0-65536, 65536 = 100%)
    pub const FILL_OPACITY: u16 = 0x0182;
    /// Fill background color
    pub const FILL_BACK_COLOR: u16 = 0x0183;
    /// Fill background opacity
    pub const FILL_BACK_OPACITY: u16 = 0x0184;
    /// Fill BLIP reference (for picture/texture fills)
    pub const FILL_BLIP: u16 = 0x4186;
    /// Fill width (for pattern fills)
    pub const FILL_WIDTH: u16 = 0x0187;
    /// Fill height (for pattern fills)
    pub const FILL_HEIGHT: u16 = 0x0188;
    /// Fill angle (for gradient fills, in degrees * 65536)
    pub const FILL_ANGLE: u16 = 0x0189;
    /// Fill focus (for gradient fills, -100 to 100)
    pub const FILL_FOCUS: u16 = 0x018A;
    /// Fill shade type
    pub const FILL_SHADE_TYPE: u16 = 0x018C;
    /// Fill rectangle right
    pub const FILL_RECT_RIGHT: u16 = 0x0193;
    /// Fill rectangle bottom
    pub const FILL_RECT_BOTTOM: u16 = 0x0194;
    /// Fill boolean properties (no fill, hit test, etc.)
    pub const FILL_STYLE_BOOL: u16 = 0x01BF;
}

/// Line properties
pub mod line_prop {
    /// Line color
    pub const LINE_COLOR: u16 = 0x01C0;
    /// Line opacity (0-65536)
    pub const LINE_OPACITY: u16 = 0x01C1;
    /// Line background color
    pub const LINE_BACK_COLOR: u16 = 0x01C2;
    /// Line width in EMUs
    pub const LINE_WIDTH: u16 = 0x01CB;
    /// Line style (single, double, etc.)
    pub const LINE_STYLE: u16 = 0x01CD;
    /// Line dash style
    pub const LINE_DASH_STYLE: u16 = 0x01CE;
    /// Line start arrow head
    pub const LINE_START_ARROW_HEAD: u16 = 0x01D0;
    /// Line end arrow head
    pub const LINE_END_ARROW_HEAD: u16 = 0x01D1;
    /// Line start arrow width
    pub const LINE_START_ARROW_WIDTH: u16 = 0x01D2;
    /// Line start arrow length
    pub const LINE_START_ARROW_LENGTH: u16 = 0x01D3;
    /// Line end arrow width
    pub const LINE_END_ARROW_WIDTH: u16 = 0x01D4;
    /// Line end arrow length
    pub const LINE_END_ARROW_LENGTH: u16 = 0x01D5;
    /// Line join style
    pub const LINE_JOIN_STYLE: u16 = 0x01D6;
    /// Line end cap style
    pub const LINE_END_CAP_STYLE: u16 = 0x01D7;
    /// Line boolean properties
    pub const LINE_STYLE_BOOL: u16 = 0x01FF;
}

/// Shadow properties
pub mod shadow_prop {
    /// Shadow type
    pub const SHADOW_TYPE: u16 = 0x0200;
    /// Shadow color
    pub const SHADOW_COLOR: u16 = 0x0201;
    /// Shadow highlight color
    pub const SHADOW_HIGHLIGHT: u16 = 0x0202;
    /// Shadow X offset
    pub const SHADOW_OFFSET_X: u16 = 0x0205;
    /// Shadow Y offset
    pub const SHADOW_OFFSET_Y: u16 = 0x0206;
    /// Shadow opacity (0-65536)
    pub const SHADOW_OPACITY: u16 = 0x0204;
    /// Shadow boolean properties
    pub const SHADOW_STYLE_BOOL: u16 = 0x023F;
}

// =============================================================================
// Fill Types
// =============================================================================

/// Fill type values
#[repr(u32)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum FillType {
    /// Solid fill
    #[default]
    Solid = 0,
    /// Pattern fill
    Pattern = 1,
    /// Texture fill
    Texture = 2,
    /// Picture fill
    Picture = 3,
    /// Shade (gradient) fill
    Shade = 4,
    /// Shade from center
    ShadeCenter = 5,
    /// Shade from shape
    ShadeShape = 6,
    /// Shade from scale
    ShadeScale = 7,
    /// Shade from title
    ShadeTitle = 8,
    /// Background fill
    Background = 9,
}

// =============================================================================
// Line Styles
// =============================================================================

/// Line style values
#[repr(u32)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum LineStyle {
    /// Single line
    #[default]
    Simple = 0,
    /// Double lines
    Double = 1,
    /// Thick-thin double lines
    ThickThin = 2,
    /// Thin-thick double lines
    ThinThick = 3,
    /// Triple lines
    Triple = 4,
}

/// Line dash style values
#[repr(u32)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum LineDashStyle {
    /// Solid line
    #[default]
    Solid = 0,
    /// Dash-dot pattern
    DashSys = 1,
    /// Dot pattern
    DotSys = 2,
    /// Dash-dot-dot pattern
    DashDotSys = 3,
    /// Dash-dot-dot pattern (system)
    DashDotDotSys = 4,
    /// Dot (round)
    DotGel = 5,
    /// Dash
    Dash = 6,
    /// Long dash
    LongDash = 7,
    /// Dash-dot
    DashDot = 8,
    /// Long dash-dot
    LongDashDot = 9,
    /// Long dash-dot-dot
    LongDashDotDot = 10,
}

/// Line cap style
#[repr(u32)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum LineCapStyle {
    /// Round cap
    #[default]
    Round = 0,
    /// Square cap
    Square = 1,
    /// Flat cap
    Flat = 2,
}

/// Line join style
#[repr(u32)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum LineJoinStyle {
    /// Bevel join
    Bevel = 0,
    /// Miter join
    #[default]
    Miter = 1,
    /// Round join
    Round = 2,
}

// =============================================================================
// Arrow Styles
// =============================================================================

/// Arrow head style
#[repr(u32)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum ArrowStyle {
    /// No arrow
    #[default]
    None = 0,
    /// Triangle arrow
    Triangle = 1,
    /// Stealth arrow
    Stealth = 2,
    /// Diamond arrow
    Diamond = 3,
    /// Oval arrow
    Oval = 4,
    /// Open arrow
    Open = 5,
}

/// Arrow size (width or length)
#[repr(u32)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum ArrowSize {
    /// Small
    Small = 0,
    /// Medium
    #[default]
    Medium = 1,
    /// Large
    Large = 2,
}

// =============================================================================
// Color
// =============================================================================

/// Shape color (same format as TextColor but separate for clarity)
#[derive(Debug, Clone, Copy)]
pub struct ShapeColor {
    /// Red component (0-255)
    pub r: u8,
    /// Green component (0-255)
    pub g: u8,
    /// Blue component (0-255)
    pub b: u8,
    /// Use scheme color instead of RGB
    pub use_scheme: bool,
    /// Scheme color index (if use_scheme is true)
    pub scheme_index: u8,
}

impl ShapeColor {
    /// Transparent (no fill/line)
    pub const TRANSPARENT: Self = Self {
        r: 0,
        g: 0,
        b: 0,
        use_scheme: false,
        scheme_index: 0,
    };

    /// Black
    pub const BLACK: Self = Self::rgb(0, 0, 0);
    /// White
    pub const WHITE: Self = Self::rgb(255, 255, 255);
    /// Red
    pub const RED: Self = Self::rgb(255, 0, 0);
    /// Green
    pub const GREEN: Self = Self::rgb(0, 255, 0);
    /// Blue
    pub const BLUE: Self = Self::rgb(0, 0, 255);
    /// Yellow
    pub const YELLOW: Self = Self::rgb(255, 255, 0);
    /// Orange
    pub const ORANGE: Self = Self::rgb(255, 165, 0);
    /// Gray
    pub const GRAY: Self = Self::rgb(128, 128, 128);
    /// Light gray
    pub const LIGHT_GRAY: Self = Self::rgb(192, 192, 192);

    /// Create an RGB color
    pub const fn rgb(r: u8, g: u8, b: u8) -> Self {
        Self {
            r,
            g,
            b,
            use_scheme: false,
            scheme_index: 0,
        }
    }

    /// Create from hex value (0xRRGGBB)
    pub const fn from_hex(hex: u32) -> Self {
        Self::rgb(
            ((hex >> 16) & 0xFF) as u8,
            ((hex >> 8) & 0xFF) as u8,
            (hex & 0xFF) as u8,
        )
    }

    /// Create a scheme color reference
    pub const fn scheme(index: u8) -> Self {
        Self {
            r: 0,
            g: 0,
            b: 0,
            use_scheme: true,
            scheme_index: index,
        }
    }

    /// Scheme fill color (index 4)
    pub const fn scheme_fill() -> Self {
        Self::scheme(4)
    }

    /// Scheme line color (index 1)
    pub const fn scheme_line() -> Self {
        Self::scheme(1)
    }

    /// Scheme shadow color (index 2)
    pub const fn scheme_shadow() -> Self {
        Self::scheme(2)
    }

    /// Convert to RGB format for PPT (PowerPoint uses RedGreenBlue format)
    pub fn to_rgbx(&self) -> u32 {
        if self.use_scheme {
            0x0800_0000 | (self.scheme_index as u32)
        } else {
            // PPT uses RGB: Red in byte 0, Green in byte 1, Blue in byte 2
            (self.r as u32) | ((self.g as u32) << 8) | ((self.b as u32) << 16)
        }
    }
}

impl Default for ShapeColor {
    fn default() -> Self {
        Self::scheme_fill()
    }
}

// =============================================================================
// Fill Style
// =============================================================================

/// Fill style configuration
#[derive(Debug, Clone)]
pub struct FillStyle {
    /// Fill type
    pub fill_type: FillType,
    /// Primary fill color
    pub color: ShapeColor,
    /// Background color (for patterns/gradients)
    pub back_color: Option<ShapeColor>,
    /// Opacity (0-100, 100 = fully opaque)
    pub opacity: u8,
    /// Whether fill is enabled
    pub enabled: bool,
    /// Gradient angle in degrees (for gradient fills)
    pub gradient_angle: Option<i16>,
    /// Picture BLIP index (for picture fills)
    pub picture_index: Option<u32>,
}

impl FillStyle {
    /// No fill
    pub fn none() -> Self {
        Self {
            fill_type: FillType::Solid,
            color: ShapeColor::TRANSPARENT,
            back_color: None,
            opacity: 0,
            enabled: false,
            gradient_angle: None,
            picture_index: None,
        }
    }

    /// Solid fill with color
    pub fn solid(color: ShapeColor) -> Self {
        Self {
            fill_type: FillType::Solid,
            color,
            back_color: None,
            opacity: 100,
            enabled: true,
            gradient_angle: None,
            picture_index: None,
        }
    }

    /// Solid fill from RGB
    pub fn solid_rgb(r: u8, g: u8, b: u8) -> Self {
        Self::solid(ShapeColor::rgb(r, g, b))
    }

    /// Solid fill from hex
    pub fn solid_hex(hex: u32) -> Self {
        Self::solid(ShapeColor::from_hex(hex))
    }

    /// Gradient fill
    pub fn gradient(color1: ShapeColor, color2: ShapeColor, angle: i16) -> Self {
        Self {
            fill_type: FillType::Shade,
            color: color1,
            back_color: Some(color2),
            opacity: 100,
            enabled: true,
            gradient_angle: Some(angle),
            picture_index: None,
        }
    }

    /// Picture fill
    pub fn picture(blip_index: u32) -> Self {
        Self {
            fill_type: FillType::Picture,
            color: ShapeColor::TRANSPARENT,
            back_color: None,
            opacity: 100,
            enabled: true,
            gradient_angle: None,
            picture_index: Some(blip_index),
        }
    }

    /// Set opacity (0-100)
    pub fn with_opacity(mut self, opacity: u8) -> Self {
        self.opacity = opacity.min(100);
        self
    }

    /// Build Escher properties for this fill
    pub fn build_properties(&self) -> Vec<(u16, u32)> {
        let mut props = Vec::new();

        if !self.enabled {
            // No fill - set the boolean property
            props.push((fill_prop::FILL_STYLE_BOOL, 0x0001_0000)); // fNoFillHitTest = false, fFilled = false
            return props;
        }

        // Fill type
        props.push((fill_prop::FILL_TYPE, self.fill_type as u32));

        // Fill color
        props.push((fill_prop::FILL_COLOR, self.color.to_rgbx()));

        // Opacity (convert 0-100 to 0-65536)
        if self.opacity < 100 {
            let opacity_value = ((self.opacity as u32) * 65536) / 100;
            props.push((fill_prop::FILL_OPACITY, opacity_value));
        }

        // Background color
        if let Some(back) = &self.back_color {
            props.push((fill_prop::FILL_BACK_COLOR, back.to_rgbx()));
        }

        // Gradient angle
        if let Some(angle) = self.gradient_angle {
            let angle_value = (angle as i32 * 65536) as u32;
            props.push((fill_prop::FILL_ANGLE, angle_value));
        }

        // Picture BLIP reference
        if let Some(blip_idx) = self.picture_index {
            props.push((fill_prop::FILL_BLIP, blip_idx));
        }

        // Boolean properties (filled, hit test)
        props.push((fill_prop::FILL_STYLE_BOOL, 0x0011_0001)); // fNoFillHitTest = false, fFilled = true

        props
    }
}

impl Default for FillStyle {
    fn default() -> Self {
        Self::solid(ShapeColor::scheme_fill())
    }
}

// =============================================================================
// Line Style
// =============================================================================

/// Line style configuration
#[derive(Debug, Clone)]
pub struct LineStyleConfig {
    /// Line color
    pub color: ShapeColor,
    /// Line width in EMUs (914400 EMUs = 1 inch)
    pub width: u32,
    /// Line style (single, double, etc.)
    pub style: LineStyle,
    /// Dash style
    pub dash: LineDashStyle,
    /// Line cap style
    pub cap: LineCapStyle,
    /// Line join style
    pub join: LineJoinStyle,
    /// Opacity (0-100)
    pub opacity: u8,
    /// Whether line is enabled
    pub enabled: bool,
    /// Start arrow style
    pub start_arrow: ArrowStyle,
    /// End arrow style
    pub end_arrow: ArrowStyle,
    /// Start arrow width
    pub start_arrow_width: ArrowSize,
    /// Start arrow length
    pub start_arrow_length: ArrowSize,
    /// End arrow width
    pub end_arrow_width: ArrowSize,
    /// End arrow length
    pub end_arrow_length: ArrowSize,
}

impl LineStyleConfig {
    /// No line
    pub fn none() -> Self {
        Self {
            color: ShapeColor::TRANSPARENT,
            width: 0,
            style: LineStyle::Simple,
            dash: LineDashStyle::Solid,
            cap: LineCapStyle::Round,
            join: LineJoinStyle::Miter,
            opacity: 0,
            enabled: false,
            start_arrow: ArrowStyle::None,
            end_arrow: ArrowStyle::None,
            start_arrow_width: ArrowSize::Medium,
            start_arrow_length: ArrowSize::Medium,
            end_arrow_width: ArrowSize::Medium,
            end_arrow_length: ArrowSize::Medium,
        }
    }

    /// Default line (1pt black)
    pub fn default_line() -> Self {
        Self {
            color: ShapeColor::scheme_line(),
            width: EMUS_PER_PT as u32,
            style: LineStyle::Simple,
            dash: LineDashStyle::Solid,
            cap: LineCapStyle::Round,
            join: LineJoinStyle::Miter,
            opacity: 100,
            enabled: true,
            start_arrow: ArrowStyle::None,
            end_arrow: ArrowStyle::None,
            start_arrow_width: ArrowSize::Medium,
            start_arrow_length: ArrowSize::Medium,
            end_arrow_width: ArrowSize::Medium,
            end_arrow_length: ArrowSize::Medium,
        }
    }

    /// Create line with color and width in points
    pub fn with_color_and_width(color: ShapeColor, width_pt: f32) -> Self {
        Self {
            color,
            width: pt_f32_to_emu_u32(width_pt),
            style: LineStyle::Simple,
            dash: LineDashStyle::Solid,
            cap: LineCapStyle::Round,
            join: LineJoinStyle::Miter,
            opacity: 100,
            enabled: true,
            start_arrow: ArrowStyle::None,
            end_arrow: ArrowStyle::None,
            start_arrow_width: ArrowSize::Medium,
            start_arrow_length: ArrowSize::Medium,
            end_arrow_width: ArrowSize::Medium,
            end_arrow_length: ArrowSize::Medium,
        }
    }

    /// Set width in points
    pub fn width_pt(mut self, points: f32) -> Self {
        self.width = pt_f32_to_emu_u32(points);
        self
    }

    /// Set dash style
    pub fn dashed(mut self, style: LineDashStyle) -> Self {
        self.dash = style;
        self
    }

    /// Set start arrow
    pub fn start_arrow(mut self, style: ArrowStyle) -> Self {
        self.start_arrow = style;
        self
    }

    /// Set end arrow
    pub fn end_arrow(mut self, style: ArrowStyle) -> Self {
        self.end_arrow = style;
        self
    }

    /// Build Escher properties for this line
    pub fn build_properties(&self) -> Vec<(u16, u32)> {
        let mut props = Vec::new();

        if !self.enabled {
            // No line
            props.push((line_prop::LINE_STYLE_BOOL, 0x0008_0000)); // fNoLine = true
            return props;
        }

        // Line color
        props.push((line_prop::LINE_COLOR, self.color.to_rgbx()));

        // Line width
        props.push((line_prop::LINE_WIDTH, self.width));

        // Opacity
        if self.opacity < 100 {
            let opacity_value = ((self.opacity as u32) * 65536) / 100;
            props.push((line_prop::LINE_OPACITY, opacity_value));
        }

        // Line style
        if self.style != LineStyle::Simple {
            props.push((line_prop::LINE_STYLE, self.style as u32));
        }

        // Dash style
        if self.dash != LineDashStyle::Solid {
            props.push((line_prop::LINE_DASH_STYLE, self.dash as u32));
        }

        // Cap style
        if self.cap != LineCapStyle::Round {
            props.push((line_prop::LINE_END_CAP_STYLE, self.cap as u32));
        }

        // Join style
        if self.join != LineJoinStyle::Miter {
            props.push((line_prop::LINE_JOIN_STYLE, self.join as u32));
        }

        // Start arrow
        if self.start_arrow != ArrowStyle::None {
            props.push((line_prop::LINE_START_ARROW_HEAD, self.start_arrow as u32));
            props.push((
                line_prop::LINE_START_ARROW_WIDTH,
                self.start_arrow_width as u32,
            ));
            props.push((
                line_prop::LINE_START_ARROW_LENGTH,
                self.start_arrow_length as u32,
            ));
        }

        // End arrow
        if self.end_arrow != ArrowStyle::None {
            props.push((line_prop::LINE_END_ARROW_HEAD, self.end_arrow as u32));
            props.push((line_prop::LINE_END_ARROW_WIDTH, self.end_arrow_width as u32));
            props.push((
                line_prop::LINE_END_ARROW_LENGTH,
                self.end_arrow_length as u32,
            ));
        }

        // Boolean properties (line enabled)
        props.push((line_prop::LINE_STYLE_BOOL, 0x0008_0008)); // fLine = true

        props
    }
}

impl Default for LineStyleConfig {
    fn default() -> Self {
        Self::default_line()
    }
}

// =============================================================================
// Shadow Style
// =============================================================================

/// Shadow type
#[repr(u32)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum ShadowType {
    /// No shadow
    #[default]
    None = 0,
    /// Offset shadow
    Offset = 1,
    /// Double shadow
    Double = 2,
    /// Emboss shadow
    Emboss = 3,
}

/// Shadow style configuration
#[derive(Debug, Clone)]
pub struct ShadowStyle {
    /// Shadow type
    pub shadow_type: ShadowType,
    /// Shadow color
    pub color: ShapeColor,
    /// X offset in EMUs
    pub offset_x: i32,
    /// Y offset in EMUs
    pub offset_y: i32,
    /// Opacity (0-100)
    pub opacity: u8,
    /// Whether shadow is enabled
    pub enabled: bool,
}

impl ShadowStyle {
    /// No shadow
    pub fn none() -> Self {
        Self {
            shadow_type: ShadowType::None,
            color: ShapeColor::GRAY,
            offset_x: 0,
            offset_y: 0,
            opacity: 0,
            enabled: false,
        }
    }

    /// Default drop shadow
    pub fn drop_shadow() -> Self {
        Self {
            shadow_type: ShadowType::Offset,
            color: ShapeColor::scheme_shadow(),
            offset_x: 25400, // 2pt
            offset_y: 25400, // 2pt
            opacity: 50,
            enabled: true,
        }
    }

    /// Custom shadow
    pub fn custom(color: ShapeColor, offset_x: i32, offset_y: i32, opacity: u8) -> Self {
        Self {
            shadow_type: ShadowType::Offset,
            color,
            offset_x,
            offset_y,
            opacity,
            enabled: true,
        }
    }

    /// Build Escher properties for this shadow
    pub fn build_properties(&self) -> Vec<(u16, u32)> {
        let mut props = Vec::new();

        if !self.enabled {
            props.push((shadow_prop::SHADOW_STYLE_BOOL, 0x0002_0000)); // fShadow = false
            return props;
        }

        // Shadow type
        props.push((shadow_prop::SHADOW_TYPE, self.shadow_type as u32));

        // Shadow color
        props.push((shadow_prop::SHADOW_COLOR, self.color.to_rgbx()));

        // Offsets
        props.push((shadow_prop::SHADOW_OFFSET_X, self.offset_x as u32));
        props.push((shadow_prop::SHADOW_OFFSET_Y, self.offset_y as u32));

        // Opacity
        let opacity_value = ((self.opacity as u32) * 65536) / 100;
        props.push((shadow_prop::SHADOW_OPACITY, opacity_value));

        // Boolean properties (shadow enabled)
        props.push((shadow_prop::SHADOW_STYLE_BOOL, 0x0002_0002)); // fShadow = true

        props
    }
}

impl Default for ShadowStyle {
    fn default() -> Self {
        Self::none()
    }
}

// =============================================================================
// Combined Shape Style
// =============================================================================

/// Combined shape styling (fill + line + shadow)
#[derive(Debug, Clone, Default)]
pub struct ShapeStyle {
    /// Fill style
    pub fill: FillStyle,
    /// Line style
    pub line: LineStyleConfig,
    /// Shadow style
    pub shadow: ShadowStyle,
}

impl ShapeStyle {
    /// Create new empty style
    pub fn new() -> Self {
        Self::default()
    }

    /// Set fill
    pub fn with_fill(mut self, fill: FillStyle) -> Self {
        self.fill = fill;
        self
    }

    /// Set line
    pub fn with_line(mut self, line: LineStyleConfig) -> Self {
        self.line = line;
        self
    }

    /// Set shadow
    pub fn with_shadow(mut self, shadow: ShadowStyle) -> Self {
        self.shadow = shadow;
        self
    }

    /// No fill, default line
    pub fn no_fill() -> Self {
        Self {
            fill: FillStyle::none(),
            line: LineStyleConfig::default_line(),
            shadow: ShadowStyle::none(),
        }
    }

    /// Solid fill, no line
    pub fn solid_no_line(color: ShapeColor) -> Self {
        Self {
            fill: FillStyle::solid(color),
            line: LineStyleConfig::none(),
            shadow: ShadowStyle::none(),
        }
    }

    /// Build all Escher properties
    pub fn build_properties(&self) -> Vec<(u16, u32)> {
        let mut props = Vec::new();
        props.extend(self.fill.build_properties());
        props.extend(self.line.build_properties());
        props.extend(self.shadow.build_properties());
        props
    }
}

// =============================================================================
// Tests
// =============================================================================

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_shape_color() {
        // Red = RGB(255, 0, 0) -> BGR: B=0, G=0, R=255 -> (0) | (0<<8) | (255<<16) = 0x00FF0000
        let red = ShapeColor::RED;
        assert_eq!(red.to_rgbx(), 0x00FF0000);

        // 0x336699 = RGB(0x33, 0x66, 0x99) -> BGR: B=0x99, G=0x66, R=0x33
        // = (0x99) | (0x66<<8) | (0x33<<16) = 0x00336699
        let hex = ShapeColor::from_hex(0x336699);
        assert_eq!(hex.r, 0x33);
        assert_eq!(hex.g, 0x66);
        assert_eq!(hex.b, 0x99);
        assert_eq!(hex.to_rgbx(), 0x00336699);

        let scheme = ShapeColor::scheme_fill();
        assert!(scheme.use_scheme);
        assert_eq!(scheme.to_rgbx(), 0x08000004);
    }

    #[test]
    fn test_fill_style() {
        let fill = FillStyle::solid_rgb(255, 0, 0);
        let props = fill.build_properties();
        assert!(!props.is_empty());
        assert!(props.iter().any(|(id, _)| *id == fill_prop::FILL_COLOR));
    }

    #[test]
    fn test_line_style() {
        let line = LineStyleConfig::default_line().width_pt(2.0);
        assert_eq!(line.width, 25400); // 2pt in EMUs
        let props = line.build_properties();
        assert!(props.iter().any(|(id, _)| *id == line_prop::LINE_WIDTH));
    }

    #[test]
    fn test_combined_style() {
        let style = ShapeStyle::new()
            .with_fill(FillStyle::solid(ShapeColor::BLUE))
            .with_line(LineStyleConfig::none());
        let props = style.build_properties();
        assert!(!props.is_empty());
    }
}
