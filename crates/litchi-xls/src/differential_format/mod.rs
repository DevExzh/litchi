//! BIFF8 global differential formats (`DXF`) and formatting properties (`XFProps`).

use super::{
    BorderStyle, Error, FillPattern, FontCharset, FontEscapement, FontFamily, FontUnderline,
    HorizontalAlignment, ReadingOrder, Result, TextRotation, VerticalAlignment,
};

mod codec;
mod validation;

#[cfg(test)]
mod tests;

pub(crate) use codec::{validate_frt_header, write_frt_header};
use validation::validate_unit_interval;

pub(crate) const DXF_RECORD_TYPE: u16 = 0x088D;
const FRT_HEADER_LEN: usize = 12;
const FIXED_PAYLOAD_LEN: usize = FRT_HEADER_LEN + 2 + 4;
const MAX_BIFF8_PAYLOAD_LEN: usize = 8_224;
const MAX_XF_PROPERTIES: usize = 2_048;

fn invalid(message: impl Into<String>) -> Error {
    Error::InvalidRecord {
        record_type: DXF_RECORD_TYPE,
        message: message.into(),
    }
}

fn read_bytes<const N: usize>(data: &[u8], offset: usize, field: &str) -> Result<[u8; N]> {
    let end = offset
        .checked_add(N)
        .ok_or_else(|| invalid(format!("truncated {field}")))?;
    let bytes = data
        .get(offset..end)
        .ok_or_else(|| invalid(format!("truncated {field}")))?;
    bytes
        .try_into()
        .map_err(|_| invalid(format!("truncated {field}")))
}

fn read_u8(data: &[u8], offset: usize, field: &str) -> Result<u8> {
    data.get(offset)
        .copied()
        .ok_or_else(|| invalid(format!("truncated {field}")))
}

fn read_u16(data: &[u8], offset: usize, field: &str) -> Result<u16> {
    Ok(u16::from_le_bytes(read_bytes::<2>(data, offset, field)?))
}

fn read_u32(data: &[u8], offset: usize, field: &str) -> Result<u32> {
    Ok(u32::from_le_bytes(read_bytes::<4>(data, offset, field)?))
}

fn read_i16(data: &[u8], offset: usize, field: &str) -> Result<i16> {
    Ok(read_u16(data, offset, field)? as i16)
}

fn read_f64(data: &[u8], offset: usize, field: &str) -> Result<f64> {
    Ok(f64::from_le_bytes(read_bytes::<8>(data, offset, field)?))
}

/// A theme color slot used by an extended formatting property.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum ThemeColor {
    Dark1,
    Light1,
    Dark2,
    Light2,
    Accent1,
    Accent2,
    Accent3,
    Accent4,
    Accent5,
    Accent6,
    Hyperlink,
    FollowedHyperlink,
}

/// The source used to resolve an extended formatting color.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum XfColorSource {
    Automatic,
    Indexed(u8),
    Rgb,
    Theme(ThemeColor),
    NotSet,
}

/// An `XFPropColor`, including its resolved RGBA cache and tint.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct XfColor {
    source: XfColorSource,
    tint: i16,
    rgba: [u8; 4],
    ignored_index: u8,
}

impl XfColor {
    pub fn try_new(source: XfColorSource, tint: i16, rgba: [u8; 4]) -> Result<Self> {
        if tint == i16::MIN {
            return Err(invalid("XFPropColor tint cannot equal -32768"));
        }
        if let XfColorSource::Indexed(index) = source
            && !matches!(index, 0..=65 | 72)
        {
            return Err(invalid(format!("invalid indexed XF color {index}")));
        }
        let ignored_index = match source {
            XfColorSource::Indexed(index) => index,
            XfColorSource::Theme(theme) => theme.to_byte(),
            _ => 0,
        };
        Ok(Self {
            source,
            tint,
            rgba,
            ignored_index,
        })
    }

    pub const fn source(&self) -> XfColorSource {
        self.source
    }
    pub const fn tint(&self) -> i16 {
        self.tint
    }
    pub const fn rgba(&self) -> [u8; 4] {
        self.rgba
    }
}

/// Border formatting stored by an `XFProp`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct XfBorder {
    color: XfColor,
    style: BorderStyle,
}

impl XfBorder {
    pub const fn new(color: XfColor, style: BorderStyle) -> Self {
        Self { color, style }
    }
    pub const fn color(&self) -> XfColor {
        self.color
    }
    pub const fn style(&self) -> BorderStyle {
        self.style
    }
}

/// Gradient fill parameters stored by an `XFProp`.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct XfGradient {
    rectangular: bool,
    degree: f64,
    fill_to_left: f64,
    fill_to_right: f64,
    fill_to_top: f64,
    fill_to_bottom: f64,
}

impl XfGradient {
    pub fn linear(degree: f64) -> Result<Self> {
        if !degree.is_finite() {
            return Err(invalid("linear gradient degree must be finite"));
        }
        Ok(Self {
            rectangular: false,
            degree,
            fill_to_left: 0.0,
            fill_to_right: 0.0,
            fill_to_top: 0.0,
            fill_to_bottom: 0.0,
        })
    }

    pub fn rectangular(left: f64, right: f64, top: f64, bottom: f64) -> Result<Self> {
        validate_unit_interval(left, "left")?;
        validate_unit_interval(right, "right")?;
        validate_unit_interval(top, "top")?;
        validate_unit_interval(bottom, "bottom")?;
        Ok(Self {
            rectangular: true,
            degree: 0.0,
            fill_to_left: left,
            fill_to_right: right,
            fill_to_top: top,
            fill_to_bottom: bottom,
        })
    }

    pub const fn is_rectangular(&self) -> bool {
        self.rectangular
    }
    pub const fn degree(&self) -> f64 {
        self.degree
    }
    pub const fn fill_to_left(&self) -> f64 {
        self.fill_to_left
    }
    pub const fn fill_to_right(&self) -> f64 {
        self.fill_to_right
    }
    pub const fn fill_to_top(&self) -> f64 {
        self.fill_to_top
    }
    pub const fn fill_to_bottom(&self) -> f64 {
        self.fill_to_bottom
    }
}

/// One color stop in an extended gradient fill.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct XfGradientStop {
    position: f64,
    color: XfColor,
    unused: u16,
}

impl XfGradientStop {
    pub fn try_new(position: f64, color: XfColor) -> Result<Self> {
        validate_unit_interval(position, "stop")?;
        Ok(Self {
            position,
            color,
            unused: 0,
        })
    }
    pub const fn position(&self) -> f64 {
        self.position
    }
    pub const fn color(&self) -> XfColor {
        self.color
    }
}

/// Font weight stored by an extended formatting property.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum XfFontWeight {
    Normal,
    Bold,
}

/// Theme-font scheme stored by an extended formatting property.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum XfFontScheme {
    None,
    Major,
    Minor,
    NotSpecified,
}

/// One typed entry in an `XFProps` array.
#[derive(Debug, Clone, PartialEq)]
pub enum XfProperty {
    FillPattern(FillPattern),
    ForegroundColor(XfColor),
    BackgroundColor(XfColor),
    Gradient(XfGradient),
    GradientStop(XfGradientStop),
    TextColor(XfColor),
    TopBorder(XfBorder),
    BottomBorder(XfBorder),
    LeftBorder(XfBorder),
    RightBorder(XfBorder),
    DiagonalBorder(XfBorder),
    VerticalBorder(XfBorder),
    HorizontalBorder(XfBorder),
    DiagonalUp(bool),
    DiagonalDown(bool),
    /// `None` represents the specification's explicit "alignment not specified" value.
    HorizontalAlignment(Option<HorizontalAlignment>),
    VerticalAlignment(VerticalAlignment),
    TextRotation(TextRotation),
    AbsoluteIndent(u16),
    ReadingOrder(ReadingOrder),
    WrapText(bool),
    JustifyDistributed(bool),
    ShrinkToFit(bool),
    Merged(bool),
    FontName(String),
    FontWeight(XfFontWeight),
    FontUnderline(FontUnderline),
    FontEscapement(FontEscapement),
    FontItalic(bool),
    FontStrikethrough(bool),
    FontOutline(bool),
    FontShadow(bool),
    FontCondensed(bool),
    FontExtended(bool),
    FontCharset(FontCharset),
    FontFamily(FontFamily),
    FontSizeTwips(u32),
    FontScheme(XfFontScheme),
    NumberFormatCode(String),
    NumberFormatId(u16),
    RelativeIndent(Option<i16>),
    Locked(bool),
    Hidden(bool),
}

/// The complete ordered formatting-property array embedded in a DXF.
#[derive(Debug, Clone, PartialEq, Default)]
pub struct XfProperties {
    properties: Vec<XfProperty>,
}

impl XfProperties {
    pub fn try_new(properties: Vec<XfProperty>) -> Result<Self> {
        let value = Self { properties };
        validation::validate_properties(&value)?;
        Ok(value)
    }

    pub fn properties(&self) -> &[XfProperty] {
        &self.properties
    }

    pub fn is_empty(&self) -> bool {
        self.properties.is_empty()
    }
}

/// A global BIFF8 differential format referenced by table-style elements.
#[derive(Debug, Clone, PartialEq)]
pub struct DifferentialFormat {
    new_border: bool,
    properties: XfProperties,
    unused_flags: u16,
}

impl DifferentialFormat {
    pub fn try_new(new_border: bool, properties: Vec<XfProperty>) -> Result<Self> {
        let value = Self {
            new_border,
            properties: XfProperties::try_new(properties)?,
            unused_flags: 0,
        };
        validation::validate_format(&value)?;
        Ok(value)
    }

    pub const fn has_new_border(&self) -> bool {
        self.new_border
    }
    pub fn properties(&self) -> &XfProperties {
        &self.properties
    }
}
