//! BIFF8 global differential formats (`DXF`) and formatting properties (`XFProps`).

use super::{
    BorderStyle, Error, FillPattern, FontCharset, FontEscapement, FontFamily, FontUnderline,
    HorizontalAlignment, ReadingOrder, Result, TextRotation, VerticalAlignment,
};

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

impl ThemeColor {
    fn from_byte(value: u8) -> Result<Self> {
        match value {
            0 => Ok(Self::Dark1),
            1 => Ok(Self::Light1),
            2 => Ok(Self::Dark2),
            3 => Ok(Self::Light2),
            4 => Ok(Self::Accent1),
            5 => Ok(Self::Accent2),
            6 => Ok(Self::Accent3),
            7 => Ok(Self::Accent4),
            8 => Ok(Self::Accent5),
            9 => Ok(Self::Accent6),
            10 => Ok(Self::Hyperlink),
            11 => Ok(Self::FollowedHyperlink),
            _ => Err(invalid(format!("reserved theme color {value}"))),
        }
    }

    const fn to_byte(self) -> u8 {
        match self {
            Self::Dark1 => 0,
            Self::Light1 => 1,
            Self::Dark2 => 2,
            Self::Light2 => 3,
            Self::Accent1 => 4,
            Self::Accent2 => 5,
            Self::Accent3 => 6,
            Self::Accent4 => 7,
            Self::Accent5 => 8,
            Self::Accent6 => 9,
            Self::Hyperlink => 10,
            Self::FollowedHyperlink => 11,
        }
    }
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

    pub(crate) fn parse(data: &[u8]) -> Result<Self> {
        if data.len() != 8 {
            return Err(invalid(format!(
                "XFPropColor has {} bytes; expected 8",
                data.len()
            )));
        }
        let [flags, index, tint_low, tint_high, red, green, blue, alpha] =
            read_bytes::<8>(data, 0, "XFPropColor")?;
        if flags & 0x01 == 0 {
            return Err(invalid("XFPropColor fValidRGBA must be set"));
        }
        let source = match flags >> 1 {
            0 => XfColorSource::Automatic,
            1 => {
                if !matches!(index, 0..=65 | 72) {
                    return Err(invalid(format!("invalid indexed XF color {index}")));
                }
                XfColorSource::Indexed(index)
            },
            2 => XfColorSource::Rgb,
            3 => XfColorSource::Theme(ThemeColor::from_byte(index)?),
            4 => XfColorSource::NotSet,
            value => return Err(invalid(format!("reserved XF color source {value}"))),
        };
        let tint = i16::from_le_bytes([tint_low, tint_high]);
        if tint == i16::MIN {
            return Err(invalid("XFPropColor tint cannot equal -32768"));
        }
        Ok(Self {
            source,
            tint,
            rgba: [red, green, blue, alpha],
            ignored_index: index,
        })
    }

    pub(crate) fn write_to(&self, output: &mut Vec<u8>) {
        let (kind, index) = match self.source {
            XfColorSource::Automatic => (0, self.ignored_index),
            XfColorSource::Indexed(index) => (1, index),
            XfColorSource::Rgb => (2, self.ignored_index),
            XfColorSource::Theme(theme) => (3, theme.to_byte()),
            XfColorSource::NotSet => (4, self.ignored_index),
        };
        output.push(0x01 | (kind << 1));
        output.push(index);
        output.extend_from_slice(&self.tint.to_le_bytes());
        output.extend_from_slice(&self.rgba);
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

    pub(crate) fn parse(data: &[u8]) -> Result<Self> {
        if data.len() != 44 {
            return Err(invalid(format!(
                "XFPropGradient has {} bytes; expected 44",
                data.len()
            )));
        }
        let rectangular = match read_u32(data, 0, "XFPropGradient.type")? {
            0 => false,
            1 => true,
            value => return Err(invalid(format!("invalid gradient type {value}"))),
        };
        let value = Self {
            rectangular,
            degree: read_f64(data, 4, "XFPropGradient.numDegree")?,
            fill_to_left: read_f64(data, 12, "XFPropGradient.numFillToLeft")?,
            fill_to_right: read_f64(data, 20, "XFPropGradient.numFillToRight")?,
            fill_to_top: read_f64(data, 28, "XFPropGradient.numFillToTop")?,
            fill_to_bottom: read_f64(data, 36, "XFPropGradient.numFillToBottom")?,
        };
        if !value.degree.is_finite() {
            return Err(invalid("gradient degree must be finite"));
        }
        for (coordinate, name) in [
            (value.fill_to_left, "left"),
            (value.fill_to_right, "right"),
            (value.fill_to_top, "top"),
            (value.fill_to_bottom, "bottom"),
        ] {
            validate_unit_interval(coordinate, name)?;
        }
        Ok(value)
    }

    pub(crate) fn write_to(&self, output: &mut Vec<u8>) {
        output.extend_from_slice(&(self.rectangular as u32).to_le_bytes());
        output.extend_from_slice(&self.degree.to_le_bytes());
        output.extend_from_slice(&self.fill_to_left.to_le_bytes());
        output.extend_from_slice(&self.fill_to_right.to_le_bytes());
        output.extend_from_slice(&self.fill_to_top.to_le_bytes());
        output.extend_from_slice(&self.fill_to_bottom.to_le_bytes());
    }
}

fn validate_unit_interval(value: f64, field: &str) -> Result<()> {
    if !value.is_finite() || !(0.0..=1.0).contains(&value) {
        return Err(invalid(format!(
            "gradient {field} coordinate must be between 0.0 and 1.0"
        )));
    }
    Ok(())
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

    pub(crate) fn parse(data: &[u8]) -> Result<Self> {
        if data.len() != 18 {
            return Err(invalid(format!(
                "XFPropGradientStop has {} bytes; expected 18",
                data.len()
            )));
        }
        let stop = Self {
            position: read_f64(data, 2, "XFPropGradientStop.numPosition")?,
            color: XfColor::parse(&read_bytes::<8>(data, 10, "XFPropGradientStop.color")?)?,
            unused: read_u16(data, 0, "XFPropGradientStop.unused")?,
        };
        validate_unit_interval(stop.position, "stop")?;
        Ok(stop)
    }

    pub(crate) fn write_to(&self, output: &mut Vec<u8>) {
        output.extend_from_slice(&self.unused.to_le_bytes());
        output.extend_from_slice(&self.position.to_le_bytes());
        self.color.write_to(output);
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

impl XfProperty {
    fn property_type(&self) -> u16 {
        match self {
            Self::FillPattern(_) => 0x0000,
            Self::ForegroundColor(_) => 0x0001,
            Self::BackgroundColor(_) => 0x0002,
            Self::Gradient(_) => 0x0003,
            Self::GradientStop(_) => 0x0004,
            Self::TextColor(_) => 0x0005,
            Self::TopBorder(_) => 0x0006,
            Self::BottomBorder(_) => 0x0007,
            Self::LeftBorder(_) => 0x0008,
            Self::RightBorder(_) => 0x0009,
            Self::DiagonalBorder(_) => 0x000A,
            Self::VerticalBorder(_) => 0x000B,
            Self::HorizontalBorder(_) => 0x000C,
            Self::DiagonalUp(_) => 0x000D,
            Self::DiagonalDown(_) => 0x000E,
            Self::HorizontalAlignment(_) => 0x000F,
            Self::VerticalAlignment(_) => 0x0010,
            Self::TextRotation(_) => 0x0011,
            Self::AbsoluteIndent(_) => 0x0012,
            Self::ReadingOrder(_) => 0x0013,
            Self::WrapText(_) => 0x0014,
            Self::JustifyDistributed(_) => 0x0015,
            Self::ShrinkToFit(_) => 0x0016,
            Self::Merged(_) => 0x0017,
            Self::FontName(_) => 0x0018,
            Self::FontWeight(_) => 0x0019,
            Self::FontUnderline(_) => 0x001A,
            Self::FontEscapement(_) => 0x001B,
            Self::FontItalic(_) => 0x001C,
            Self::FontStrikethrough(_) => 0x001D,
            Self::FontOutline(_) => 0x001E,
            Self::FontShadow(_) => 0x001F,
            Self::FontCondensed(_) => 0x0020,
            Self::FontExtended(_) => 0x0021,
            Self::FontCharset(_) => 0x0022,
            Self::FontFamily(_) => 0x0023,
            Self::FontSizeTwips(_) => 0x0024,
            Self::FontScheme(_) => 0x0025,
            Self::NumberFormatCode(_) => 0x0026,
            Self::NumberFormatId(_) => 0x0029,
            Self::RelativeIndent(_) => 0x002A,
            Self::Locked(_) => 0x002B,
            Self::Hidden(_) => 0x002C,
        }
    }

    fn parse(property_type: u16, data: &[u8]) -> Result<Self> {
        let exact = |expected: usize| {
            if data.len() == expected {
                Ok(())
            } else {
                Err(invalid(format!(
                    "XFProp 0x{property_type:04X} has {} data bytes; expected {expected}",
                    data.len()
                )))
            }
        };
        match property_type {
            0x0000 => {
                exact(1)?;
                Ok(Self::FillPattern(parse_fill_pattern(read_u8(
                    data,
                    0,
                    "XFPropFillPattern.pattern",
                )?)?))
            },
            0x0001 => Ok(Self::ForegroundColor(XfColor::parse(data)?)),
            0x0002 => Ok(Self::BackgroundColor(XfColor::parse(data)?)),
            0x0003 => Ok(Self::Gradient(XfGradient::parse(data)?)),
            0x0004 => Ok(Self::GradientStop(XfGradientStop::parse(data)?)),
            0x0005 => Ok(Self::TextColor(XfColor::parse(data)?)),
            0x0006..=0x000C => {
                exact(10)?;
                let border = XfBorder {
                    color: XfColor::parse(&read_bytes::<8>(data, 0, "XFPropBorder.color")?)?,
                    style: parse_border_style(read_u16(data, 8, "XFPropBorder.dgBorder")?)?,
                };
                Ok(match property_type {
                    0x0006 => Self::TopBorder(border),
                    0x0007 => Self::BottomBorder(border),
                    0x0008 => Self::LeftBorder(border),
                    0x0009 => Self::RightBorder(border),
                    0x000A => Self::DiagonalBorder(border),
                    0x000B => Self::VerticalBorder(border),
                    _ => Self::HorizontalBorder(border),
                })
            },
            0x000D | 0x000E | 0x0014..=0x0017 | 0x001C..=0x0021 | 0x002B | 0x002C => {
                exact(1)?;
                let value = parse_bool(read_u8(data, 0, "XFProp.Boolean")?, property_type)?;
                Ok(match property_type {
                    0x000D => Self::DiagonalUp(value),
                    0x000E => Self::DiagonalDown(value),
                    0x0014 => Self::WrapText(value),
                    0x0015 => Self::JustifyDistributed(value),
                    0x0016 => Self::ShrinkToFit(value),
                    0x0017 => Self::Merged(value),
                    0x001C => Self::FontItalic(value),
                    0x001D => Self::FontStrikethrough(value),
                    0x001E => Self::FontOutline(value),
                    0x001F => Self::FontShadow(value),
                    0x0020 => Self::FontCondensed(value),
                    0x0021 => Self::FontExtended(value),
                    0x002B => Self::Locked(value),
                    _ => Self::Hidden(value),
                })
            },
            0x000F => {
                exact(1)?;
                Ok(Self::HorizontalAlignment(parse_horizontal_alignment(
                    read_u8(data, 0, "XFProp.horizontal alignment")?,
                )?))
            },
            0x0010 => {
                exact(1)?;
                Ok(Self::VerticalAlignment(parse_vertical_alignment(read_u8(
                    data,
                    0,
                    "XFProp.vertical alignment",
                )?)?))
            },
            0x0011 => {
                exact(1)?;
                Ok(Self::TextRotation(parse_rotation(read_u8(
                    data,
                    0,
                    "XFProp.text rotation",
                )?)?))
            },
            0x0012 => {
                exact(2)?;
                let value = read_u16(data, 0, "absolute indent")?;
                if value > 15 {
                    return Err(invalid(format!("absolute indent {value} exceeds 15")));
                }
                Ok(Self::AbsoluteIndent(value))
            },
            0x0013 => {
                exact(1)?;
                Ok(Self::ReadingOrder(parse_reading_order(read_u8(
                    data,
                    0,
                    "XFProp.reading order",
                )?)?))
            },
            0x0018 => Ok(Self::FontName(parse_lp_wide_string(data)?)),
            0x0019 => {
                exact(2)?;
                Ok(Self::FontWeight(match read_u16(data, 0, "font weight")? {
                    400 => XfFontWeight::Normal,
                    700 => XfFontWeight::Bold,
                    value => return Err(invalid(format!("reserved font weight {value}"))),
                }))
            },
            0x001A => {
                exact(2)?;
                Ok(Self::FontUnderline(parse_underline(read_u16(
                    data,
                    0,
                    "underline",
                )?)?))
            },
            0x001B => {
                exact(2)?;
                Ok(Self::FontEscapement(parse_escapement(read_u16(
                    data, 0, "script",
                )?)?))
            },
            0x0022 => {
                exact(1)?;
                Ok(Self::FontCharset(parse_charset(read_u8(
                    data,
                    0,
                    "XFProp.font charset",
                )?)?))
            },
            0x0023 => {
                exact(1)?;
                Ok(Self::FontFamily(parse_font_family(read_u8(
                    data,
                    0,
                    "XFProp.font family",
                )?)?))
            },
            0x0024 => {
                exact(4)?;
                let value = read_u32(data, 0, "font size")?;
                if !(20..=8191).contains(&value) {
                    return Err(invalid(format!(
                        "font size {value} is outside 20..=8191 twips"
                    )));
                }
                Ok(Self::FontSizeTwips(value))
            },
            0x0025 => {
                exact(1)?;
                Ok(Self::FontScheme(
                    match read_u8(data, 0, "XFProp.font scheme")? {
                        0 => XfFontScheme::None,
                        1 => XfFontScheme::Major,
                        2 => XfFontScheme::Minor,
                        0xFF => XfFontScheme::NotSpecified,
                        value => return Err(invalid(format!("reserved font scheme {value}"))),
                    },
                ))
            },
            0x0026 => Ok(Self::NumberFormatCode(parse_number_format_code(data)?)),
            0x0029 => {
                exact(2)?;
                Ok(Self::NumberFormatId(read_u16(data, 0, "number format id")?))
            },
            0x002A => {
                exact(2)?;
                let value = read_i16(data, 0, "relative indent")?;
                if value == 255 {
                    Ok(Self::RelativeIndent(None))
                } else if (-15..=15).contains(&value) {
                    Ok(Self::RelativeIndent(Some(value)))
                } else {
                    Err(invalid(format!("relative indent {value} is invalid")))
                }
            },
            _ => Err(invalid(format!(
                "reserved XF property type 0x{property_type:04X}"
            ))),
        }
    }

    fn data_bytes(&self) -> Result<Vec<u8>> {
        let mut data = Vec::new();
        match self {
            Self::FillPattern(value) => data.push(fill_pattern_byte(*value)),
            Self::ForegroundColor(value)
            | Self::BackgroundColor(value)
            | Self::TextColor(value) => value.write_to(&mut data),
            Self::Gradient(value) => value.write_to(&mut data),
            Self::GradientStop(value) => value.write_to(&mut data),
            Self::TopBorder(value)
            | Self::BottomBorder(value)
            | Self::LeftBorder(value)
            | Self::RightBorder(value)
            | Self::DiagonalBorder(value)
            | Self::VerticalBorder(value)
            | Self::HorizontalBorder(value) => {
                value.color.write_to(&mut data);
                data.extend_from_slice(&border_style_u16(value.style).to_le_bytes());
            },
            Self::DiagonalUp(value)
            | Self::DiagonalDown(value)
            | Self::WrapText(value)
            | Self::JustifyDistributed(value)
            | Self::ShrinkToFit(value)
            | Self::Merged(value)
            | Self::FontItalic(value)
            | Self::FontStrikethrough(value)
            | Self::FontOutline(value)
            | Self::FontShadow(value)
            | Self::FontCondensed(value)
            | Self::FontExtended(value)
            | Self::Locked(value)
            | Self::Hidden(value) => data.push(u8::from(*value)),
            Self::HorizontalAlignment(value) => data.push(horizontal_alignment_byte(*value)),
            Self::VerticalAlignment(value) => data.push(vertical_alignment_byte(*value)),
            Self::TextRotation(value) => data.push(rotation_byte(*value)?),
            Self::AbsoluteIndent(value) => {
                if *value > 15 {
                    return Err(invalid(format!("absolute indent {value} exceeds 15")));
                }
                data.extend_from_slice(&value.to_le_bytes());
            },
            Self::ReadingOrder(value) => data.push(reading_order_byte(*value)),
            Self::FontName(value) => write_lp_wide_string(value, &mut data)?,
            Self::FontWeight(value) => data.extend_from_slice(
                &match value {
                    XfFontWeight::Normal => 400u16,
                    XfFontWeight::Bold => 700u16,
                }
                .to_le_bytes(),
            ),
            Self::FontUnderline(value) => {
                data.extend_from_slice(&underline_u16(*value).to_le_bytes())
            },
            Self::FontEscapement(value) => {
                data.extend_from_slice(&escapement_u16(*value).to_le_bytes())
            },
            Self::FontCharset(value) => data.push(charset_byte(*value)),
            Self::FontFamily(value) => data.push(font_family_byte(*value)),
            Self::FontSizeTwips(value) => {
                if !(20..=8191).contains(value) {
                    return Err(invalid(format!(
                        "font size {value} is outside 20..=8191 twips"
                    )));
                }
                data.extend_from_slice(&value.to_le_bytes());
            },
            Self::FontScheme(value) => data.push(match value {
                XfFontScheme::None => 0,
                XfFontScheme::Major => 1,
                XfFontScheme::Minor => 2,
                XfFontScheme::NotSpecified => 0xFF,
            }),
            Self::NumberFormatCode(value) => write_number_format_code(value, &mut data)?,
            Self::NumberFormatId(value) => data.extend_from_slice(&value.to_le_bytes()),
            Self::RelativeIndent(value) => {
                let value = match value {
                    Some(value) if (-15..=15).contains(value) => *value,
                    Some(value) => {
                        return Err(invalid(format!("relative indent {value} is invalid")));
                    },
                    None => 255,
                };
                data.extend_from_slice(&value.to_le_bytes());
            },
        }
        Ok(data)
    }
}

/// The complete ordered formatting-property array embedded in a DXF.
#[derive(Debug, Clone, PartialEq, Default)]
pub struct XfProperties {
    properties: Vec<XfProperty>,
}

impl XfProperties {
    pub fn try_new(properties: Vec<XfProperty>) -> Result<Self> {
        let value = Self { properties };
        value.validate()?;
        Ok(value)
    }

    pub fn properties(&self) -> &[XfProperty] {
        &self.properties
    }

    pub fn is_empty(&self) -> bool {
        self.properties.is_empty()
    }

    pub(crate) fn parse(data: &[u8]) -> Result<Self> {
        if data.len() < 4 {
            return Err(invalid("truncated XFProps header"));
        }
        if read_u16(data, 0, "XFProps.reserved")? != 0 {
            return Err(invalid("XFProps reserved field must be zero"));
        }
        let count = usize::from(read_u16(data, 2, "XFProps.cprops")?);
        if count > MAX_XF_PROPERTIES {
            return Err(invalid(format!(
                "XFProps count {count} exceeds resource cap {MAX_XF_PROPERTIES}"
            )));
        }
        let mut offset = 4usize;
        let mut properties = Vec::with_capacity(count);
        for _ in 0..count {
            let property_type = read_u16(data, offset, "XFProp.xfPropType")?;
            let size_offset = offset
                .checked_add(2)
                .ok_or_else(|| invalid("XFProp header range overflows"))?;
            let size = usize::from(read_u16(data, size_offset, "XFProp.cb")?);
            if size < 4 {
                return Err(invalid(format!(
                    "XFProp size {size} is smaller than its header"
                )));
            }
            let data_offset = offset
                .checked_add(4)
                .ok_or_else(|| invalid("XFProp data range overflows"))?;
            let end = offset
                .checked_add(size)
                .ok_or_else(|| invalid("XFProp range overflows"))?;
            let blob = data
                .get(data_offset..end)
                .ok_or_else(|| invalid("truncated XFProp data"))?;
            properties.push(XfProperty::parse(property_type, blob)?);
            offset = end;
        }
        if offset != data.len() {
            return Err(invalid(
                "XFProps count does not consume its payload exactly",
            ));
        }
        Self::try_new(properties)
    }

    fn validate(&self) -> Result<()> {
        if self.properties.len() > MAX_XF_PROPERTIES || self.properties.len() > u16::MAX as usize {
            return Err(invalid(format!(
                "XFProps count exceeds resource cap {MAX_XF_PROPERTIES}"
            )));
        }
        let has_pattern = self
            .properties
            .iter()
            .any(|property| matches!(property, XfProperty::FillPattern(_)));
        let has_gradient = self.properties.iter().any(|property| {
            matches!(
                property,
                XfProperty::Gradient(_) | XfProperty::GradientStop(_)
            )
        });
        if has_pattern && has_gradient {
            return Err(invalid(
                "XFProps cannot combine pattern-fill and gradient properties",
            ));
        }
        let mut preceding_gradient = false;
        let mut distributed = false;
        let mut horizontal_distributed = false;
        for property in &self.properties {
            match property {
                XfProperty::Gradient(_) => preceding_gradient = true,
                XfProperty::GradientStop(_) if !preceding_gradient => {
                    return Err(invalid("gradient stop has no preceding gradient property"));
                },
                XfProperty::JustifyDistributed(true) => distributed = true,
                XfProperty::HorizontalAlignment(Some(HorizontalAlignment::Distributed)) => {
                    horizontal_distributed = true;
                },
                _ => {},
            }
            property.data_bytes()?;
        }
        if distributed && !horizontal_distributed {
            return Err(invalid(
                "justify-distributed requires distributed horizontal alignment",
            ));
        }
        Ok(())
    }

    pub(crate) fn to_bytes(&self) -> Result<Vec<u8>> {
        self.validate()?;
        let mut data = Vec::new();
        data.extend_from_slice(&0u16.to_le_bytes());
        let property_count = u16::try_from(self.properties.len())
            .map_err(|_| invalid("XFProps count exceeds BIFF u16"))?;
        data.extend_from_slice(&property_count.to_le_bytes());
        for property in &self.properties {
            let blob = property.data_bytes()?;
            let property_size = 4usize
                .checked_add(blob.len())
                .ok_or_else(|| invalid("XFProp size overflows"))?;
            let size = u16::try_from(property_size)
                .map_err(|_| invalid("XFProp size exceeds BIFF u16"))?;
            data.extend_from_slice(&property.property_type().to_le_bytes());
            data.extend_from_slice(&size.to_le_bytes());
            data.extend_from_slice(&blob);
        }
        Ok(data)
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
        value.validate()?;
        Ok(value)
    }

    pub const fn has_new_border(&self) -> bool {
        self.new_border
    }
    pub fn properties(&self) -> &XfProperties {
        &self.properties
    }

    pub fn parse_payload(data: &[u8]) -> Result<Self> {
        if !(FIXED_PAYLOAD_LEN..=MAX_BIFF8_PAYLOAD_LEN).contains(&data.len()) {
            return Err(invalid(format!(
                "DXF payload has {} bytes; expected {FIXED_PAYLOAD_LEN}..={MAX_BIFF8_PAYLOAD_LEN}",
                data.len()
            )));
        }
        validate_frt_header(data, DXF_RECORD_TYPE)?;
        let flags = read_u16(data, FRT_HEADER_LEN, "DXF flags")?;
        if flags & !0x0007 != 0 {
            return Err(invalid("DXF reserved flag bits must be zero"));
        }
        let value = Self {
            new_border: flags & 0x0002 != 0,
            unused_flags: flags & 0x0005,
            properties: XfProperties::parse(
                data.get(FRT_HEADER_LEN + 2..)
                    .ok_or_else(|| invalid("truncated DXF properties"))?,
            )?,
        };
        value.validate()?;
        Ok(value)
    }

    pub fn to_payload(&self) -> Result<Vec<u8>> {
        self.validate()?;
        let properties = self.properties.to_bytes()?;
        let size = FRT_HEADER_LEN
            .checked_add(2)
            .and_then(|size| size.checked_add(properties.len()))
            .ok_or_else(|| invalid("DXF payload size overflows"))?;
        if size > MAX_BIFF8_PAYLOAD_LEN {
            return Err(invalid("DXF exceeds the BIFF8 record payload cap"));
        }
        let mut data = Vec::with_capacity(size);
        write_frt_header(&mut data, DXF_RECORD_TYPE);
        data.extend_from_slice(
            &(self.unused_flags | if self.new_border { 2 } else { 0 }).to_le_bytes(),
        );
        data.extend_from_slice(&properties);
        Ok(data)
    }

    pub fn to_record_bytes(&self) -> Result<Vec<u8>> {
        let payload = self.to_payload()?;
        let record_len = 4usize
            .checked_add(payload.len())
            .ok_or_else(|| invalid("DXF record size overflows"))?;
        let payload_len =
            u16::try_from(payload.len()).map_err(|_| invalid("DXF payload exceeds BIFF u16"))?;
        let mut data = Vec::with_capacity(record_len);
        data.extend_from_slice(&DXF_RECORD_TYPE.to_le_bytes());
        data.extend_from_slice(&payload_len.to_le_bytes());
        data.extend_from_slice(&payload);
        Ok(data)
    }

    fn validate(&self) -> Result<()> {
        self.properties.validate()?;
        if !self.new_border
            && self.properties.properties.iter().any(|property| {
                matches!(
                    property,
                    XfProperty::VerticalBorder(_) | XfProperty::HorizontalBorder(_)
                )
            })
        {
            return Err(invalid("internal border properties require fNewBorder"));
        }
        Ok(())
    }
}

pub(crate) fn validate_frt_header(data: &[u8], record_type: u16) -> Result<()> {
    let header = data.get(..FRT_HEADER_LEN).ok_or(Error::InvalidRecord {
        record_type,
        message: "truncated FrtHeader".to_string(),
    })?;
    let actual_record_type = header
        .get(..2)
        .and_then(|bytes| <[u8; 2]>::try_from(bytes).ok())
        .map(u16::from_le_bytes)
        .ok_or(Error::InvalidRecord {
            record_type,
            message: "truncated FrtHeader".to_string(),
        })?;
    if actual_record_type != record_type {
        return Err(Error::InvalidRecord {
            record_type,
            message: "future-record type does not match record header".to_string(),
        });
    }
    if header
        .get(2..FRT_HEADER_LEN)
        .is_none_or(|reserved| reserved.iter().any(|byte| *byte != 0))
    {
        return Err(Error::InvalidRecord {
            record_type,
            message: "future-record reserved fields must be zero".to_string(),
        });
    }
    Ok(())
}

pub(crate) fn write_frt_header(output: &mut Vec<u8>, record_type: u16) {
    output.extend_from_slice(&record_type.to_le_bytes());
    output.extend_from_slice(&[0; 10]);
}

fn parse_bool(value: u8, property_type: u16) -> Result<bool> {
    match value {
        0 => Ok(false),
        1 => Ok(true),
        _ => Err(invalid(format!(
            "XFProp 0x{property_type:04X} Boolean is {value}"
        ))),
    }
}

fn parse_fill_pattern(value: u8) -> Result<FillPattern> {
    Ok(match value {
        0 => FillPattern::None,
        1 => FillPattern::Solid,
        2 => FillPattern::MediumGray,
        3 => FillPattern::DarkGray,
        4 => FillPattern::LightGray,
        5 => FillPattern::DarkHorizontal,
        6 => FillPattern::DarkVertical,
        7 => FillPattern::DarkDown,
        8 => FillPattern::DarkUp,
        9 => FillPattern::DarkGrid,
        10 => FillPattern::DarkTrellis,
        11 => FillPattern::LightHorizontal,
        12 => FillPattern::LightVertical,
        13 => FillPattern::LightDown,
        14 => FillPattern::LightUp,
        15 => FillPattern::LightGrid,
        16 => FillPattern::LightTrellis,
        17 => FillPattern::Gray125,
        18 => FillPattern::Gray0625,
        _ => return Err(invalid(format!("reserved fill pattern {value}"))),
    })
}

fn fill_pattern_byte(value: FillPattern) -> u8 {
    match value {
        FillPattern::None => 0,
        FillPattern::Solid => 1,
        FillPattern::MediumGray => 2,
        FillPattern::DarkGray => 3,
        FillPattern::LightGray => 4,
        FillPattern::DarkHorizontal => 5,
        FillPattern::DarkVertical => 6,
        FillPattern::DarkDown => 7,
        FillPattern::DarkUp => 8,
        FillPattern::DarkGrid => 9,
        FillPattern::DarkTrellis => 10,
        FillPattern::LightHorizontal => 11,
        FillPattern::LightVertical => 12,
        FillPattern::LightDown => 13,
        FillPattern::LightUp => 14,
        FillPattern::LightGrid => 15,
        FillPattern::LightTrellis => 16,
        FillPattern::Gray125 => 17,
        FillPattern::Gray0625 => 18,
    }
}

fn parse_border_style(value: u16) -> Result<BorderStyle> {
    Ok(match value {
        0 => BorderStyle::None,
        1 => BorderStyle::Thin,
        2 => BorderStyle::Medium,
        3 => BorderStyle::Dashed,
        4 => BorderStyle::Dotted,
        5 => BorderStyle::Thick,
        6 => BorderStyle::Double,
        7 => BorderStyle::Hair,
        8 => BorderStyle::MediumDashed,
        9 => BorderStyle::DashDot,
        10 => BorderStyle::MediumDashDot,
        11 => BorderStyle::DashDotDot,
        12 => BorderStyle::MediumDashDotDot,
        13 => BorderStyle::SlantedDashDot,
        _ => return Err(invalid(format!("reserved border style {value}"))),
    })
}

fn border_style_u16(value: BorderStyle) -> u16 {
    match value {
        BorderStyle::None => 0,
        BorderStyle::Thin => 1,
        BorderStyle::Medium => 2,
        BorderStyle::Dashed => 3,
        BorderStyle::Dotted => 4,
        BorderStyle::Thick => 5,
        BorderStyle::Double => 6,
        BorderStyle::Hair => 7,
        BorderStyle::MediumDashed => 8,
        BorderStyle::DashDot => 9,
        BorderStyle::MediumDashDot => 10,
        BorderStyle::DashDotDot => 11,
        BorderStyle::MediumDashDotDot => 12,
        BorderStyle::SlantedDashDot => 13,
    }
}

fn parse_horizontal_alignment(value: u8) -> Result<Option<HorizontalAlignment>> {
    Ok(match value {
        0 => Some(HorizontalAlignment::General),
        1 => Some(HorizontalAlignment::Left),
        2 => Some(HorizontalAlignment::Center),
        3 => Some(HorizontalAlignment::Right),
        4 => Some(HorizontalAlignment::Fill),
        5 => Some(HorizontalAlignment::Justify),
        6 => Some(HorizontalAlignment::CenterAcrossSelection),
        7 => Some(HorizontalAlignment::Distributed),
        0xFF => None,
        _ => return Err(invalid(format!("reserved horizontal alignment {value}"))),
    })
}

fn horizontal_alignment_byte(value: Option<HorizontalAlignment>) -> u8 {
    match value {
        Some(HorizontalAlignment::General) => 0,
        Some(HorizontalAlignment::Left) => 1,
        Some(HorizontalAlignment::Center) => 2,
        Some(HorizontalAlignment::Right) => 3,
        Some(HorizontalAlignment::Fill) => 4,
        Some(HorizontalAlignment::Justify) => 5,
        Some(HorizontalAlignment::CenterAcrossSelection) => 6,
        Some(HorizontalAlignment::Distributed) => 7,
        None => 0xFF,
    }
}

fn parse_vertical_alignment(value: u8) -> Result<VerticalAlignment> {
    Ok(match value {
        0 => VerticalAlignment::Top,
        1 => VerticalAlignment::Center,
        2 => VerticalAlignment::Bottom,
        3 => VerticalAlignment::Justify,
        4 => VerticalAlignment::Distributed,
        _ => return Err(invalid(format!("reserved vertical alignment {value}"))),
    })
}

fn vertical_alignment_byte(value: VerticalAlignment) -> u8 {
    match value {
        VerticalAlignment::Top => 0,
        VerticalAlignment::Center => 1,
        VerticalAlignment::Bottom => 2,
        VerticalAlignment::Justify => 3,
        VerticalAlignment::Distributed => 4,
    }
}

fn parse_rotation(value: u8) -> Result<TextRotation> {
    match value {
        0 => Ok(TextRotation::None),
        1..=90 => Ok(TextRotation::CounterClockwise(value)),
        91..=180 => Ok(TextRotation::Clockwise(value - 90)),
        255 => Ok(TextRotation::Vertical),
        _ => Err(invalid(format!("reserved text rotation {value}"))),
    }
}

fn rotation_byte(value: TextRotation) -> Result<u8> {
    match value {
        TextRotation::None => Ok(0),
        TextRotation::CounterClockwise(value @ 1..=90) => Ok(value),
        TextRotation::Clockwise(value @ 1..=90) => Ok(value + 90),
        TextRotation::Vertical => Ok(255),
        _ => Err(invalid("text rotation degrees must be between 1 and 90")),
    }
}

fn parse_reading_order(value: u8) -> Result<ReadingOrder> {
    match value {
        0 => Ok(ReadingOrder::Context),
        1 => Ok(ReadingOrder::LeftToRight),
        2 => Ok(ReadingOrder::RightToLeft),
        _ => Err(invalid(format!("reserved reading order {value}"))),
    }
}

fn reading_order_byte(value: ReadingOrder) -> u8 {
    match value {
        ReadingOrder::Context => 0,
        ReadingOrder::LeftToRight => 1,
        ReadingOrder::RightToLeft => 2,
    }
}

fn parse_underline(value: u16) -> Result<FontUnderline> {
    match value {
        0 => Ok(FontUnderline::None),
        1 => Ok(FontUnderline::Single),
        2 => Ok(FontUnderline::Double),
        0x21 => Ok(FontUnderline::SingleAccounting),
        0x22 => Ok(FontUnderline::DoubleAccounting),
        _ => Err(invalid(format!("reserved underline style {value}"))),
    }
}

fn underline_u16(value: FontUnderline) -> u16 {
    match value {
        FontUnderline::None => 0,
        FontUnderline::Single => 1,
        FontUnderline::Double => 2,
        FontUnderline::SingleAccounting => 0x21,
        FontUnderline::DoubleAccounting => 0x22,
    }
}

fn parse_escapement(value: u16) -> Result<FontEscapement> {
    match value {
        0 => Ok(FontEscapement::Normal),
        1 => Ok(FontEscapement::Superscript),
        2 => Ok(FontEscapement::Subscript),
        _ => Err(invalid(format!("reserved font escapement {value}"))),
    }
}

fn escapement_u16(value: FontEscapement) -> u16 {
    match value {
        FontEscapement::Normal => 0,
        FontEscapement::Superscript => 1,
        FontEscapement::Subscript => 2,
    }
}

fn parse_charset(value: u8) -> Result<FontCharset> {
    match value {
        0 => Ok(FontCharset::Ansi),
        1 => Ok(FontCharset::Default),
        2 => Ok(FontCharset::Symbol),
        77 => Ok(FontCharset::Mac),
        128 => Ok(FontCharset::ShiftJis),
        129 => Ok(FontCharset::Korean),
        130 => Ok(FontCharset::Johab),
        134 => Ok(FontCharset::Gb2312),
        136 => Ok(FontCharset::ChineseBig5),
        161 => Ok(FontCharset::Greek),
        162 => Ok(FontCharset::Turkish),
        163 => Ok(FontCharset::Vietnamese),
        177 => Ok(FontCharset::Hebrew),
        178 => Ok(FontCharset::Arabic),
        186 => Ok(FontCharset::Baltic),
        204 => Ok(FontCharset::Russian),
        222 => Ok(FontCharset::Thai),
        238 => Ok(FontCharset::EastEurope),
        255 => Ok(FontCharset::Oem),
        _ => Err(invalid(format!(
            "unsupported LOGFONT character set {value}"
        ))),
    }
}

fn charset_byte(value: FontCharset) -> u8 {
    match value {
        FontCharset::Ansi => 0,
        FontCharset::Default => 1,
        FontCharset::Symbol => 2,
        FontCharset::Mac => 77,
        FontCharset::ShiftJis => 128,
        FontCharset::Korean => 129,
        FontCharset::Johab => 130,
        FontCharset::Gb2312 => 134,
        FontCharset::ChineseBig5 => 136,
        FontCharset::Greek => 161,
        FontCharset::Turkish => 162,
        FontCharset::Vietnamese => 163,
        FontCharset::Hebrew => 177,
        FontCharset::Arabic => 178,
        FontCharset::Baltic => 186,
        FontCharset::Russian => 204,
        FontCharset::Thai => 222,
        FontCharset::EastEurope => 238,
        FontCharset::Oem => 255,
    }
}

fn parse_font_family(value: u8) -> Result<FontFamily> {
    match value {
        0 => Ok(FontFamily::NotApplicable),
        1 => Ok(FontFamily::Roman),
        2 => Ok(FontFamily::Swiss),
        3 => Ok(FontFamily::Modern),
        4 => Ok(FontFamily::Script),
        5 => Ok(FontFamily::Decorative),
        _ => Err(invalid(format!("reserved font family {value}"))),
    }
}

fn font_family_byte(value: FontFamily) -> u8 {
    match value {
        FontFamily::NotApplicable => 0,
        FontFamily::Roman => 1,
        FontFamily::Swiss => 2,
        FontFamily::Modern => 3,
        FontFamily::Script => 4,
        FontFamily::Decorative => 5,
    }
}

fn parse_lp_wide_string(data: &[u8]) -> Result<String> {
    let count = usize::from(read_u16(data, 0, "LPWideString.cchCharacters")?);
    if count > 32 {
        return Err(invalid(
            "font name LPWideString is malformed or exceeds 32 characters",
        ));
    }
    let byte_len = count
        .checked_mul(2)
        .ok_or_else(|| invalid("font name LPWideString size overflows"))?;
    let end = 2usize
        .checked_add(byte_len)
        .ok_or_else(|| invalid("font name LPWideString range overflows"))?;
    if data.len() != end {
        return Err(invalid(
            "font name LPWideString is malformed or exceeds 32 characters",
        ));
    }
    decode_utf16(
        data.get(2..end)
            .ok_or_else(|| invalid("truncated font name LPWideString"))?,
        "font name",
    )
}

fn write_lp_wide_string(value: &str, data: &mut Vec<u8>) -> Result<()> {
    let units = value.encode_utf16().collect::<Vec<_>>();
    if units.len() > 32 {
        return Err(invalid("font name exceeds 32 UTF-16 code units"));
    }
    let count =
        u16::try_from(units.len()).map_err(|_| invalid("font name length exceeds BIFF u16"))?;
    data.extend_from_slice(&count.to_le_bytes());
    data.extend(units.into_iter().flat_map(u16::to_le_bytes));
    Ok(())
}

fn parse_number_format_code(data: &[u8]) -> Result<String> {
    if data.len() < 2 {
        return Err(invalid("truncated number-format string"));
    }
    let count = usize::from(read_u16(data, 0, "number-format cch")?);
    if !(1..=255).contains(&count) {
        return Err(invalid("number-format string length must be 1..=255"));
    }
    let char_len = count
        .checked_mul(2)
        .ok_or_else(|| invalid("number-format string size overflows"))?;
    let end = 2usize
        .checked_add(char_len)
        .ok_or_else(|| invalid("number-format string range overflows"))?;
    if end != data.len() {
        return Err(invalid(
            "number-format string has trailing or truncated UTF-16 data",
        ));
    }
    decode_utf16(
        data.get(2..end)
            .ok_or_else(|| invalid("truncated number-format string"))?,
        "number-format string",
    )
}

fn write_number_format_code(value: &str, data: &mut Vec<u8>) -> Result<()> {
    let units = value.encode_utf16().collect::<Vec<_>>();
    if !(1..=255).contains(&units.len()) {
        return Err(invalid("number-format string length must be 1..=255"));
    }
    let count = u16::try_from(units.len())
        .map_err(|_| invalid("number-format string length exceeds BIFF u16"))?;
    data.extend_from_slice(&count.to_le_bytes());
    data.extend(units.into_iter().flat_map(u16::to_le_bytes));
    Ok(())
}

fn decode_utf16(data: &[u8], field: &str) -> Result<String> {
    if !data.len().is_multiple_of(2) {
        return Err(invalid(format!("{field} has an odd byte length")));
    }
    let mut units = Vec::with_capacity(data.len() / 2);
    for chunk in data.chunks_exact(2) {
        let bytes = <[u8; 2]>::try_from(chunk)
            .map_err(|_| invalid(format!("{field} has an invalid UTF-16 unit")))?;
        units.push(u16::from_le_bytes(bytes));
    }
    char::decode_utf16(units)
        .collect::<Result<String, _>>()
        .map_err(|_| invalid(format!("{field} contains invalid UTF-16")))
}

#[cfg(test)]
mod tests {
    use super::*;

    fn color() -> XfColor {
        XfColor::try_new(
            XfColorSource::Theme(ThemeColor::Accent2),
            100,
            [1, 2, 3, 255],
        )
        .unwrap()
    }

    #[test]
    fn typed_dxf_round_trips_representative_property_families() {
        let dxf = DifferentialFormat::try_new(
            true,
            vec![
                XfProperty::Gradient(XfGradient::linear(45.0).unwrap()),
                XfProperty::GradientStop(XfGradientStop::try_new(0.5, color()).unwrap()),
                XfProperty::TopBorder(XfBorder::new(color(), BorderStyle::Thin)),
                XfProperty::VerticalBorder(XfBorder::new(color(), BorderStyle::Dashed)),
                XfProperty::HorizontalAlignment(Some(HorizontalAlignment::Distributed)),
                XfProperty::JustifyDistributed(true),
                XfProperty::TextRotation(TextRotation::Clockwise(30)),
                XfProperty::FontName("Aptos".to_string()),
                XfProperty::FontWeight(XfFontWeight::Bold),
                XfProperty::FontUnderline(FontUnderline::Double),
                XfProperty::FontSizeTwips(220),
                XfProperty::NumberFormatCode("0.00".to_string()),
                XfProperty::RelativeIndent(Some(-2)),
                XfProperty::Locked(true),
            ],
        )
        .unwrap();
        let payload = dxf.to_payload().unwrap();
        assert_eq!(DifferentialFormat::parse_payload(&payload).unwrap(), dxf);
        let record = dxf.to_record_bytes().unwrap();
        assert_eq!(&record[..2], &[0x8D, 0x08]);
    }

    #[test]
    fn xfprop_color_uses_low_flag_bit_and_high_seven_type_bits() {
        // Apache POI producer forms: bit 7 is clear, while fValidRGBA in bit 0 is set.
        let rgb = [0x05, 0xFF, 0x00, 0x00, 0xFF, 0xC7, 0xCE, 0xFF];
        let parsed = XfColor::parse(&rgb).unwrap();
        assert_eq!(parsed.source(), XfColorSource::Rgb);
        let mut encoded = Vec::new();
        parsed.write_to(&mut encoded);
        assert_eq!(encoded, rgb);

        let theme = [0x07, 0x04, 0x65, 0x66, 0xDC, 0xE6, 0xF1, 0xFF];
        let parsed = XfColor::parse(&theme).unwrap();
        assert_eq!(parsed.source(), XfColorSource::Theme(ThemeColor::Accent1));
        let mut encoded = Vec::new();
        parsed.write_to(&mut encoded);
        assert_eq!(encoded, theme);

        let indexed = [0x03, 0x40, 0, 0, 1, 2, 3, 4];
        assert_eq!(
            XfColor::parse(&indexed).unwrap().source(),
            XfColorSource::Indexed(0x40)
        );
        let automatic = [0x01, 0xAA, 0, 0, 1, 2, 3, 4];
        assert_eq!(
            XfColor::parse(&automatic).unwrap().source(),
            XfColorSource::Automatic
        );
        let not_set = [0x09, 0xAA, 0, 0, 1, 2, 3, 4];
        assert_eq!(
            XfColor::parse(&not_set).unwrap().source(),
            XfColorSource::NotSet
        );
    }

    #[test]
    fn xfprop_color_rejects_clear_valid_flag_and_invalid_type_data() {
        assert!(XfColor::parse(&[0x04, 0, 0, 0, 0, 0, 0, 0]).is_err());
        assert!(XfColor::parse(&[0x0B, 0, 0, 0, 0, 0, 0, 0]).is_err());
        assert!(XfColor::parse(&[0x03, 66, 0, 0, 0, 0, 0, 0]).is_err());
        assert!(XfColor::parse(&[0x07, 12, 0, 0, 0, 0, 0, 0]).is_err());
        assert!(XfColor::parse(&[0x05, 0, 0x00, 0x80, 0, 0, 0, 0]).is_err());
    }

    #[test]
    fn rejects_hostile_headers_sizes_flags_and_property_relationships() {
        let empty = DifferentialFormat::try_new(false, vec![])
            .unwrap()
            .to_payload()
            .unwrap();
        assert!(DifferentialFormat::parse_payload(&empty[..17]).is_err());
        let mut bad = empty.clone();
        bad[0] = 0;
        assert!(DifferentialFormat::parse_payload(&bad).is_err());
        let mut bad = empty.clone();
        bad[14] = 1;
        assert!(DifferentialFormat::parse_payload(&bad).is_err());
        let mut bad = empty;
        bad[16..18].copy_from_slice(&1u16.to_le_bytes());
        assert!(DifferentialFormat::parse_payload(&bad).is_err());

        assert!(
            DifferentialFormat::try_new(
                false,
                vec![XfProperty::VerticalBorder(XfBorder::new(
                    color(),
                    BorderStyle::Thin,
                ))],
            )
            .is_err()
        );
        assert!(
            DifferentialFormat::try_new(
                false,
                vec![
                    XfProperty::FillPattern(FillPattern::Solid),
                    XfProperty::Gradient(XfGradient::linear(0.0).unwrap()),
                ],
            )
            .is_err()
        );
        assert!(
            DifferentialFormat::try_new(false, vec![XfProperty::JustifyDistributed(true)],)
                .is_err()
        );
    }

    #[test]
    fn fixed_width_reads_reject_offset_overflow_without_panicking() {
        assert!(matches!(
            std::panic::catch_unwind(|| read_u16(&[], usize::MAX, "u16")),
            Ok(Err(_))
        ));
        assert!(matches!(
            std::panic::catch_unwind(|| read_u32(&[], usize::MAX, "u32")),
            Ok(Err(_))
        ));
        assert!(matches!(
            std::panic::catch_unwind(|| read_f64(&[], usize::MAX, "f64")),
            Ok(Err(_))
        ));
    }

    #[test]
    fn malformed_fixed_width_properties_return_errors_without_panicking() {
        let empty = DifferentialFormat::try_new(false, vec![])
            .unwrap()
            .to_payload()
            .unwrap();
        for property_type in [
            0x0000u16, 0x0001, 0x0003, 0x0004, 0x0006, 0x000D, 0x000F, 0x0010, 0x0011, 0x0012,
            0x0013, 0x0018, 0x0019, 0x001A, 0x001B, 0x0022, 0x0023, 0x0024, 0x0025, 0x0029, 0x002A,
        ] {
            let mut payload = empty.clone();
            payload[16..18].copy_from_slice(&1u16.to_le_bytes());
            payload.extend_from_slice(&property_type.to_le_bytes());
            payload.extend_from_slice(&4u16.to_le_bytes());
            let parsed = std::panic::catch_unwind(|| DifferentialFormat::parse_payload(&payload));
            assert!(
                matches!(parsed, Ok(Err(_))),
                "property type 0x{property_type:04X} did not reject empty data"
            );
        }
    }

    #[test]
    fn oversized_dxf_writes_are_rejected_without_truncating_record_length() {
        let dxf = DifferentialFormat::try_new(
            false,
            vec![XfProperty::WrapText(false); MAX_XF_PROPERTIES],
        )
        .unwrap();
        assert!(dxf.to_payload().is_err());
        assert!(dxf.to_record_bytes().is_err());
    }

    #[test]
    fn enforces_resource_caps() {
        assert!(
            XfProperties::try_new(vec![XfProperty::WrapText(false); MAX_XF_PROPERTIES + 1])
                .is_err()
        );
        let huge = "x".repeat(256);
        assert!(
            DifferentialFormat::try_new(false, vec![XfProperty::NumberFormatCode(huge)]).is_err()
        );
    }

    #[test]
    fn number_format_code_matches_producer_wide_string_bytes() {
        let producer = [
            0x05, 0x00, 0x22, 0x00, 0x24, 0x00, 0x22, 0x00, 0x30, 0x00, 0x30, 0x00,
        ];
        assert_eq!(parse_number_format_code(&producer).unwrap(), "\"$\"00");

        let mut encoded = Vec::new();
        write_number_format_code("\"$\"00", &mut encoded).unwrap();
        assert_eq!(encoded, producer);
    }

    #[test]
    fn number_format_code_rejects_malformed_wide_strings() {
        assert!(parse_number_format_code(&[0x00, 0x00]).is_err());
        assert!(parse_number_format_code(&[0x00, 0x01]).is_err());
        assert!(parse_number_format_code(&[0x02, 0x00, 0x30, 0x00]).is_err());
        assert!(parse_number_format_code(&[0x01, 0x00, 0x30, 0x00, 0x00]).is_err());
        assert!(parse_number_format_code(&[0x01, 0x00, 0x00, 0xD8]).is_err());

        // A BIFF XLUnicodeString flags byte is not part of this XFProp payload.
        assert!(parse_number_format_code(&[0x01, 0x00, 0x01, 0x30, 0x00]).is_err());
    }
}
