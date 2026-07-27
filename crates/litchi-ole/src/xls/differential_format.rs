//! BIFF8 global differential formats (`DXF`) and formatting properties (`XFProps`).

use super::{
    XlsBorderStyle, XlsError, XlsFillPattern, XlsFontCharset, XlsFontEscapement, XlsFontFamily,
    XlsFontUnderline, XlsHorizontalAlignment, XlsReadingOrder, XlsResult, XlsTextRotation,
    XlsVerticalAlignment,
};

pub(crate) const DXF_RECORD_TYPE: u16 = 0x088D;
const FRT_HEADER_LEN: usize = 12;
const FIXED_PAYLOAD_LEN: usize = FRT_HEADER_LEN + 2 + 4;
const MAX_BIFF8_PAYLOAD_LEN: usize = 8_224;
const MAX_XF_PROPERTIES: usize = 2_048;

fn invalid(message: impl Into<String>) -> XlsError {
    XlsError::InvalidRecord {
        record_type: DXF_RECORD_TYPE,
        message: message.into(),
    }
}

fn read_u16(data: &[u8], offset: usize, field: &str) -> XlsResult<u16> {
    let bytes = data
        .get(offset..offset + 2)
        .ok_or_else(|| invalid(format!("truncated {field}")))?;
    Ok(u16::from_le_bytes([bytes[0], bytes[1]]))
}

fn read_u32(data: &[u8], offset: usize, field: &str) -> XlsResult<u32> {
    let bytes = data
        .get(offset..offset + 4)
        .ok_or_else(|| invalid(format!("truncated {field}")))?;
    Ok(u32::from_le_bytes(bytes.try_into().unwrap()))
}

fn read_i16(data: &[u8], offset: usize, field: &str) -> XlsResult<i16> {
    Ok(read_u16(data, offset, field)? as i16)
}

fn read_f64(data: &[u8], offset: usize, field: &str) -> XlsResult<f64> {
    let bytes = data
        .get(offset..offset + 8)
        .ok_or_else(|| invalid(format!("truncated {field}")))?;
    Ok(f64::from_le_bytes(bytes.try_into().unwrap()))
}

/// A theme color slot used by an extended formatting property.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum XlsThemeColor {
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

impl XlsThemeColor {
    fn from_byte(value: u8) -> XlsResult<Self> {
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
pub enum XlsXfColorSource {
    Automatic,
    Indexed(u8),
    Rgb,
    Theme(XlsThemeColor),
    NotSet,
}

/// An `XFPropColor`, including its resolved RGBA cache and tint.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct XlsXfColor {
    source: XlsXfColorSource,
    tint: i16,
    rgba: [u8; 4],
    ignored_index: u8,
}

impl XlsXfColor {
    pub fn try_new(source: XlsXfColorSource, tint: i16, rgba: [u8; 4]) -> XlsResult<Self> {
        if tint == i16::MIN {
            return Err(invalid("XFPropColor tint cannot equal -32768"));
        }
        if let XlsXfColorSource::Indexed(index) = source {
            if !matches!(index, 0..=65 | 72) {
                return Err(invalid(format!("invalid indexed XF color {index}")));
            }
        }
        let ignored_index = match source {
            XlsXfColorSource::Indexed(index) => index,
            XlsXfColorSource::Theme(theme) => theme.to_byte(),
            _ => 0,
        };
        Ok(Self {
            source,
            tint,
            rgba,
            ignored_index,
        })
    }

    pub const fn source(&self) -> XlsXfColorSource {
        self.source
    }
    pub const fn tint(&self) -> i16 {
        self.tint
    }
    pub const fn rgba(&self) -> [u8; 4] {
        self.rgba
    }

    pub(crate) fn parse(data: &[u8]) -> XlsResult<Self> {
        if data.len() != 8 {
            return Err(invalid(format!(
                "XFPropColor has {} bytes; expected 8",
                data.len()
            )));
        }
        if data[0] & 0x01 == 0 {
            return Err(invalid("XFPropColor fValidRGBA must be set"));
        }
        let source = match data[0] >> 1 {
            0 => XlsXfColorSource::Automatic,
            1 => {
                if !matches!(data[1], 0..=65 | 72) {
                    return Err(invalid(format!("invalid indexed XF color {}", data[1])));
                }
                XlsXfColorSource::Indexed(data[1])
            },
            2 => XlsXfColorSource::Rgb,
            3 => XlsXfColorSource::Theme(XlsThemeColor::from_byte(data[1])?),
            4 => XlsXfColorSource::NotSet,
            value => return Err(invalid(format!("reserved XF color source {value}"))),
        };
        let tint = i16::from_le_bytes([data[2], data[3]]);
        if tint == i16::MIN {
            return Err(invalid("XFPropColor tint cannot equal -32768"));
        }
        Ok(Self {
            source,
            tint,
            rgba: data[4..8].try_into().unwrap(),
            ignored_index: data[1],
        })
    }

    pub(crate) fn write_to(&self, output: &mut Vec<u8>) {
        let (kind, index) = match self.source {
            XlsXfColorSource::Automatic => (0, self.ignored_index),
            XlsXfColorSource::Indexed(index) => (1, index),
            XlsXfColorSource::Rgb => (2, self.ignored_index),
            XlsXfColorSource::Theme(theme) => (3, theme.to_byte()),
            XlsXfColorSource::NotSet => (4, self.ignored_index),
        };
        output.push(0x01 | (kind << 1));
        output.push(index);
        output.extend_from_slice(&self.tint.to_le_bytes());
        output.extend_from_slice(&self.rgba);
    }
}

/// Border formatting stored by an `XFProp`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct XlsXfBorder {
    color: XlsXfColor,
    style: XlsBorderStyle,
}

impl XlsXfBorder {
    pub const fn new(color: XlsXfColor, style: XlsBorderStyle) -> Self {
        Self { color, style }
    }
    pub const fn color(&self) -> XlsXfColor {
        self.color
    }
    pub const fn style(&self) -> XlsBorderStyle {
        self.style
    }
}

/// Gradient fill parameters stored by an `XFProp`.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct XlsXfGradient {
    rectangular: bool,
    degree: f64,
    fill_to_left: f64,
    fill_to_right: f64,
    fill_to_top: f64,
    fill_to_bottom: f64,
}

impl XlsXfGradient {
    pub fn linear(degree: f64) -> XlsResult<Self> {
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

    pub fn rectangular(left: f64, right: f64, top: f64, bottom: f64) -> XlsResult<Self> {
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

    pub(crate) fn parse(data: &[u8]) -> XlsResult<Self> {
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

fn validate_unit_interval(value: f64, field: &str) -> XlsResult<()> {
    if !value.is_finite() || !(0.0..=1.0).contains(&value) {
        return Err(invalid(format!(
            "gradient {field} coordinate must be between 0.0 and 1.0"
        )));
    }
    Ok(())
}

/// One color stop in an extended gradient fill.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct XlsXfGradientStop {
    position: f64,
    color: XlsXfColor,
    unused: u16,
}

impl XlsXfGradientStop {
    pub fn try_new(position: f64, color: XlsXfColor) -> XlsResult<Self> {
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
    pub const fn color(&self) -> XlsXfColor {
        self.color
    }

    pub(crate) fn parse(data: &[u8]) -> XlsResult<Self> {
        if data.len() != 18 {
            return Err(invalid(format!(
                "XFPropGradientStop has {} bytes; expected 18",
                data.len()
            )));
        }
        let stop = Self {
            position: read_f64(data, 2, "XFPropGradientStop.numPosition")?,
            color: XlsXfColor::parse(&data[10..18])?,
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
pub enum XlsXfFontWeight {
    Normal,
    Bold,
}

/// Theme-font scheme stored by an extended formatting property.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum XlsXfFontScheme {
    None,
    Major,
    Minor,
    NotSpecified,
}

/// One typed entry in an `XFProps` array.
#[derive(Debug, Clone, PartialEq)]
pub enum XlsXfProperty {
    FillPattern(XlsFillPattern),
    ForegroundColor(XlsXfColor),
    BackgroundColor(XlsXfColor),
    Gradient(XlsXfGradient),
    GradientStop(XlsXfGradientStop),
    TextColor(XlsXfColor),
    TopBorder(XlsXfBorder),
    BottomBorder(XlsXfBorder),
    LeftBorder(XlsXfBorder),
    RightBorder(XlsXfBorder),
    DiagonalBorder(XlsXfBorder),
    VerticalBorder(XlsXfBorder),
    HorizontalBorder(XlsXfBorder),
    DiagonalUp(bool),
    DiagonalDown(bool),
    /// `None` represents the specification's explicit "alignment not specified" value.
    HorizontalAlignment(Option<XlsHorizontalAlignment>),
    VerticalAlignment(XlsVerticalAlignment),
    TextRotation(XlsTextRotation),
    AbsoluteIndent(u16),
    ReadingOrder(XlsReadingOrder),
    WrapText(bool),
    JustifyDistributed(bool),
    ShrinkToFit(bool),
    Merged(bool),
    FontName(String),
    FontWeight(XlsXfFontWeight),
    FontUnderline(XlsFontUnderline),
    FontEscapement(XlsFontEscapement),
    FontItalic(bool),
    FontStrikethrough(bool),
    FontOutline(bool),
    FontShadow(bool),
    FontCondensed(bool),
    FontExtended(bool),
    FontCharset(XlsFontCharset),
    FontFamily(XlsFontFamily),
    FontSizeTwips(u32),
    FontScheme(XlsXfFontScheme),
    NumberFormatCode(String),
    NumberFormatId(u16),
    RelativeIndent(Option<i16>),
    Locked(bool),
    Hidden(bool),
}

impl XlsXfProperty {
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

    fn parse(property_type: u16, data: &[u8]) -> XlsResult<Self> {
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
                Ok(Self::FillPattern(parse_fill_pattern(data[0])?))
            },
            0x0001 => Ok(Self::ForegroundColor(XlsXfColor::parse(data)?)),
            0x0002 => Ok(Self::BackgroundColor(XlsXfColor::parse(data)?)),
            0x0003 => Ok(Self::Gradient(XlsXfGradient::parse(data)?)),
            0x0004 => Ok(Self::GradientStop(XlsXfGradientStop::parse(data)?)),
            0x0005 => Ok(Self::TextColor(XlsXfColor::parse(data)?)),
            0x0006..=0x000C => {
                exact(10)?;
                let border = XlsXfBorder {
                    color: XlsXfColor::parse(&data[..8])?,
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
                let value = parse_bool(data[0], property_type)?;
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
                    data[0],
                )?))
            },
            0x0010 => {
                exact(1)?;
                Ok(Self::VerticalAlignment(parse_vertical_alignment(data[0])?))
            },
            0x0011 => {
                exact(1)?;
                Ok(Self::TextRotation(parse_rotation(data[0])?))
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
                Ok(Self::ReadingOrder(parse_reading_order(data[0])?))
            },
            0x0018 => Ok(Self::FontName(parse_lp_wide_string(data)?)),
            0x0019 => {
                exact(2)?;
                Ok(Self::FontWeight(match read_u16(data, 0, "font weight")? {
                    400 => XlsXfFontWeight::Normal,
                    700 => XlsXfFontWeight::Bold,
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
                Ok(Self::FontCharset(parse_charset(data[0])?))
            },
            0x0023 => {
                exact(1)?;
                Ok(Self::FontFamily(parse_font_family(data[0])?))
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
                Ok(Self::FontScheme(match data[0] {
                    0 => XlsXfFontScheme::None,
                    1 => XlsXfFontScheme::Major,
                    2 => XlsXfFontScheme::Minor,
                    0xFF => XlsXfFontScheme::NotSpecified,
                    value => return Err(invalid(format!("reserved font scheme {value}"))),
                }))
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

    fn data_bytes(&self) -> XlsResult<Vec<u8>> {
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
                    XlsXfFontWeight::Normal => 400u16,
                    XlsXfFontWeight::Bold => 700u16,
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
                XlsXfFontScheme::None => 0,
                XlsXfFontScheme::Major => 1,
                XlsXfFontScheme::Minor => 2,
                XlsXfFontScheme::NotSpecified => 0xFF,
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
pub struct XlsXfProperties {
    properties: Vec<XlsXfProperty>,
}

impl XlsXfProperties {
    pub fn try_new(properties: Vec<XlsXfProperty>) -> XlsResult<Self> {
        let value = Self { properties };
        value.validate()?;
        Ok(value)
    }

    pub fn properties(&self) -> &[XlsXfProperty] {
        &self.properties
    }

    pub fn is_empty(&self) -> bool {
        self.properties.is_empty()
    }

    pub(crate) fn parse(data: &[u8]) -> XlsResult<Self> {
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
            let size = usize::from(read_u16(data, offset + 2, "XFProp.cb")?);
            if size < 4 {
                return Err(invalid(format!(
                    "XFProp size {size} is smaller than its header"
                )));
            }
            let end = offset
                .checked_add(size)
                .ok_or_else(|| invalid("XFProp range overflows"))?;
            let blob = data
                .get(offset + 4..end)
                .ok_or_else(|| invalid("truncated XFProp data"))?;
            properties.push(XlsXfProperty::parse(property_type, blob)?);
            offset = end;
        }
        if offset != data.len() {
            return Err(invalid(
                "XFProps count does not consume its payload exactly",
            ));
        }
        Self::try_new(properties)
    }

    fn validate(&self) -> XlsResult<()> {
        if self.properties.len() > MAX_XF_PROPERTIES || self.properties.len() > u16::MAX as usize {
            return Err(invalid(format!(
                "XFProps count exceeds resource cap {MAX_XF_PROPERTIES}"
            )));
        }
        let has_pattern = self
            .properties
            .iter()
            .any(|property| matches!(property, XlsXfProperty::FillPattern(_)));
        let has_gradient = self.properties.iter().any(|property| {
            matches!(
                property,
                XlsXfProperty::Gradient(_) | XlsXfProperty::GradientStop(_)
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
                XlsXfProperty::Gradient(_) => preceding_gradient = true,
                XlsXfProperty::GradientStop(_) if !preceding_gradient => {
                    return Err(invalid("gradient stop has no preceding gradient property"));
                },
                XlsXfProperty::JustifyDistributed(true) => distributed = true,
                XlsXfProperty::HorizontalAlignment(Some(XlsHorizontalAlignment::Distributed)) => {
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

    pub(crate) fn to_bytes(&self) -> XlsResult<Vec<u8>> {
        self.validate()?;
        let mut data = Vec::new();
        data.extend_from_slice(&0u16.to_le_bytes());
        data.extend_from_slice(&(self.properties.len() as u16).to_le_bytes());
        for property in &self.properties {
            let blob = property.data_bytes()?;
            let size = u16::try_from(4usize + blob.len())
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
pub struct XlsDifferentialFormat {
    new_border: bool,
    properties: XlsXfProperties,
    unused_flags: u16,
}

impl XlsDifferentialFormat {
    pub fn try_new(new_border: bool, properties: Vec<XlsXfProperty>) -> XlsResult<Self> {
        let value = Self {
            new_border,
            properties: XlsXfProperties::try_new(properties)?,
            unused_flags: 0,
        };
        value.validate()?;
        Ok(value)
    }

    pub const fn has_new_border(&self) -> bool {
        self.new_border
    }
    pub fn properties(&self) -> &XlsXfProperties {
        &self.properties
    }

    pub fn parse_payload(data: &[u8]) -> XlsResult<Self> {
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
            properties: XlsXfProperties::parse(&data[FRT_HEADER_LEN + 2..])?,
        };
        value.validate()?;
        Ok(value)
    }

    pub fn to_payload(&self) -> XlsResult<Vec<u8>> {
        self.validate()?;
        let properties = self.properties.to_bytes()?;
        let size = FRT_HEADER_LEN + 2 + properties.len();
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

    pub fn to_record_bytes(&self) -> XlsResult<Vec<u8>> {
        let payload = self.to_payload()?;
        let mut data = Vec::with_capacity(4 + payload.len());
        data.extend_from_slice(&DXF_RECORD_TYPE.to_le_bytes());
        data.extend_from_slice(&(payload.len() as u16).to_le_bytes());
        data.extend_from_slice(&payload);
        Ok(data)
    }

    fn validate(&self) -> XlsResult<()> {
        self.properties.validate()?;
        if !self.new_border
            && self.properties.properties.iter().any(|property| {
                matches!(
                    property,
                    XlsXfProperty::VerticalBorder(_) | XlsXfProperty::HorizontalBorder(_)
                )
            })
        {
            return Err(invalid("internal border properties require fNewBorder"));
        }
        Ok(())
    }
}

pub(crate) fn validate_frt_header(data: &[u8], record_type: u16) -> XlsResult<()> {
    if data.len() < FRT_HEADER_LEN {
        return Err(XlsError::InvalidRecord {
            record_type,
            message: "truncated FrtHeader".to_string(),
        });
    }
    if u16::from_le_bytes([data[0], data[1]]) != record_type {
        return Err(XlsError::InvalidRecord {
            record_type,
            message: "future-record type does not match record header".to_string(),
        });
    }
    if data[2..12].iter().any(|byte| *byte != 0) {
        return Err(XlsError::InvalidRecord {
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

fn parse_bool(value: u8, property_type: u16) -> XlsResult<bool> {
    match value {
        0 => Ok(false),
        1 => Ok(true),
        _ => Err(invalid(format!(
            "XFProp 0x{property_type:04X} Boolean is {value}"
        ))),
    }
}

fn parse_fill_pattern(value: u8) -> XlsResult<XlsFillPattern> {
    Ok(match value {
        0 => XlsFillPattern::None,
        1 => XlsFillPattern::Solid,
        2 => XlsFillPattern::MediumGray,
        3 => XlsFillPattern::DarkGray,
        4 => XlsFillPattern::LightGray,
        5 => XlsFillPattern::DarkHorizontal,
        6 => XlsFillPattern::DarkVertical,
        7 => XlsFillPattern::DarkDown,
        8 => XlsFillPattern::DarkUp,
        9 => XlsFillPattern::DarkGrid,
        10 => XlsFillPattern::DarkTrellis,
        11 => XlsFillPattern::LightHorizontal,
        12 => XlsFillPattern::LightVertical,
        13 => XlsFillPattern::LightDown,
        14 => XlsFillPattern::LightUp,
        15 => XlsFillPattern::LightGrid,
        16 => XlsFillPattern::LightTrellis,
        17 => XlsFillPattern::Gray125,
        18 => XlsFillPattern::Gray0625,
        _ => return Err(invalid(format!("reserved fill pattern {value}"))),
    })
}

fn fill_pattern_byte(value: XlsFillPattern) -> u8 {
    match value {
        XlsFillPattern::None => 0,
        XlsFillPattern::Solid => 1,
        XlsFillPattern::MediumGray => 2,
        XlsFillPattern::DarkGray => 3,
        XlsFillPattern::LightGray => 4,
        XlsFillPattern::DarkHorizontal => 5,
        XlsFillPattern::DarkVertical => 6,
        XlsFillPattern::DarkDown => 7,
        XlsFillPattern::DarkUp => 8,
        XlsFillPattern::DarkGrid => 9,
        XlsFillPattern::DarkTrellis => 10,
        XlsFillPattern::LightHorizontal => 11,
        XlsFillPattern::LightVertical => 12,
        XlsFillPattern::LightDown => 13,
        XlsFillPattern::LightUp => 14,
        XlsFillPattern::LightGrid => 15,
        XlsFillPattern::LightTrellis => 16,
        XlsFillPattern::Gray125 => 17,
        XlsFillPattern::Gray0625 => 18,
    }
}

fn parse_border_style(value: u16) -> XlsResult<XlsBorderStyle> {
    Ok(match value {
        0 => XlsBorderStyle::None,
        1 => XlsBorderStyle::Thin,
        2 => XlsBorderStyle::Medium,
        3 => XlsBorderStyle::Dashed,
        4 => XlsBorderStyle::Dotted,
        5 => XlsBorderStyle::Thick,
        6 => XlsBorderStyle::Double,
        7 => XlsBorderStyle::Hair,
        8 => XlsBorderStyle::MediumDashed,
        9 => XlsBorderStyle::DashDot,
        10 => XlsBorderStyle::MediumDashDot,
        11 => XlsBorderStyle::DashDotDot,
        12 => XlsBorderStyle::MediumDashDotDot,
        13 => XlsBorderStyle::SlantedDashDot,
        _ => return Err(invalid(format!("reserved border style {value}"))),
    })
}

fn border_style_u16(value: XlsBorderStyle) -> u16 {
    match value {
        XlsBorderStyle::None => 0,
        XlsBorderStyle::Thin => 1,
        XlsBorderStyle::Medium => 2,
        XlsBorderStyle::Dashed => 3,
        XlsBorderStyle::Dotted => 4,
        XlsBorderStyle::Thick => 5,
        XlsBorderStyle::Double => 6,
        XlsBorderStyle::Hair => 7,
        XlsBorderStyle::MediumDashed => 8,
        XlsBorderStyle::DashDot => 9,
        XlsBorderStyle::MediumDashDot => 10,
        XlsBorderStyle::DashDotDot => 11,
        XlsBorderStyle::MediumDashDotDot => 12,
        XlsBorderStyle::SlantedDashDot => 13,
    }
}

fn parse_horizontal_alignment(value: u8) -> XlsResult<Option<XlsHorizontalAlignment>> {
    Ok(match value {
        0 => Some(XlsHorizontalAlignment::General),
        1 => Some(XlsHorizontalAlignment::Left),
        2 => Some(XlsHorizontalAlignment::Center),
        3 => Some(XlsHorizontalAlignment::Right),
        4 => Some(XlsHorizontalAlignment::Fill),
        5 => Some(XlsHorizontalAlignment::Justify),
        6 => Some(XlsHorizontalAlignment::CenterAcrossSelection),
        7 => Some(XlsHorizontalAlignment::Distributed),
        0xFF => None,
        _ => return Err(invalid(format!("reserved horizontal alignment {value}"))),
    })
}

fn horizontal_alignment_byte(value: Option<XlsHorizontalAlignment>) -> u8 {
    match value {
        Some(XlsHorizontalAlignment::General) => 0,
        Some(XlsHorizontalAlignment::Left) => 1,
        Some(XlsHorizontalAlignment::Center) => 2,
        Some(XlsHorizontalAlignment::Right) => 3,
        Some(XlsHorizontalAlignment::Fill) => 4,
        Some(XlsHorizontalAlignment::Justify) => 5,
        Some(XlsHorizontalAlignment::CenterAcrossSelection) => 6,
        Some(XlsHorizontalAlignment::Distributed) => 7,
        None => 0xFF,
    }
}

fn parse_vertical_alignment(value: u8) -> XlsResult<XlsVerticalAlignment> {
    Ok(match value {
        0 => XlsVerticalAlignment::Top,
        1 => XlsVerticalAlignment::Center,
        2 => XlsVerticalAlignment::Bottom,
        3 => XlsVerticalAlignment::Justify,
        4 => XlsVerticalAlignment::Distributed,
        _ => return Err(invalid(format!("reserved vertical alignment {value}"))),
    })
}

fn vertical_alignment_byte(value: XlsVerticalAlignment) -> u8 {
    match value {
        XlsVerticalAlignment::Top => 0,
        XlsVerticalAlignment::Center => 1,
        XlsVerticalAlignment::Bottom => 2,
        XlsVerticalAlignment::Justify => 3,
        XlsVerticalAlignment::Distributed => 4,
    }
}

fn parse_rotation(value: u8) -> XlsResult<XlsTextRotation> {
    match value {
        0 => Ok(XlsTextRotation::None),
        1..=90 => Ok(XlsTextRotation::CounterClockwise(value)),
        91..=180 => Ok(XlsTextRotation::Clockwise(value - 90)),
        255 => Ok(XlsTextRotation::Vertical),
        _ => Err(invalid(format!("reserved text rotation {value}"))),
    }
}

fn rotation_byte(value: XlsTextRotation) -> XlsResult<u8> {
    match value {
        XlsTextRotation::None => Ok(0),
        XlsTextRotation::CounterClockwise(value @ 1..=90) => Ok(value),
        XlsTextRotation::Clockwise(value @ 1..=90) => Ok(value + 90),
        XlsTextRotation::Vertical => Ok(255),
        _ => Err(invalid("text rotation degrees must be between 1 and 90")),
    }
}

fn parse_reading_order(value: u8) -> XlsResult<XlsReadingOrder> {
    match value {
        0 => Ok(XlsReadingOrder::Context),
        1 => Ok(XlsReadingOrder::LeftToRight),
        2 => Ok(XlsReadingOrder::RightToLeft),
        _ => Err(invalid(format!("reserved reading order {value}"))),
    }
}

fn reading_order_byte(value: XlsReadingOrder) -> u8 {
    match value {
        XlsReadingOrder::Context => 0,
        XlsReadingOrder::LeftToRight => 1,
        XlsReadingOrder::RightToLeft => 2,
    }
}

fn parse_underline(value: u16) -> XlsResult<XlsFontUnderline> {
    match value {
        0 => Ok(XlsFontUnderline::None),
        1 => Ok(XlsFontUnderline::Single),
        2 => Ok(XlsFontUnderline::Double),
        0x21 => Ok(XlsFontUnderline::SingleAccounting),
        0x22 => Ok(XlsFontUnderline::DoubleAccounting),
        _ => Err(invalid(format!("reserved underline style {value}"))),
    }
}

fn underline_u16(value: XlsFontUnderline) -> u16 {
    match value {
        XlsFontUnderline::None => 0,
        XlsFontUnderline::Single => 1,
        XlsFontUnderline::Double => 2,
        XlsFontUnderline::SingleAccounting => 0x21,
        XlsFontUnderline::DoubleAccounting => 0x22,
    }
}

fn parse_escapement(value: u16) -> XlsResult<XlsFontEscapement> {
    match value {
        0 => Ok(XlsFontEscapement::Normal),
        1 => Ok(XlsFontEscapement::Superscript),
        2 => Ok(XlsFontEscapement::Subscript),
        _ => Err(invalid(format!("reserved font escapement {value}"))),
    }
}

fn escapement_u16(value: XlsFontEscapement) -> u16 {
    match value {
        XlsFontEscapement::Normal => 0,
        XlsFontEscapement::Superscript => 1,
        XlsFontEscapement::Subscript => 2,
    }
}

fn parse_charset(value: u8) -> XlsResult<XlsFontCharset> {
    match value {
        0 => Ok(XlsFontCharset::Ansi),
        1 => Ok(XlsFontCharset::Default),
        2 => Ok(XlsFontCharset::Symbol),
        77 => Ok(XlsFontCharset::Mac),
        128 => Ok(XlsFontCharset::ShiftJis),
        129 => Ok(XlsFontCharset::Korean),
        130 => Ok(XlsFontCharset::Johab),
        134 => Ok(XlsFontCharset::Gb2312),
        136 => Ok(XlsFontCharset::ChineseBig5),
        161 => Ok(XlsFontCharset::Greek),
        162 => Ok(XlsFontCharset::Turkish),
        163 => Ok(XlsFontCharset::Vietnamese),
        177 => Ok(XlsFontCharset::Hebrew),
        178 => Ok(XlsFontCharset::Arabic),
        186 => Ok(XlsFontCharset::Baltic),
        204 => Ok(XlsFontCharset::Russian),
        222 => Ok(XlsFontCharset::Thai),
        238 => Ok(XlsFontCharset::EastEurope),
        255 => Ok(XlsFontCharset::Oem),
        _ => Err(invalid(format!(
            "unsupported LOGFONT character set {value}"
        ))),
    }
}

fn charset_byte(value: XlsFontCharset) -> u8 {
    match value {
        XlsFontCharset::Ansi => 0,
        XlsFontCharset::Default => 1,
        XlsFontCharset::Symbol => 2,
        XlsFontCharset::Mac => 77,
        XlsFontCharset::ShiftJis => 128,
        XlsFontCharset::Korean => 129,
        XlsFontCharset::Johab => 130,
        XlsFontCharset::Gb2312 => 134,
        XlsFontCharset::ChineseBig5 => 136,
        XlsFontCharset::Greek => 161,
        XlsFontCharset::Turkish => 162,
        XlsFontCharset::Vietnamese => 163,
        XlsFontCharset::Hebrew => 177,
        XlsFontCharset::Arabic => 178,
        XlsFontCharset::Baltic => 186,
        XlsFontCharset::Russian => 204,
        XlsFontCharset::Thai => 222,
        XlsFontCharset::EastEurope => 238,
        XlsFontCharset::Oem => 255,
    }
}

fn parse_font_family(value: u8) -> XlsResult<XlsFontFamily> {
    match value {
        0 => Ok(XlsFontFamily::NotApplicable),
        1 => Ok(XlsFontFamily::Roman),
        2 => Ok(XlsFontFamily::Swiss),
        3 => Ok(XlsFontFamily::Modern),
        4 => Ok(XlsFontFamily::Script),
        5 => Ok(XlsFontFamily::Decorative),
        _ => Err(invalid(format!("reserved font family {value}"))),
    }
}

fn font_family_byte(value: XlsFontFamily) -> u8 {
    match value {
        XlsFontFamily::NotApplicable => 0,
        XlsFontFamily::Roman => 1,
        XlsFontFamily::Swiss => 2,
        XlsFontFamily::Modern => 3,
        XlsFontFamily::Script => 4,
        XlsFontFamily::Decorative => 5,
    }
}

fn parse_lp_wide_string(data: &[u8]) -> XlsResult<String> {
    let count = usize::from(read_u16(data, 0, "LPWideString.cchCharacters")?);
    if count > 32 || data.len() != 2 + count * 2 {
        return Err(invalid(
            "font name LPWideString is malformed or exceeds 32 characters",
        ));
    }
    decode_utf16(&data[2..], "font name")
}

fn write_lp_wide_string(value: &str, data: &mut Vec<u8>) -> XlsResult<()> {
    let units = value.encode_utf16().collect::<Vec<_>>();
    if units.len() > 32 {
        return Err(invalid("font name exceeds 32 UTF-16 code units"));
    }
    data.extend_from_slice(&(units.len() as u16).to_le_bytes());
    data.extend(units.into_iter().flat_map(u16::to_le_bytes));
    Ok(())
}

fn parse_number_format_code(data: &[u8]) -> XlsResult<String> {
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
    decode_utf16(&data[2..], "number-format string")
}

fn write_number_format_code(value: &str, data: &mut Vec<u8>) -> XlsResult<()> {
    let units = value.encode_utf16().collect::<Vec<_>>();
    if !(1..=255).contains(&units.len()) {
        return Err(invalid("number-format string length must be 1..=255"));
    }
    data.extend_from_slice(&(units.len() as u16).to_le_bytes());
    data.extend(units.into_iter().flat_map(u16::to_le_bytes));
    Ok(())
}

fn decode_utf16(data: &[u8], field: &str) -> XlsResult<String> {
    if data.len() % 2 != 0 {
        return Err(invalid(format!("{field} has an odd byte length")));
    }
    let units = data
        .chunks_exact(2)
        .map(|chunk| u16::from_le_bytes([chunk[0], chunk[1]]));
    char::decode_utf16(units)
        .collect::<Result<String, _>>()
        .map_err(|_| invalid(format!("{field} contains invalid UTF-16")))
}

#[cfg(test)]
mod tests {
    use super::*;

    fn color() -> XlsXfColor {
        XlsXfColor::try_new(
            XlsXfColorSource::Theme(XlsThemeColor::Accent2),
            100,
            [1, 2, 3, 255],
        )
        .unwrap()
    }

    #[test]
    fn typed_dxf_round_trips_representative_property_families() {
        let dxf = XlsDifferentialFormat::try_new(
            true,
            vec![
                XlsXfProperty::Gradient(XlsXfGradient::linear(45.0).unwrap()),
                XlsXfProperty::GradientStop(XlsXfGradientStop::try_new(0.5, color()).unwrap()),
                XlsXfProperty::TopBorder(XlsXfBorder::new(color(), XlsBorderStyle::Thin)),
                XlsXfProperty::VerticalBorder(XlsXfBorder::new(color(), XlsBorderStyle::Dashed)),
                XlsXfProperty::HorizontalAlignment(Some(XlsHorizontalAlignment::Distributed)),
                XlsXfProperty::JustifyDistributed(true),
                XlsXfProperty::TextRotation(XlsTextRotation::Clockwise(30)),
                XlsXfProperty::FontName("Aptos".to_string()),
                XlsXfProperty::FontWeight(XlsXfFontWeight::Bold),
                XlsXfProperty::FontUnderline(XlsFontUnderline::Double),
                XlsXfProperty::FontSizeTwips(220),
                XlsXfProperty::NumberFormatCode("0.00".to_string()),
                XlsXfProperty::RelativeIndent(Some(-2)),
                XlsXfProperty::Locked(true),
            ],
        )
        .unwrap();
        let payload = dxf.to_payload().unwrap();
        assert_eq!(XlsDifferentialFormat::parse_payload(&payload).unwrap(), dxf);
        let record = dxf.to_record_bytes().unwrap();
        assert_eq!(&record[..2], &[0x8D, 0x08]);
    }

    #[test]
    fn xfprop_color_uses_low_flag_bit_and_high_seven_type_bits() {
        // Apache POI producer forms: bit 7 is clear, while fValidRGBA in bit 0 is set.
        let rgb = [0x05, 0xFF, 0x00, 0x00, 0xFF, 0xC7, 0xCE, 0xFF];
        let parsed = XlsXfColor::parse(&rgb).unwrap();
        assert_eq!(parsed.source(), XlsXfColorSource::Rgb);
        let mut encoded = Vec::new();
        parsed.write_to(&mut encoded);
        assert_eq!(encoded, rgb);

        let theme = [0x07, 0x04, 0x65, 0x66, 0xDC, 0xE6, 0xF1, 0xFF];
        let parsed = XlsXfColor::parse(&theme).unwrap();
        assert_eq!(
            parsed.source(),
            XlsXfColorSource::Theme(XlsThemeColor::Accent1)
        );
        let mut encoded = Vec::new();
        parsed.write_to(&mut encoded);
        assert_eq!(encoded, theme);

        let indexed = [0x03, 0x40, 0, 0, 1, 2, 3, 4];
        assert_eq!(
            XlsXfColor::parse(&indexed).unwrap().source(),
            XlsXfColorSource::Indexed(0x40)
        );
        let automatic = [0x01, 0xAA, 0, 0, 1, 2, 3, 4];
        assert_eq!(
            XlsXfColor::parse(&automatic).unwrap().source(),
            XlsXfColorSource::Automatic
        );
        let not_set = [0x09, 0xAA, 0, 0, 1, 2, 3, 4];
        assert_eq!(
            XlsXfColor::parse(&not_set).unwrap().source(),
            XlsXfColorSource::NotSet
        );
    }

    #[test]
    fn xfprop_color_rejects_clear_valid_flag_and_invalid_type_data() {
        assert!(XlsXfColor::parse(&[0x04, 0, 0, 0, 0, 0, 0, 0]).is_err());
        assert!(XlsXfColor::parse(&[0x0B, 0, 0, 0, 0, 0, 0, 0]).is_err());
        assert!(XlsXfColor::parse(&[0x03, 66, 0, 0, 0, 0, 0, 0]).is_err());
        assert!(XlsXfColor::parse(&[0x07, 12, 0, 0, 0, 0, 0, 0]).is_err());
        assert!(XlsXfColor::parse(&[0x05, 0, 0x00, 0x80, 0, 0, 0, 0]).is_err());
    }

    #[test]
    fn rejects_hostile_headers_sizes_flags_and_property_relationships() {
        let empty = XlsDifferentialFormat::try_new(false, vec![])
            .unwrap()
            .to_payload()
            .unwrap();
        assert!(XlsDifferentialFormat::parse_payload(&empty[..17]).is_err());
        let mut bad = empty.clone();
        bad[0] = 0;
        assert!(XlsDifferentialFormat::parse_payload(&bad).is_err());
        let mut bad = empty.clone();
        bad[14] = 1;
        assert!(XlsDifferentialFormat::parse_payload(&bad).is_err());
        let mut bad = empty;
        bad[16..18].copy_from_slice(&1u16.to_le_bytes());
        assert!(XlsDifferentialFormat::parse_payload(&bad).is_err());

        assert!(
            XlsDifferentialFormat::try_new(
                false,
                vec![XlsXfProperty::VerticalBorder(XlsXfBorder::new(
                    color(),
                    XlsBorderStyle::Thin,
                ))],
            )
            .is_err()
        );
        assert!(
            XlsDifferentialFormat::try_new(
                false,
                vec![
                    XlsXfProperty::FillPattern(XlsFillPattern::Solid),
                    XlsXfProperty::Gradient(XlsXfGradient::linear(0.0).unwrap()),
                ],
            )
            .is_err()
        );
        assert!(
            XlsDifferentialFormat::try_new(false, vec![XlsXfProperty::JustifyDistributed(true)],)
                .is_err()
        );
    }

    #[test]
    fn enforces_resource_caps() {
        assert!(
            XlsXfProperties::try_new(vec![XlsXfProperty::WrapText(false); MAX_XF_PROPERTIES + 1])
                .is_err()
        );
        let huge = "x".repeat(256);
        assert!(
            XlsDifferentialFormat::try_new(false, vec![XlsXfProperty::NumberFormatCode(huge)])
                .is_err()
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
