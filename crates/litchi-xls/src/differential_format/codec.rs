//! BIFF8 wire codec for DXF and `XFProp` payloads.
//!
//! This module owns all byte-level parsing and emission. The parent module
//! remains the semantic facade exposed to the rest of the XLS crate.

use super::validation::validate_unit_interval;
use super::{
    BorderStyle, DXF_RECORD_TYPE, DifferentialFormat, Error, FIXED_PAYLOAD_LEN, FRT_HEADER_LEN,
    FillPattern, FontCharset, FontEscapement, FontFamily, FontUnderline, HorizontalAlignment,
    MAX_BIFF8_PAYLOAD_LEN, MAX_XF_PROPERTIES, ReadingOrder, Result, TextRotation, ThemeColor,
    VerticalAlignment, XfBorder, XfColor, XfColorSource, XfFontScheme, XfFontWeight, XfGradient,
    XfGradientStop, XfProperties, XfProperty, invalid, read_bytes, read_f64, read_i16, read_u8,
    read_u16, read_u32,
};

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

    pub(super) const fn to_byte(self) -> u8 {
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

impl XfColor {
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

impl XfGradient {
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
        output.extend_from_slice(&u32::from(self.rectangular).to_le_bytes());
        output.extend_from_slice(&self.degree.to_le_bytes());
        output.extend_from_slice(&self.fill_to_left.to_le_bytes());
        output.extend_from_slice(&self.fill_to_right.to_le_bytes());
        output.extend_from_slice(&self.fill_to_top.to_le_bytes());
        output.extend_from_slice(&self.fill_to_bottom.to_le_bytes());
    }
}

impl XfGradientStop {
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

    pub(crate) fn data_bytes(&self) -> Result<Vec<u8>> {
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
                data.extend_from_slice(&underline_u16(*value).to_le_bytes());
            },
            Self::FontEscapement(value) => {
                data.extend_from_slice(&escapement_u16(*value).to_le_bytes());
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

impl XfProperties {
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

    pub(crate) fn to_bytes(&self) -> Result<Vec<u8>> {
        super::validation::validate_properties(self)?;
        let mut data = Vec::new();
        data.extend_from_slice(&0u16.to_le_bytes());
        let property_count = u16::try_from(self.properties.len())
            .map_err(|_error| invalid("XFProps count exceeds BIFF u16"))?;
        data.extend_from_slice(&property_count.to_le_bytes());
        for property in &self.properties {
            let blob = property.data_bytes()?;
            let property_size = 4usize
                .checked_add(blob.len())
                .ok_or_else(|| invalid("XFProp size overflows"))?;
            let size = u16::try_from(property_size)
                .map_err(|_error| invalid("XFProp size exceeds BIFF u16"))?;
            data.extend_from_slice(&property.property_type().to_le_bytes());
            data.extend_from_slice(&size.to_le_bytes());
            data.extend_from_slice(&blob);
        }
        Ok(data)
    }
}

impl DifferentialFormat {
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
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
        super::validation::validate_format(&value)?;
        Ok(value)
    }

    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn to_payload(&self) -> Result<Vec<u8>> {
        super::validation::validate_format(self)?;
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

    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn to_record_bytes(&self) -> Result<Vec<u8>> {
        let payload = self.to_payload()?;
        let record_len = 4usize
            .checked_add(payload.len())
            .ok_or_else(|| invalid("DXF record size overflows"))?;
        let payload_len = u16::try_from(payload.len())
            .map_err(|_error| invalid("DXF payload exceeds BIFF u16"))?;
        let mut data = Vec::with_capacity(record_len);
        data.extend_from_slice(&DXF_RECORD_TYPE.to_le_bytes());
        data.extend_from_slice(&payload_len.to_le_bytes());
        data.extend_from_slice(&payload);
        Ok(data)
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
    let count = u16::try_from(units.len())
        .map_err(|_error| invalid("font name length exceeds BIFF u16"))?;
    data.extend_from_slice(&count.to_le_bytes());
    data.extend(units.into_iter().flat_map(u16::to_le_bytes));
    Ok(())
}

pub(crate) fn parse_number_format_code(data: &[u8]) -> Result<String> {
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

pub(crate) fn write_number_format_code(value: &str, data: &mut Vec<u8>) -> Result<()> {
    let units = value.encode_utf16().collect::<Vec<_>>();
    if !(1..=255).contains(&units.len()) {
        return Err(invalid("number-format string length must be 1..=255"));
    }
    let count = u16::try_from(units.len())
        .map_err(|_error| invalid("number-format string length exceeds BIFF u16"))?;
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
            .map_err(|_error| invalid(format!("{field} has an invalid UTF-16 unit")))?;
        units.push(u16::from_le_bytes(bytes));
    }
    char::decode_utf16(units)
        .collect::<Result<String, _>>()
        .map_err(|_error| invalid(format!("{field} contains invalid UTF-16")))
}
