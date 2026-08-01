//! BIFF8 `XFExt` record (MS-XLS 2.4.355): formatting property extensions
//! (`ExtProp`, MS-XLS 2.5.108) attached to an `XF` record.
//!
//! Extensions cover theme-aware `FullColorExt` colors (MS-XLS 2.5.155),
//! gradient fills (`XFExtGradient`, MS-XLS 2.5.280), the theme-font scheme,
//! and the text indentation level. Extensions of a type or size the
//! specification does not define are preserved verbatim.

use super::differential_format::{XlsXfFontScheme, XlsXfGradient, XlsXfGradientStop};
use super::{XlsError, XlsResult};

/// Record type of the `XFExt` record.
pub(crate) const XF_EXT_RECORD_TYPE: u16 = 0x087D;

/// Size in bytes of an `FrtHeader` (MS-XLS 2.5.135).
const FRT_HEADER_LEN: usize = 12;
/// Size in bytes of the fixed portion following the `FrtHeader`:
/// reserved1 + ixfe + reserved2 + cexts.
const FIXED_TAIL_LEN: usize = 8;
/// Largest legal `ixfe` value (MS-XLS 2.4.355).
const MAX_XF_INDEX: u16 = 4050;
/// Resource cap on the number of `ExtProp` entries in one record.
const MAX_EXT_PROPS: usize = 256;
/// Size in bytes of a `FullColorExt` (MS-XLS 2.5.155).
const FULL_COLOR_EXT_LEN: usize = 16;
/// Size in bytes of the `XFPropGradient` inside an `XFExtGradient`.
const GRADIENT_LEN: usize = 44;
/// Size in bytes of one `GradStop` inside an `XFExtGradient`.
const GRAD_STOP_LEN: usize = 18;
/// Largest legal gradient-stop count (MS-XLS 2.5.280).
const MAX_GRADIENT_STOPS: usize = 256;
/// Largest legal text indentation level (MS-XLS 2.5.108).
const MAX_INDENT: u16 = 250;

// ExtProp extType values (MS-XLS 2.5.108).
const EXT_FILL_FOREGROUND: u16 = 0x0004;
const EXT_FILL_BACKGROUND: u16 = 0x0005;
const EXT_FILL_GRADIENT: u16 = 0x0006;
const EXT_BORDER_TOP: u16 = 0x0007;
const EXT_BORDER_BOTTOM: u16 = 0x0008;
const EXT_BORDER_LEFT: u16 = 0x0009;
const EXT_BORDER_RIGHT: u16 = 0x000A;
const EXT_BORDER_DIAGONAL: u16 = 0x000B;
const EXT_TEXT_COLOR: u16 = 0x000D;
const EXT_FONT_SCHEME: u16 = 0x000E;
const EXT_INDENT: u16 = 0x000F;

fn invalid(message: impl Into<String>) -> XlsError {
    XlsError::InvalidRecord {
        record_type: XF_EXT_RECORD_TYPE,
        message: message.into(),
    }
}

/// How the color data of a `FullColorExt` is stored (`XColorType`).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum XlsFullColorType {
    /// Automatic color; `value` is zero.
    Automatic,
    /// Color-table index (`IcvXF`) in the low bytes of `value`.
    Indexed,
    /// `LongRGBA` in `value`.
    Rgb,
    /// `ColorTheme` in `value`.
    Theme,
    /// Color not set; `value` is zero.
    NotSet,
}

impl XlsFullColorType {
    fn from_code(value: u16) -> XlsResult<Self> {
        Ok(match value {
            0x0000 => Self::Automatic,
            0x0001 => Self::Indexed,
            0x0002 => Self::Rgb,
            0x0003 => Self::Theme,
            0x0004 => Self::NotSet,
            value => return Err(invalid(format!("reserved XColorType {value}"))),
        })
    }

    fn code(self) -> u16 {
        match self {
            Self::Automatic => 0x0000,
            Self::Indexed => 0x0001,
            Self::Rgb => 0x0002,
            Self::Theme => 0x0003,
            Self::NotSet => 0x0004,
        }
    }
}

/// A `FullColorExt` (MS-XLS 2.5.155): a theme-aware extended color.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct XlsFullColorExt {
    color_type: XlsFullColorType,
    /// Tint; positive lightens, negative darkens.
    tint: i16,
    /// Raw `xclrValue` color data (index, RGBA, or theme selector).
    value: u32,
    /// Trailing undefined bytes, preserved verbatim.
    unused: [u8; 8],
}

impl XlsFullColorExt {
    /// A color; `Automatic` and `NotSet` colors must carry a zero value.
    pub fn try_new(color_type: XlsFullColorType, tint: i16, value: u32) -> XlsResult<Self> {
        if matches!(
            color_type,
            XlsFullColorType::Automatic | XlsFullColorType::NotSet
        ) && value != 0
        {
            return Err(invalid(
                "automatic and not-set colors must carry a zero value",
            ));
        }
        Ok(Self {
            color_type,
            tint,
            value,
            unused: [0; 8],
        })
    }

    pub const fn color_type(&self) -> XlsFullColorType {
        self.color_type
    }
    pub const fn tint(&self) -> i16 {
        self.tint
    }
    pub const fn value(&self) -> u32 {
        self.value
    }

    fn parse(data: &[u8]) -> XlsResult<Self> {
        if data.len() != FULL_COLOR_EXT_LEN {
            return Err(invalid(format!(
                "FullColorExt has {} bytes; expected {FULL_COLOR_EXT_LEN}",
                data.len()
            )));
        }
        let value = Self {
            color_type: XlsFullColorType::from_code(u16::from_le_bytes([data[0], data[1]]))?,
            tint: i16::from_le_bytes([data[2], data[3]]),
            value: u32::from_le_bytes(data[4..8].try_into().expect("length checked")),
            unused: data[8..16].try_into().expect("length checked"),
        };
        if matches!(
            value.color_type,
            XlsFullColorType::Automatic | XlsFullColorType::NotSet
        ) && value.value != 0
        {
            return Err(invalid(
                "automatic and not-set colors must carry a zero value",
            ));
        }
        Ok(value)
    }

    fn write_to(&self, output: &mut Vec<u8>) {
        output.extend_from_slice(&self.color_type.code().to_le_bytes());
        output.extend_from_slice(&self.tint.to_le_bytes());
        output.extend_from_slice(&self.value.to_le_bytes());
        output.extend_from_slice(&self.unused);
    }
}

/// One typed entry in an `XFExt` `rgExt` array (MS-XLS 2.5.108).
#[derive(Debug, Clone, PartialEq)]
pub enum XlsExtProp {
    /// Cell interior foreground color (0x0004).
    FillForegroundColor(XlsFullColorExt),
    /// Cell interior background color (0x0005).
    FillBackgroundColor(XlsFullColorExt),
    /// Cell interior gradient fill (0x0006).
    FillGradient {
        /// Gradient parameters.
        gradient: XlsXfGradient,
        /// Gradient color stops.
        stops: Vec<XlsXfGradientStop>,
    },
    /// Top cell border color (0x0007).
    BorderTopColor(XlsFullColorExt),
    /// Bottom cell border color (0x0008).
    BorderBottomColor(XlsFullColorExt),
    /// Left cell border color (0x0009).
    BorderLeftColor(XlsFullColorExt),
    /// Right cell border color (0x000A).
    BorderRightColor(XlsFullColorExt),
    /// Diagonal cell border color (0x000B).
    BorderDiagonalColor(XlsFullColorExt),
    /// Cell text color (0x000D).
    TextColor(XlsFullColorExt),
    /// Theme-font scheme (0x000E).
    FontScheme(XlsXfFontScheme),
    /// Text indentation level (0x000F), at most 250.
    Indent(u16),
    /// An extension the specification does not define, or a known extension
    /// with a malformed payload; the bytes are preserved verbatim.
    Unknown {
        /// Raw `extType`.
        ext_type: u16,
        /// Raw `extPropData`.
        data: Vec<u8>,
    },
}

impl XlsExtProp {
    fn ext_type(&self) -> u16 {
        match self {
            Self::FillForegroundColor(_) => EXT_FILL_FOREGROUND,
            Self::FillBackgroundColor(_) => EXT_FILL_BACKGROUND,
            Self::FillGradient { .. } => EXT_FILL_GRADIENT,
            Self::BorderTopColor(_) => EXT_BORDER_TOP,
            Self::BorderBottomColor(_) => EXT_BORDER_BOTTOM,
            Self::BorderLeftColor(_) => EXT_BORDER_LEFT,
            Self::BorderRightColor(_) => EXT_BORDER_RIGHT,
            Self::BorderDiagonalColor(_) => EXT_BORDER_DIAGONAL,
            Self::TextColor(_) => EXT_TEXT_COLOR,
            Self::FontScheme(_) => EXT_FONT_SCHEME,
            Self::Indent(_) => EXT_INDENT,
            Self::Unknown { ext_type, .. } => *ext_type,
        }
    }

    fn parse(ext_type: u16, data: &[u8]) -> Self {
        let unknown = || Self::Unknown {
            ext_type,
            data: data.to_vec(),
        };
        let color = |build: fn(XlsFullColorExt) -> Self| {
            XlsFullColorExt::parse(data).map_or_else(|_| unknown(), build)
        };
        match ext_type {
            EXT_FILL_FOREGROUND => color(Self::FillForegroundColor),
            EXT_FILL_BACKGROUND => color(Self::FillBackgroundColor),
            EXT_BORDER_TOP => color(Self::BorderTopColor),
            EXT_BORDER_BOTTOM => color(Self::BorderBottomColor),
            EXT_BORDER_LEFT => color(Self::BorderLeftColor),
            EXT_BORDER_RIGHT => color(Self::BorderRightColor),
            EXT_BORDER_DIAGONAL => color(Self::BorderDiagonalColor),
            EXT_TEXT_COLOR => color(Self::TextColor),
            EXT_FILL_GRADIENT => Self::parse_gradient(data).unwrap_or_else(|_| unknown()),
            EXT_FONT_SCHEME => {
                if data.len() == 2 {
                    match u16::from_le_bytes([data[0], data[1]]) {
                        0 => Self::FontScheme(XlsXfFontScheme::None),
                        1 => Self::FontScheme(XlsXfFontScheme::Major),
                        2 => Self::FontScheme(XlsXfFontScheme::Minor),
                        _ => unknown(),
                    }
                } else {
                    unknown()
                }
            },
            EXT_INDENT => {
                if data.len() == 2 {
                    let indent = u16::from_le_bytes([data[0], data[1]]);
                    if indent <= MAX_INDENT {
                        Self::Indent(indent)
                    } else {
                        unknown()
                    }
                } else {
                    unknown()
                }
            },
            _ => unknown(),
        }
    }

    fn parse_gradient(data: &[u8]) -> XlsResult<Self> {
        if data.len() < GRADIENT_LEN + 4 {
            return Err(invalid("XFExtGradient is truncated"));
        }
        let gradient = XlsXfGradient::parse(&data[..GRADIENT_LEN])?;
        let stop_count = u32::from_le_bytes(
            data[GRADIENT_LEN..GRADIENT_LEN + 4]
                .try_into()
                .expect("sliced"),
        ) as usize;
        if stop_count > MAX_GRADIENT_STOPS {
            return Err(invalid("XFExtGradient stop count exceeds 256"));
        }
        let expected = GRADIENT_LEN + 4 + stop_count * GRAD_STOP_LEN;
        if data.len() != expected {
            return Err(invalid(
                "XFExtGradient stop count does not match its payload",
            ));
        }
        let mut stops = Vec::with_capacity(stop_count);
        for chunk in data[GRADIENT_LEN + 4..].chunks_exact(GRAD_STOP_LEN) {
            stops.push(XlsXfGradientStop::parse(chunk)?);
        }
        Ok(Self::FillGradient { gradient, stops })
    }

    fn data_bytes(&self) -> XlsResult<Vec<u8>> {
        let mut data = Vec::new();
        match self {
            Self::FillForegroundColor(value)
            | Self::FillBackgroundColor(value)
            | Self::BorderTopColor(value)
            | Self::BorderBottomColor(value)
            | Self::BorderLeftColor(value)
            | Self::BorderRightColor(value)
            | Self::BorderDiagonalColor(value)
            | Self::TextColor(value) => value.write_to(&mut data),
            Self::FillGradient { gradient, stops } => {
                if stops.len() > MAX_GRADIENT_STOPS {
                    return Err(invalid("XFExtGradient stop count exceeds 256"));
                }
                gradient.write_to(&mut data);
                data.extend_from_slice(&(stops.len() as u32).to_le_bytes());
                for stop in stops {
                    stop.write_to(&mut data);
                }
            },
            Self::FontScheme(value) => {
                let code = match value {
                    XlsXfFontScheme::None => 0u16,
                    XlsXfFontScheme::Major => 1,
                    XlsXfFontScheme::Minor => 2,
                    XlsXfFontScheme::NotSpecified => {
                        return Err(invalid("ExtProp FontScheme has no not-specified value"));
                    },
                };
                data.extend_from_slice(&code.to_le_bytes());
            },
            Self::Indent(value) => {
                if *value > MAX_INDENT {
                    return Err(invalid(format!("indentation level {value} exceeds 250")));
                }
                data.extend_from_slice(&value.to_le_bytes());
            },
            Self::Unknown { data: raw, .. } => data.extend_from_slice(raw),
        }
        Ok(data)
    }
}

/// Typed `XFExt` record content (MS-XLS 2.4.355).
#[derive(Debug, Clone, PartialEq)]
pub struct XlsXfExt {
    xf_index: u16,
    properties: Vec<XlsExtProp>,
}

impl XlsXfExt {
    /// Extensions for the `XF` record at `xf_index` (at most 4050).
    pub fn try_new(xf_index: u16, properties: Vec<XlsExtProp>) -> XlsResult<Self> {
        if xf_index > MAX_XF_INDEX {
            return Err(invalid(format!("XFExt index {xf_index} exceeds 4050")));
        }
        if properties.len() > MAX_EXT_PROPS {
            return Err(invalid("XFExt property count exceeds the resource cap"));
        }
        Ok(Self {
            xf_index,
            properties,
        })
    }

    /// Index of the `XF` record these properties extend.
    pub const fn xf_index(&self) -> u16 {
        self.xf_index
    }

    /// The formatting property extensions, in record order.
    pub fn properties(&self) -> &[XlsExtProp] {
        &self.properties
    }

    /// Parse an `XFExt` record payload.
    pub(crate) fn parse(data: &[u8]) -> XlsResult<Self> {
        if data.len() < FRT_HEADER_LEN + FIXED_TAIL_LEN {
            return Err(XlsError::InvalidLength {
                expected: FRT_HEADER_LEN + FIXED_TAIL_LEN,
                found: data.len(),
            });
        }
        if u16::from_le_bytes([data[0], data[1]]) != XF_EXT_RECORD_TYPE {
            return Err(invalid("XFExt FrtHeader.rt mismatch"));
        }
        let xf_index = u16::from_le_bytes([data[14], data[15]]);
        if xf_index > MAX_XF_INDEX {
            return Err(invalid(format!("XFExt index {xf_index} exceeds 4050")));
        }
        let count = usize::from(u16::from_le_bytes([data[18], data[19]]));
        if count > MAX_EXT_PROPS {
            return Err(invalid("XFExt property count exceeds the resource cap"));
        }
        let mut offset = FRT_HEADER_LEN + FIXED_TAIL_LEN;
        let mut properties = Vec::with_capacity(count);
        for _ in 0..count {
            if offset + 4 > data.len() {
                return Err(invalid("truncated ExtProp header"));
            }
            let ext_type = u16::from_le_bytes([data[offset], data[offset + 1]]);
            let size = usize::from(u16::from_le_bytes([data[offset + 2], data[offset + 3]]));
            if size < 4 {
                return Err(invalid("ExtProp size is smaller than its header"));
            }
            let end = offset
                .checked_add(size)
                .ok_or_else(|| invalid("ExtProp overflow"))?;
            let blob = data
                .get(offset + 4..end)
                .ok_or_else(|| invalid("truncated ExtProp data"))?;
            properties.push(XlsExtProp::parse(ext_type, blob));
            offset = end;
        }
        if offset != data.len() {
            return Err(invalid(
                "XFExt property count does not consume its payload exactly",
            ));
        }
        Ok(Self {
            xf_index,
            properties,
        })
    }

    /// Serialize back to a complete `XFExt` record payload.
    pub(crate) fn to_payload(&self) -> XlsResult<Vec<u8>> {
        if self.properties.len() > u16::MAX as usize {
            return Err(invalid("XFExt property count exceeds u16"));
        }
        let mut payload = Vec::new();
        payload.extend_from_slice(&XF_EXT_RECORD_TYPE.to_le_bytes());
        payload.extend_from_slice(&[0; FRT_HEADER_LEN - 2]);
        payload.extend_from_slice(&[0; 2]); // reserved1
        payload.extend_from_slice(&self.xf_index.to_le_bytes());
        payload.extend_from_slice(&[0; 2]); // reserved2
        payload.extend_from_slice(&(self.properties.len() as u16).to_le_bytes());
        for property in &self.properties {
            let blob = property.data_bytes()?;
            let size =
                u16::try_from(4 + blob.len()).map_err(|_| invalid("ExtProp size exceeds u16"))?;
            payload.extend_from_slice(&property.ext_type().to_le_bytes());
            payload.extend_from_slice(&size.to_le_bytes());
            payload.extend_from_slice(&blob);
        }
        Ok(payload)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn color_property(ext_type: u16) -> Vec<u8> {
        let mut data = Vec::new();
        data.extend_from_slice(&ext_type.to_le_bytes());
        data.extend_from_slice(&20u16.to_le_bytes()); // 4 header + 16 data
        data.extend_from_slice(&3u16.to_le_bytes()); // theme
        data.extend_from_slice(&(-2i16).to_le_bytes()); // tint
        data.extend_from_slice(&1u32.to_le_bytes()); // theme slot 1
        data.extend_from_slice(&[0xEE; 8]); // unused, preserved
        data
    }

    fn record(xf_index: u16, properties: &[Vec<u8>]) -> Vec<u8> {
        let mut data = Vec::new();
        data.extend_from_slice(&XF_EXT_RECORD_TYPE.to_le_bytes());
        data.extend_from_slice(&[0; 12]);
        data.extend_from_slice(&xf_index.to_le_bytes());
        data.extend_from_slice(&[0; 2]);
        data.extend_from_slice(&(properties.len() as u16).to_le_bytes());
        for property in properties {
            data.extend_from_slice(property);
        }
        data
    }

    #[test]
    fn parses_color_and_scalar_properties() {
        let mut indent = Vec::new();
        indent.extend_from_slice(&EXT_INDENT.to_le_bytes());
        indent.extend_from_slice(&6u16.to_le_bytes());
        indent.extend_from_slice(&12u16.to_le_bytes());
        let data = record(21, &[color_property(EXT_FILL_FOREGROUND), indent]);
        let parsed = XlsXfExt::parse(&data).unwrap();
        assert_eq!(parsed.xf_index(), 21);
        match parsed.properties() {
            [
                XlsExtProp::FillForegroundColor(color),
                XlsExtProp::Indent(12),
            ] => {
                assert_eq!(color.color_type(), XlsFullColorType::Theme);
                assert_eq!(color.tint(), -2);
                assert_eq!(color.value(), 1);
            },
            other => panic!("unexpected properties: {other:?}"),
        }
        assert_eq!(parsed.to_payload().unwrap(), data);
    }

    #[test]
    fn parses_gradient_fill() {
        let gradient = XlsXfGradient::linear(90.0).unwrap();
        let color = super::super::differential_format::XlsXfColor::try_new(
            super::super::differential_format::XlsXfColorSource::Rgb,
            0,
            [0xFF, 0, 0, 0xFF],
        )
        .unwrap();
        let stop = XlsXfGradientStop::try_new(0.5, color).unwrap();
        let property = XlsExtProp::FillGradient {
            gradient,
            stops: vec![stop],
        };
        let mut blob = property.data_bytes().unwrap();
        let mut entry = Vec::new();
        entry.extend_from_slice(&EXT_FILL_GRADIENT.to_le_bytes());
        entry.extend_from_slice(&((4 + blob.len()) as u16).to_le_bytes());
        entry.append(&mut blob);
        let data = record(0, &[entry]);
        let parsed = XlsXfExt::parse(&data).unwrap();
        assert_eq!(parsed.properties(), &[property]);
        assert_eq!(parsed.to_payload().unwrap(), data);
    }

    #[test]
    fn preserves_unknown_and_malformed_extensions() {
        // Unknown extType.
        let mut unknown = Vec::new();
        unknown.extend_from_slice(&0x00F0u16.to_le_bytes());
        unknown.extend_from_slice(&7u16.to_le_bytes());
        unknown.extend_from_slice(&[1, 2, 3]);
        // Malformed known extType (1-byte FontScheme).
        let mut malformed = Vec::new();
        malformed.extend_from_slice(&EXT_FONT_SCHEME.to_le_bytes());
        malformed.extend_from_slice(&5u16.to_le_bytes());
        malformed.push(0x02);
        let data = record(5, &[unknown, malformed]);
        let parsed = XlsXfExt::parse(&data).unwrap();
        assert!(matches!(
            parsed.properties(),
            [
                XlsExtProp::Unknown {
                    ext_type: 0x00F0,
                    ..
                },
                XlsExtProp::Unknown {
                    ext_type: EXT_FONT_SCHEME,
                    ..
                }
            ]
        ));
        assert_eq!(parsed.to_payload().unwrap(), data);
    }

    #[test]
    fn rejects_malformed_records() {
        // Truncated.
        assert!(XlsXfExt::parse(&[0; 10]).is_err());
        // Wrong FrtHeader.rt.
        let mut wrong_rt = record(0, &[]);
        wrong_rt[0..2].copy_from_slice(&0x087Cu16.to_le_bytes());
        assert!(XlsXfExt::parse(&wrong_rt).is_err());
        // Index above 4050.
        assert!(XlsXfExt::parse(&record(4051, &[])).is_err());
        // Declared count not consuming the payload.
        let mut trailing = record(0, &[]);
        trailing[18] = 1;
        assert!(XlsXfExt::parse(&trailing).is_err());
        // Constructor validation.
        assert!(XlsXfExt::try_new(4051, Vec::new()).is_err());
        assert!(XlsFullColorExt::try_new(XlsFullColorType::Automatic, 0, 7).is_err());
    }
}
