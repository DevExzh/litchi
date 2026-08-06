//! Semantic DrawingML color values.

use std::fmt;

use crate::Result;

/// A checked six-digit sRGB value.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[repr(transparent)]
#[must_use]
pub struct Rgb([u8; 3]);

impl Rgb {
    /// Construct an sRGB value from its red, green, and blue channels.
    #[inline]
    pub const fn new(red: u8, green: u8, blue: u8) -> Self {
        Self([red, green, blue])
    }

    /// Return the three channels in red/green/blue order.
    #[inline]
    pub const fn channels(self) -> [u8; 3] {
        self.0
    }

    /// Return the red channel.
    #[inline]
    pub const fn red(self) -> u8 {
        self.0[0]
    }

    /// Return the green channel.
    #[inline]
    pub const fn green(self) -> u8 {
        self.0[1]
    }

    /// Return the blue channel.
    #[inline]
    pub const fn blue(self) -> u8 {
        self.0[2]
    }

    /// Parse the exact six-digit hexadecimal sRGB lexical form.
    pub fn parse(value: &str) -> Result<Self> {
        if value.len() != 6 || !value.bytes().all(|byte| byte.is_ascii_hexdigit()) {
            return Err(crate::Error::Invalid(
                "DrawingML sRGB colors must contain six hexadecimal digits".into(),
            ));
        }
        let bytes = value.as_bytes();
        Ok(Self::new(
            hex_pair(bytes[0], bytes[1]),
            hex_pair(bytes[2], bytes[3]),
            hex_pair(bytes[4], bytes[5]),
        ))
    }

    /// Return the canonical uppercase hexadecimal lexical form.
    #[inline]
    pub fn to_hex(self) -> String {
        let mut value = String::with_capacity(6);
        write_hex(&mut value, self);
        value
    }
}

impl From<[u8; 3]> for Rgb {
    #[inline]
    fn from(value: [u8; 3]) -> Self {
        Self(value)
    }
}

impl From<Rgb> for [u8; 3] {
    #[inline]
    fn from(value: Rgb) -> Self {
        value.0
    }
}

impl fmt::Display for Rgb {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        for channel in self.0 {
            write!(formatter, "{channel:02X}")?;
        }
        Ok(())
    }
}

/// A checked `ST_SchemeColorVal` token.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[repr(u8)]
#[must_use]
pub enum Scheme {
    Background,
    Text,
    Background2,
    Text2,
    Accent1,
    Accent2,
    Accent3,
    Accent4,
    Accent5,
    Accent6,
    Hyperlink,
    FollowedHyperlink,
    Dark1,
    Light1,
    Dark2,
    Light2,
    Placeholder,
}

impl Scheme {
    /// Return the exact DrawingML token.
    #[inline]
    pub const fn token(self) -> &'static str {
        match self {
            Self::Background => "bg1",
            Self::Text => "tx1",
            Self::Background2 => "bg2",
            Self::Text2 => "tx2",
            Self::Accent1 => "accent1",
            Self::Accent2 => "accent2",
            Self::Accent3 => "accent3",
            Self::Accent4 => "accent4",
            Self::Accent5 => "accent5",
            Self::Accent6 => "accent6",
            Self::Hyperlink => "hlink",
            Self::FollowedHyperlink => "folHlink",
            Self::Dark1 => "dk1",
            Self::Light1 => "lt1",
            Self::Dark2 => "dk2",
            Self::Light2 => "lt2",
            Self::Placeholder => "phClr",
        }
    }

    /// Parse an exact DrawingML scheme-color token.
    #[inline]
    pub fn from_token(value: &str) -> Option<Self> {
        Some(match value {
            "bg1" => Self::Background,
            "tx1" => Self::Text,
            "bg2" => Self::Background2,
            "tx2" => Self::Text2,
            "accent1" => Self::Accent1,
            "accent2" => Self::Accent2,
            "accent3" => Self::Accent3,
            "accent4" => Self::Accent4,
            "accent5" => Self::Accent5,
            "accent6" => Self::Accent6,
            "hlink" => Self::Hyperlink,
            "folHlink" => Self::FollowedHyperlink,
            "dk1" => Self::Dark1,
            "lt1" => Self::Light1,
            "dk2" => Self::Dark2,
            "lt2" => Self::Light2,
            "phClr" => Self::Placeholder,
            _ => return None,
        })
    }
}

/// A bounded, well-formed DrawingML color fragment that this semantic owner
/// does not model yet.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
#[must_use]
pub struct Unknown {
    xml: Box<[u8]>,
}

impl Unknown {
    /// Validate and retain one complete color fragment.
    pub fn from_xml(xml: &[u8]) -> Result<Self> {
        Ok(Self {
            xml: super::codec::validated_fragment(xml)?
                .to_vec()
                .into_boxed_slice(),
        })
    }

    /// Return the retained fragment without copying it.
    #[inline]
    pub fn as_xml(&self) -> &[u8] {
        &self.xml
    }

    pub(super) fn from_validated(xml: &[u8]) -> Self {
        Self { xml: xml.into() }
    }
}

/// A format-neutral DrawingML color choice.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Value {
    /// Direct sRGB color (`a:srgbClr`).
    Rgb(Rgb),
    /// Theme-bound color (`a:schemeClr`).
    Scheme(Scheme),
    /// A valid color choice or transform not modeled by this shared owner.
    Unknown(Unknown),
}

impl Value {
    /// Construct a direct sRGB color.
    #[inline]
    pub const fn rgb(value: Rgb) -> Self {
        Self::Rgb(value)
    }

    /// Construct a theme-bound color.
    #[inline]
    pub const fn scheme(value: Scheme) -> Self {
        Self::Scheme(value)
    }

    /// Construct a checked opaque color fragment.
    #[inline]
    pub fn unknown(xml: &[u8]) -> Result<Self> {
        Ok(Self::Unknown(Unknown::from_xml(xml)?))
    }

    /// Return the direct sRGB value, when this is a typed RGB choice.
    #[inline]
    pub const fn as_rgb(&self) -> Option<Rgb> {
        match self {
            Self::Rgb(value) => Some(*value),
            Self::Scheme(_) | Self::Unknown(_) => None,
        }
    }

    /// Return the scheme token, when this is a typed theme-bound choice.
    #[inline]
    pub const fn as_scheme(&self) -> Option<Scheme> {
        match self {
            Self::Scheme(value) => Some(*value),
            Self::Rgb(_) | Self::Unknown(_) => None,
        }
    }

    /// Return whether this value is retained as an opaque fragment.
    #[inline]
    pub const fn is_unknown(&self) -> bool {
        matches!(self, Self::Unknown(_))
    }
}

pub(super) fn write_hex(output: &mut String, value: Rgb) {
    for channel in value.0 {
        let _ = fmt::Write::write_fmt(output, format_args!("{channel:02X}"));
    }
}

const fn hex_pair(high: u8, low: u8) -> u8 {
    (hex_digit(high) << 4) | hex_digit(low)
}

const fn hex_digit(value: u8) -> u8 {
    match value {
        b'0'..=b'9' => value - b'0',
        b'a'..=b'f' => value - b'a' + 10,
        b'A'..=b'F' => value - b'A' + 10,
        _ => 0,
    }
}
