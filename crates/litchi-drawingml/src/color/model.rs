//! Semantic `DrawingML` color values.

use std::fmt;

use super::validation;

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
    #[must_use]
    pub const fn channels(self) -> [u8; 3] {
        self.0
    }

    /// Return the red channel.
    #[inline]
    #[must_use]
    pub const fn red(self) -> u8 {
        self.0[0]
    }

    /// Return the green channel.
    #[inline]
    #[must_use]
    pub const fn green(self) -> u8 {
        self.0[1]
    }

    /// Return the blue channel.
    #[inline]
    #[must_use]
    pub const fn blue(self) -> u8 {
        self.0[2]
    }

    /// Parse the exact six-digit hexadecimal sRGB lexical form.
    pub fn parse(value: &str) -> crate::Result<Self> {
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
    #[must_use]
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

macro_rules! percentage_value {
    (
        $(#[$meta:meta])*
        $name:ident,
        $value:ty,
        $validate:ident,
        $parse:ident,
        $kind:literal
    ) => {
        $(#[$meta])*
        #[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
        #[repr(transparent)]
        #[must_use]
        pub struct $name($value);

        impl $name {
            /// Construct a checked DrawingML scalar from its canonical raw value.
            pub fn new(value: $value) -> crate::Result<Self> {
                Ok(Self(validation::$validate(value, $kind)?))
            }

            /// Parse a raw thousandth-of-a-percent or Office percent-sign value.
            pub fn parse(value: &str) -> crate::Result<Self> {
                Ok(Self(validation::$parse(value, $kind)?))
            }

            /// Return the canonical raw value in thousandths of a percent.
            #[inline]
            pub const fn value(self) -> $value {
                self.0
            }
        }
    };
}

percentage_value!(
    /// A signed `ST_Percentage` value in thousandths of a percent.
    Percentage,
    i32,
    percentage,
    parse_percentage,
    "percentages"
);
percentage_value!(
    /// A signed `ST_FixedPercentage` value in thousandths of a percent.
    FixedPercentage,
    i32,
    percentage,
    parse_percentage,
    "fixed percentages"
);
percentage_value!(
    /// A non-negative `ST_PositivePercentage` value in thousandths of a percent.
    PositivePercentage,
    u32,
    positive_percentage,
    parse_positive_percentage,
    "positive percentages"
);
percentage_value!(
    /// A non-negative `ST_PositiveFixedPercentage` value in thousandths of a percent.
    PositiveFixedPercentage,
    u32,
    positive_percentage,
    parse_positive_percentage,
    "positive fixed percentages"
);

/// A signed `ST_Angle` value in 60000ths of a degree.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
#[repr(transparent)]
#[must_use]
pub struct Angle(i32);

impl Angle {
    /// Construct a signed angle. The underlying schema type is an `xsd:int`.
    #[inline]
    pub const fn new(value: i32) -> Self {
        Self(value)
    }

    /// Parse an angle in 60000ths of a degree.
    pub fn parse(value: &str) -> crate::Result<Self> {
        Ok(Self(validation::parse_angle(value)?))
    }

    /// Return the raw angle value.
    #[inline]
    #[must_use]
    pub const fn value(self) -> i32 {
        self.0
    }
}

/// A checked `ST_PositiveFixedAngle` value in 60000ths of a degree.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
#[repr(transparent)]
#[must_use]
pub struct PositiveAngle(u32);

impl PositiveAngle {
    /// Construct an angle in the schema's inclusive 0–360 degree range.
    pub fn new(value: u32) -> crate::Result<Self> {
        Ok(Self(validation::positive_angle(value)?))
    }

    /// Parse an angle in 60000ths of a degree.
    pub fn parse(value: &str) -> crate::Result<Self> {
        Ok(Self(validation::parse_positive_angle(value)?))
    }

    /// Return the raw angle value.
    #[inline]
    #[must_use]
    pub const fn value(self) -> u32 {
        self.0
    }
}

/// A checked `scrgbClr` color with percentage channels.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[must_use]
pub struct ScRgb {
    red: PositiveFixedPercentage,
    green: PositiveFixedPercentage,
    blue: PositiveFixedPercentage,
}

impl ScRgb {
    /// Construct a scRGB color from raw thousandths-of-a-percent channels.
    pub fn new(red: u32, green: u32, blue: u32) -> crate::Result<Self> {
        Ok(Self::from_values(
            PositiveFixedPercentage::new(red)?,
            PositiveFixedPercentage::new(green)?,
            PositiveFixedPercentage::new(blue)?,
        ))
    }

    /// Construct a scRGB color from checked channels.
    #[inline]
    pub const fn from_values(
        red: PositiveFixedPercentage,
        green: PositiveFixedPercentage,
        blue: PositiveFixedPercentage,
    ) -> Self {
        Self { red, green, blue }
    }

    /// Return the red channel.
    #[inline]
    pub const fn red(self) -> PositiveFixedPercentage {
        self.red
    }

    /// Return the green channel.
    #[inline]
    pub const fn green(self) -> PositiveFixedPercentage {
        self.green
    }

    /// Return the blue channel.
    #[inline]
    pub const fn blue(self) -> PositiveFixedPercentage {
        self.blue
    }
}

/// A checked `hslClr` color.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[must_use]
pub struct Hsl {
    hue: PositiveAngle,
    saturation: Percentage,
    luminance: Percentage,
}

impl Hsl {
    /// Construct an HSL color from raw schema values.
    pub fn new(hue: u32, saturation: i32, luminance: i32) -> crate::Result<Self> {
        Ok(Self::from_values(
            PositiveAngle::new(hue)?,
            Percentage::new(saturation)?,
            Percentage::new(luminance)?,
        ))
    }

    /// Construct an HSL color from checked schema values.
    #[inline]
    pub const fn from_values(
        hue: PositiveAngle,
        saturation: Percentage,
        luminance: Percentage,
    ) -> Self {
        Self {
            hue,
            saturation,
            luminance,
        }
    }

    /// Return the hue in 60000ths of a degree.
    #[inline]
    pub const fn hue(self) -> PositiveAngle {
        self.hue
    }

    /// Return the saturation in thousandths of a percent.
    #[inline]
    pub const fn saturation(self) -> Percentage {
        self.saturation
    }

    /// Return the luminance in thousandths of a percent.
    #[inline]
    pub const fn luminance(self) -> Percentage {
        self.luminance
    }
}

const SYSTEM_VALUES: &[&str] = &[
    "scrollBar",
    "background",
    "activeCaption",
    "inactiveCaption",
    "menu",
    "window",
    "windowFrame",
    "menuText",
    "windowText",
    "captionText",
    "activeBorder",
    "inactiveBorder",
    "appWorkspace",
    "highlight",
    "highlightText",
    "btnFace",
    "btnShadow",
    "grayText",
    "btnText",
    "inactiveCaptionText",
    "btnHighlight",
    "3dDkShadow",
    "3dLight",
    "infoText",
    "infoBk",
    "hotLight",
    "gradientActiveCaption",
    "gradientInactiveCaption",
    "menuHilight",
    "menuBar",
];

/// A checked `sysClr` token and its optional cached RGB value.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
#[must_use]
pub struct System {
    value: Box<str>,
    last: Option<Rgb>,
}

impl System {
    /// Construct a system color from an `ST_SystemColorVal` token.
    pub fn new(value: impl AsRef<str>, last: Option<Rgb>) -> crate::Result<Self> {
        let value = value.as_ref();
        if !SYSTEM_VALUES.contains(&value) {
            return Err(crate::Error::Invalid(format!(
                "unknown DrawingML system color token: {value:?}"
            )));
        }
        Ok(Self {
            value: value.to_owned().into_boxed_str(),
            last,
        })
    }

    /// Parse a system color without a cached RGB value.
    #[must_use]
    pub fn from_token(value: &str) -> Option<Self> {
        Self::new(value, None).ok()
    }

    /// Return the exact `ST_SystemColorVal` token.
    #[inline]
    #[must_use]
    pub fn token(&self) -> &str {
        &self.value
    }

    /// Return the optional `lastClr` value.
    #[inline]
    #[must_use]
    pub const fn last_rgb(&self) -> Option<Rgb> {
        self.last
    }
}

const PRESET_VALUES: &[&str] = &[
    "aliceBlue",
    "antiqueWhite",
    "aqua",
    "aquamarine",
    "azure",
    "beige",
    "bisque",
    "black",
    "blanchedAlmond",
    "blue",
    "blueViolet",
    "brown",
    "burlyWood",
    "cadetBlue",
    "chartreuse",
    "chocolate",
    "coral",
    "cornflowerBlue",
    "cornsilk",
    "crimson",
    "cyan",
    "darkBlue",
    "darkCyan",
    "darkGoldenrod",
    "darkGray",
    "darkGreen",
    "darkGrey",
    "darkKhaki",
    "darkMagenta",
    "darkOliveGreen",
    "darkOrange",
    "darkOrchid",
    "darkRed",
    "darkSalmon",
    "darkSeaGreen",
    "darkSlateBlue",
    "darkSlateGray",
    "darkSlateGrey",
    "darkTurquoise",
    "darkViolet",
    "deepPink",
    "deepSkyBlue",
    "dimGray",
    "dimGrey",
    "dodgerBlue",
    "firebrick",
    "floralWhite",
    "forestGreen",
    "fuchsia",
    "gainsboro",
    "ghostWhite",
    "gold",
    "goldenrod",
    "gray",
    "green",
    "greenYellow",
    "grey",
    "honeydew",
    "hotPink",
    "indianRed",
    "indigo",
    "ivory",
    "khaki",
    "lavender",
    "lavenderBlush",
    "lawnGreen",
    "lemonChiffon",
    "lightBlue",
    "lightCoral",
    "lightCyan",
    "lightGoldenrodYellow",
    "lightGray",
    "lightGreen",
    "lightGrey",
    "lightPink",
    "lightSalmon",
    "lightSeaGreen",
    "lightSkyBlue",
    "lightSlateGray",
    "lightSlateGrey",
    "lightSteelBlue",
    "lightYellow",
    "lime",
    "limeGreen",
    "linen",
    "magenta",
    "maroon",
    "mediumAquamarine",
    "mediumBlue",
    "mediumOrchid",
    "mediumPurple",
    "mediumSeaGreen",
    "mediumSlateBlue",
    "mediumSpringGreen",
    "mediumTurquoise",
    "mediumVioletRed",
    "midnightBlue",
    "mintCream",
    "mistyRose",
    "moccasin",
    "navajoWhite",
    "navy",
    "oldLace",
    "olive",
    "oliveDrab",
    "orange",
    "orangeRed",
    "orchid",
    "paleGoldenrod",
    "paleGreen",
    "paleTurquoise",
    "paleVioletRed",
    "papayaWhip",
    "peachPuff",
    "peru",
    "pink",
    "plum",
    "powderBlue",
    "purple",
    "red",
    "rosyBrown",
    "royalBlue",
    "saddleBrown",
    "salmon",
    "sandyBrown",
    "seaGreen",
    "seaShell",
    "sienna",
    "silver",
    "skyBlue",
    "slateBlue",
    "slateGray",
    "slateGrey",
    "snow",
    "springGreen",
    "steelBlue",
    "tan",
    "teal",
    "thistle",
    "tomato",
    "turquoise",
    "violet",
    "wheat",
    "white",
    "whiteSmoke",
    "yellow",
    "yellowGreen",
    "transparent",
];

/// A checked `prstClr` token.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
#[must_use]
pub struct Preset(Box<str>);

impl Preset {
    /// Construct a preset color from an `ST_PresetColorVal` token.
    pub fn new(value: impl AsRef<str>) -> crate::Result<Self> {
        let value = value.as_ref();
        if !PRESET_VALUES.contains(&value) {
            return Err(crate::Error::Invalid(format!(
                "unknown DrawingML preset color token: {value:?}"
            )));
        }
        Ok(Self(value.to_owned().into_boxed_str()))
    }

    /// Parse a preset color token.
    #[must_use]
    pub fn from_token(value: &str) -> Option<Self> {
        Self::new(value).ok()
    }

    /// Return the exact `ST_PresetColorVal` token.
    #[inline]
    #[must_use]
    pub fn token(&self) -> &str {
        &self.0
    }
}

/// A typed `DrawingML` color choice without transforms.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Base {
    /// Direct sRGB color (`a:srgbClr`).
    Rgb(Rgb),
    /// Theme-bound color (`a:schemeClr`).
    Scheme(Scheme),
    /// Scaled RGB color (`a:scrgbClr`).
    ScRgb(ScRgb),
    /// Hue/saturation/luminance color (`a:hslClr`).
    Hsl(Hsl),
    /// Operating-system-bound color (`a:sysClr`).
    System(System),
    /// Preset color (`a:prstClr`).
    Preset(Preset),
}

impl From<Rgb> for Base {
    #[inline]
    fn from(value: Rgb) -> Self {
        Self::Rgb(value)
    }
}

impl From<Scheme> for Base {
    #[inline]
    fn from(value: Scheme) -> Self {
        Self::Scheme(value)
    }
}

impl From<ScRgb> for Base {
    #[inline]
    fn from(value: ScRgb) -> Self {
        Self::ScRgb(value)
    }
}

impl From<Hsl> for Base {
    #[inline]
    fn from(value: Hsl) -> Self {
        Self::Hsl(value)
    }
}

impl From<System> for Base {
    #[inline]
    fn from(value: System) -> Self {
        Self::System(value)
    }
}

impl From<Preset> for Base {
    #[inline]
    fn from(value: Preset) -> Self {
        Self::Preset(value)
    }
}

/// An ordered, typed `DrawingML` color transform sequence.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Transform {
    /// Set opacity (`a:alpha`).
    Alpha(PositiveFixedPercentage),
    /// Modulate opacity (`a:alphaMod`).
    AlphaMod(PositivePercentage),
    /// Offset opacity (`a:alphaOff`).
    AlphaOff(FixedPercentage),
    /// Set the blue channel (`a:blue`).
    Blue(Percentage),
    /// Modulate the blue channel (`a:blueMod`).
    BlueMod(Percentage),
    /// Offset the blue channel (`a:blueOff`).
    BlueOff(Percentage),
    /// Complement (`a:comp`).
    Complement,
    /// Apply sRGB gamma (`a:gamma`).
    Gamma,
    /// Convert to grayscale (`a:gray`).
    Gray,
    /// Set the green channel (`a:green`).
    Green(Percentage),
    /// Modulate the green channel (`a:greenMod`).
    GreenMod(Percentage),
    /// Offset the green channel (`a:greenOff`).
    GreenOff(Percentage),
    /// Set the hue (`a:hue`).
    Hue(PositiveAngle),
    /// Modulate the hue (`a:hueMod`).
    HueMod(PositivePercentage),
    /// Offset the hue (`a:hueOff`).
    HueOff(Angle),
    /// Inverse (`a:inv`).
    Inverse,
    /// Apply inverse sRGB gamma (`a:invGamma`).
    InverseGamma,
    /// Set the luminance (`a:lum`).
    Lum(Percentage),
    /// Modulate the luminance (`a:lumMod`).
    LumMod(Percentage),
    /// Offset the luminance (`a:lumOff`).
    LumOff(Percentage),
    /// Set the red channel (`a:red`).
    Red(Percentage),
    /// Modulate the red channel (`a:redMod`).
    RedMod(Percentage),
    /// Offset the red channel (`a:redOff`).
    RedOff(Percentage),
    /// Set the saturation (`a:sat`).
    Sat(Percentage),
    /// Modulate the saturation (`a:satMod`).
    SatMod(Percentage),
    /// Offset the saturation (`a:satOff`).
    SatOff(Percentage),
    /// Darken (`a:shade`).
    Shade(PositiveFixedPercentage),
    /// Lighten (`a:tint`).
    Tint(PositiveFixedPercentage),
}

/// A color choice plus an ordered transform sequence.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
#[must_use]
pub struct Transformed {
    base: Base,
    transforms: Box<[Transform]>,
}

impl Transformed {
    /// Construct a transformed color. The sequence must contain at least one
    /// transform and may not exceed [`super::MAX_TRANSFORMS`].
    pub fn new<I>(base: Base, transforms: I) -> crate::Result<Self>
    where
        I: IntoIterator<Item = Transform>,
    {
        let transforms: Box<[_]> = transforms.into_iter().collect();
        if transforms.is_empty() {
            return Err(crate::Error::Invalid(
                "DrawingML transformed colors require at least one transform".into(),
            ));
        }
        if transforms.len() > super::MAX_TRANSFORMS {
            return Err(crate::Error::Limit {
                resource: "DrawingML color transforms",
                limit: super::MAX_TRANSFORMS,
            });
        }
        Ok(Self { base, transforms })
    }

    /// Return the untransformed color choice.
    #[inline]
    #[must_use]
    pub const fn base(&self) -> &Base {
        &self.base
    }

    /// Return transforms in their source order.
    #[inline]
    #[must_use]
    pub fn transforms(&self) -> &[Transform] {
        &self.transforms
    }
}

/// A format-neutral `DrawingML` color choice.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Value {
    /// Direct sRGB color (`a:srgbClr`).
    Rgb(Rgb),
    /// Theme-bound color (`a:schemeClr`).
    Scheme(Scheme),
    /// Scaled RGB color (`a:scrgbClr`).
    ScRgb(ScRgb),
    /// Hue/saturation/luminance color (`a:hslClr`).
    Hsl(Hsl),
    /// Operating-system-bound color (`a:sysClr`).
    System(System),
    /// Preset color (`a:prstClr`).
    Preset(Preset),
    /// A typed color plus an ordered transform sequence.
    Transformed(Transformed),
    /// A valid color choice or transform not modeled by this shared owner.
    Unknown(Unknown),
}

impl Value {
    /// Construct a direct sRGB color.
    #[inline]
    #[must_use]
    pub const fn rgb(value: Rgb) -> Self {
        Self::Rgb(value)
    }

    /// Construct a theme-bound color.
    #[inline]
    #[must_use]
    pub const fn scheme(value: Scheme) -> Self {
        Self::Scheme(value)
    }

    /// Construct a scaled RGB color.
    #[inline]
    #[must_use]
    pub const fn scrgb(value: ScRgb) -> Self {
        Self::ScRgb(value)
    }

    /// Construct an HSL color.
    #[inline]
    #[must_use]
    pub const fn hsl(value: Hsl) -> Self {
        Self::Hsl(value)
    }

    /// Construct a system color.
    #[inline]
    #[must_use]
    pub fn system(value: System) -> Self {
        Self::System(value)
    }

    /// Construct a preset color.
    #[inline]
    #[must_use]
    pub fn preset(value: Preset) -> Self {
        Self::Preset(value)
    }

    /// Construct a typed color with an ordered transform sequence.
    pub fn transformed<I>(base: impl Into<Base>, transforms: I) -> crate::Result<Self>
    where
        I: IntoIterator<Item = Transform>,
    {
        Ok(Self::Transformed(Transformed::new(
            base.into(),
            transforms,
        )?))
    }

    /// Construct a checked opaque color fragment.
    #[inline]
    pub fn unknown(xml: &[u8]) -> crate::Result<Self> {
        Ok(Self::Unknown(Unknown::from_xml(xml)?))
    }

    /// Return the direct sRGB value, when this is an untransformed RGB choice.
    #[inline]
    #[must_use]
    pub const fn as_rgb(&self) -> Option<Rgb> {
        match self {
            Self::Rgb(value) => Some(*value),
            Self::Scheme(_)
            | Self::ScRgb(_)
            | Self::Hsl(_)
            | Self::System(_)
            | Self::Preset(_)
            | Self::Transformed(_)
            | Self::Unknown(_) => None,
        }
    }

    /// Return the scheme token, when this is an untransformed theme choice.
    #[inline]
    #[must_use]
    pub const fn as_scheme(&self) -> Option<Scheme> {
        match self {
            Self::Scheme(value) => Some(*value),
            Self::Rgb(_)
            | Self::ScRgb(_)
            | Self::Hsl(_)
            | Self::System(_)
            | Self::Preset(_)
            | Self::Transformed(_)
            | Self::Unknown(_) => None,
        }
    }

    /// Return the transformed representation, if present.
    #[inline]
    #[must_use]
    pub const fn as_transformed(&self) -> Option<&Transformed> {
        match self {
            Self::Transformed(value) => Some(value),
            Self::Rgb(_)
            | Self::Scheme(_)
            | Self::ScRgb(_)
            | Self::Hsl(_)
            | Self::System(_)
            | Self::Preset(_)
            | Self::Unknown(_) => None,
        }
    }

    /// Return whether this value is retained as an opaque fragment.
    #[inline]
    #[must_use]
    pub const fn is_unknown(&self) -> bool {
        matches!(self, Self::Unknown(_))
    }

    pub(super) fn from_base(base: Base) -> Self {
        match base {
            Base::Rgb(value) => Self::Rgb(value),
            Base::Scheme(value) => Self::Scheme(value),
            Base::ScRgb(value) => Self::ScRgb(value),
            Base::Hsl(value) => Self::Hsl(value),
            Base::System(value) => Self::System(value),
            Base::Preset(value) => Self::Preset(value),
        }
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
    /// Return the exact `DrawingML` token.
    #[inline]
    #[must_use]
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

    /// Parse an exact `DrawingML` scheme-color token.
    #[inline]
    #[must_use]
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

/// A bounded, well-formed `DrawingML` color fragment that this semantic owner
/// does not model yet.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
#[must_use]
pub struct Unknown {
    xml: Box<[u8]>,
}

impl Unknown {
    /// Validate and retain one complete color fragment.
    pub fn from_xml(xml: &[u8]) -> crate::Result<Self> {
        Ok(Self {
            xml: validation::validated_fragment(xml)?
                .to_vec()
                .into_boxed_slice(),
        })
    }

    /// Return the retained fragment without copying it.
    #[inline]
    #[must_use]
    pub fn as_xml(&self) -> &[u8] {
        &self.xml
    }

    pub(super) fn from_validated(xml: &[u8]) -> Self {
        Self { xml: xml.into() }
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
