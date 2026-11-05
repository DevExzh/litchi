//! Typed values shared by the direct-property readers, and the lexical
//! parsers that produce them. The conventions are documented on the
//! `properties` module.

use crate::color::Theme;

/// A length in twentieths of a point (twips, the `dxa` unit).
///
/// ```
/// use litchi_docx::properties::Twips;
///
/// assert_eq!(Twips(1440).0, 1440); // one inch
/// ```
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash, Default)]
pub struct Twips(pub i64);

/// A length in half points: font sizes, kerning thresholds and baseline
/// positions.
///
/// ```
/// use litchi_docx::properties::HalfPoints;
///
/// assert_eq!(HalfPoints(22).0, 22); // an 11 pt font
/// ```
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash, Default)]
pub struct HalfPoints(pub i64);

/// A length in eighths of a point: border widths (`w:sz` of a border).
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash, Default)]
pub struct EighthPoints(pub i64);

/// A length in whole points: the spacing between a border and its content
/// (`w:space`).
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash, Default)]
pub struct Points(pub i64);

/// Generate a closed WordprocessingML vocabulary with exact token
/// conversions.
macro_rules! vocabulary {
    (
        $(#[$meta:meta])*
        $vis:vis enum $name:ident {
            $(
                $(#[$variant_meta:meta])*
                $variant:ident = $token:literal,
            )+
        }
    ) => {
        $(#[$meta])*
        #[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
        $vis enum $name {
            $(
                $(#[$variant_meta])*
                $variant,
            )+
        }

        impl $name {
            /// Every value, in schema order.
            pub const ALL: &'static [Self] = &[$(Self::$variant,)+];

            /// Parse the exact, case-sensitive WordprocessingML token.
            #[must_use]
            pub fn from_xml(value: &str) -> Option<Self> {
                match value {
                    $($token => Some(Self::$variant),)+
                    _ => None,
                }
            }

            /// Return the exact WordprocessingML token.
            #[must_use]
            pub const fn as_str(self) -> &'static str {
                match self {
                    $(Self::$variant => $token,)+
                }
            }
        }
    };
}

pub(super) use vocabulary;

/// An `ST_HexColor` value: automatic, or explicit red, green and blue.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum HexColor {
    /// `auto`: the consumer chooses, normally black on light backgrounds.
    Auto,
    /// Explicit red, green and blue components.
    Rgb([u8; 3]),
}

/// A direct colour (`CT_Color`): the stored value and an optional theme
/// reference with its transforms, kept as authored.
///
/// When a theme colour is present Word renders the theme colour (after the
/// tint or shade) and keeps the stored value as a fallback.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Color {
    /// The stored value (`w:val`).
    pub value: HexColor,
    /// The theme colour reference (`w:themeColor`).
    pub theme: Option<Theme>,
    /// The theme tint transform (`w:themeTint`).
    pub theme_tint: Option<u8>,
    /// The theme shade transform (`w:themeShade`).
    pub theme_shade: Option<u8>,
}

vocabulary! {
    /// A built-in text highlight colour (`ST_HighlightColor`).
    pub enum Highlight {
        /// `black`
        Black = "black",
        /// `blue`
        Blue = "blue",
        /// `cyan`
        Cyan = "cyan",
        /// `green`
        Green = "green",
        /// `magenta`
        Magenta = "magenta",
        /// `red`
        Red = "red",
        /// `yellow`
        Yellow = "yellow",
        /// `white`
        White = "white",
        /// `darkBlue`
        DarkBlue = "darkBlue",
        /// `darkCyan`
        DarkCyan = "darkCyan",
        /// `darkGreen`
        DarkGreen = "darkGreen",
        /// `darkMagenta`
        DarkMagenta = "darkMagenta",
        /// `darkRed`
        DarkRed = "darkRed",
        /// `darkYellow`
        DarkYellow = "darkYellow",
        /// `darkGray`
        DarkGray = "darkGray",
        /// `lightGray`
        LightGray = "lightGray",
        /// `none`: explicitly no highlight.
        None = "none",
    }
}

vocabulary! {
    /// Which font slot renders characters shared by several scripts
    /// (`w:rFonts/@w:hint`).
    ///
    /// The schema lists `default` and `eastAsia`; Word also writes and reads
    /// `cs` ([MS-OI29500]).
    pub enum FontHint {
        /// `default`
        Default = "default",
        /// `eastAsia`
        EastAsia = "eastAsia",
        /// `cs`: the complex-script slot.
        ComplexScript = "cs",
    }
}

vocabulary! {
    /// A theme font reference (`ST_Theme`), kept instead of the resolved
    /// family name.
    pub enum ThemeFont {
        /// `majorEastAsia`
        MajorEastAsia = "majorEastAsia",
        /// `majorBidi`
        MajorBidi = "majorBidi",
        /// `majorAscii`
        MajorAscii = "majorAscii",
        /// `majorHAnsi`
        MajorHAnsi = "majorHAnsi",
        /// `minorEastAsia`
        MinorEastAsia = "minorEastAsia",
        /// `minorBidi`
        MinorBidi = "minorBidi",
        /// `minorAscii`
        MinorAscii = "minorAscii",
        /// `minorHAnsi`
        MinorHAnsi = "minorHAnsi",
    }
}

/// The direct font slots of a run (`w:rFonts`).
///
/// Every slot is exactly what the run states; an unstated slot is `None` and
/// inherits from the style chain. A theme slot (`*_theme`) takes precedence
/// over the literal family of the same slot when both are present.
#[derive(Clone, Default, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub struct Fonts {
    /// `w:hint`
    pub hint: Option<FontHint>,
    /// `w:ascii`: the family for ASCII characters.
    pub ascii: Option<String>,
    /// `w:hAnsi`: the family for other Latin characters.
    pub h_ansi: Option<String>,
    /// `w:eastAsia`
    pub east_asia: Option<String>,
    /// `w:cs`: the complex-script family.
    pub complex_script: Option<String>,
    /// `w:asciiTheme`
    pub ascii_theme: Option<ThemeFont>,
    /// `w:hAnsiTheme`
    pub h_ansi_theme: Option<ThemeFont>,
    /// `w:eastAsiaTheme`
    pub east_asia_theme: Option<ThemeFont>,
    /// `w:cstheme`
    pub complex_script_theme: Option<ThemeFont>,
}

/// The direct language tags of a run (`w:lang`), as authored (a BCP 47 tag
/// or a hexadecimal language identifier).
#[derive(Clone, Default, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub struct Language {
    /// `w:val`: the Latin-script language.
    pub latin: Option<String>,
    /// `w:eastAsia`
    pub east_asia: Option<String>,
    /// `w:bidi`: the complex-script language.
    pub bidi: Option<String>,
}

vocabulary! {
    /// The vertical placement of a run (`w:vertAlign`).
    pub enum VerticalAlign {
        /// `baseline`
        Baseline = "baseline",
        /// `superscript`
        Superscript = "superscript",
        /// `subscript`
        Subscript = "subscript",
    }
}

vocabulary! {
    /// An East Asian emphasis mark (`w:em`).
    pub enum Emphasis {
        /// `none`
        None = "none",
        /// `dot`
        Dot = "dot",
        /// `comma`
        Comma = "comma",
        /// `circle`
        Circle = "circle",
        /// `underDot`
        UnderDot = "underDot",
    }
}

vocabulary! {
    /// A legacy animated text effect (`w:effect`).
    pub enum TextEffect {
        /// `blinkBackground`
        BlinkBackground = "blinkBackground",
        /// `lights`
        Lights = "lights",
        /// `antsBlack`
        AntsBlack = "antsBlack",
        /// `antsRed`
        AntsRed = "antsRed",
        /// `shimmer`
        Shimmer = "shimmer",
        /// `sparkle`
        Sparkle = "sparkle",
        /// `none`
        None = "none",
    }
}

vocabulary! {
    /// A shading pattern (`ST_Shd`).
    pub enum ShadingPattern {
        /// `nil`: no shading.
        Nil = "nil",
        /// `clear`: the fill colour only.
        Clear = "clear",
        /// `solid`: the pattern colour only.
        Solid = "solid",
        /// `horzStripe`
        HorizontalStripe = "horzStripe",
        /// `vertStripe`
        VerticalStripe = "vertStripe",
        /// `reverseDiagStripe`
        ReverseDiagonalStripe = "reverseDiagStripe",
        /// `diagStripe`
        DiagonalStripe = "diagStripe",
        /// `horzCross`
        HorizontalCross = "horzCross",
        /// `diagCross`
        DiagonalCross = "diagCross",
        /// `thinHorzStripe`
        ThinHorizontalStripe = "thinHorzStripe",
        /// `thinVertStripe`
        ThinVerticalStripe = "thinVertStripe",
        /// `thinReverseDiagStripe`
        ThinReverseDiagonalStripe = "thinReverseDiagStripe",
        /// `thinDiagStripe`
        ThinDiagonalStripe = "thinDiagStripe",
        /// `thinHorzCross`
        ThinHorizontalCross = "thinHorzCross",
        /// `thinDiagCross`
        ThinDiagonalCross = "thinDiagCross",
        /// `pct5`
        Percent5 = "pct5",
        /// `pct10`
        Percent10 = "pct10",
        /// `pct12`
        Percent12 = "pct12",
        /// `pct15`
        Percent15 = "pct15",
        /// `pct20`
        Percent20 = "pct20",
        /// `pct25`
        Percent25 = "pct25",
        /// `pct30`
        Percent30 = "pct30",
        /// `pct35`
        Percent35 = "pct35",
        /// `pct37`
        Percent37 = "pct37",
        /// `pct40`
        Percent40 = "pct40",
        /// `pct45`
        Percent45 = "pct45",
        /// `pct50`
        Percent50 = "pct50",
        /// `pct55`
        Percent55 = "pct55",
        /// `pct60`
        Percent60 = "pct60",
        /// `pct62`
        Percent62 = "pct62",
        /// `pct65`
        Percent65 = "pct65",
        /// `pct70`
        Percent70 = "pct70",
        /// `pct75`
        Percent75 = "pct75",
        /// `pct80`
        Percent80 = "pct80",
        /// `pct85`
        Percent85 = "pct85",
        /// `pct87`
        Percent87 = "pct87",
        /// `pct90`
        Percent90 = "pct90",
        /// `pct95`
        Percent95 = "pct95",
    }
}

/// Direct shading (`CT_Shd`) of a run, paragraph, table or cell.
///
/// The pattern is drawn in the pattern colour over the fill colour; each
/// colour keeps its theme reference and transforms as authored.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub struct Shading {
    /// `w:val`
    pub pattern: ShadingPattern,
    /// `w:color`: the pattern colour.
    pub color: Option<HexColor>,
    /// `w:themeColor`
    pub theme_color: Option<Theme>,
    /// `w:themeTint`
    pub theme_tint: Option<u8>,
    /// `w:themeShade`
    pub theme_shade: Option<u8>,
    /// `w:fill`: the background colour.
    pub fill: Option<HexColor>,
    /// `w:themeFill`
    pub theme_fill: Option<Theme>,
    /// `w:themeFillTint`
    pub theme_fill_tint: Option<u8>,
    /// `w:themeFillShade`
    pub theme_fill_shade: Option<u8>,
}

impl Shading {
    /// Shading with `pattern` and no colour, the starting point for writing
    /// one with [`crate::document::Edit::format_paragraph`] or
    /// [`crate::document::Edit::format_run`].
    ///
    /// ```
    /// use litchi_docx::properties::{HexColor, Shading, ShadingPattern};
    ///
    /// let mut shading = Shading::new(ShadingPattern::Clear);
    /// shading.fill = Some(HexColor::Rgb([0xD9, 0xD9, 0xD9]));
    /// assert_eq!(shading.pattern, ShadingPattern::Clear);
    /// ```
    #[must_use]
    pub const fn new(pattern: ShadingPattern) -> Self {
        Self {
            pattern,
            color: None,
            theme_color: None,
            theme_tint: None,
            theme_shade: None,
            fill: None,
            theme_fill: None,
            theme_fill_tint: None,
            theme_fill_shade: None,
        }
    }
}

/// Every `ST_Border` token, sorted by byte value for binary search.
const BORDER_STYLES: [&str; 193] = [
    "apples",
    "archedScallops",
    "babyPacifier",
    "babyRattle",
    "balloons3Colors",
    "balloonsHotAir",
    "basicBlackDashes",
    "basicBlackDots",
    "basicBlackSquares",
    "basicThinLines",
    "basicWhiteDashes",
    "basicWhiteDots",
    "basicWhiteSquares",
    "basicWideInline",
    "basicWideMidline",
    "basicWideOutline",
    "bats",
    "birds",
    "birdsFlight",
    "cabins",
    "cakeSlice",
    "candyCorn",
    "celticKnotwork",
    "certificateBanner",
    "chainLink",
    "champagneBottle",
    "checkedBarBlack",
    "checkedBarColor",
    "checkered",
    "christmasTree",
    "circlesLines",
    "circlesRectangles",
    "classicalWave",
    "clocks",
    "compass",
    "confetti",
    "confettiGrays",
    "confettiOutline",
    "confettiStreamers",
    "confettiWhite",
    "cornerTriangles",
    "couponCutoutDashes",
    "couponCutoutDots",
    "crazyMaze",
    "creaturesButterfly",
    "creaturesFish",
    "creaturesInsects",
    "creaturesLadyBug",
    "crossStitch",
    "cup",
    "custom",
    "dashDotStroked",
    "dashSmallGap",
    "dashed",
    "decoArch",
    "decoArchColor",
    "decoBlocks",
    "diamondsGray",
    "dotDash",
    "dotDotDash",
    "dotted",
    "double",
    "doubleD",
    "doubleDiamonds",
    "doubleWave",
    "earth1",
    "earth2",
    "earth3",
    "eclipsingSquares1",
    "eclipsingSquares2",
    "eggsBlack",
    "fans",
    "film",
    "firecrackers",
    "flowersBlockPrint",
    "flowersDaisies",
    "flowersModern1",
    "flowersModern2",
    "flowersPansy",
    "flowersRedRose",
    "flowersRoses",
    "flowersTeacup",
    "flowersTiny",
    "gems",
    "gingerbreadMan",
    "gradient",
    "handmade1",
    "handmade2",
    "heartBalloon",
    "heartGray",
    "hearts",
    "heebieJeebies",
    "holly",
    "houseFunky",
    "hypnotic",
    "iceCreamCones",
    "inset",
    "lightBulb",
    "lightning1",
    "lightning2",
    "mapPins",
    "mapleLeaf",
    "mapleMuffins",
    "marquee",
    "marqueeToothed",
    "moons",
    "mosaic",
    "musicNotes",
    "nil",
    "none",
    "northwest",
    "outset",
    "ovals",
    "packages",
    "palmsBlack",
    "palmsColor",
    "paperClips",
    "papyrus",
    "partyFavor",
    "partyGlass",
    "pencils",
    "people",
    "peopleHats",
    "peopleWaving",
    "poinsettias",
    "postageStamp",
    "pumpkin1",
    "pushPinNote1",
    "pushPinNote2",
    "pyramids",
    "pyramidsAbove",
    "quadrants",
    "rings",
    "safari",
    "sawtooth",
    "sawtoothGray",
    "scaredCat",
    "seattle",
    "shadowedSquares",
    "shapes1",
    "shapes2",
    "sharksTeeth",
    "shorebirdTracks",
    "single",
    "skyrocket",
    "snowflakeFancy",
    "snowflakes",
    "sombrero",
    "southwest",
    "stars",
    "stars3d",
    "starsBlack",
    "starsShadowed",
    "starsTop",
    "sun",
    "swirligig",
    "thick",
    "thickThinLargeGap",
    "thickThinMediumGap",
    "thickThinSmallGap",
    "thinThickLargeGap",
    "thinThickMediumGap",
    "thinThickSmallGap",
    "thinThickThinLargeGap",
    "thinThickThinMediumGap",
    "thinThickThinSmallGap",
    "threeDEmboss",
    "threeDEngrave",
    "tornPaper",
    "tornPaperBlack",
    "trees",
    "triangle1",
    "triangle2",
    "triangleCircle1",
    "triangleCircle2",
    "triangleParty",
    "triangles",
    "triple",
    "twistedLines1",
    "twistedLines2",
    "vine",
    "wave",
    "waveline",
    "weavingAngles",
    "weavingBraid",
    "weavingRibbon",
    "weavingStrips",
    "whiteFlowers",
    "woodwork",
    "xIllusions",
    "zanyTriangles",
    "zigZag",
    "zigZagStitch",
];

const _: () = {
    // The table is strictly sorted by bytes, so binary search finds every
    // token, and its length fits the one-byte index.
    let mut index = 1;
    while index < BORDER_STYLES.len() {
        let previous = BORDER_STYLES[index - 1].as_bytes();
        let current = BORDER_STYLES[index].as_bytes();
        let mut byte = 0;
        let mut ordered = false;
        loop {
            if byte == previous.len() {
                ordered = byte < current.len();
                break;
            }
            if byte == current.len() {
                break;
            }
            if previous[byte] != current[byte] {
                ordered = previous[byte] < current[byte];
                break;
            }
            byte += 1;
        }
        assert!(ordered, "BORDER_STYLES must be strictly sorted");
        index += 1;
    }
    assert!(BORDER_STYLES.len() <= u8::MAX as usize + 1);
};

/// A border line style (`ST_Border`): one of the schema's 193 line and art
/// border tokens.
///
/// The value is a checked token rather than an enum because most of the
/// vocabulary is decorative art borders that consumers treat uniformly.
///
/// ```
/// use litchi_docx::properties::BorderStyle;
///
/// let style = BorderStyle::from_xml("single").unwrap();
/// assert_eq!(style.as_str(), "single");
/// assert!(!style.is_none());
/// assert!(BorderStyle::from_xml("nil").unwrap().is_none());
/// assert_eq!(BorderStyle::from_xml("Single"), None);
/// ```
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct BorderStyle(u8);

impl BorderStyle {
    /// Parse the exact, case-sensitive `ST_Border` token.
    #[must_use]
    pub fn from_xml(value: &str) -> Option<Self> {
        BORDER_STYLES
            .binary_search(&value)
            .ok()
            .and_then(|index| u8::try_from(index).ok())
            .map(Self)
    }

    /// Return the exact `ST_Border` token.
    #[must_use]
    pub fn as_str(self) -> &'static str {
        BORDER_STYLES[usize::from(self.0)]
    }

    /// Whether the style draws no line (`nil` or `none`).
    #[must_use]
    pub fn is_none(self) -> bool {
        matches!(self.as_str(), "nil" | "none")
    }
}

/// One direct border (`CT_Border`) of a run, paragraph, table or cell.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub struct Border {
    /// `w:val`
    pub style: BorderStyle,
    /// `w:color`; the schema default is `auto`.
    pub color: Option<HexColor>,
    /// `w:themeColor`
    pub theme_color: Option<Theme>,
    /// `w:themeTint`
    pub theme_tint: Option<u8>,
    /// `w:themeShade`
    pub theme_shade: Option<u8>,
    /// `w:sz`: the line width.
    pub size: Option<EighthPoints>,
    /// `w:space`: the distance from the content; the schema default is 0.
    pub space: Option<Points>,
    /// `w:shadow`
    pub shadow: Option<bool>,
    /// `w:frame`
    pub frame: Option<bool>,
}

impl Border {
    /// A border of `style` with no other attribute, the starting point for
    /// writing one.
    ///
    /// ```
    /// use litchi_docx::properties::{Border, BorderStyle, EighthPoints};
    ///
    /// let mut border = Border::new(BorderStyle::from_xml("single").unwrap());
    /// border.size = Some(EighthPoints(8));
    /// assert_eq!(border.style.as_str(), "single");
    /// ```
    #[must_use]
    pub const fn new(style: BorderStyle) -> Self {
        Self {
            style,
            color: None,
            theme_color: None,
            theme_tint: None,
            theme_shade: None,
            size: None,
            space: None,
            shadow: None,
            frame: None,
        }
    }
}

// ---------------------------------------------------------------------------
// Lexical parsers. Each returns `None` for a value outside its schema domain.
// ---------------------------------------------------------------------------

/// Trim the whitespace an `xsd:boolean`, `xsd:integer` or `xsd:decimal`
/// lexical form may carry (their whitespace facet is `collapse`).
fn collapse(value: &str) -> &str {
    value.trim_matches(|character| matches!(character, ' ' | '\t' | '\r' | '\n'))
}

/// `ST_OnOff`: `xsd:boolean` plus `on` and `off`.
pub(super) fn on_off(value: &str) -> Option<bool> {
    match collapse(value) {
        "true" | "1" | "on" => Some(true),
        "false" | "0" | "off" => Some(false),
        _ => None,
    }
}

/// `xsd:integer` (and `ST_DecimalNumber`) as an `i64`.
pub(super) fn integer(value: &str) -> Option<i64> {
    let value = collapse(value);
    let (negative, digits) = match value.as_bytes().first() {
        Some(b'-') => (true, &value[1..]),
        Some(b'+') => (false, &value[1..]),
        _ => (false, value),
    };
    if digits.is_empty() || !digits.bytes().all(|byte| byte.is_ascii_digit()) {
        return None;
    }
    let magnitude = digits.parse::<i128>().ok()?;
    i64::try_from(if negative { -magnitude } else { magnitude }).ok()
}

/// `xsd:integer` narrowed to `u32`; negative values are out of domain.
pub(super) fn unsigned(value: &str) -> Option<u32> {
    integer(value).and_then(|number| u32::try_from(number).ok())
}

/// Hexadecimal digits of a fixed width (`ST_UcharHexNumber` is two,
/// `ST_ShortHexNumber` four, `ST_HexColorRGB` six).
fn hex<const DIGITS: usize>(value: &str) -> Option<u32> {
    let value = collapse(value);
    if value.len() != DIGITS || !value.bytes().all(|byte| byte.is_ascii_hexdigit()) {
        return None;
    }
    u32::from_str_radix(value, 16).ok()
}

/// `ST_UcharHexNumber`: one byte as two hexadecimal digits.
pub(super) fn hex_byte(value: &str) -> Option<u8> {
    hex::<2>(value).and_then(|number| u8::try_from(number).ok())
}

/// `ST_ShortHexNumber`: two bytes as four hexadecimal digits.
pub(super) fn hex_short(value: &str) -> Option<u16> {
    hex::<4>(value).and_then(|number| u16::try_from(number).ok())
}

/// `ST_HexColor`: `auto` or six hexadecimal digits.
pub(super) fn hex_color(value: &str) -> Option<HexColor> {
    if collapse(value) == "auto" {
        return Some(HexColor::Auto);
    }
    let rgb = hex::<6>(value)?;
    let [_, red, green, blue] = rgb.to_be_bytes();
    Some(HexColor::Rgb([red, green, blue]))
}

/// `ST_ThemeColor`.
pub(super) fn theme_color(value: &str) -> Option<Theme> {
    Theme::parse(value)
}

/// `ST_TextScale`: a percentage from 0 to 600, written as an integer or with
/// a trailing `%`.
pub(super) fn text_scale(value: &str) -> Option<u16> {
    let value = collapse(value);
    let digits = value.strip_suffix('%').unwrap_or(value);
    if digits.is_empty() || !digits.bytes().all(|byte| byte.is_ascii_digit()) {
        return None;
    }
    let digits = digits.trim_start_matches('0');
    if digits.len() > 3 {
        return None;
    }
    let scale = if digits.is_empty() {
        0
    } else {
        digits.parse::<u16>().ok()?
    };
    (scale <= 600).then_some(scale)
}

/// The unit a measure is stored in, as a count per point.
#[derive(Clone, Copy)]
pub(super) enum Unit {
    Twips,
    HalfPoints,
    EighthPoints,
    Points,
}

impl Unit {
    const fn per_point(self) -> i128 {
        match self {
            Self::Twips => 20,
            Self::HalfPoints => 2,
            Self::EighthPoints => 8,
            Self::Points => 1,
        }
    }
}

/// Whether a measure may be negative.
#[derive(Clone, Copy, PartialEq, Eq)]
pub(super) enum Sign {
    Signed,
    Unsigned,
}

/// The most significant digits a number may carry. Thirty digits keep every
/// intermediate product inside `i128`.
const MAX_DIGITS: usize = 30;

/// A decimal number as an exact fraction `mantissa / scale`.
struct Decimal {
    negative: bool,
    mantissa: i128,
    scale: i128,
}

/// Parse `[+-]?(digits(.digits?)?|.digits)` (`xsd:decimal`). With `strict`,
/// the stricter universal-measure and percentage grammar applies instead:
/// no `+`, digits before the point, and digits after a point that is
/// written.
fn decimal(number: &str, strict: bool) -> Option<Decimal> {
    let (negative, digits) = match number.as_bytes().first() {
        Some(b'-') => (true, &number[1..]),
        Some(b'+') if !strict => (false, &number[1..]),
        _ => (false, number),
    };
    let (whole, fraction, point) = match digits.split_once('.') {
        Some((whole, fraction)) => (whole, fraction, true),
        None => (digits, "", false),
    };
    if (whole.is_empty() && fraction.is_empty())
        || !whole.bytes().all(|byte| byte.is_ascii_digit())
        || !fraction.bytes().all(|byte| byte.is_ascii_digit())
        || (strict && (whole.is_empty() || (point && fraction.is_empty())))
    {
        return None;
    }
    let whole = whole.trim_start_matches('0');
    let fraction = fraction.trim_end_matches('0');
    if whole.len() + fraction.len() > MAX_DIGITS {
        return None;
    }
    let mut mantissa: i128 = 0;
    for byte in whole.bytes().chain(fraction.bytes()) {
        mantissa = mantissa * 10 + i128::from(byte - b'0');
    }
    let scale = 10_i128.checked_pow(u32::try_from(fraction.len()).ok()?)?;
    Some(Decimal {
        negative,
        mantissa,
        scale,
    })
}

/// `negative * numerator / denominator`, rounded half away from zero, as an
/// `i64`. Both operands are non-negative magnitudes.
fn rounded(negative: bool, numerator: i128, denominator: i128) -> Option<i64> {
    let mut magnitude = numerator.checked_div(denominator)?;
    if numerator.checked_rem(denominator)?.checked_mul(2)? >= denominator {
        magnitude = magnitude.checked_add(1)?;
    }
    i64::try_from(if negative { -magnitude } else { magnitude }).ok()
}

/// A WordprocessingML measure: a number in `unit` (`xsd:integer`, or
/// `xsd:decimal` as the unsigned measures allow), or a universal measure in
/// `mm`, `cm`, `in`, `pt`, `pc` or `pi`. The value is converted exactly and
/// rounded half away from zero to a whole `unit`.
pub(super) fn measure(value: &str, unit: Unit, sign: Sign) -> Option<i64> {
    let value = collapse(value);
    // A universal-measure suffix and the size of its unit in points.
    let suffix = value.len().checked_sub(2).and_then(|split| {
        let (number, suffix) = (value.get(..split)?, value.get(split..)?);
        let points: (i128, i128) = match suffix {
            "mm" => (360, 127),
            "cm" => (3600, 127),
            "in" => (72, 1),
            "pt" => (1, 1),
            "pc" | "pi" => (12, 1),
            _ => return None,
        };
        Some((number, points))
    });
    let result = match suffix {
        Some((number, (points, per))) => {
            let number = decimal(number, true)?;
            rounded(
                number.negative,
                number
                    .mantissa
                    .checked_mul(points)?
                    .checked_mul(unit.per_point())?,
                number.scale.checked_mul(per)?,
            )?
        },
        None => {
            let number = decimal(value, false)?;
            rounded(number.negative, number.mantissa, number.scale)?
        },
    };
    if sign == Sign::Unsigned && result < 0 {
        return None;
    }
    Some(result)
}

/// An `ST_Percentage` (`"50%"`), in fiftieths of a percent: `"50%"` is 2500.
pub(super) fn percent_fiftieths(value: &str) -> Option<i64> {
    let number = decimal(collapse(value).strip_suffix('%')?, true)?;
    rounded(
        number.negative,
        number.mantissa.checked_mul(50)?,
        number.scale,
    )
}

/// A plain `xsd:decimal`, rounded half away from zero to an integer.
pub(super) fn number(value: &str) -> Option<i64> {
    let number = decimal(collapse(value), false)?;
    rounded(number.negative, number.mantissa, number.scale)
}

impl std::fmt::Debug for Fonts {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.debug_struct("Fonts").finish_non_exhaustive()
    }
}

impl std::fmt::Debug for Language {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.debug_struct("Language").finish_non_exhaustive()
    }
}
