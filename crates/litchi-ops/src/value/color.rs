//! Colors in OfficeCLI's vocabulary (`OC/Core/ParseHelpers.cs`,
//! `OC/Core/DrawingColorBuilder.cs`, the Excel and Word color setters).

use std::fmt;
use std::str::FromStr;

use super::error::{ErrorKind, ValueError, What};
use super::number::{DotNetDouble, Styles, parse_f64, parse_i32, parse_i64};
use super::text::{find_folded, starts_with_ci, strip_prefix_ci, strip_suffix_ci};

pub(super) mod names;

use names::{DRAWING_SCHEMES, NAMED, SPREADSHEET_THEMES, SYSTEMS, TRANSFORMS, WORD_SCHEMES};

/// Which color vocabulary a property speaks. The same text can mean
/// different things in different file parts (`accent1+tint40` is a DrawingML
/// tint transform in a slide and a signed SpreadsheetML tint in a cell), so
/// every color parse names its context.
///
/// ```
/// use litchi_ops::value::{Color, ColorContext};
///
/// let slide = Color::parse("accent1+tint40", ColorContext::Drawing)?;
/// let cell = Color::parse("accent1+tint40", ColorContext::Spreadsheet)?;
/// assert_eq!(slide.transforms().len(), 1);
/// assert_eq!(cell.spreadsheet_tint(), Some(0.4));
/// assert!(Color::parse("auto", ColorContext::Word).is_ok());
/// assert!(Color::parse("auto", ColorContext::Drawing).is_err());
/// # Ok::<(), litchi_ops::Error>(())
/// ```
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
#[non_exhaustive]
pub enum ColorContext {
    /// DrawingML colors (`a:srgbClr`, `a:schemeClr`, `a:sysClr` with
    /// transforms): slide shapes and text, chart series, drawing shapes in
    /// workbooks. OfficeCLI `DrawingColorBuilder.BuildColorElement`.
    #[default]
    Drawing,
    /// SpreadsheetML cell colors (ARGB or a theme slot, with an optional
    /// signed `tint`): `font.color`, `fill`, tab colors. OfficeCLI
    /// `ExcelStyleManager`.
    Spreadsheet,
    /// WordprocessingML text colors (`w:color`: hex, `auto`, or a theme
    /// color). OfficeCLI `WordHandler` run `color`.
    Word,
}

/// An sRGB color with alpha (255 is opaque).
///
/// [`Display`](fmt::Display) prints `#RRGGBB`, or `#RRGGBBAA` (CSS order) when
/// the alpha is not 255, as OfficeCLI's `FormatHexColor` prints colors.
///
/// ```
/// use litchi_ops::value::Rgba;
///
/// assert_eq!(Rgba::named("CornflowerBlue"), Some(Rgba::new(0x64, 0x95, 0xED)));
/// assert_eq!(Rgba::preset("dkBlue"), Rgba::named("darkblue"));
/// assert_eq!(Rgba::new(255, 0, 0).with_alpha(0x80).to_string(), "#FF000080");
/// ```
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Rgba {
    red: u8,
    green: u8,
    blue: u8,
    alpha: u8,
}

impl Rgba {
    /// An opaque color.
    #[must_use]
    pub const fn new(red: u8, green: u8, blue: u8) -> Self {
        Self {
            red,
            green,
            blue,
            alpha: 255,
        }
    }

    /// The same color with `alpha` (0 is transparent, 255 opaque).
    #[must_use]
    pub const fn with_alpha(self, alpha: u8) -> Self {
        Self { alpha, ..self }
    }

    /// The red channel.
    #[must_use]
    pub const fn red(self) -> u8 {
        self.red
    }

    /// The green channel.
    #[must_use]
    pub const fn green(self) -> u8 {
        self.green
    }

    /// The blue channel.
    #[must_use]
    pub const fn blue(self) -> u8 {
        self.blue
    }

    /// The alpha channel (255 is opaque).
    #[must_use]
    pub const fn alpha(self) -> u8 {
        self.alpha
    }

    /// `true` when the alpha is 255.
    #[must_use]
    pub const fn is_opaque(self) -> bool {
        self.alpha == 255
    }

    /// The DrawingML `a:alpha` value OfficeCLI writes for this alpha, in
    /// thousandths of a percent (`alpha / 255 × 100000`, truncated), or
    /// `None` when the color is opaque.
    ///
    /// ```
    /// use litchi_ops::value::Rgba;
    ///
    /// assert_eq!(Rgba::new(0, 0, 0).with_alpha(0x80).drawing_alpha(), Some(50_196));
    /// assert_eq!(Rgba::new(0, 0, 0).drawing_alpha(), None);
    /// ```
    #[must_use]
    #[allow(
        clippy::cast_possible_truncation,
        clippy::cast_sign_loss,
        reason = "OfficeCLI truncates a value in 0..=100000"
    )]
    pub fn drawing_alpha(self) -> Option<u32> {
        (self.alpha != 255).then(|| (f64::from(self.alpha) / 255.0 * 100_000.0) as u32)
    }

    /// A named color, ASCII case-insensitively: the 148 CSS named colors,
    /// `transparent` (black with alpha 0), and the 190 DrawingML preset
    /// colors (`ST_PresetColorVal`, including the legacy `dk*`, `lt*` and
    /// `med*` spellings).
    #[must_use]
    pub fn named(name: &str) -> Option<Self> {
        find_folded(NAMED, |named| named.name, name).map(names::Named::rgba)
    }

    /// The color of a DrawingML preset token (`a:prstClr/@val`), ASCII
    /// case-insensitively. CSS names that are not preset tokens
    /// (`rebeccapurple`, `transparent`) are not presets.
    #[must_use]
    pub fn preset(token: &str) -> Option<Self> {
        find_folded(NAMED, |named| named.name, token)
            .filter(|named| named.preset.is_some())
            .map(names::Named::rgba)
    }
}

impl fmt::Display for Rgba {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(
            formatter,
            "#{:02X}{:02X}{:02X}",
            self.red, self.green, self.blue
        )?;
        if self.alpha != 255 {
            write!(formatter, "{:02X}", self.alpha)?;
        }
        Ok(())
    }
}

/// A DrawingML or WordprocessingML theme color: one of the 12 slots of a
/// theme's color scheme, or a slot reached through the color map
/// (`text1`/`tx1`, `background1`/`bg1`), or the placeholder color `phClr`.
///
/// It prints with the long names OfficeCLI's `NormalizeSchemeColorName`
/// emits (`dark1`, `text1`, `hyperlink`), which are also WordprocessingML's
/// `ST_ThemeColor` names.
///
/// ```
/// use litchi_ops::value::Scheme;
///
/// assert_eq!(Scheme::Dark1.name(), "dark1");
/// assert_eq!(Scheme::Dark1.drawing_token(), "dk1");
/// ```
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Scheme {
    /// `bg1`, `background1`: the first background color of the color map.
    Background1,
    /// `tx1`, `text1`: the first text color of the color map.
    Text1,
    /// `bg2`, `background2`.
    Background2,
    /// `tx2`, `text2`.
    Text2,
    /// `accent1`.
    Accent1,
    /// `accent2`.
    Accent2,
    /// `accent3`.
    Accent3,
    /// `accent4`.
    Accent4,
    /// `accent5`.
    Accent5,
    /// `accent6`.
    Accent6,
    /// `hlink`, `hyperlink`.
    Hyperlink,
    /// `folHlink`, `followedHyperlink`.
    FollowedHyperlink,
    /// `phClr`: the color of the style matrix entry being applied.
    Placeholder,
    /// `dk1`, `dark1`.
    Dark1,
    /// `lt1`, `light1`.
    Light1,
    /// `dk2`, `dark2`.
    Dark2,
    /// `lt2`, `light2`.
    Light2,
}

impl Scheme {
    /// The canonical name Litchi and OfficeCLI print.
    #[must_use]
    pub const fn name(self) -> &'static str {
        match self {
            Self::Background1 => "background1",
            Self::Text1 => "text1",
            Self::Background2 => "background2",
            Self::Text2 => "text2",
            Self::Accent1 => "accent1",
            Self::Accent2 => "accent2",
            Self::Accent3 => "accent3",
            Self::Accent4 => "accent4",
            Self::Accent5 => "accent5",
            Self::Accent6 => "accent6",
            Self::Hyperlink => "hyperlink",
            Self::FollowedHyperlink => "followedHyperlink",
            Self::Placeholder => "phClr",
            Self::Dark1 => "dark1",
            Self::Light1 => "light1",
            Self::Dark2 => "dark2",
            Self::Light2 => "light2",
        }
    }

    /// The DrawingML `ST_SchemeColorVal` token (`a:schemeClr/@val`).
    #[must_use]
    pub const fn drawing_token(self) -> &'static str {
        match self {
            Self::Background1 => "bg1",
            Self::Text1 => "tx1",
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
            Self::Placeholder => "phClr",
            Self::Dark1 => "dk1",
            Self::Light1 => "lt1",
            Self::Dark2 => "dk2",
            Self::Light2 => "lt2",
        }
    }
}

/// A SpreadsheetML theme slot (`CT_Color/@theme`, 0 to 11). Excel numbers
/// the slots with light before dark, and prints them with the short names
/// OfficeCLI's `ExcelThemeIndexToName` emits (`lt1`, `dk1`, `hlink`).
///
/// ```
/// use litchi_ops::value::ThemeSlot;
///
/// assert_eq!(ThemeSlot::from_index(4), Some(ThemeSlot::Accent1));
/// assert_eq!(ThemeSlot::Light1.name(), "lt1");
/// assert_eq!(ThemeSlot::Light1.index(), 0);
/// ```
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum ThemeSlot {
    /// Index 0, `lt1`.
    Light1,
    /// Index 1, `dk1`.
    Dark1,
    /// Index 2, `lt2`.
    Light2,
    /// Index 3, `dk2`.
    Dark2,
    /// Index 4, `accent1`.
    Accent1,
    /// Index 5, `accent2`.
    Accent2,
    /// Index 6, `accent3`.
    Accent3,
    /// Index 7, `accent4`.
    Accent4,
    /// Index 8, `accent5`.
    Accent5,
    /// Index 9, `accent6`.
    Accent6,
    /// Index 10, `hlink`.
    Hyperlink,
    /// Index 11, `folHlink`.
    FollowedHyperlink,
}

impl ThemeSlot {
    const ALL: [Self; 12] = [
        Self::Light1,
        Self::Dark1,
        Self::Light2,
        Self::Dark2,
        Self::Accent1,
        Self::Accent2,
        Self::Accent3,
        Self::Accent4,
        Self::Accent5,
        Self::Accent6,
        Self::Hyperlink,
        Self::FollowedHyperlink,
    ];

    /// The slot with SpreadsheetML theme index `index`.
    #[must_use]
    pub const fn from_index(index: u32) -> Option<Self> {
        if index < Self::ALL.len() as u32 {
            Some(Self::ALL[index as usize])
        } else {
            None
        }
    }

    /// The SpreadsheetML theme index.
    #[must_use]
    pub const fn index(self) -> u32 {
        self as u32
    }

    /// The name Litchi and OfficeCLI print.
    #[must_use]
    pub const fn name(self) -> &'static str {
        match self {
            Self::Light1 => "lt1",
            Self::Dark1 => "dk1",
            Self::Light2 => "lt2",
            Self::Dark2 => "dk2",
            Self::Accent1 => "accent1",
            Self::Accent2 => "accent2",
            Self::Accent3 => "accent3",
            Self::Accent4 => "accent4",
            Self::Accent5 => "accent5",
            Self::Accent6 => "accent6",
            Self::Hyperlink => "hlink",
            Self::FollowedHyperlink => "folHlink",
        }
    }
}

/// A DrawingML system color (`a:sysClr/@val`, `ST_SystemColorVal`): a color
/// the host operating system resolves.
///
/// ```
/// use litchi_ops::value::System;
///
/// assert_eq!(System::from_token("WINDOWTEXT"), Some(System::WindowText));
/// assert_eq!(System::WindowText.token(), "windowText");
/// ```
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum System {
    /// `scrollBar`.
    ScrollBar,
    /// `background`.
    Background,
    /// `activeCaption`.
    ActiveCaption,
    /// `inactiveCaption`.
    InactiveCaption,
    /// `menu`.
    Menu,
    /// `window`.
    Window,
    /// `windowFrame`.
    WindowFrame,
    /// `menuText`.
    MenuText,
    /// `windowText`.
    WindowText,
    /// `captionText`.
    CaptionText,
    /// `activeBorder`.
    ActiveBorder,
    /// `inactiveBorder`.
    InactiveBorder,
    /// `appWorkspace`.
    AppWorkspace,
    /// `highlight`.
    Highlight,
    /// `highlightText`.
    HighlightText,
    /// `btnFace`.
    ButtonFace,
    /// `btnShadow`.
    ButtonShadow,
    /// `grayText`.
    GrayText,
    /// `btnText`.
    ButtonText,
    /// `inactiveCaptionText`.
    InactiveCaptionText,
    /// `btnHighlight`.
    ButtonHighlight,
    /// `3dDkShadow`.
    DarkShadow3d,
    /// `3dLight`.
    Light3d,
    /// `infoText`.
    InfoText,
    /// `infoBk`.
    InfoBackground,
    /// `hotLight`.
    HotLight,
    /// `gradientActiveCaption`.
    GradientActiveCaption,
    /// `gradientInactiveCaption`.
    GradientInactiveCaption,
    /// `menuHighlight`.
    MenuHighlight,
    /// `menuBar`.
    MenuBar,
}

impl System {
    /// The `ST_SystemColorVal` token, which is also the printed form.
    #[must_use]
    pub const fn token(self) -> &'static str {
        match self {
            Self::ScrollBar => "scrollBar",
            Self::Background => "background",
            Self::ActiveCaption => "activeCaption",
            Self::InactiveCaption => "inactiveCaption",
            Self::Menu => "menu",
            Self::Window => "window",
            Self::WindowFrame => "windowFrame",
            Self::MenuText => "menuText",
            Self::WindowText => "windowText",
            Self::CaptionText => "captionText",
            Self::ActiveBorder => "activeBorder",
            Self::InactiveBorder => "inactiveBorder",
            Self::AppWorkspace => "appWorkspace",
            Self::Highlight => "highlight",
            Self::HighlightText => "highlightText",
            Self::ButtonFace => "btnFace",
            Self::ButtonShadow => "btnShadow",
            Self::GrayText => "grayText",
            Self::ButtonText => "btnText",
            Self::InactiveCaptionText => "inactiveCaptionText",
            Self::ButtonHighlight => "btnHighlight",
            Self::DarkShadow3d => "3dDkShadow",
            Self::Light3d => "3dLight",
            Self::InfoText => "infoText",
            Self::InfoBackground => "infoBk",
            Self::HotLight => "hotLight",
            Self::GradientActiveCaption => "gradientActiveCaption",
            Self::GradientInactiveCaption => "gradientInactiveCaption",
            Self::MenuHighlight => "menuHighlight",
            Self::MenuBar => "menuBar",
        }
    }

    /// The system color named `token`, ASCII case-insensitively.
    #[must_use]
    pub fn from_token(token: &str) -> Option<Self> {
        find_folded(SYSTEMS, |(name, _)| name, token).map(|&(_, system)| system)
    }

    /// The `lastClr` OfficeCLI writes with a system color it builds, for
    /// the nine system colors it knows (`window`, `windowText`,
    /// `background`, `windowFrame`, `highlight`, `highlightText`, `btnFace`,
    /// `btnText`, `grayText`).
    #[must_use]
    pub const fn default_last_color(self) -> Option<Rgba> {
        match self {
            Self::Window | Self::Background | Self::HighlightText => {
                Some(Rgba::new(0xFF, 0xFF, 0xFF))
            },
            Self::WindowText | Self::WindowFrame | Self::ButtonText => Some(Rgba::new(0, 0, 0)),
            Self::Highlight => Some(Rgba::new(0x00, 0x78, 0xD7)),
            Self::ButtonFace => Some(Rgba::new(0xF0, 0xF0, 0xF0)),
            Self::GrayText => Some(Rgba::new(0x6D, 0x6D, 0x6D)),
            _ => None,
        }
    }
}

/// A DrawingML color transform OfficeCLI's vocabulary names.
///
/// ```
/// use litchi_ops::value::TransformKind;
///
/// assert_eq!(TransformKind::LumMod.name(), "lumMod");
/// assert_eq!(TransformKind::Shade.range(), (0, 100_000));
/// ```
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum TransformKind {
    /// `lumMod`: luminance modulation (signed, up to ±1000%).
    LumMod,
    /// `lumOff`: luminance offset.
    LumOff,
    /// `satMod`: saturation modulation.
    SatMod,
    /// `satOff`: saturation offset.
    SatOff,
    /// `hueMod`: hue modulation.
    HueMod,
    /// `hueOff`: hue offset.
    HueOff,
    /// `shade`: darkening toward black (0% to 100%).
    Shade,
    /// `tint`: lightening toward white (0% to 100%).
    Tint,
    /// `alpha`: opacity (0% to 100%).
    Alpha,
}

impl TransformKind {
    /// The element name (`lumMod`), which is also the printed name.
    #[must_use]
    pub const fn name(self) -> &'static str {
        match self {
            Self::LumMod => "lumMod",
            Self::LumOff => "lumOff",
            Self::SatMod => "satMod",
            Self::SatOff => "satOff",
            Self::HueMod => "hueMod",
            Self::HueOff => "hueOff",
            Self::Shade => "shade",
            Self::Tint => "tint",
            Self::Alpha => "alpha",
        }
    }

    /// The accepted range of the value in thousandths of a percent:
    /// 0 to 100 000 for `shade`, `tint` and `alpha`
    /// (`ST_PositiveFixedPercentage`), -1 000 000 to 1 000 000 for the
    /// modulations and offsets (`ST_Percentage`, bounded as OfficeCLI bounds
    /// it).
    #[must_use]
    pub const fn range(self) -> (i32, i32) {
        match self {
            Self::Shade | Self::Tint | Self::Alpha => (0, 100_000),
            _ => (-1_000_000, 1_000_000),
        }
    }
}

/// The canonical transform names, for messages and suggestions.
pub(super) const TRANSFORM_NAMES: [&str; 9] = [
    "lumMod", "lumOff", "satMod", "satOff", "hueMod", "hueOff", "shade", "tint", "alpha",
];

/// The percentage and raw ranges of the transform `name` as message text.
pub(super) fn transform_range_text(name: &str) -> (&'static str, &'static str) {
    match find_folded(TRANSFORMS, |(spelling, _)| spelling, name).map(|&(_, kind)| kind) {
        Some(TransformKind::Shade | TransformKind::Tint | TransformKind::Alpha) => {
            ("0 to 100", "0 to 100000")
        },
        _ => ("-1000 to 1000", "-1000000 to 1000000"),
    }
}

/// One color transform: a kind and a value in thousandths of a percent
/// (OOXML's `val` attribute; `75000` is 75%).
///
/// It prints as OfficeCLI's suffix: `lumMod75` for whole percentages, and
/// `lumMod=75500` (the raw form OfficeCLI also reads) otherwise.
///
/// ```
/// use litchi_ops::value::{Transform, TransformKind};
///
/// let darker = Transform::new(TransformKind::LumMod, 75_000).expect("in range");
/// assert_eq!(darker.to_string(), "lumMod75");
/// assert_eq!(Transform::new(TransformKind::LumMod, 75_500).expect("in range").to_string(), "lumMod=75500");
/// assert!(Transform::new(TransformKind::Shade, -1).is_none());
/// ```
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Transform {
    kind: TransformKind,
    value: i32,
}

impl Transform {
    /// A transform, or `None` when `value` is outside the kind's
    /// [`range`](TransformKind::range).
    #[must_use]
    pub const fn new(kind: TransformKind, value: i32) -> Option<Self> {
        let (low, high) = kind.range();
        if value >= low && value <= high {
            Some(Self { kind, value })
        } else {
            None
        }
    }

    /// The kind.
    #[must_use]
    pub const fn kind(self) -> TransformKind {
        self.kind
    }

    /// The value in thousandths of a percent.
    #[must_use]
    pub const fn value(self) -> i32 {
        self.value
    }
}

impl fmt::Display for Transform {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        let name = self.kind.name();
        if self.value % 1000 == 0 {
            write!(formatter, "{name}{}", self.value / 1000)
        } else {
            write!(formatter, "{name}={}", self.value)
        }
    }
}

/// What a color is before its transforms or tint.
///
/// ```
/// use litchi_ops::value::{Base, Color, ColorContext, Scheme};
///
/// let color = Color::parse("accent2+lumMod50", ColorContext::Drawing)?;
/// assert_eq!(color.base(), Base::Scheme(Scheme::Accent2));
/// # Ok::<(), litchi_ops::Error>(())
/// ```
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Base {
    /// An sRGB color with alpha.
    Rgb(Rgba),
    /// A DrawingML or WordprocessingML theme color.
    Scheme(Scheme),
    /// A SpreadsheetML theme slot.
    Theme(ThemeSlot),
    /// A DrawingML system color.
    System(System),
    /// WordprocessingML's automatic text color (`w:color/@val="auto"`).
    Auto,
}

/// A color: a [`Base`], up to [`Color::MAX_TRANSFORMS`] DrawingML transforms,
/// and, for spreadsheet cells, a signed tint.
///
/// [`Color::parse`] reads every form OfficeCLI accepts in a
/// [`ColorContext`]: hex with 3, 4 (`#RGBA`), 6 or 8 digits (`#RRGGBBAA`
/// with a `#`, `AARRGGBB` without), the CSS named colors and `transparent`,
/// the DrawingML preset colors, `rgb()`/`rgba()`/`hsl()`/`hsla()` (comma or
/// CSS level 4 space syntax), theme names, `+transform` suffixes
/// (`accent1+lumMod75+lumOff25`, `lumMod=75000`, or `accent1:lumMod75`) in
/// DrawingML, and `+tint` suffixes (`accent1+tint40`, `dk2+tint-25`) in
/// spreadsheet cells. Matching ignores ASCII case and surrounding white
/// space; nothing is allocated.
///
/// [`Display`](fmt::Display) prints the canonical form, which parses back to
/// an equal color in the same context. Theme colors stay references with
/// their transforms; they are never resolved to RGB.
///
/// ```
/// use litchi_ops::value::{Color, ColorContext, Rgba};
///
/// let red = Color::parse("rgb(255, 0, 0)", ColorContext::Drawing)?;
/// assert_eq!(red, Color::rgb(255, 0, 0));
/// assert_eq!(red.to_string(), "#FF0000");
///
/// let darker = Color::parse("Accent1 + lumMod75", ColorContext::Drawing)?;
/// assert_eq!(darker.to_string(), "accent1+lumMod75");
///
/// let tinted = Color::parse("bg1+tint=-25", ColorContext::Spreadsheet)?;
/// assert_eq!(tinted.to_string(), "lt1+tint-25");
/// assert_eq!(tinted.spreadsheet_tint(), Some(-0.25));
///
/// assert_eq!(Color::parse("#80FF0000", ColorContext::Drawing)?.base(),
///     litchi_ops::value::Base::Rgb(Rgba::new(0x80, 0xFF, 0x00).with_alpha(0x00)));
/// # Ok::<(), litchi_ops::Error>(())
/// ```
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Color {
    base: Base,
    transforms: [Transform; Color::MAX_TRANSFORMS],
    len: u8,
    /// The spreadsheet tint in billionths (`-1e9..=1e9`); 0 is no tint.
    tint: i32,
}

/// The filler of unused transform slots. Unused slots always hold it, so the
/// derived equality and hash see only the used transforms.
const NO_TRANSFORM: Transform = Transform {
    kind: TransformKind::LumMod,
    value: 0,
};

/// OfficeCLI's `ExcelTintEpsilon`: tints closer than this are the same.
const TINT_EPSILON: f64 = 0.005;

impl Color {
    /// The most transforms one color holds. OOXML allows longer chains; no
    /// Office application writes more than four.
    pub const MAX_TRANSFORMS: usize = 8;

    /// WordprocessingML's automatic color.
    pub const AUTO: Self = Self::new(Base::Auto);

    /// A color with `base` and no transform or tint.
    #[must_use]
    pub const fn new(base: Base) -> Self {
        Self {
            base,
            transforms: [NO_TRANSFORM; Self::MAX_TRANSFORMS],
            len: 0,
            tint: 0,
        }
    }

    /// An opaque sRGB color.
    #[must_use]
    pub const fn rgb(red: u8, green: u8, blue: u8) -> Self {
        Self::new(Base::Rgb(Rgba::new(red, green, blue)))
    }

    /// An sRGB color with alpha (255 is opaque).
    #[must_use]
    pub const fn rgba(red: u8, green: u8, blue: u8, alpha: u8) -> Self {
        Self::new(Base::Rgb(Rgba::new(red, green, blue).with_alpha(alpha)))
    }

    /// A DrawingML or WordprocessingML theme color.
    #[must_use]
    pub const fn scheme(scheme: Scheme) -> Self {
        Self::new(Base::Scheme(scheme))
    }

    /// A SpreadsheetML theme slot.
    #[must_use]
    pub const fn theme(slot: ThemeSlot) -> Self {
        Self::new(Base::Theme(slot))
    }

    /// A DrawingML system color.
    #[must_use]
    pub const fn system(system: System) -> Self {
        Self::new(Base::System(system))
    }

    /// The base color.
    #[must_use]
    pub const fn base(&self) -> Base {
        self.base
    }

    /// The transforms, in order.
    #[must_use]
    pub fn transforms(&self) -> &[Transform] {
        &self.transforms[..usize::from(self.len)]
    }

    /// The color with `transform` appended, or `None` when it already holds
    /// [`Color::MAX_TRANSFORMS`].
    #[must_use]
    pub fn with_transform(mut self, transform: Transform) -> Option<Self> {
        let slot = self.transforms.get_mut(usize::from(self.len))?;
        *slot = transform;
        self.len += 1;
        Some(self)
    }

    /// The color with `transforms` appended in order, or `None` when they
    /// would exceed [`Color::MAX_TRANSFORMS`]. This is how a reader that holds
    /// a DrawingML transform list builds the color; nothing is allocated.
    ///
    /// ```
    /// use litchi_ops::value::{Color, Scheme, Transform, TransformKind};
    ///
    /// let lum = [
    ///     Transform::new(TransformKind::LumMod, 75_000).expect("in range"),
    ///     Transform::new(TransformKind::LumOff, 25_000).expect("in range"),
    /// ];
    /// let color = Color::scheme(Scheme::Accent1).with_transforms(&lum).expect("two fit");
    /// assert_eq!(color.to_string(), "accent1+lumMod75+lumOff25");
    /// assert_eq!(Color::scheme(Scheme::Accent1).with_transforms(&[lum[0]; 9]), None);
    /// ```
    #[must_use]
    pub fn with_transforms(self, transforms: &[Transform]) -> Option<Self> {
        transforms
            .iter()
            .try_fold(self, |color, &transform| color.with_transform(transform))
    }

    /// The SpreadsheetML tint (`CT_Color/@tint`, -1.0 to 1.0; negative
    /// darkens), when there is one. Tints are kept to 1e-9.
    #[must_use]
    pub fn spreadsheet_tint(&self) -> Option<f64> {
        (self.tint != 0).then(|| f64::from(self.tint) / 1e9)
    }

    /// The color with a SpreadsheetML tint, or `None` unless `tint` is a
    /// finite value from -1.0 to 1.0. A zero tint removes it.
    #[must_use]
    #[allow(
        clippy::cast_possible_truncation,
        reason = "the scaled tint is rounded and within ±1e9"
    )]
    pub fn with_spreadsheet_tint(mut self, tint: f64) -> Option<Self> {
        if !(-1.0..=1.0).contains(&tint) {
            return None;
        }
        self.tint = (tint * 1e9).round() as i32;
        Some(self)
    }

    /// Semantic equality: the same base and transforms, and tints within
    /// OfficeCLI's tolerance of 0.005 (`ExcelTintMatches`). Query
    /// equality on color properties uses this.
    #[must_use]
    pub fn matches(&self, other: &Self) -> bool {
        self.base == other.base
            && self.transforms() == other.transforms()
            && (self.spreadsheet_tint().unwrap_or(0.0) - other.spreadsheet_tint().unwrap_or(0.0))
                .abs()
                < TINT_EPSILON
    }

    /// Parses `text` in `context`.
    ///
    /// # Errors
    ///
    /// Returns an `invalid_value` [`ValueError`] when the text is not a color
    /// of the context. Its suggestion names the nearest color or theme name.
    pub fn parse(text: &str, context: ColorContext) -> Result<Self, ValueError<'_>> {
        let what = What::Color(context);
        let trimmed = text.trim();
        if trimmed.is_empty() {
            return Err(ValueError::new(what, ErrorKind::Empty, text));
        }
        let parsed = match context {
            ColorContext::Drawing => parse_drawing(trimmed),
            ColorContext::Spreadsheet => parse_spreadsheet(trimmed),
            ColorContext::Word => parse_word(trimmed),
        };
        parsed.map_err(|(kind, part)| ValueError::new(what, kind, text).at(part))
    }
}

impl fmt::Display for Color {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self.base {
            Base::Rgb(rgba) => rgba.fmt(formatter)?,
            Base::Scheme(scheme) => formatter.write_str(scheme.name())?,
            Base::Theme(slot) => formatter.write_str(slot.name())?,
            Base::System(system) => formatter.write_str(system.token())?,
            Base::Auto => formatter.write_str("auto")?,
        }
        for transform in self.transforms() {
            write!(formatter, "+{transform}")?;
        }
        if let Some(tint) = self.spreadsheet_tint() {
            // OfficeCLI `ExcelThemeNameWithTint`: omitted below the tolerance,
            // else the percentage rounded to two decimals (ties to even).
            if tint.abs() >= TINT_EPSILON {
                let percent = (tint * 100.0 * 100.0).round_ties_even() / 100.0;
                write!(formatter, "+tint{}", DotNetDouble(percent))?;
            }
        }
        Ok(())
    }
}

impl FromStr for Color {
    type Err = crate::Error;

    /// Parses a DrawingML color ([`ColorContext::Drawing`]).
    fn from_str(text: &str) -> Result<Self, Self::Err> {
        Self::parse(text, ColorContext::Drawing).map_err(crate::Error::from)
    }
}

/// A rejected part of a color: why, and which part of the text.
type Fault<'a> = (ErrorKind, &'a str);

/// OfficeCLI `DrawingColorBuilder.BuildColorElement`: a theme, system or RGB
/// base, then `+` transforms. `:` also separates a single transform after a
/// theme name (`accent1:lumMod75`).
fn parse_drawing(text: &str) -> Result<Color, Fault<'_>> {
    let split = match text.find('+') {
        Some(index) if index > 0 => Some(index),
        _ => text
            .find(':')
            .filter(|&index| index > 0 && drawing_scheme(text[..index].trim()).is_some()),
    };
    let (base_text, chain) = match split {
        Some(index) => (text[..index].trim(), Some(&text[index + 1..])),
        None => (text, None),
    };
    let base = if let Some(scheme) = drawing_scheme(base_text) {
        Base::Scheme(scheme)
    } else if let Some(system) = System::from_token(base_text) {
        Base::System(system)
    } else if base_text.eq_ignore_ascii_case("auto") {
        return Err((ErrorKind::Auto, base_text));
    } else {
        Base::Rgb(parse_value(base_text)?)
    };
    let mut color = Color::new(base);
    if let Some(chain) = chain {
        for token in chain
            .split('+')
            .map(str::trim)
            .filter(|token| !token.is_empty())
        {
            let transform = parse_transform(token)?;
            color = color
                .with_transform(transform)
                .ok_or((ErrorKind::TooManyTransforms, token))?;
        }
    }
    Ok(color)
}

/// A DrawingML theme name; leading `#`s are ignored, as OfficeCLI's
/// `TryParseSchemeColor` ignores them.
fn drawing_scheme(text: &str) -> Option<Scheme> {
    let name = text.trim_start_matches('#');
    find_folded(DRAWING_SCHEMES, |(spelling, _)| spelling, name).map(|&(_, scheme)| scheme)
}

/// One `name<percent>` or `name=<raw>` transform token.
fn parse_transform(token: &str) -> Result<Transform, Fault<'_>> {
    let name_end = token
        .bytes()
        .position(|byte| !byte.is_ascii_alphabetic())
        .unwrap_or(token.len());
    if name_end == 0 {
        return Err((ErrorKind::TransformName, token));
    }
    if name_end == token.len() {
        return Err((ErrorKind::TransformValue, token));
    }
    let kind = find_folded(TRANSFORMS, |(spelling, _)| spelling, &token[..name_end])
        .map(|&(_, kind)| kind)
        .ok_or((ErrorKind::UnknownTransform, token))?;
    let rest = &token[name_end..];
    let (raw, number) = match rest.strip_prefix('=') {
        Some(number) => (true, number),
        None => (false, rest),
    };
    let number = parse_i32(number).ok_or((ErrorKind::TransformNumber, token))?;
    let value = if raw {
        number
    } else {
        number
            .checked_mul(1000)
            .ok_or((ErrorKind::TransformRange, token))?
    };
    Transform::new(kind, value).ok_or((ErrorKind::TransformRange, token))
}

/// OfficeCLI's Excel color setters: a theme slot or an ARGB value, with an
/// optional `+tint<percent>` suffix (`ParseHelpers.SplitExcelColorTint`).
fn parse_spreadsheet(text: &str) -> Result<Color, Fault<'_>> {
    let plus = text.rfind('+').filter(|&index| index > 0);
    let (base_text, tint) = match plus {
        Some(index) if starts_with_ci(text[index + 1..].trim_start(), "tint") => {
            let suffix = text[index + 1..].trim_start();
            let percent = parse_f64(suffix[4..].trim_start_matches('='), Styles::FLOAT)
                .ok_or((ErrorKind::Tint, suffix))?;
            if !(-100.0..=100.0).contains(&percent) {
                return Err((ErrorKind::TintRange, suffix));
            }
            (text[..index].trim(), percent / 100.0)
        },
        // Any other `+` stays in the base, as in OfficeCLI (`rgb(+5,0,0)`).
        _ => (text, 0.0),
    };
    let base = match find_folded(SPREADSHEET_THEMES, |(spelling, _)| spelling, base_text) {
        Some(&(_, slot)) => Base::Theme(slot),
        None => Base::Rgb(parse_value(base_text).map_err(|fault| {
            transform_suffix(text, |head| spreadsheet_base(head).is_some()).unwrap_or(fault)
        })?),
    };
    Color::new(base)
        .with_spreadsheet_tint(tint)
        .ok_or((ErrorKind::TintRange, text))
}

/// A spreadsheet base color: a theme slot or an RGB value.
fn spreadsheet_base(text: &str) -> Option<Base> {
    match find_folded(SPREADSHEET_THEMES, |(spelling, _)| spelling, text) {
        Some(&(_, slot)) => Some(Base::Theme(slot)),
        None => parse_value(text).ok().map(Base::Rgb),
    }
}

/// When `text` is a valid base followed by a `+` suffix the context does not
/// take (`accent1+lumMod75` in a cell), the [`ErrorKind::NoTransforms`]
/// fault naming that suffix.
fn transform_suffix<'a>(text: &'a str, is_base: impl Fn(&str) -> bool) -> Option<Fault<'a>> {
    let index = text.find('+').filter(|&index| index > 0)?;
    is_base(text[..index].trim()).then(|| (ErrorKind::NoTransforms, &text[index + 1..]))
}

/// OfficeCLI's Word run `color`: `auto`, a theme name (not `#`-prefixed),
/// or an RGB value.
fn parse_word(text: &str) -> Result<Color, Fault<'_>> {
    if text.contains(';') {
        return Err((ErrorKind::ThemeTail, text));
    }
    word_base(text).map(Color::new).ok_or_else(|| {
        transform_suffix(text, |head| word_base(head).is_some())
            .unwrap_or_else(|| parse_value(text).err().unwrap_or((ErrorKind::Syntax, text)))
    })
}

/// A Word base color: `auto`, a theme name, or an RGB value.
fn word_base(text: &str) -> Option<Base> {
    if text.eq_ignore_ascii_case("auto") {
        return Some(Base::Auto);
    }
    if !text.starts_with('#') {
        if let Some(&(_, scheme)) = find_folded(WORD_SCHEMES, |(spelling, _)| spelling, text) {
            return Some(Base::Scheme(scheme));
        }
    }
    parse_value(text).ok().map(Base::Rgb)
}

/// OfficeCLI `TryResolveColorInput` then the hex forms of
/// `NormalizeArgbColor`/`SanitizeColorForOoxml`: a named color, an `rgb()`
/// or `hsl()` function, or hex.
fn parse_value(text: &str) -> Result<Rgba, Fault<'_>> {
    if let Some(named) = find_folded(NAMED, |named| named.name, text) {
        return Ok(named.rgba());
    }
    if let Some(function) = parse_function(text) {
        return function;
    }
    parse_hex(text).ok_or((ErrorKind::Syntax, text))
}

/// Hex: 3 digits (`F00`), 4 digits after `#` with a non-zero alpha digit
/// (`#F00A`, CSS `#RGBA`), 6 digits, or 8 digits: `#RRGGBBAA` (CSS order)
/// with a `#`, `AARRGGBB` (OOXML order) without one.
fn parse_hex(text: &str) -> Option<Rgba> {
    let hashed = text.starts_with('#');
    let hex = text.trim_start_matches('#').as_bytes();
    if !hex.iter().all(u8::is_ascii_hexdigit) {
        return None;
    }
    let digit = |index: usize| hex_digit(hex[index]);
    let pair = |index: usize| (hex_digit(hex[index]) << 4) | hex_digit(hex[index + 1]);
    match hex.len() {
        3 => Some(Rgba::new(digit(0) * 17, digit(1) * 17, digit(2) * 17)),
        4 if hashed && hex[3] != b'0' => {
            Some(Rgba::new(digit(0) * 17, digit(1) * 17, digit(2) * 17).with_alpha(digit(3) * 17))
        },
        6 => Some(Rgba::new(pair(0), pair(2), pair(4))),
        8 if hashed => Some(Rgba::new(pair(0), pair(2), pair(4)).with_alpha(pair(6))),
        8 => Some(Rgba::new(pair(2), pair(4), pair(6)).with_alpha(pair(0))),
        _ => None,
    }
}

const fn hex_digit(byte: u8) -> u8 {
    match byte {
        b'0'..=b'9' => byte - b'0',
        b'a'..=b'f' => byte - b'a' + 10,
        b'A'..=b'F' => byte - b'A' + 10,
        _ => 0,
    }
}

/// `rgb(…)`, `rgba(…)`, `hsl(…)` or `hsla(…)`. `None` when the text is not
/// shaped like a function, so the caller tries hex next (OfficeCLI's regex
/// fails and falls through the same way).
fn parse_function(text: &str) -> Option<Result<Rgba, Fault<'_>>> {
    let (hsl, rest) = if let Some(rest) =
        strip_prefix_ci(text, "rgba(").or_else(|| strip_prefix_ci(text, "rgb("))
    {
        (false, rest)
    } else if let Some(rest) =
        strip_prefix_ci(text, "hsla(").or_else(|| strip_prefix_ci(text, "hsl("))
    {
        (true, rest)
    } else {
        return None;
    };
    let (components, alpha) = split_arguments(rest.strip_suffix(')')?)?;
    Some(if hsl {
        hsl_color(components, alpha)
    } else {
        rgb_color(components, alpha)
    })
}

/// Splits `c1 c2 c3 [/ a]` or `c1, c2, c3[, a]` the way OfficeCLI's pattern
/// `\s*C\s*[,\s]\s*C\s*[,\s]\s*C\s*(?:[,/]\s*C\s*)?` does, where a component
/// `C` is a run without white space, `,` or `/`.
fn split_arguments(body: &str) -> Option<([&str; 3], Option<&str>)> {
    fn component<'a>(rest: &mut &'a str) -> Option<&'a str> {
        let end = rest
            .find(|c: char| c.is_whitespace() || c == ',' || c == '/')
            .unwrap_or(rest.len());
        if end == 0 {
            return None;
        }
        let (taken, remaining) = rest.split_at(end);
        *rest = remaining;
        Some(taken)
    }
    fn separator(rest: &mut &str) -> Option<()> {
        let mut remaining = rest.trim_start();
        if let Some(after) = remaining.strip_prefix(',') {
            remaining = after.trim_start();
        }
        if remaining.len() == rest.len() {
            return None;
        }
        *rest = remaining;
        Some(())
    }
    let mut rest = body.trim_start();
    let first = component(&mut rest)?;
    separator(&mut rest)?;
    let second = component(&mut rest)?;
    separator(&mut rest)?;
    let third = component(&mut rest)?;
    rest = rest.trim_start();
    let alpha = match rest.strip_prefix([',', '/']) {
        Some(after) => {
            let mut after = after.trim_start();
            let alpha = component(&mut after)?;
            rest = after.trim_start();
            Some(alpha)
        },
        None => None,
    };
    rest.is_empty().then_some(([first, second, third], alpha))
}

/// Rounds a channel to a byte, ties to even (.NET `Math.Round`).
#[allow(
    clippy::cast_possible_truncation,
    clippy::cast_sign_loss,
    reason = "the value is rounded and clamped to 0..=255 first"
)]
fn to_byte(value: f64) -> u8 {
    value.round_ties_even().clamp(0.0, 255.0) as u8
}

fn rgb_color<'a>(components: [&'a str; 3], alpha: Option<&'a str>) -> Result<Rgba, Fault<'a>> {
    let [red, green, blue] = components;
    let color = Rgba::new(rgb_channel(red)?, rgb_channel(green)?, rgb_channel(blue)?);
    Ok(color.with_alpha(alpha.map_or(Ok(255), alpha_channel)?))
}

/// OfficeCLI `ParseRgbComponent`: 0-255, or 0%-100%.
fn rgb_channel(token: &str) -> Result<u8, Fault<'_>> {
    let channel = match token.strip_suffix('%') {
        Some(percent) => parse_f64(percent, Styles::FLOAT)
            .filter(|percent| (0.0..=100.0).contains(percent))
            .map(|percent| to_byte(percent / 100.0 * 255.0)),
        None => parse_i64(token, Styles::INTEGER).and_then(|value| u8::try_from(value).ok()),
    };
    channel.ok_or((ErrorKind::Component, token))
}

/// OfficeCLI `ParseAlphaComponent`: 0-1, or 0%-100%.
fn alpha_channel(token: &str) -> Result<u8, Fault<'_>> {
    let alpha = match token.strip_suffix('%') {
        Some(percent) => parse_f64(percent, Styles::FLOAT)
            .filter(|percent| (0.0..=100.0).contains(percent))
            .map(|percent| percent / 100.0),
        None => parse_f64(token, Styles::FLOAT).filter(|alpha| (0.0..=1.0).contains(alpha)),
    };
    alpha
        .map(|alpha| to_byte(alpha * 255.0))
        .ok_or((ErrorKind::Alpha, token))
}

fn hsl_color<'a>(components: [&'a str; 3], alpha: Option<&'a str>) -> Result<Rgba, Fault<'a>> {
    let [hue, saturation, lightness] = components;
    let hue = hue_degrees(hue)?;
    let saturation = hsl_fraction(saturation)?;
    let lightness = hsl_fraction(lightness)?;
    let color = hsl_to_rgb(hue, saturation, lightness);
    Ok(color.with_alpha(alpha.map_or(Ok(255), alpha_channel)?))
}

/// OfficeCLI `ParseHueDegrees`: a number with an optional `deg` suffix,
/// reduced to 0-360.
fn hue_degrees(token: &str) -> Result<f64, Fault<'_>> {
    let number = strip_suffix_ci(token, "deg").map_or(token, str::trim);
    let degrees = parse_f64(number, Styles::FLOAT)
        .filter(|degrees| degrees.is_finite())
        .ok_or((ErrorKind::Hue, token))?;
    let reduced = degrees % 360.0;
    Ok(if reduced < 0.0 {
        reduced + 360.0
    } else {
        reduced
    })
}

/// OfficeCLI `ParsePercent01`: a percentage from 0% to 100%, as 0-1.
fn hsl_fraction(token: &str) -> Result<f64, Fault<'_>> {
    token
        .strip_suffix('%')
        .and_then(|percent| parse_f64(percent, Styles::FLOAT))
        .filter(|percent| (0.0..=100.0).contains(percent))
        .map(|percent| percent / 100.0)
        .ok_or((ErrorKind::HslPercent, token))
}

/// CSS Color Level 3 HSL to RGB, in OfficeCLI's `HslToRgb` operation order.
fn hsl_to_rgb(hue: f64, saturation: f64, lightness: f64) -> Rgba {
    let chroma = (1.0 - (2.0 * lightness - 1.0).abs()) * saturation;
    let sector = hue / 60.0;
    let x = chroma * (1.0 - (sector % 2.0 - 1.0).abs());
    let (red, green, blue) = if sector < 1.0 {
        (chroma, x, 0.0)
    } else if sector < 2.0 {
        (x, chroma, 0.0)
    } else if sector < 3.0 {
        (0.0, chroma, x)
    } else if sector < 4.0 {
        (0.0, x, chroma)
    } else if sector < 5.0 {
        (x, 0.0, chroma)
    } else {
        (chroma, 0.0, x)
    };
    let m = lightness - chroma / 2.0;
    Rgba::new(
        to_byte((red + m) * 255.0),
        to_byte((green + m) * 255.0),
        to_byte((blue + m) * 255.0),
    )
}

/// A hint for a theme or system color name, or `none`, that `context` does
/// not accept although another context does.
pub(super) fn context_hint(context: ColorContext, text: &str) -> Option<&'static str> {
    let text = text.trim();
    if text.eq_ignore_ascii_case("none") {
        return Some("'none' is not a color; a property that can be cleared accepts 'none' itself");
    }
    let drawing = drawing_scheme(text).is_some() || System::from_token(text).is_some();
    let word = !text.starts_with('#')
        && find_folded(WORD_SCHEMES, |(spelling, _)| spelling, text).is_some();
    let sheet = find_folded(SPREADSHEET_THEMES, |(spelling, _)| spelling, text).is_some();
    match context {
        ColorContext::Spreadsheet if drawing || word => Some(
            "cell colors take the theme slots lt1, dk1, lt2, dk2, accent1 to accent6, hlink and folHlink",
        ),
        ColorContext::Word if drawing || sheet => Some(
            "Word text colors take auto, a hex or named color, or a theme color such as dark1, text1 or accent1 (without '#')",
        ),
        ColorContext::Drawing if word => Some(
            "DrawingML colors take scheme names (accent1 to accent6, dk1/dark1, lt1/light1, tx1/text1, bg1/background1, hlink, folHlink, phClr) or system colors such as windowText",
        ),
        _ => None,
    }
}

/// Every color and theme name `context` accepts, with the canonical name it
/// means, for "did you mean" suggestions.
pub(super) fn names(context: ColorContext) -> impl Iterator<Item = (&'static str, &'static str)> {
    let (schemes, themes, systems): (&[(&str, Scheme)], &[(&str, ThemeSlot)], &[(&str, System)]) =
        match context {
            ColorContext::Drawing => (DRAWING_SCHEMES, &[], SYSTEMS),
            ColorContext::Spreadsheet => (&[], SPREADSHEET_THEMES, &[]),
            ColorContext::Word => (WORD_SCHEMES, &[], &[]),
        };
    NAMED
        .iter()
        .map(|named| (named.name, named.name))
        .chain(
            schemes
                .iter()
                .map(|&(spelling, scheme)| (spelling, scheme.name())),
        )
        .chain(
            themes
                .iter()
                .map(|&(spelling, slot)| (spelling, slot.name())),
        )
        .chain(
            systems
                .iter()
                .map(|&(spelling, system)| (spelling, system.token())),
        )
}
