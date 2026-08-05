//! Contextual theme values.

use crate::{Error, Result};

const MAX_NAME_CHARS: usize = 256;
const MAX_SCRIPT_COUNT: usize = 64;
const MAX_SCRIPT_CHARS: usize = 16;

/// One of the twelve DrawingML color slots.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Slot {
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

impl Slot {
    pub const ALL: [Self; 12] = [
        Self::Dark1,
        Self::Light1,
        Self::Dark2,
        Self::Light2,
        Self::Accent1,
        Self::Accent2,
        Self::Accent3,
        Self::Accent4,
        Self::Accent5,
        Self::Accent6,
        Self::Hyperlink,
        Self::FollowedHyperlink,
    ];

    pub const fn token(self) -> &'static str {
        match self {
            Self::Dark1 => "dk1",
            Self::Light1 => "lt1",
            Self::Dark2 => "dk2",
            Self::Light2 => "lt2",
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

    pub fn from_token(token: &str) -> Option<Self> {
        Self::ALL.into_iter().find(|slot| slot.token() == token)
    }
}

/// A DrawingML system-color token.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum System {
    ActiveBorder,
    ActiveCaption,
    AppWorkspace,
    Background,
    ButtonFace,
    ButtonHighlight,
    ButtonShadow,
    ButtonText,
    CaptionText,
    GradientActiveCaption,
    GradientInactiveCaption,
    GrayText,
    Highlight,
    HighlightText,
    HotLight,
    InactiveBorder,
    InactiveCaption,
    InactiveCaptionText,
    InfoBackground,
    InfoText,
    Menu,
    MenuBar,
    MenuHighlight,
    MenuText,
    ScrollBar,
    Window,
    WindowFrame,
    WindowText,
}

impl System {
    pub const fn token(self) -> &'static str {
        match self {
            Self::ActiveBorder => "activeBorder",
            Self::ActiveCaption => "activeCaption",
            Self::AppWorkspace => "appWorkspace",
            Self::Background => "background",
            Self::ButtonFace => "btnFace",
            Self::ButtonHighlight => "btnHighlight",
            Self::ButtonShadow => "btnShadow",
            Self::ButtonText => "btnText",
            Self::CaptionText => "captionText",
            Self::GradientActiveCaption => "gradientActiveCaption",
            Self::GradientInactiveCaption => "gradientInactiveCaption",
            Self::GrayText => "grayText",
            Self::Highlight => "highlight",
            Self::HighlightText => "highlightText",
            Self::HotLight => "hotLight",
            Self::InactiveBorder => "inactiveBorder",
            Self::InactiveCaption => "inactiveCaption",
            Self::InactiveCaptionText => "inactiveCaptionText",
            Self::InfoBackground => "infoBk",
            Self::InfoText => "infoText",
            Self::Menu => "menu",
            Self::MenuBar => "menuBar",
            Self::MenuHighlight => "menuHighlight",
            Self::MenuText => "menuText",
            Self::ScrollBar => "scrollBar",
            Self::Window => "window",
            Self::WindowFrame => "windowFrame",
            Self::WindowText => "windowText",
        }
    }

    pub fn from_token(token: &str) -> Option<Self> {
        Some(match token {
            "activeBorder" => Self::ActiveBorder,
            "activeCaption" => Self::ActiveCaption,
            "appWorkspace" => Self::AppWorkspace,
            "background" => Self::Background,
            "btnFace" => Self::ButtonFace,
            "btnHighlight" => Self::ButtonHighlight,
            "btnShadow" => Self::ButtonShadow,
            "btnText" => Self::ButtonText,
            "captionText" => Self::CaptionText,
            "gradientActiveCaption" => Self::GradientActiveCaption,
            "gradientInactiveCaption" => Self::GradientInactiveCaption,
            "grayText" => Self::GrayText,
            "highlight" => Self::Highlight,
            "highlightText" => Self::HighlightText,
            "hotLight" => Self::HotLight,
            "inactiveBorder" => Self::InactiveBorder,
            "inactiveCaption" => Self::InactiveCaption,
            "inactiveCaptionText" => Self::InactiveCaptionText,
            "infoBk" => Self::InfoBackground,
            "infoText" => Self::InfoText,
            "menu" => Self::Menu,
            "menuBar" => Self::MenuBar,
            "menuHighlight" => Self::MenuHighlight,
            "menuText" => Self::MenuText,
            "scrollBar" => Self::ScrollBar,
            "window" => Self::Window,
            "windowFrame" => Self::WindowFrame,
            "windowText" => Self::WindowText,
            _ => return None,
        })
    }
}

/// A color value stored in a palette slot.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum Color {
    Rgb(String),
    System { kind: System, last: Option<String> },
}

impl Color {
    pub fn rgb(value: &str) -> Result<Self> {
        Ok(Self::Rgb(hex(value)?))
    }

    pub fn system(kind: System, last: Option<&str>) -> Result<Self> {
        Ok(Self::System {
            kind,
            last: last.map(hex).transpose()?,
        })
    }
}

/// The complete twelve-slot color palette.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Palette {
    name: String,
    values: Vec<(Slot, Color)>,
}

impl Palette {
    pub fn new(name: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            values: Vec::with_capacity(Slot::ALL.len()),
        }
    }

    pub fn with(mut self, slot: Slot, color: Color) -> Self {
        if let Some(existing) = self.values.iter_mut().find(|(current, _)| *current == slot) {
            existing.1 = color;
        } else {
            self.values.push((slot, color));
        }
        self
    }

    #[inline]
    pub fn name(&self) -> &str {
        &self.name
    }

    #[inline]
    pub fn color(&self, slot: Slot) -> Option<&Color> {
        self.values
            .iter()
            .find(|(current, _)| *current == slot)
            .map(|(_, color)| color)
    }

    pub(crate) fn values(&self) -> &[(Slot, Color)] {
        &self.values
    }
}

/// One script-specific typeface.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Script {
    pub code: String,
    pub typeface: String,
}

/// Major or minor font face in a font set.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Face {
    pub latin: String,
    pub east_asian: String,
    pub complex_script: String,
    pub scripts: Vec<Script>,
}

impl Face {
    pub fn new(latin: impl Into<String>) -> Self {
        Self {
            latin: latin.into(),
            east_asian: String::new(),
            complex_script: String::new(),
            scripts: Vec::new(),
        }
    }

    pub fn east_asian(mut self, typeface: impl Into<String>) -> Self {
        self.east_asian = typeface.into();
        self
    }

    pub fn complex_script(mut self, typeface: impl Into<String>) -> Self {
        self.complex_script = typeface.into();
        self
    }

    pub fn script(mut self, code: impl Into<String>, typeface: impl Into<String>) -> Self {
        self.scripts.push(Script {
            code: code.into(),
            typeface: typeface.into(),
        });
        self
    }
}

/// Major and minor font faces.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct FontSet {
    name: String,
    major: Face,
    minor: Face,
}

impl FontSet {
    pub fn new(name: impl Into<String>, major: Face, minor: Face) -> Self {
        Self {
            name: name.into(),
            major,
            minor,
        }
    }

    #[inline]
    pub fn name(&self) -> &str {
        &self.name
    }

    #[inline]
    pub fn major(&self) -> &Face {
        &self.major
    }

    #[inline]
    pub fn minor(&self) -> &Face {
        &self.minor
    }
}

/// A parsed DrawingML theme document.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Theme {
    pub name: String,
    pub colors: Palette,
    pub fonts: FontSet,
}

/// A theme override containing optional color and font scheme replacements.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct Override {
    pub colors: Option<Palette>,
    pub fonts: Option<FontSet>,
}

impl Override {
    pub fn new() -> Self {
        Self::default()
    }

    pub fn colors(mut self, value: Palette) -> Self {
        self.colors = Some(value);
        self
    }

    pub fn fonts(mut self, value: FontSet) -> Self {
        self.fonts = Some(value);
        self
    }
}

pub(crate) fn validate_name(kind: &str, value: &str) -> Result<()> {
    if value.is_empty() || value.chars().count() > MAX_NAME_CHARS {
        return Err(invalid(format!(
            "{kind} name is empty or exceeds {MAX_NAME_CHARS} characters"
        )));
    }
    Ok(())
}

pub(crate) fn validate_palette(value: &Palette) -> Result<()> {
    validate_name("color palette", value.name())?;
    if value.values.len() != Slot::ALL.len()
        || Slot::ALL.iter().any(|slot| value.color(*slot).is_none())
    {
        return Err(invalid(
            "color palette must contain every DrawingML slot exactly once",
        ));
    }
    for (_, color) in value.values() {
        match color {
            Color::Rgb(value) => {
                let _ = hex(value)?;
            },
            Color::System {
                last: Some(value), ..
            } => {
                let _ = hex(value)?;
            },
            Color::System { last: None, .. } => {},
        }
    }
    Ok(())
}

pub(crate) fn validate_fonts(value: &FontSet) -> Result<()> {
    validate_name("font set", value.name())?;
    for (kind, face) in [("major", value.major()), ("minor", value.minor())] {
        validate_name(kind, &face.latin)?;
        if face.east_asian.chars().count() > MAX_NAME_CHARS
            || face.complex_script.chars().count() > MAX_NAME_CHARS
        {
            return Err(invalid(format!(
                "{kind} typeface exceeds {MAX_NAME_CHARS} characters"
            )));
        }
        if face.scripts.len() > MAX_SCRIPT_COUNT {
            return Err(invalid(format!("{kind} font has too many script entries")));
        }
        for script in &face.scripts {
            if script.code.is_empty()
                || script.code.chars().count() > MAX_SCRIPT_CHARS
                || !script.code.bytes().all(|b| b.is_ascii_alphanumeric())
            {
                return Err(invalid(format!("invalid {kind} script code")));
            }
            validate_name("script typeface", &script.typeface)?;
        }
    }
    Ok(())
}

pub(crate) fn hex(value: &str) -> Result<String> {
    if value.len() == 6 && value.bytes().all(|byte| byte.is_ascii_hexdigit()) {
        Ok(value.to_ascii_uppercase())
    } else {
        Err(invalid("sRGB colors must contain six hexadecimal digits"))
    }
}

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}
